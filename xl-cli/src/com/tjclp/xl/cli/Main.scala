package com.tjclp.xl.cli

import java.nio.file.Path

import cats.effect.{ExitCode, IO}
import cats.implicits.*
import cats.syntax.parallel.*
import com.monovore.decline.*
import com.monovore.decline.effect.*

import com.tjclp.xl.{*, given}
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.ooxml.XlsxReader.ReaderConfig
import com.tjclp.xl.ooxml.writer.{WriterConfig, XmlBackend}
import com.tjclp.xl.cli.commands.{
  CellCommands,
  CommentCommands,
  ImportCommands,
  ReadCommands,
  SheetCommands,
  StreamingReadCommands,
  StreamingWriteCommands,
  WorkbookCommands,
  WriteCommands
}
import com.tjclp.xl.cli.raster.{
  BatikRasterizer,
  CairoSvg,
  ImageMagick,
  RasterizerChain,
  Resvg,
  RsvgConvert
}
import com.tjclp.xl.cli.helpers.SheetResolver
import com.tjclp.xl.cli.output.Format

/** Read version from generated resource, fallback to dev */
private object BuildInfo:
  val version: String =
    val props = new java.util.Properties()
    val stream = Option(getClass.getResourceAsStream("/version.properties"))
    stream.foreach(props.load)
    Option(props.getProperty("version")).getOrElse("dev")

/**
 * XL CLI - LLM-friendly Excel operations.
 *
 * Stateless by design: each command is self-contained. Use global flags:
 *   - `-f, --file` — Input file (required)
 *   - `-s, --sheet` — Sheet name (optional, defaults to first)
 *   - `-o, --output` — Output file for mutations (required for put/putf)
 */
object Main
    extends CommandIOApp(
      name = "xl",
      header = "LLM-friendly Excel operations (stateless)",
      version = BuildInfo.version
    ):

  override def main: Opts[IO[ExitCode]] =
    // Workbook-level: only --file (no --sheet)
    // Note: --stream not supported for workbook-level commands (need full metadata)
    val workbookSubcmds = namesCmd
    val workbookOpts = (fileOpt, maxSizeOpt, workbookSubcmds).mapN { (file, maxSize, cmd) =>
      run(file, None, None, None, maxSize, false, cmd)
    }

    // Sheets command: --file required, --output optional (required for hide/show, not for list)
    // This needs its own opts chain because list doesn't need output but hide/show do
    val sheetsOpts =
      (fileOpt, outputOpt.orNone, backendOpt, maxSizeOpt, streamOpt, sheetsCmd).mapN {
        (file, output, backend, maxSize, stream, cmd) =>
          run(file, None, output, backend, maxSize, stream, cmd)
      }

    // Headless commands: --file is optional (for constant formulas like =1+1, =PI())
    // Note: --stream not supported for eval (needs formula analysis)
    // evala requires --file (array formulas need sheet context)
    val headlessOpts = (fileOpt.orNone, sheetOpt, maxSizeOpt, evalCmd orElse evalArrayCmd).mapN {
      (fileOpt, sheet, maxSize, cmd) =>
        runHeadless(fileOpt, sheet, maxSize, cmd)
    }

    // Sheet-level read-only: --file and --sheet (no --output)
    val sheetReadOnlySubcmds =
      boundsCmd orElse viewCmd orElse cellCmd orElse searchCmd orElse statsCmd

    val sheetReadOnlyOpts = (fileOpt, sheetOpt, maxSizeOpt, streamOpt, sheetReadOnlySubcmds).mapN {
      (file, sheet, maxSize, stream, cmd) =>
        run(file, sheet, None, None, maxSize, stream, cmd)
    }

    // Sheet-level write: --file, --sheet, and --output (required)
    // Sheet-level write: --file, --sheet, and --output (required)
    // --stream enables O(1) output memory with style preservation (hybrid streaming)
    val sheetWriteSubcmds =
      putCmd orElse putfCmd orElse styleCmd orElse rowCmd orElse colCmd orElse autoFitCmd orElse batchCmd orElse importCmd orElse addSheetCmd orElse removeSheetCmd orElse renameSheetCmd orElse moveSheetCmd orElse copySheetCmd orElse mergeCmd orElse unmergeCmd orElse commentCmd orElse removeCommentCmd orElse clearCmd orElse fillCmd orElse sortCmd

    val sheetWriteOpts =
      (fileOpt, sheetOpt, outputOpt, backendOpt, maxSizeOpt, streamOpt, sheetWriteSubcmds).mapN {
        (file, sheet, out, backend, maxSize, stream, cmd) =>
          run(file, sheet, Some(out), backend, maxSize, stream, cmd)
      }

    // Standalone: no --file required (creates new files)
    val standaloneOpts = newCmd.map { case (outPath, sheetName, sheets, backend) =>
      runStandalone(outPath, sheetName, sheets, backend)
    }

    // Info commands: no file required
    val infoOpts = functionsCmd.map(_ => runInfo())
    val rasterOpts = rasterizersCmd.map(_ => runRasterizers())

    rasterOpts orElse infoOpts orElse standaloneOpts orElse headlessOpts orElse sheetsOpts orElse workbookOpts orElse sheetReadOnlyOpts orElse sheetWriteOpts

  // ==========================================================================
  // Global options
  // ==========================================================================

  private val fileOpt =
    Opts.option[Path]("file", "Excel file to operate on (required)", "f")

  private val sheetOpt =
    Opts
      .option[String]("sheet", "Sheet to select (required for sheet-level operations)", "s")
      .orNone

  private val outputOpt =
    Opts.option[Path]("output", "Output file (required)", "o")

  private val backendOpt: Opts[Option[XmlBackend]] =
    Opts
      .option[String]("backend", "XML backend: scalaxml (default, stable) or saxstax (faster)")
      .mapValidated {
        case "scalaxml" | "scala-xml" | "xml" =>
          cats.data.Validated.valid(XmlBackend.ScalaXml)
        case "saxstax" | "sax-stax" | "stax" =>
          cats.data.Validated.valid(XmlBackend.SaxStax)
        case other =>
          cats.data.Validated.invalidNel(
            s"Unknown backend: $other. Use 'scalaxml' (default) or 'saxstax' (faster)"
          )
      }
      .orNone

  private val maxSizeOpt: Opts[Option[Long]] =
    Opts
      .option[Long](
        "max-size",
        "Max uncompressed size in MB for in-memory load (default: 100, 0 = unlimited). Use for large files when --stream is not supported."
      )
      .orNone

  private val streamOpt: Opts[Boolean] =
    Opts
      .flag(
        "stream",
        "Use O(1) memory streaming for large files (100k+ rows). Supports: search, stats, bounds, view (markdown/csv/json). 7-8x faster than in-memory."
      )
      .orFalse

  // ==========================================================================
  // Command definitions
  // ==========================================================================

  private val rangeArg = Opts.argument[String]("range")
  private val refArg = Opts.argument[String]("ref")
  private val valueArg = Opts.argument[String]("value")
  // Alternative flag for values starting with - (e.g., --value=-5)
  private val valueOpt = Opts.option[String]("value", "Cell value (use for negative numbers)", "v")
  private val patternArg = Opts.argument[String]("pattern")

  private val formulasOpt = Opts.flag("formulas", "Show formulas instead of values").orFalse
  private val evalOpt = Opts.flag("eval", "Evaluate formulas (compute live values)").orFalse
  private val strictOpt =
    Opts.flag("strict", "Fail on formula evaluation errors (use with --eval)").orFalse
  private val limitOpt = Opts.option[Int]("limit", "Maximum rows to display").withDefault(50)
  private val formatOpt = Opts
    .option[String]("format", "Output format: markdown, html, svg, json, csv, png, jpeg, webp, pdf")
    .withDefault("markdown")
    .mapValidated { s =>
      s.toLowerCase match
        case "markdown" | "md" => cats.data.Validated.valid(ViewFormat.Markdown)
        case "html" => cats.data.Validated.valid(ViewFormat.Html)
        case "svg" => cats.data.Validated.valid(ViewFormat.Svg)
        case "json" => cats.data.Validated.valid(ViewFormat.Json)
        case "csv" => cats.data.Validated.valid(ViewFormat.Csv)
        case "png" => cats.data.Validated.valid(ViewFormat.Png)
        case "jpeg" | "jpg" => cats.data.Validated.valid(ViewFormat.Jpeg)
        case "webp" => cats.data.Validated.valid(ViewFormat.WebP)
        case "pdf" => cats.data.Validated.valid(ViewFormat.Pdf)
        case other =>
          cats.data.Validated.invalidNel(
            s"Unknown format: $other. Use markdown, html, svg, json, csv, png, jpeg, webp, or pdf"
          )
    }
  private val printScaleOpt =
    Opts.flag("print-scale", "Apply print scaling (for PDF-like output)").orFalse
  private val gridlinesOpt =
    Opts.flag("gridlines", "Show cell gridlines in SVG output").orFalse
  private val showLabelsOpt =
    Opts.flag("show-labels", "Include column letters (A, B, C) and row numbers (1, 2, 3)").orFalse
  private val dpiOpt =
    Opts.option[Int]("dpi", "DPI for raster output (default: 144 for retina)").withDefault(144)
  private val qualityOpt =
    Opts.option[Int]("quality", "JPEG quality 1-100 (default: 90)").withDefault(90)
  private val rasterOutputOpt =
    Opts
      .option[Path](
        "raster-output",
        "Output file for raster formats (required for png/jpeg/webp/pdf)"
      )
      .orNone
  private val skipEmptyOpt =
    Opts.flag("skip-empty", "Skip empty cells (JSON) or empty rows/columns (tabular)").orFalse
  private val headerRowOpt =
    Opts
      .option[Int](
        "header-row",
        "Use values from this row as keys in JSON output (1-based row number)"
      )
      .orNone
  private val rasterizerOpt =
    Opts
      .option[String](
        "rasterizer",
        "Force specific rasterizer: batik, cairosvg, rsvg-convert, resvg, imagemagick"
      )
      .orNone
  private val sheetsFilterOpt =
    Opts
      .option[String]("sheets", "Comma-separated list of sheets to search (default: all)")
      .orNone

  // ==========================================================================
  // Extended help strings
  // ==========================================================================

  private val viewHelp = """View range in multiple formats (table, JSON, image, PDF).

USAGE:
  xl -f file.xlsx -s Sheet1 view A1:D10
  xl -f file.xlsx view "Sheet1!A1:D10"    # Qualified ref (no -s needed)

FORMATS:
  markdown (default), json, csv, html, svg, png, jpeg, webp, pdf

OUTPUT FLAGS:
  --format <fmt>      Output format
  --formulas          Show formulas instead of values
  --eval              Evaluate formulas (compute live values)
  --strict            Fail on formula evaluation errors (use with --eval)
  --skip-empty        Skip empty cells/rows
  --show-labels       Include row/column headers (A, B, C / 1, 2, 3)
  --header-row <n>    Use row N as JSON keys (1-based)

RASTER FLAGS (png/jpeg/webp/pdf):
  --raster-output <path>  Output file (required for raster formats)
  --dpi <n>               Resolution (default: 144)
  --quality <n>           JPEG quality 1-100 (default: 90)
  --rasterizer <name>     Force: batik, cairosvg, rsvg-convert, resvg, imagemagick

EXAMPLES:
  xl -f data.xlsx -s Sheet1 view A1:D10                    # Markdown table
  xl -f data.xlsx -s Sheet1 view A1:D10 --format json      # JSON array
  xl -f data.xlsx -s Sheet1 view A1:D10 --eval             # Computed values
  xl -f data.xlsx -s Sheet1 view A1:D10 --eval --strict    # Fail on eval errors
  xl -f data.xlsx -s Sheet1 view A1:D10 --formulas         # Show formulas
  xl -f data.xlsx -s Sheet1 view A1:D10 --format png --raster-output chart.png"""

  private val styleHelp =
    """Apply formatting to cells. Styles merge by default (use --replace to overwrite).

USAGE:
  xl -f file.xlsx -s Sheet1 -o out.xlsx style A1:D1 --bold --bg yellow

FONT:
  --bold, --italic, --underline
  --font-size <pt>    Font size in points
  --font-name <name>  Font family (e.g., "Arial", "Calibri")
  --fg <color>        Text color

FILL:
  --bg <color>        Background color

ALIGNMENT:
  --align <left|center|right>
  --valign <top|middle|bottom>
  --wrap              Enable text wrapping

NUMBER FORMAT:
  --format <general|number|currency|percent|date|text>

BORDERS:
  --border <none|thin|medium|thick>      All sides
  --border-top/right/bottom/left <style> Individual sides
  --border-color <color>                 Border color

COLORS:
  Named: red, blue, navy, yellow, green, white, black, orange, purple, gray
  Hex: #FF6600, #4472C4
  RGB: rgb(100,150,200)

EXAMPLES:
  xl -f f.xlsx -s S1 -o o.xlsx style A1:E1 --bold --bg navy --fg white --align center
  xl -f f.xlsx -s S1 -o o.xlsx style B2:B100 --format currency
  xl -f f.xlsx -s S1 -o o.xlsx style A1 --border thin --border-color black
  xl -f f.xlsx -s S1 -o o.xlsx style C1:C10 --replace --bg yellow  # Replace, don't merge"""

  private val importHelp = """Import CSV data with automatic type detection.

USAGE:
  xl -f file.xlsx -s Sheet1 -o out.xlsx import data.csv A1
  xl -f file.xlsx -o out.xlsx import data.csv --new-sheet "Data"

OPTIONS:
  --delimiter <char>      Field separator (default: ,)
  --encoding <enc>        Input encoding (default: UTF-8)
  --skip-header           Skip first row (treat as header, do not import)
  --no-type-inference     Treat all values as text
  --new-sheet <name>      Create new sheet for imported data

TYPE INFERENCE:
  Numbers:   100, 29.99, -5.5 → Number type
  Booleans:  true, false (case-insensitive) → Boolean type
  Dates:     2024-01-15 (ISO 8601 only) → DateTime type
  Text:      Everything else

LIMITATIONS:
  - Entire CSV loaded into memory (not streamed)
  - Recommended: <50k rows for optimal performance
  - Date formats: Only ISO 8601 (YYYY-MM-DD) supported

EXAMPLES:
  xl -f f.xlsx -s S1 -o o.xlsx import data.csv A1
  xl -f f.xlsx -o o.xlsx import data.csv --new-sheet "Imported"
  xl -f f.xlsx -s S1 -o o.xlsx import data.csv A1 --delimiter ";" --skip-header
  xl -f f.xlsx -s S1 -o o.xlsx import data.csv A1 --no-type-inference"""

  private val putHelp = """Write value(s) to cell or range.

USAGE:
  xl -f file.xlsx -s Sheet1 -o out.xlsx put <ref> <value>
  xl -f file.xlsx -s Sheet1 -o out.xlsx put <range> <value>        # Fill all
  xl -f file.xlsx -s Sheet1 -o out.xlsx put <range> <v1> <v2> ...  # Batch

MODES:
  Single:   put A1 100              → Write 100 to A1
  Fill:     put A1:A10 "TBD"        → Fill range with same value
  Batch:    put A1:C1 "X" "Y" "Z"   → Different value per cell (row-major)

NEGATIVE NUMBERS:
  Use --value flag (- is interpreted as flag prefix):
  ❌ put A1 -100              → Error: unknown flag
  ✅ put A1 --value "-100"    → Writes -100 to A1

EXAMPLES:
  xl -f f.xlsx -s S1 -o o.xlsx put A1 "Hello"
  xl -f f.xlsx -s S1 -o o.xlsx put B2:B10 0              # Fill with zeros
  xl -f f.xlsx -s S1 -o o.xlsx put A1:D1 "Q1" "Q2" "Q3" "Q4"
  xl -f f.xlsx -s S1 -o o.xlsx put A1 --value "-500"     # Negative number"""

  private val putfHelp = """Write formula(s) to cell or range with Excel-style dragging.

USAGE:
  xl -f file.xlsx -s Sheet1 -o out.xlsx putf <ref> <formula>
  xl -f file.xlsx -s Sheet1 -o out.xlsx putf <range> <formula>     # Drag
  xl -f file.xlsx -s Sheet1 -o out.xlsx putf <range> <f1> <f2> ... # Batch

FORMULA DRAGGING:
  Single formula + range → references shift automatically:
  putf B2:B10 "=A2*1.1"  →  B2: =A2*1.1, B3: =A3*1.1, B4: =A4*1.1 ...

ANCHOR MODES ($ controls shifting):
  $A$1   Absolute (never shifts)
  $A1    Column absolute, row relative
  A$1    Column relative, row absolute
  A1     Fully relative (shifts both ways)

RUNNING TOTALS:
  putf C2:C10 "=SUM(\$B\$2:B2)"  →  C2: =SUM($B$2:B2), C3: =SUM($B$2:B3) ...

BATCH (explicit, no dragging):
  putf D1:D3 "=A1+B1" "=A2*B2" "=A3-B3"  → Formulas applied as-is

EXAMPLES:
  xl -f f.xlsx -s S1 -o o.xlsx putf C1 "=A1+B1"
  xl -f f.xlsx -s S1 -o o.xlsx putf B2:B100 "=A2*1.1"
  xl -f f.xlsx -s S1 -o o.xlsx putf C2:C10 "=SUM(\$A\$1:A2)"
  xl -f f.xlsx -s S1 -o o.xlsx putf D1:D3 "=A1+B1" "=A2*B2" "=A3-B3\""""

  private val sortHelp = """Sort rows in range by one or more columns.

USAGE:
  xl -f file.xlsx -s Sheet1 -o out.xlsx sort <range> --by <col> [options]

OPTIONS:
  --by <col>        Primary sort column (required)
  --then-by <col>   Secondary sort column (repeatable)
  --desc            Sort descending (default: ascending)
  --numeric         Force numeric comparison ("10" > "9")
  --header          First row is header (exclude from sort)

BEHAVIOR:
  - Empty cells sort last
  - Formulas use cached value for sorting
  - Booleans sort as 0 (FALSE) / 1 (TRUE)
  - Rows move together (columns outside range preserved)

EXAMPLES:
  xl -f f.xlsx -s S1 -o o.xlsx sort A1:D100 --by B
  xl -f f.xlsx -s S1 -o o.xlsx sort A1:D100 --by B --desc --numeric
  xl -f f.xlsx -s S1 -o o.xlsx sort A1:D100 --by B --then-by C --header"""

  // --- Info commands (no --file required) ---

  val functionsCmd: Opts[Unit] =
    Opts.subcommand("functions", "List supported Excel functions") {
      Opts.unit
    }

  val rasterizersCmd: Opts[Unit] =
    Opts.subcommand("rasterizers", "List available SVG-to-raster backends") {
      Opts.unit
    }

  // --- Standalone commands (no --file required) ---

  private val outputArg = Opts.argument[Path]("output")
  private val sheetNameOpt =
    Opts.option[String]("sheet-name", "Sheet name (defaults to 'Sheet1')").withDefault("Sheet1")
  private val sheetsOpt: Opts[List[String]] =
    Opts.options[String]("sheet", "Sheet name (repeatable for multiple sheets)").orEmpty

  val newCmd: Opts[(Path, String, List[String], Option[XmlBackend])] =
    Opts.subcommand("new", "Create a blank xlsx file") {
      (outputArg, sheetNameOpt, sheetsOpt, backendOpt).tupled
    }

  // --- Read-only commands ---

  // --stats flag for sheets command (full mode with cell counts)
  private val statsOpt: Opts[Boolean] =
    Opts
      .flag(
        "stats",
        "Show cell and formula counts (slower, requires loading all data)"
      )
      .orFalse

  // --scan flag for bounds command (full streaming scan)
  private val scanOpt: Opts[Boolean] =
    Opts
      .flag(
        "scan",
        "Force full streaming scan for accurate bounds (slower, but accurate)"
      )
      .orFalse

  // --very flag for sheets hide subcommand
  private val veryHideOpt: Opts[Boolean] =
    Opts
      .flag("very", "Make sheet very hidden (not accessible from Excel UI, only via VBA)")
      .orFalse

  val sheetsCmd: Opts[CliCommand] = Opts.subcommand(
    "sheets",
    "Sheet operations: list, hide, show"
  ) {
    // Sheet name argument for hide/show subcommands (local to sheetsCmd)
    val targetSheetArg: Opts[String] = Opts.argument[String]("sheet-name")

    // List subcommand (explicit)
    val listSubCmd = Opts.subcommand("list", "List all sheets") {
      statsOpt.map(SheetsAction.List.apply)
    }

    // Hide subcommand
    val hideSubCmd = Opts.subcommand("hide", "Hide a sheet from the sheet tabs") {
      (targetSheetArg, veryHideOpt).mapN(SheetsAction.Hide.apply)
    }

    // Show subcommand
    val showSubCmd = Opts.subcommand("show", "Show a hidden sheet") {
      targetSheetArg.map(SheetsAction.Show.apply)
    }

    // Default to list if no subcommand (backwards compat: `xl sheets` = `xl sheets list`)
    val defaultList = statsOpt.map(SheetsAction.List.apply)

    (listSubCmd orElse hideSubCmd orElse showSubCmd orElse defaultList)
      .map(CliCommand.Sheets.apply)
  }

  val namesCmd: Opts[CliCommand] = Opts.subcommand("names", "List defined names (named ranges)") {
    Opts(CliCommand.Names)
  }

  val boundsCmd: Opts[CliCommand] = Opts.subcommand(
    "bounds",
    "Show used range (instant from dimension element, --scan for accurate scan)"
  ) {
    scanOpt.map(CliCommand.Bounds.apply)
  }

  val viewCmd: Opts[CliCommand] =
    Opts.subcommand("view", viewHelp) {
      (
        rangeArg,
        formulasOpt,
        evalOpt,
        strictOpt,
        limitOpt,
        formatOpt,
        printScaleOpt,
        gridlinesOpt,
        showLabelsOpt,
        dpiOpt,
        qualityOpt,
        rasterOutputOpt,
        skipEmptyOpt,
        headerRowOpt,
        rasterizerOpt
      )
        .mapN(CliCommand.View.apply)
    }

  private val noStyleOpt =
    Opts.flag("no-style", "Omit style information from output").orFalse

  val cellCmd: Opts[CliCommand] = Opts.subcommand("cell", "Get cell details") {
    (refArg, noStyleOpt).mapN(CliCommand.Cell.apply)
  }

  val searchCmd: Opts[CliCommand] =
    Opts.subcommand("search", "Search for cells (all sheets by default)") {
      (patternArg, limitOpt, sheetsFilterOpt).mapN(CliCommand.Search.apply)
    }

  val statsCmd: Opts[CliCommand] =
    Opts.subcommand("stats", "Calculate statistics for numeric values in range") {
      rangeArg.map(CliCommand.Stats.apply)
    }

  // --- Analyze ---

  private val formulaArg = Opts.argument[String]("formula")
  private val withOpts =
    Opts
      .options[String]("with", "Cell overrides (e.g., A1=100,B2=200). Repeatable.", "w")
      .map(_.toList)
      .withDefault(Nil)

  private def parseOverrides(withStrs: List[String]): List[String] =
    withStrs.flatMap(_.split(",").map(_.trim).filter(_.nonEmpty))

  val evalCmd: Opts[CliCommand] =
    Opts.subcommand("eval", "Evaluate formula without modifying sheet") {
      (formulaArg, withOpts).mapN { (formula, withStrs) =>
        CliCommand.Eval(formula, parseOverrides(withStrs))
      }
    }

  private val atOpt =
    Opts.option[String]("at", "Target cell for array spill (default: virtual cell)").orNone

  val evalArrayCmd: Opts[CliCommand] =
    Opts.subcommand("evala", "Evaluate array formula and display result grid") {
      (formulaArg, atOpt, withOpts).mapN { (formula, target, withStrs) =>
        CliCommand.EvalArray(formula, target, parseOverrides(withStrs))
      }
    }

  // --- Mutate (require -o) ---

  // Variadic values for put (supports single value, fill pattern, or batch values)
  private val valuesArg = Opts.arguments[String]("value")

  val putCmd: Opts[CliCommand] = Opts.subcommand("put", putHelp) {
    // Support both positional args and --value flag (for negative numbers)
    val valuesOrOpt = valueOpt.map(v => List(v)) orElse valuesArg.map(_.toList)
    (refArg, valuesOrOpt).mapN(CliCommand.Put.apply)
  }

  // Variadic formulas for putf (supports single formula, dragging, or batch formulas)
  private val formulasArg = Opts.arguments[String]("formula")

  val putfCmd: Opts[CliCommand] = Opts.subcommand("putf", putfHelp) {
    (refArg, formulasArg).mapN { (ref, formulas) =>
      CliCommand.PutFormula(ref, formulas.toList)
    }
  }

  // --- Style command options ---
  private val boldOpt = Opts.flag("bold", "Bold text").orFalse
  private val italicOpt = Opts.flag("italic", "Italic text").orFalse
  private val underlineOpt = Opts.flag("underline", "Underline text").orFalse
  private val bgOpt =
    Opts.option[String]("bg", "Background color (name, #hex, or rgb(r,g,b))").orNone
  private val fgOpt = Opts.option[String]("fg", "Text color (name, #hex, or rgb(r,g,b))").orNone
  private val fontSizeOpt = Opts.option[Double]("font-size", "Font size in points").orNone
  private val fontNameOpt = Opts.option[String]("font-name", "Font family name").orNone
  private val alignOpt = Opts.option[String]("align", "Horizontal: left, center, right").orNone
  private val valignOpt = Opts.option[String]("valign", "Vertical: top, middle, bottom").orNone
  private val wrapOpt = Opts.flag("wrap", "Enable text wrapping").orFalse
  private val numFormatOpt =
    Opts
      .option[String]("format", "Number format: general, number, currency, percent, date, text")
      .orNone
  private val borderOpt =
    Opts.option[String]("border", "Border style for all sides: none, thin, medium, thick").orNone
  private val borderTopOpt =
    Opts.option[String]("border-top", "Top border style: none, thin, medium, thick").orNone
  private val borderRightOpt =
    Opts.option[String]("border-right", "Right border style: none, thin, medium, thick").orNone
  private val borderBottomOpt =
    Opts.option[String]("border-bottom", "Bottom border style: none, thin, medium, thick").orNone
  private val borderLeftOpt =
    Opts.option[String]("border-left", "Left border style: none, thin, medium, thick").orNone
  private val borderColorOpt = Opts.option[String]("border-color", "Border color").orNone
  private val replaceOpt = Opts
    .flag("replace", "Replace entire style instead of merging with existing")
    .orFalse

  val styleCmd: Opts[CliCommand] = Opts.subcommand("style", styleHelp) {
    (
      rangeArg,
      boldOpt,
      italicOpt,
      underlineOpt,
      bgOpt,
      fgOpt,
      fontSizeOpt,
      fontNameOpt,
      alignOpt,
      valignOpt,
      wrapOpt,
      numFormatOpt,
      borderOpt,
      borderTopOpt,
      borderRightOpt,
      borderBottomOpt,
      borderLeftOpt,
      borderColorOpt,
      replaceOpt
    ).mapN(CliCommand.Style.apply)
  }

  // --- Row/Column command options ---
  private val rowArg = Opts.argument[Int]("row")
  private val colArg = Opts.argument[String]("col")
  private val heightOpt = Opts.option[Double]("height", "Row height in points").orNone
  private val widthOpt = Opts.option[Double]("width", "Column width in character units").orNone
  private val hideOpt = Opts.flag("hide", "Hide row/column").orFalse
  private val showOpt = Opts.flag("show", "Show (unhide) row/column").orFalse
  private val autoFitOpt = Opts.flag("auto-fit", "Auto-fit column width based on content").orFalse

  val rowCmd: Opts[CliCommand] = Opts.subcommand("row", "Set row properties (height, hide/show)") {
    (rowArg, heightOpt, hideOpt, showOpt).mapN(CliCommand.RowOp.apply)
  }

  val colCmd: Opts[CliCommand] =
    Opts.subcommand(
      "col",
      "Set column properties (width, hide/show, auto-fit). Supports ranges like A:F"
    ) {
      (colArg, widthOpt, hideOpt, showOpt, autoFitOpt).mapN(CliCommand.ColOp.apply)
    }

  // --- Auto-fit command ---
  private val autoFitColumnsOpt =
    Opts
      .option[String]("columns", "Column range to auto-fit (e.g., A:F). Default: all used columns")
      .orNone
  val autoFitCmd: Opts[CliCommand] =
    Opts.subcommand("autofit", "Auto-fit column widths based on content") {
      autoFitColumnsOpt.map(CliCommand.AutoFit.apply)
    }

  // --- Batch command ---
  private val batchArg = Opts.argument[String]("operations").withDefault("-")
  private val batchHelp = """Apply multiple operations atomically from JSON.

USAGE:
  xl -f in.xlsx -s Sheet1 -o out.xlsx batch ops.json
  echo '[...]' | xl -f in.xlsx -s Sheet1 -o out.xlsx batch -

OPERATIONS:
  put       {"op": "put", "ref": "A1", "value": "Hello"}
  putf      {"op": "putf", "ref": "A1", "value": "=SUM(B1:B10)"}
  style     {"op": "style", "range": "A1:D1", "bold": true, "bg": "#FFFF00"}
  merge     {"op": "merge", "range": "A1:D1"}
  unmerge   {"op": "unmerge", "range": "A1:D1"}
  colwidth  {"op": "colwidth", "col": "A", "width": 15.5}
  rowheight {"op": "rowheight", "row": 1, "height": 30}

STYLE PROPERTIES:
  Font:      bold, italic, underline, fg, fontSize, fontName
  Fill:      bg (background color, e.g., "#FFFF00" or "yellow")
  Align:     align (left/center/right), valign (top/middle/bottom), wrap
  Format:    numFormat (general/number/currency/percent/date/text)
  Border:    border (all), borderTop/Right/Bottom/Left, borderColor
  Mode:      replace (true=replace style, false=merge with existing)

EXAMPLE:
  [
    {"op": "put", "ref": "A1", "value": "Revenue Report"},
    {"op": "style", "range": "A1:D1", "bold": true, "bg": "#4472C4", "fg": "#FFFFFF", "align": "center"},
    {"op": "merge", "range": "A1:D1"},
    {"op": "colwidth", "col": "A", "width": 25},
    {"op": "put", "ref": "A2", "value": "Q1"},
    {"op": "put", "ref": "B2", "value": 1000},
    {"op": "putf", "ref": "C2", "value": "=B2*1.1"}
  ]

Operations execute in order. Use "-" to read from stdin."""

  val batchCmd: Opts[CliCommand] =
    Opts.subcommand("batch", batchHelp) {
      batchArg.map(CliCommand.Batch.apply)
    }

  // --- Import command ---
  private val csvPathArg = Opts.argument[String]("csv-file")
  private val startRefOpt = Opts.argument[String]("start-ref").orNone
  private val delimiterOpt =
    Opts.option[Char]("delimiter", "Field separator (default: comma)").withDefault(',')
  private val skipHeaderOpt =
    Opts.flag("skip-header", "Skip first row (treat as header, do not import)").orFalse
  private val encodingOpt =
    Opts.option[String]("encoding", "Input encoding (default: UTF-8)").withDefault("UTF-8")
  private val newSheetImportOpt =
    Opts.option[String]("new-sheet", "Create new sheet with this name").orNone
  private val noTypeInferenceOpt =
    Opts.flag("no-type-inference", "Treat all values as text").orFalse

  val importCmd: Opts[CliCommand] =
    Opts.subcommand("import", importHelp) {
      (
        csvPathArg,
        startRefOpt,
        delimiterOpt,
        skipHeaderOpt,
        encodingOpt,
        newSheetImportOpt,
        noTypeInferenceOpt
      )
        .mapN { (path, ref, delim, skipHeader, enc, newSh, noInfer) =>
          CliCommand.Import(path, ref, delim, skipHeader, enc, newSh, noInfer)
        }
    }

  // --- Sheet management commands ---
  private val sheetNameArg = Opts.argument[String]("name")
  private val afterOpt =
    Opts.option[String]("after", "Insert new sheet after this sheet").orNone
  private val beforeOpt =
    Opts.option[String]("before", "Insert new sheet before this sheet").orNone

  val addSheetCmd: Opts[CliCommand] =
    Opts.subcommand("add-sheet", "Add new empty sheet to workbook") {
      (sheetNameArg, afterOpt, beforeOpt).mapN(CliCommand.AddSheet.apply)
    }

  val removeSheetCmd: Opts[CliCommand] =
    Opts.subcommand("remove-sheet", "Remove sheet from workbook") {
      sheetNameArg.map(CliCommand.RemoveSheet.apply)
    }

  private val newNameArg = Opts.argument[String]("new-name")

  val renameSheetCmd: Opts[CliCommand] =
    Opts.subcommand("rename-sheet", "Rename a sheet") {
      (sheetNameArg, newNameArg).mapN(CliCommand.RenameSheet.apply)
    }

  private val toIndexOpt =
    Opts.option[Int]("to", "Move to index (0-based)").orNone

  val moveSheetCmd: Opts[CliCommand] =
    Opts.subcommand("move-sheet", "Move sheet to new position") {
      (sheetNameArg, toIndexOpt, afterOpt, beforeOpt).mapN(CliCommand.MoveSheet.apply)
    }

  val copySheetCmd: Opts[CliCommand] =
    Opts.subcommand("copy-sheet", "Copy sheet to new name") {
      (sheetNameArg, newNameArg).mapN(CliCommand.CopySheet.apply)
    }

  val mergeCmd: Opts[CliCommand] =
    Opts.subcommand("merge", "Merge cells in range") {
      rangeArg.map(CliCommand.Merge.apply)
    }

  val unmergeCmd: Opts[CliCommand] =
    Opts.subcommand("unmerge", "Unmerge cells in range") {
      rangeArg.map(CliCommand.Unmerge.apply)
    }

  // --- Comment commands ---
  private val commentTextArg = Opts.argument[String]("text")
  private val authorOpt = Opts.option[String]("author", "Comment author name").orNone

  val commentCmd: Opts[CliCommand] =
    Opts.subcommand("comment", "Add comment to cell") {
      (refArg, commentTextArg, authorOpt).mapN(CliCommand.AddComment.apply)
    }

  val removeCommentCmd: Opts[CliCommand] =
    Opts.subcommand("remove-comment", "Remove comment from cell") {
      refArg.map(CliCommand.RemoveComment.apply)
    }

  // --- Clear command ---
  private val clearAllOpt = Opts.flag("all", "Clear contents, styles, and comments").orFalse
  private val clearStylesOpt = Opts.flag("styles", "Clear styles only (reset to default)").orFalse
  private val clearCommentsOpt = Opts.flag("comments", "Clear comments only").orFalse

  val clearCmd: Opts[CliCommand] =
    Opts.subcommand("clear", "Clear cell contents, styles, or comments from range") {
      (rangeArg, clearAllOpt, clearStylesOpt, clearCommentsOpt).mapN(CliCommand.Clear.apply)
    }

  // --- Fill command ---
  private val sourceArg = Opts.argument[String]("source")
  private val targetArg = Opts.argument[String]("target")
  private val rightOpt = Opts.flag("right", "Fill rightward instead of downward").orFalse

  val fillCmd: Opts[CliCommand] =
    Opts.subcommand("fill", "Fill cells with source value/formula (Excel Ctrl+D/Ctrl+R)") {
      (sourceArg, targetArg, rightOpt).mapN { (source, target, right) =>
        val direction = if right then FillDirection.Right else FillDirection.Down
        CliCommand.Fill(source, target, direction)
      }
    }

  // --- Sort command ---
  private val byOpt = Opts.option[String]("by", "Primary sort column (required)", "b")
  private val descOpt = Opts.flag("desc", "Sort descending (default: ascending)").orFalse
  private val numericSortOpt = Opts.flag("numeric", "Force numeric comparison").orFalse
  private val thenByOpts = Opts.options[String]("then-by", "Additional sort column(s)").orEmpty
  private val sortHeaderOpt =
    Opts.flag("header", "First row is header (exclude from sort)").orFalse

  val sortCmd: Opts[CliCommand] =
    Opts.subcommand("sort", sortHelp) {
      (rangeArg, byOpt, descOpt, numericSortOpt, thenByOpts, sortHeaderOpt).mapN {
        (range, by, desc, numeric, thenBy, header) =>
          val direction =
            if desc then SortDirection.Descending else SortDirection.Ascending
          val mode = if numeric then SortMode.Numeric else SortMode.Alphanumeric
          val primaryKey = SortKey(by, direction, mode)
          // Secondary keys inherit direction and mode from primary
          val secondaryKeys = thenBy.map(col => SortKey(col, direction, mode)).toList
          CliCommand.Sort(range, primaryKey :: secondaryKeys, header)
      }
    }

  // ==========================================================================
  // Command execution
  // ==========================================================================

  private def run(
    filePath: Path,
    sheetNameOpt: Option[String],
    outputOpt: Option[Path],
    backendOpt: Option[XmlBackend],
    maxSizeOpt: Option[Long],
    stream: Boolean,
    cmd: CliCommand
  ): IO[ExitCode] =
    execute(filePath, sheetNameOpt, outputOpt, backendOpt, maxSizeOpt, stream, cmd).attempt
      .flatMap {
        case Right(output) =>
          IO.println(output).as(ExitCode.Success)
        case Left(err) =>
          IO.println(Format.errorSimple(err.getMessage)).as(ExitCode.Error)
      }

  private def runInfo(): IO[ExitCode] =
    IO.println(formatFunctionList()).as(ExitCode.Success)

  private def runRasterizers(): IO[ExitCode] =
    formatRasterizerList().flatMap { (output, hasWorking) =>
      IO.println(output).as(ExitCode.Success)
    }

  /**
   * Check all rasterizers and format a status table.
   *
   * Returns (formatted output, true if at least one rasterizer works). Checks run in parallel for
   * better performance.
   *
   * Status terminology:
   *   - available: Works correctly
   *   - missing: Binary not in PATH
   *   - broken: Found but non-functional (e.g., delegate missing)
   *   - unavailable: Cannot be used in current environment (e.g., Batik on native-image)
   */
  private def formatRasterizerList(): IO[(String, Boolean)] =
    // Run all availability checks in parallel
    (
      BatikRasterizer.isAvailable,
      CairoSvg.isAvailable,
      RsvgConvert.isAvailable,
      Resvg.isAvailable,
      ImageMagick.isAvailable,
      ImageMagick.diagnostics
    ).parMapN {
      (batikAvail, cairoAvail, rsvgAvail, resvgAvail, imageMagickAvail, imageMagickDiag) =>
        val sb = new StringBuilder
        sb.append("SVG Rasterizer Status\n")
        sb.append("=" * 60 + "\n\n")
        sb.append(f"${"Backend"}%-14s | ${"Status"}%-11s | ${"Notes"}\n")
        sb.append("-" * 60 + "\n")

        // Batik - "unavailable" when AWT not present (native image)
        val batikStatus = if batikAvail then "available" else "unavailable"
        val batikNote = "Built-in (requires AWT, not native image)"
        sb.append(f"${"batik"}%-14s | ${batikStatus}%-11s | $batikNote\n")

        // CairoSvg
        val cairoStatus = if cairoAvail then "available" else "missing"
        val cairoNote = if cairoAvail then "pip install cairosvg" else "Not in PATH"
        sb.append(f"${"cairosvg"}%-14s | ${cairoStatus}%-11s | $cairoNote\n")

        // rsvg-convert
        val rsvgStatus = if rsvgAvail then "available" else "missing"
        val rsvgNote = if rsvgAvail then "librsvg2-bin" else "Not in PATH"
        sb.append(f"${"rsvg-convert"}%-14s | ${rsvgStatus}%-11s | $rsvgNote\n")

        // resvg
        val resvgStatus = if resvgAvail then "available" else "missing"
        val resvgNote = if resvgAvail then "cargo install resvg" else "Not in PATH"
        sb.append(f"${"resvg"}%-14s | ${resvgStatus}%-11s | $resvgNote\n")

        // ImageMagick (with delegate check)
        // "broken" = found but delegate missing, "missing" = not in PATH
        val imStatus =
          if imageMagickAvail then "available"
          else if imageMagickDiag.contains("missing") then "broken"
          else "missing"
        val imNote =
          val cleaned = imageMagickDiag.replaceAll("ImageMagick \\d+ \\((magick|convert)\\) ", "")
          if cleaned.length > 40 then cleaned.take(37) + "..." else cleaned
        sb.append(f"${"imagemagick"}%-14s | ${imStatus}%-11s | $imNote\n")

        sb.append("\n")

        val anyAvailable = batikAvail || cairoAvail || rsvgAvail || resvgAvail || imageMagickAvail
        if anyAvailable then
          sb.append("At least one rasterizer is available for PNG/JPEG/PDF export.\n")
        else
          sb.append("WARNING: No rasterizers available! PNG/JPEG/PDF export will fail.\n")
          sb.append("\nInstall one of:\n")
          sb.append("  pip install cairosvg           # Python, most portable\n")
          sb.append("  apt install librsvg2-bin       # rsvg-convert, fast\n")
          sb.append("  cargo install resvg            # Rust, best quality\n")
          sb.append("  apt install imagemagick        # Last resort\n")

        (sb.toString, anyAvailable)
    }

  private def formatFunctionList(): String =
    // Dynamically get all functions from the registry
    val names = FunctionRegistry.allNames
    val count = names.size

    val sb = new StringBuilder
    sb.append(s"Supported Excel Functions ($count total)\n")
    sb.append("=" * 40 + "\n\n")

    // Display in columns (5 per row)
    names.grouped(5).foreach { row =>
      sb.append(row.map(n => f"$n%-14s").mkString("  "))
      sb.append("\n")
    }

    sb.append("\nUsage: xl eval \"=FUNCTION(args)\"\n")
    sb.append("Example: xl eval \"=SUM(1,2,3)\" or xl -f data.xlsx eval \"=SUM(A1:A10)\"\n")
    sb.toString

  private def runHeadless(
    filePathOpt: Option[Path],
    sheetNameOpt: Option[String],
    maxSizeOpt: Option[Long],
    cmd: CliCommand
  ): IO[ExitCode] =
    val excel = ExcelIO.instance[IO]
    val readerConfig = buildReaderConfig(maxSizeOpt)
    val workbookIO: IO[Workbook] = filePathOpt match
      case Some(filePath) => excel.readWith(filePath, readerConfig)
      case None => IO.pure(Workbook(Vector.empty)) // Truly empty workbook for constant formulas

    (for
      wb <- workbookIO
      sheet <- SheetResolver.resolveSheet(wb, sheetNameOpt)
      result <- cmd match
        case CliCommand.Eval(formulaStr, overrides) =>
          ReadCommands.eval(wb, sheet, formulaStr, overrides)
        case CliCommand.EvalArray(formulaStr, targetRef, overrides) =>
          ReadCommands.evalArray(wb, sheet, formulaStr, targetRef, overrides)
        case other =>
          IO.raiseError(new Exception(s"Unexpected headless command: $other"))
    yield result).attempt.flatMap {
      case Right(output) =>
        IO.println(output).as(ExitCode.Success)
      case Left(err) =>
        IO.println(Format.errorSimple(err.getMessage)).as(ExitCode.Error)
    }

  private def runStandalone(
    outPath: Path,
    sheetName: String,
    sheets: List[String],
    backendOpt: Option[XmlBackend]
  ): IO[ExitCode] =
    val config = backendOpt.fold(WriterConfig.default)(b => WriterConfig(backend = b))
    (for
      // --sheet takes precedence over --sheet-name; if neither, default to "Sheet1"
      names <- sheets match
        case Nil =>
          IO.fromEither(SheetName(sheetName).left.map(e => new Exception(e))).map(List(_))
        case list =>
          list.traverse(n => IO.fromEither(SheetName(n).left.map(e => new Exception(e))))
      wb = Workbook(names.map(Sheet(_)).toVector)
      _ <- ExcelIO.instance[IO].writeWith(wb, outPath, config)
    yield
      val sheetList = names.map(_.value).mkString(", ")
      s"Created ${outPath.toAbsolutePath} with ${names.size} sheet(s): $sheetList"
    ).attempt.flatMap {
      case Right(output) =>
        IO.println(output).as(ExitCode.Success)
      case Left(err) =>
        IO.println(Format.errorSimple(err.getMessage)).as(ExitCode.Error)
    }

  private def execute(
    filePath: Path,
    sheetNameOpt: Option[String],
    outputOpt: Option[Path],
    backendOpt: Option[XmlBackend],
    maxSizeOpt: Option[Long],
    stream: Boolean,
    cmd: CliCommand
  ): IO[String] =
    // Handle metadata-only commands (instant for any file size)
    cmd match
      // Sheets: list/hide/show operations
      case CliCommand.Sheets(action) =>
        action match
          case SheetsAction.List(stats) =>
            if stats then
              // Full mode: load workbook and get cell counts
              val excel = ExcelIO.instance[IO]
              val readerConfig = buildReaderConfig(maxSizeOpt)
              excel.readWith(filePath, readerConfig).flatMap(wb => WorkbookCommands.sheets(wb))
            else
              // Quick mode: metadata only (instant)
              WorkbookCommands.sheetsQuick(filePath)
          case SheetsAction.Hide(name, veryHide) =>
            // Hide requires loading workbook and writing output
            requireOutputAction(outputOpt, "sheets hide") { outputPath =>
              val excel = ExcelIO.instance[IO]
              val readerConfig = buildReaderConfig(maxSizeOpt)
              val config = buildWriterConfig(backendOpt)
              excel.readWith(filePath, readerConfig).flatMap { wb =>
                SheetCommands.hideSheet(wb, name, veryHide, outputPath, config, stream)
              }
            }
          case SheetsAction.Show(name) =>
            // Show requires loading workbook and writing output
            requireOutputAction(outputOpt, "sheets show") { outputPath =>
              val excel = ExcelIO.instance[IO]
              val readerConfig = buildReaderConfig(maxSizeOpt)
              val config = buildWriterConfig(backendOpt)
              excel.readWith(filePath, readerConfig).flatMap { wb =>
                SheetCommands.showSheet(wb, name, outputPath, config, stream)
              }
            }

      // Names: always lightweight (defined names don't require cell data)
      case CliCommand.Names =>
        WorkbookCommands.namesLight(filePath)

      // Bounds: dimension-first by default (--scan for full scan)
      case CliCommand.Bounds(scan) =>
        if scan then
          // Full streaming scan (accurate)
          StreamingReadCommands.boundsScan(filePath, sheetNameOpt)
        else
          // Dimension element (instant)
          StreamingReadCommands.boundsDimension(filePath, sheetNameOpt)

      // Other commands: regular execution path
      case _ =>
        // For write commands: stream flag enables O(1) output memory (hybrid streaming)
        // For read commands: stream flag enables O(1) input memory (true streaming)
        val isReadCmd = cmd match
          case _: CliCommand.Search | _: CliCommand.Stats | _: CliCommand.Bounds |
              _: CliCommand.View | _: CliCommand.Cell =>
            true
          case _ => false

        // Check for streaming write commands (true O(1) memory transform)
        val isStreamingWriteCmd = cmd match
          case _: CliCommand.Style => true
          case _: CliCommand.Put => true
          case _: CliCommand.PutFormula => true
          case _: CliCommand.Batch => true
          case _ => false

        if stream && isReadCmd then executeStreaming(filePath, sheetNameOpt, cmd)
        else if stream && isStreamingWriteCmd then
          executeStreamingWrite(filePath, sheetNameOpt, outputOpt, cmd)
        else
          val excel = ExcelIO.instance[IO]
          val readerConfig = buildReaderConfig(maxSizeOpt)
          for
            wb <- excel.readWith(filePath, readerConfig)
            sheet <- SheetResolver.resolveSheet(wb, sheetNameOpt)
            result <- executeCommand(wb, sheet, outputOpt, backendOpt, stream, cmd)
          yield result

  /** Execute command using streaming mode (O(1) memory). */
  private def executeStreaming(
    filePath: Path,
    sheetNameOpt: Option[String],
    cmd: CliCommand
  ): IO[String] = cmd match
    case CliCommand.Search(pattern, limit, sheetsFilter) =>
      StreamingReadCommands.search(filePath, sheetNameOpt, pattern, limit, sheetsFilter)

    case CliCommand.Stats(refStr) =>
      StreamingReadCommands.stats(filePath, sheetNameOpt, refStr)

    case CliCommand.Bounds(scan) =>
      if scan then StreamingReadCommands.boundsScan(filePath, sheetNameOpt)
      else StreamingReadCommands.boundsDimension(filePath, sheetNameOpt)

    case CliCommand.View(
          rangeStr,
          showFormulas,
          evalFormulas,
          _, // strict - not used in streaming mode
          limit,
          format,
          _,
          _,
          showLabels,
          _,
          _,
          _,
          skipEmpty,
          headerRow,
          _
        ) =>
      if evalFormulas then
        IO.raiseError(
          new Exception(
            "--eval is not supported with --stream (streaming view uses cached values only)"
          )
        )
      else
        StreamingReadCommands.view(
          filePath,
          sheetNameOpt,
          rangeStr,
          showFormulas,
          limit,
          format,
          showLabels,
          skipEmpty,
          headerRow
        )

    case CliCommand.Cell(refStr, noStyle) =>
      StreamingReadCommands.cell(filePath, sheetNameOpt, refStr, noStyle)

    case _ =>
      IO.raiseError(
        new Exception(
          "--stream not supported for this command. Supported: search, stats, bounds, view (markdown/csv/json only), cell"
        )
      )

  /** Execute streaming write command (O(1) memory transform). */
  private def executeStreamingWrite(
    filePath: Path,
    sheetNameOpt: Option[String],
    outputOpt: Option[Path],
    cmd: CliCommand
  ): IO[String] = cmd match
    case CliCommand.Style(
          rangeStr,
          bold,
          italic,
          underline,
          bg,
          fg,
          fontSize,
          fontName,
          align,
          valign,
          wrap,
          numFormat,
          border,
          borderTop,
          borderRight,
          borderBottom,
          borderLeft,
          borderColor,
          replace
        ) =>
      outputOpt match
        case None =>
          IO.raiseError(new Exception("--output is required for style command"))
        case Some(outputPath) =>
          StreamingWriteCommands.style(
            filePath,
            outputPath,
            sheetNameOpt,
            rangeStr,
            bold,
            italic,
            underline,
            bg,
            fg,
            fontSize,
            fontName,
            align,
            valign,
            wrap,
            numFormat,
            border,
            borderTop,
            borderRight,
            borderBottom,
            borderLeft,
            borderColor,
            replace
          )

    case CliCommand.Put(refStr, values) =>
      outputOpt match
        case None =>
          IO.raiseError(new Exception("--output is required for put command"))
        case Some(outputPath) =>
          StreamingWriteCommands.put(filePath, outputPath, sheetNameOpt, refStr, values)

    case CliCommand.PutFormula(refStr, formulas) =>
      outputOpt match
        case None =>
          IO.raiseError(new Exception("--output is required for putf command"))
        case Some(outputPath) =>
          StreamingWriteCommands.putFormula(filePath, outputPath, sheetNameOpt, refStr, formulas)

    case CliCommand.Batch(source) =>
      outputOpt match
        case None =>
          IO.raiseError(new Exception("--output is required for batch command"))
        case Some(outputPath) =>
          StreamingWriteCommands.batch(filePath, outputPath, sheetNameOpt, source)

    case _ =>
      IO.raiseError(
        new Exception(
          "--stream for write commands only supports: put, putf, style, batch"
        )
      )

  private def executeCommand(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    outputOpt: Option[Path],
    backendOpt: Option[XmlBackend],
    stream: Boolean,
    cmd: CliCommand
  ): IO[String] = cmd match
    // Workbook commands (these are now handled in execute() before reaching here)
    case CliCommand.Sheets(action) =>
      action match
        case SheetsAction.List(_) => WorkbookCommands.sheets(wb)
        case SheetsAction.Hide(name, veryHide) =>
          requireOutput(outputOpt, backendOpt, stream)(
            SheetCommands.hideSheet(wb, name, veryHide, _, _, _)
          )
        case SheetsAction.Show(name) =>
          requireOutput(outputOpt, backendOpt, stream)(
            SheetCommands.showSheet(wb, name, _, _, _)
          )

    case CliCommand.Names =>
      WorkbookCommands.names(wb)

    // Read commands (bounds is now handled in execute() before reaching here)
    case CliCommand.Bounds(_) =>
      ReadCommands.bounds(wb, sheetOpt)

    case CliCommand.View(
          rangeStr,
          showFormulas,
          evalFormulas,
          strict,
          limit,
          format,
          printScale,
          showGridlines,
          showLabels,
          dpi,
          quality,
          rasterOutput,
          skipEmpty,
          headerRow,
          rasterizer
        ) =>
      ReadCommands.view(
        wb,
        sheetOpt,
        rangeStr,
        showFormulas,
        evalFormulas,
        strict,
        limit,
        format,
        printScale,
        showGridlines,
        showLabels,
        dpi,
        quality,
        rasterOutput,
        skipEmpty,
        headerRow,
        rasterizer
      )

    case CliCommand.Cell(refStr, noStyle) =>
      ReadCommands.cell(wb, sheetOpt, refStr, noStyle)

    case CliCommand.Search(pattern, limit, sheetsFilter) =>
      ReadCommands.search(wb, sheetOpt, pattern, limit, sheetsFilter)

    case CliCommand.Stats(refStr) =>
      ReadCommands.stats(wb, sheetOpt, refStr)

    case CliCommand.Eval(formulaStr, overrides) =>
      ReadCommands.eval(wb, sheetOpt, formulaStr, overrides)

    case CliCommand.EvalArray(formulaStr, targetRef, overrides) =>
      ReadCommands.evalArray(wb, sheetOpt, formulaStr, targetRef, overrides)

    // Write commands (require output)
    case CliCommand.Put(refStr, values) =>
      requireOutput(outputOpt, backendOpt, stream)(
        WriteCommands.put(wb, sheetOpt, refStr, values, _, _, _)
      )

    case CliCommand.PutFormula(refStr, formulas) =>
      requireOutput(outputOpt, backendOpt, stream)(
        WriteCommands.putFormula(wb, sheetOpt, refStr, formulas, _, _, _)
      )

    case CliCommand.Style(
          rangeStr,
          bold,
          italic,
          underline,
          bg,
          fg,
          fontSize,
          fontName,
          align,
          valign,
          wrap,
          numFormat,
          border,
          borderTop,
          borderRight,
          borderBottom,
          borderLeft,
          borderColor,
          replace
        ) =>
      requireOutput(outputOpt, backendOpt, stream) { (outputPath, config, streamWrite) =>
        WriteCommands.style(
          wb,
          sheetOpt,
          rangeStr,
          bold,
          italic,
          underline,
          bg,
          fg,
          fontSize,
          fontName,
          align,
          valign,
          wrap,
          numFormat,
          border,
          borderTop,
          borderRight,
          borderBottom,
          borderLeft,
          borderColor,
          replace,
          outputPath,
          config,
          streamWrite
        )
      }

    case CliCommand.RowOp(rowNum, height, hide, show) =>
      requireOutput(outputOpt, backendOpt, stream)(
        WriteCommands.row(wb, sheetOpt, rowNum, height, hide, show, _, _, _)
      )

    case CliCommand.ColOp(colStr, width, hide, show, autoFit) =>
      requireOutput(outputOpt, backendOpt, stream)(
        WriteCommands.col(wb, sheetOpt, colStr, width, hide, show, autoFit, _, _, _)
      )

    case CliCommand.Batch(source) =>
      requireOutput(outputOpt, backendOpt, stream)(
        WriteCommands.batch(wb, sheetOpt, source, _, _, _)
      )

    case CliCommand.Import(csvPath, startRefOpt, delim, skipHeader, enc, newSheetOpt, noInfer) =>
      requireOutput(outputOpt, backendOpt, stream) { (outputPath, writerConfig, streamWrite) =>
        ImportCommands.importCsv(
          wb,
          sheetOpt,
          csvPath,
          startRefOpt,
          delim,
          skipHeader,
          enc,
          newSheetOpt,
          noInfer,
          outputPath,
          writerConfig,
          streamWrite
        )
      }

    // Sheet management commands
    case CliCommand.AddSheet(name, afterOpt, beforeOpt) =>
      requireOutput(outputOpt, backendOpt, stream)(
        SheetCommands.addSheet(wb, name, afterOpt, beforeOpt, _, _, _)
      )

    case CliCommand.RemoveSheet(name) =>
      requireOutput(outputOpt, backendOpt, stream)(SheetCommands.removeSheet(wb, name, _, _, _))

    case CliCommand.RenameSheet(oldName, newName) =>
      requireOutput(outputOpt, backendOpt, stream)(
        SheetCommands.renameSheet(wb, oldName, newName, _, _, _)
      )

    case CliCommand.MoveSheet(name, toIndexOpt, afterOpt, beforeOpt) =>
      requireOutput(outputOpt, backendOpt, stream)(
        SheetCommands.moveSheet(wb, name, toIndexOpt, afterOpt, beforeOpt, _, _, _)
      )

    case CliCommand.CopySheet(sourceName, targetName) =>
      requireOutput(outputOpt, backendOpt, stream)(
        SheetCommands.copySheet(wb, sourceName, targetName, _, _, _)
      )

    // Cell commands
    case CliCommand.Merge(rangeStr) =>
      requireOutput(outputOpt, backendOpt, stream)(
        CellCommands.merge(wb, sheetOpt, rangeStr, _, _, _)
      )

    case CliCommand.Unmerge(rangeStr) =>
      requireOutput(outputOpt, backendOpt, stream)(
        CellCommands.unmerge(wb, sheetOpt, rangeStr, _, _, _)
      )

    case CliCommand.AddComment(refStr, text, author) =>
      requireOutput(outputOpt, backendOpt, stream)(
        CommentCommands.addComment(wb, sheetOpt, refStr, text, author, _, _, _)
      )

    case CliCommand.RemoveComment(refStr) =>
      requireOutput(outputOpt, backendOpt, stream)(
        CommentCommands.removeComment(wb, sheetOpt, refStr, _, _, _)
      )

    case CliCommand.Clear(rangeStr, all, styles, comments) =>
      requireOutput(outputOpt, backendOpt, stream)(
        CellCommands.clear(wb, sheetOpt, rangeStr, all, styles, comments, _, _, _)
      )

    case CliCommand.Fill(source, target, direction) =>
      requireOutput(outputOpt, backendOpt, stream)(
        WriteCommands.fill(wb, sheetOpt, source, target, direction, _, _, _)
      )

    case CliCommand.AutoFit(columnsOpt) =>
      requireOutput(outputOpt, backendOpt, stream)(
        WriteCommands.autoFit(wb, sheetOpt, columnsOpt, _, _, _)
      )

    case CliCommand.Sort(rangeStr, sortKeys, hasHeader) =>
      requireOutput(outputOpt, backendOpt, stream)(
        WriteCommands.sort(wb, sheetOpt, rangeStr, sortKeys, hasHeader, _, _, _)
      )

  // ==========================================================================
  // Helpers
  // ==========================================================================

  private def requireOutput(
    outputOpt: Option[Path],
    backendOpt: Option[XmlBackend],
    stream: Boolean = false
  )(f: (Path, WriterConfig, Boolean) => IO[String]): IO[String] =
    val config = backendOpt.fold(WriterConfig.default)(b => WriterConfig(backend = b))
    outputOpt.fold(IO.raiseError[String](new Exception("Internal: output required")))(path =>
      f(path, config, stream)
    )

  /** Require output path or raise user-friendly error, providing path to action */
  private def requireOutputAction(outputOpt: Option[Path], commandName: String)(
    f: Path => IO[String]
  ): IO[String] =
    outputOpt match
      case Some(path) => f(path)
      case None =>
        IO.raiseError(new Exception(s"$commandName requires -o/--output"))

  /** Build WriterConfig from CLI backend option */
  private def buildWriterConfig(backendOpt: Option[XmlBackend]): WriterConfig =
    backendOpt.fold(WriterConfig.default)(b => WriterConfig(backend = b))

  /** Build ReaderConfig from CLI maxSize option (in MB). 0 means unlimited. */
  private def buildReaderConfig(maxSizeOpt: Option[Long]): ReaderConfig =
    maxSizeOpt match
      case Some(0) => ReaderConfig.permissive
      case Some(mb) => ReaderConfig.default.copy(maxUncompressedSize = mb * 1024 * 1024)
      case None => ReaderConfig.default
