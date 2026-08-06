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
  ChartCommands,
  CommentCommands,
  DiffCommands,
  FilterCommands,
  ImportCommands,
  LintCommands,
  ReadCommands,
  SheetCommands,
  StreamingReadCommands,
  StreamingWriteCommands,
  WorkbookCommands,
  WriteCommands
}
import com.tjclp.xl.ooxml.lint.WorkbookLint
import com.tjclp.xl.cli.raster.{
  BatikRasterizer,
  CairoSvg,
  ImageMagick,
  NativeImage,
  RasterizerChain,
  Resvg,
  RsvgConvert
}
import com.tjclp.xl.cli.helpers.{BatchParser, SheetResolver}
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
 *   - `--no-recalc` / `--preserve-caches` — write verbs only (GH-468): apply the edit and
 *     recalculate nothing. Non-structural verbs keep every cached formula value in the file; the
 *     structural verbs leave the caches their edit invalidated uncached (see [[WritePolicy]])
 *   - `--strict` — write verbs only (GH-496): exit 1 when the write's recalculation reports formula
 *     errors, non-convergence, or data-table seed warnings
 *
 * Global flags precede the verb: `xl -f in.xlsx -o out.xlsx --strict batch -`. (`view --eval
 * --strict` is a separate, subcommand-scoped flag of the same name.)
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
      run(file, None, None, None, None, maxSize, false, cmd)
    }

    // Sheets command: --file required, --output optional (required for hide/show, not for list)
    // This needs its own opts chain because list doesn't need output but hide/show do
    val sheetsOpts =
      (fileOpt, outputOpt.orNone, inPlaceOpt, backendOpt, maxSizeOpt, streamOpt, sheetsCmd).mapN {
        (file, outOpt, inPlace, backend, maxSize, stream, cmd) =>
          runWithOutput(outOpt, inPlace, file) { (out, display) =>
            runResult(file, None, out, display, backend, maxSize, stream, cmd)
          }
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
      boundsCmd orElse viewCmd orElse cellCmd orElse searchCmd orElse statsCmd orElse filterCmd

    val sheetReadOnlyOpts = (fileOpt, sheetOpt, maxSizeOpt, streamOpt, sheetReadOnlySubcmds).mapN {
      (file, sheet, maxSize, stream, cmd) =>
        run(file, sheet, None, None, None, maxSize, stream, cmd)
    }

    // Sheet-level write: --file, --sheet, and --output (required)
    // --stream uses SAX/StAX workbook writes for modifying commands.
    val sheetWriteSubcmds =
      putCmd orElse putfCmd orElse styleCmd orElse rowCmd orElse colCmd orElse groupRowsCmd orElse groupColsCmd orElse ungroupRowsCmd orElse ungroupColsCmd orElse autoFitCmd orElse batchCmd orElse recalcCmd orElse importCmd orElse importMdCmd orElse addSheetCmd orElse removeSheetCmd orElse renameSheetCmd orElse moveSheetCmd orElse copySheetCmd orElse mergeCmd orElse unmergeCmd orElse commentCmd orElse removeCommentCmd orElse clearCmd orElse fillCmd orElse sortCmd orElse freezeCmd orElse unfreezeCmd orElse copyCmd orElse nameCmd orElse insertRowsCmd orElse deleteRowsCmd orElse insertColsCmd orElse deleteColsCmd orElse chartCmd orElse addImageCmd orElse sheetViewCmd orElse tabColorCmd orElse autoFilterCmd orElse pageSetupCmd orElse headerFooterCmd orElse cfCmd

    val sheetWriteOpts =
      (
        fileOpt,
        sheetOpt,
        outputOpt.orNone,
        inPlaceOpt,
        backendOpt,
        maxSizeOpt,
        streamOpt,
        writePolicyOpt,
        sheetWriteSubcmds
      ).mapN { (file, sheet, outOpt, inPlace, backend, maxSize, stream, policy, cmd) =>
        runWithOutput(outOpt, inPlace, file) { (out, display) =>
          runResult(file, sheet, out, display, backend, maxSize, stream, cmd, policy)
        }
      }

    // Standalone: no --file required (creates new files)
    val standaloneOpts = newCmd.map { case (outPath, sheetName, sheets, backend) =>
      runStandalone(outPath, sheetName, sheets, backend)
    }

    // Diff: compares -f against -g (two inputs, no output); custom exit codes
    val diffOpts = (fileOpt, sheetOpt, maxSizeOpt, diffCmd).mapN { (file, sheet, maxSize, cmd) =>
      cmd match
        case CliCommand.Diff(file2, format) => runDiff(file, file2, sheet, maxSize, format)
        case other =>
          IO.println(Format.errorSimple(s"Unexpected diff command: $other")).as(ExitCode.Error)
    }

    // Lint: raw-zip structural validation (GH-397, no output file); custom exit codes.
    // The file arrives via -f or positionally (GH-422); exactly one form must be used.
    val lintOpts = (fileOpt.orNone, lintCmd).mapN { case (flagFile, (cmd, positional)) =>
      cmd match
        case CliCommand.Lint(format) =>
          resolveLintFile(flagFile, positional) match
            case Right(file) => runLint(file, format)
            case Left(msg) => IO.println(Format.errorSimple(msg)).as(ExitCode(2))
        case other =>
          IO.println(Format.errorSimple(s"Unexpected lint command: $other")).as(ExitCode.Error)
    }

    // Info commands: no file required
    val infoOpts = functionsCmd.map(_ => runInfo())
    val rasterOpts = rasterizersCmd.map(_ => runRasterizers())

    // Batch dry-run: only needs batch source, no --file or --output
    val dryRunFlag =
      Opts.flag("dry-run", "Validate batch JSON without writing")
    val batchDryRunOpts =
      Opts
        .subcommand("batch", batchHelp) {
          (batchArg, dryRunFlag).mapN((src, _) => src)
        }
        .map(src => batchDryRun(src).flatMap(IO.println).as(ExitCode.Success))

    rasterOpts orElse infoOpts orElse standaloneOpts orElse diffOpts orElse lintOpts orElse headlessOpts orElse sheetsOpts orElse workbookOpts orElse sheetReadOnlyOpts orElse batchDryRunOpts orElse sheetWriteOpts

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

  private val inPlaceOpt: Opts[Boolean] =
    Opts.flag("in-place", "Edit file in-place (same as -o matching -f)", "i").orFalse

  /**
   * GH-468: skip the trailing recalculation of a write. Two spellings because the field asked for
   * both — `--no-recalc` says what it does, `--preserve-caches` says why you want it.
   */
  private val noRecalcOpt: Opts[Boolean] =
    (
      Opts
        .flag(
          "no-recalc",
          "Apply the edit without recalculating. put/putf/fill/copy/batch keep every cached value in the file; the structural verbs (insert/delete rows/cols) leave every formula the edit invalidated UNCACHED rather than re-stamp a stale number, and report both counts. Use when the caches come from another engine."
        )
        .orFalse,
      Opts.flag("preserve-caches", "Alias for --no-recalc").orFalse
    ).mapN(_ || _)

  /** GH-496: promote a write's advisory recalculation warnings to exit 1 (CI gate). */
  private val strictWriteOpt: Opts[Boolean] =
    Opts
      .flag(
        "strict",
        "Exit 1 when a write's recalculation reports formula errors, non-convergence, or data-table seed warnings (default: advisory, exit 0). The output file is written either way."
      )
      .orFalse

  /** Cross-cutting write posture (GH-468/GH-496), parsed before the verb like -f/-o/--stream. */
  private val writePolicyOpt: Opts[WritePolicy] =
    (noRecalcOpt, strictWriteOpt).mapN(WritePolicy.apply)

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
  private val limitOpt = Opts
    .option[Int]("limit", "Maximum rows to display (default: 50; 0 = no limit)")
    .withDefault(50)
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
  private val skipHiddenOpt =
    Opts
      .flag(
        "skip-hidden",
        "Omit hidden rows/columns (default: render them, marked) — markdown/json/csv only"
      )
      .orFalse
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
  --limit <n>         Max rows to display (default: 50; 0 = no limit).
                      When output is clipped, markdown appends a "… showing X of Y rows"
                      trailer; json adds "truncated"/"totalRows" fields (with --stream the
                      notice goes to stderr instead — streaming json stays a bare array);
                      csv/svg note on stderr; html notes on stderr and appends an HTML
                      comment; raster formats append the notice to the "Exported:" line.
  --formulas          Show formulas instead of values
  --eval              Evaluate formulas (compute live values)
  --strict            Fail on formula evaluation errors (use with --eval)
  --skip-empty        Skip empty cells/rows
  --skip-hidden       Omit hidden rows/columns. Default renders them (an explicitly
                      requested range never silently loses cells) with a marker:
                      markdown appends a "note: …" trailer, csv notes on stderr, json
                      carries "hiddenRows"/"hiddenCols". With --skip-hidden the same
                      marker names what was dropped. Data formats only — html/svg/raster
                      mirror Excel's display and always omit hidden lines.
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

  private val importMdHelp = """Import a GFM markdown table with automatic type detection.

USAGE:
  xl -f file.xlsx -s Sheet1 -o out.xlsx import-md table.md --start A1
  cat table.md | xl -f file.xlsx -s Sheet1 -o out.xlsx import-md -
  xl -f file.xlsx -o out.xlsx import-md table.md --new-sheet "Data"

OPTIONS:
  --start <ref>           Top-left cell for the table (default: A1)
  --skip-header           Skip the table's header row (do not import it)
  --no-type-inference     Treat all values as text
  --new-sheet <name>      Create new sheet for imported data

FORMAT:
  GFM pipe tables: header row, delimiter row (|---|---|), body rows.
  Outer pipes optional; \| inside a cell is a literal pipe; cells are trimmed.
  Alignment markers (:--- left, :---: center, ---: right) become cell alignment.

TYPE DETECTION (same smart detection as batch put):
  Currency:  $1,234.56 → Number + Currency format
  Percent:   45.5% → 0.455 + Percent format
  Dates:     2025-01-15 (ISO 8601) → DateTime + Date format
  Numbers:   100, 29.99, -5.5 | Booleans: true/false | Text: everything else

NOTES:
  - The first table found in the input is imported (preamble text is skipped)
  - Input is read as UTF-8; entire table loads in memory
  - Use "-" to read from stdin

Docs: docs/reference/cli.md (import-md section)

EXAMPLES:
  xl -f f.xlsx -s S1 -o o.xlsx import-md table.md
  xl -f f.xlsx -s S1 -o o.xlsx import-md table.md --start C5 --skip-header
  echo "| A | B |\n|---|---|\n| 1 | 2 |" | xl -f f.xlsx -s S1 -o o.xlsx import-md -"""

  private val putHelp = """Write value(s) to cell or range.

USAGE:
  xl -f file.xlsx -s Sheet1 -o out.xlsx put <ref> <value>
  xl -f file.xlsx -s Sheet1 -o out.xlsx put <range> <value>        # Fill all
  xl -f file.xlsx -s Sheet1 -o out.xlsx put <range> <v1> <v2> ...  # Batch

MODES:
  Single:   put A1 100              → Write 100 to A1
  Fill:     put A1:A10 "TBD"        → Fill range with same value
  Batch:    put A1:C1 "X" "Y" "Z"   → Different value per cell (row-major)

TYPE DETECTION:
  Currency, percent, ISO dates, numbers, and booleans are detected automatically.
  Use --no-detect to preserve every input as text.

NEGATIVE NUMBERS:
  Use --value flag (- is interpreted as flag prefix):
  ❌ put A1 -100              → Error: unknown flag
  ✅ put A1 --value "-100"    → Writes -100 to A1

EXAMPLES:
  xl -f f.xlsx -s S1 -o o.xlsx put A1 "Hello"
  xl -f f.xlsx -s S1 -o o.xlsx put A1 2025-01-15
  xl -f f.xlsx -s S1 -o o.xlsx put A1 2025-01-15 --no-detect
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

  private val diffHelp = """Compare two workbooks and report cell, style, and structure differences.

USAGE:
  xl -f old.xlsx diff -g new.xlsx
  xl -f old.xlsx -s Sheet1 diff -g new.xlsx          # Single sheet only
  xl -f old.xlsx diff -g new.xlsx --format json      # Stable JSON schema

COMPARES (per sheet, refs in A1):
  - Changed cells: value, formula text, resolved style (styleChanged flag)
  - Added / removed cells
  - Sheets added / removed
  - Merged-range, comment, and hyperlink deltas

NOTES:
  - Formula cells compare by formula text (cached values are derived, ignored)
  - Styles compare RESOLVED formatting, not raw style ids
  - Both files load in memory (--max-size applies to each)

EXIT CODES (diff-tool convention):
  0 = files are identical
  1 = differences found
  2 = error (unreadable file, bad sheet filter, ...)

Docs: docs/reference/cli.md (diff section)

EXAMPLES:
  xl -f v1.xlsx diff -g v2.xlsx
  xl -f v1.xlsx diff -g v2.xlsx --format json | jq '.sheets[0].changed'
  xl -f v1.xlsx diff -g v2.xlsx && echo "no changes\""""

  // --- Diff command (GH-137) ---

  private val file2Opt =
    Opts.option[Path]("file2", "Second file to compare against (required)", "g")

  private val diffFormatOpt: Opts[DiffFormat] =
    Opts
      .option[String]("format", "Output format: markdown (default), json")
      .withDefault("markdown")
      .mapValidated { s =>
        s.toLowerCase match
          case "markdown" | "md" => cats.data.Validated.valid(DiffFormat.Markdown)
          case "json" => cats.data.Validated.valid(DiffFormat.Json)
          case other =>
            cats.data.Validated.invalidNel(s"Unknown format: $other. Use markdown or json")
      }

  val diffCmd: Opts[CliCommand] =
    Opts.subcommand("diff", diffHelp) {
      (file2Opt, diffFormatOpt).mapN(CliCommand.Diff.apply)
    }

  // --- Lint command (GH-397) ---

  private val lintHelp =
    """Validate workbook package structure against the Excel-repair classes.

Lints the RAW ZIP PARTS (never the parsed model — a full load would repair
the very structure being checked). Flags what Excel repairs loudly but every
lenient reader accepts silently:
  - Child-element order in xl/workbook.xml (CT_Workbook) and each worksheet
    (CT_Worksheet) vs the ECMA-376 schema sequence
  - r:id references (sheet, externalReference, pivotCache, drawing,
    legacyDrawing, hyperlink, tablePart, ...) that do not resolve in the
    paired .rels, resolve to a relationship of the wrong type, or target a
    part missing from the package
  - ref/sqref/dimension tokens past row 1048576 or column XFD
  - data-table records whose grid was torn by an unguarded edit, and
    uncached table interiors in a calcMode="autoNoTable" book (they open
    BLANK — refresh them with `xl recalc --tables`)
  - formula text stored with a leading '=' inside <f> (non-spec; strict
    readers misread it — re-writing the file with xl heals it)

USAGE:
  xl lint report.xlsx
  xl -f report.xlsx lint                     # Equivalent flag form
  xl lint report.xlsx --format json          # Stable machine-readable schema

FINDING CATEGORIES:
  child-order | unresolved-rel-id | wrong-rel-type | missing-part |
  missing-content-type | ref-out-of-bounds | data-table-torn |
  data-table-unseeded | formula-leading-equals

EXIT CODES:
  0 = no findings (package structure is clean)
  1 = findings reported
  2 = error (unreadable file, missing/malformed core part, ...)

Docs: xl lint is read-only; it never repairs or rewrites the file.

EXAMPLES:
  xl lint deliverable.xlsx && echo "safe to send"
  xl lint deliverable.xlsx --format json | jq '.findings'"""

  private val lintFormatOpt: Opts[LintFormat] =
    Opts
      .option[String]("format", "Output format: text (default), json")
      .withDefault("text")
      .mapValidated { s =>
        s.toLowerCase match
          case "text" => cats.data.Validated.valid(LintFormat.Text)
          case "json" => cats.data.Validated.valid(LintFormat.Json)
          case other =>
            cats.data.Validated.invalidNel(s"Unknown format: $other. Use text or json")
      }

  /**
   * Lint reads exactly one file and writes nothing, so it also accepts the file as a positional
   * argument — `xl lint report.xlsx` — alongside the global `-f` form (GH-422).
   */
  val lintCmd: Opts[(CliCommand, Option[Path])] =
    Opts.subcommand("lint", lintHelp) {
      (lintFormatOpt, Opts.argument[Path]("file").orNone).mapN((format, positional) =>
        (CliCommand.Lint(format), positional)
      )
    }

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

  val nameCmd: Opts[CliCommand] = Opts.subcommand("name", "Manage named ranges: add, rm") {
    val nameArg = Opts.argument[String]("name")
    val refArg = Opts.argument[String]("refers-to")
    val addSub =
      Opts.subcommand("add", "Add or replace a named range (e.g. name add Tax 'Sheet1!$A$1')") {
        (nameArg, refArg).mapN(NameAction.Add.apply)
      }
    val rmSub = Opts.subcommand("rm", "Remove a named range") {
      nameArg.map(NameAction.Remove.apply)
    }
    (addSub orElse rmSub).map(CliCommand.Name.apply)
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
        rasterizerOpt,
        skipHiddenOpt
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

  // --- Filter command (GH-134, phase 1) ---

  private val filterHelp = """Filter rows of the used range with a --where predicate (read-only).

USAGE:
  xl -f data.xlsx -s Sheet1 filter --where "B > 100"
  xl -f data.xlsx -s Sheet1 filter --where "Price > 100" --header
  xl -f data.xlsx -s Sheet1 filter --where "A LIKE 'Widget%'" --columns A,C:E --format csv

PREDICATE GRAMMAR (keywords case-insensitive):
  Comparisons:  B > 100, A = 'Widget', C != TRUE   (= != <> > >= < <=)
  Wildcards:    A LIKE 'Widget%'                   (% matches any run)
  Ranges:       B BETWEEN 10 AND 100               (inclusive)
  Sets:         A IN ('x', 'y', 'z')
  Blanks:       A IS EMPTY / A IS NOT EMPTY
  Logic:        AND, OR, NOT, parentheses          (NOT > AND > OR)

SEMANTICS:
  - Columns are letters (A, B) or header names with --header (first used row;
    header names win over letters on collision, matched case-insensitively)
  - Numbers compare numerically, strings case-insensitively, booleans to TRUE/FALSE
  - Type mismatch (e.g. text cell vs number literal) = row doesn't match, never an error
  - Formula cells compare by cached value

OPTIONS:
  --where <pred>      Filter predicate (required)
  --columns <spec>    Output columns, e.g. A,C:E (default: all used columns)
  --limit <n>         Max rows to display (default: 50)
  --format <fmt>      markdown (default), csv, json
  --header            First used row holds column names (excluded from matching)

NOTES:
  - Read-only: does not modify the file (no -o needed)
  - Loads the workbook in memory; --stream is not supported (use --max-size for large files)
  - Output rows keep their original row numbers

Docs: docs/reference/cli.md (filter section)

EXAMPLES:
  xl -f sales.xlsx -s Q1 filter --where "Revenue > 10000 AND Region = 'EMEA'" --header
  xl -f data.xlsx -s Sheet1 filter --where "B BETWEEN 10 AND 99" --format json
  xl -f data.xlsx -s Sheet1 filter --where "A IS NOT EMPTY" --columns A:C --limit 200"""

  private val whereOpt =
    Opts.option[String]("where", "Filter predicate (e.g. \"B > 100 AND C = 'x'\")")
  private val filterColumnsOpt =
    Opts.option[String]("columns", "Columns to output, e.g. A,C:E (default: all used)").orNone
  private val filterLimitOpt =
    Opts.option[Int]("limit", "Maximum matching rows to display").withDefault(50)
  private val filterFormatOpt: Opts[FilterFormat] =
    Opts
      .option[String]("format", "Output format: markdown (default), csv, json")
      .withDefault("markdown")
      .mapValidated { s =>
        s.toLowerCase match
          case "markdown" | "md" => cats.data.Validated.valid(FilterFormat.Markdown)
          case "csv" => cats.data.Validated.valid(FilterFormat.Csv)
          case "json" => cats.data.Validated.valid(FilterFormat.Json)
          case other =>
            cats.data.Validated.invalidNel(s"Unknown format: $other. Use markdown, csv, or json")
      }
  private val filterHeaderOpt =
    Opts
      .flag("header", "Treat the first used row as column names (excluded from matching)")
      .orFalse

  val filterCmd: Opts[CliCommand] =
    Opts.subcommand("filter", filterHelp) {
      (whereOpt, filterColumnsOpt, filterLimitOpt, filterFormatOpt, filterHeaderOpt)
        .mapN(CliCommand.Filter.apply)
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

  private val csvOpt: Opts[Boolean] =
    Opts
      .flag(
        "csv",
        "Split a single comma-separated value across the target range (count must match)"
      )
      .orFalse

  private val noDetectOpt: Opts[Boolean] =
    Opts
      .flag(
        "no-detect",
        "Preserve put values as text instead of detecting currency, percent, dates, or numbers"
      )
      .orFalse

  val putCmd: Opts[CliCommand] = Opts.subcommand("put", putHelp) {
    // Support both positional args and --value flag (for negative numbers)
    val valuesOrOpt = valueOpt.map(v => List(v)) orElse valuesArg.map(_.toList)
    (refArg, valuesOrOpt, csvOpt, noDetectOpt).mapN { (ref, values, csvSplit, noDetect) =>
      CliCommand.Put(ref, values, csvSplit, detect = !noDetect)
    }
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

  // --- Row/column outline grouping commands (GH-421) ---
  private val groupLevelOpt =
    Opts.option[Int]("level", "Outline level (1-7, default 1)").withDefault(1)
  private val groupCollapsedOpt =
    Opts
      .flag(
        "collapsed",
        "Collapse the group: members are hidden and the summary row/column after the group " +
          "gets the +/- marker"
      )
      .orFalse
  private val groupRowsArg = Opts.argument[String]("rows")
  private val groupColsArg = Opts.argument[String]("cols")

  val groupRowsCmd: Opts[CliCommand] =
    Opts.subcommand("group-rows", "Group rows into a collapsible outline (e.g., 10:20)") {
      (groupRowsArg, groupLevelOpt, groupCollapsedOpt).mapN(CliCommand.GroupRows.apply)
    }

  val groupColsCmd: Opts[CliCommand] =
    Opts.subcommand("group-cols", "Group columns into a collapsible outline (e.g., E:H)") {
      (groupColsArg, groupLevelOpt, groupCollapsedOpt).mapN(CliCommand.GroupCols.apply)
    }

  val ungroupRowsCmd: Opts[CliCommand] =
    Opts.subcommand(
      "ungroup-rows",
      "Remove outline grouping from rows (rows hidden by a collapse stay hidden; " +
        "use `row <n> --show`)"
    ) {
      groupRowsArg.map(CliCommand.UngroupRows.apply)
    }

  val ungroupColsCmd: Opts[CliCommand] =
    Opts.subcommand(
      "ungroup-cols",
      "Remove outline grouping from columns (columns hidden by a collapse stay hidden; " +
        "use `col <letter> --show`)"
    ) {
      groupColsArg.map(CliCommand.UngroupCols.apply)
    }

  // --- Batch command ---
  private val batchArg = Opts.argument[String]("operations").withDefault("-")
  private val batchHelp = """Apply multiple operations atomically from JSON.

USAGE:
  xl -f in.xlsx -s Sheet1 -o out.xlsx batch ops.json
  echo '[...]' | xl -f in.xlsx -s Sheet1 -o out.xlsx batch -

OPERATIONS:
  put       {"op": "put", "ref": "A1", "value": "Hello"}
  putf      {"op": "putf", "ref": "A1", "value": "=SUM(B1:B10)"}  (also accepts "formula";
            optional "format" applies a number format to the formula cell(s))
  style     {"op": "style", "range": "A1:D1", "bold": true, "bg": "#FFFF00"}
  merge     {"op": "merge", "range": "A1:D1"}
  unmerge   {"op": "unmerge", "range": "A1:D1"}
  colwidth  {"op": "colwidth", "col": "A", "width": 15.5}
  rowheight {"op": "rowheight", "row": 1, "height": 30}

ROW/COLUMN GROUPING (GH-421):
  group-rows    {"op": "group-rows", "rows": "10:20", "level": 1, "collapsed": false}
  group-cols    {"op": "group-cols", "cols": "E:H", "level": 1, "collapsed": false}
  ungroup-rows  {"op": "ungroup-rows", "rows": "10:20"}
  ungroup-cols  {"op": "ungroup-cols", "cols": "E:H"}
                (collapsed hides the members and marks the summary row/col after the group)

APPEARANCE & PRINT SETUP (GH-358):
  sheet-view    {"op": "sheet-view", "gridlines": false, "zoom": 85, "tabSelected": true}
  tab-color     {"op": "tab-color", "color": "#1F4E79"}  (or {"clear": true};
                colors: named, #hex, rgb(r,g,b), theme:accent1[:tint])
  autofilter    {"op": "autofilter", "range": "A1:M29"}  (or {"clear": true} to strip
                the sheet's autoFilter, even one preserved from the source file)
  page-setup    {"op": "page-setup", "orientation": "landscape", "scale": 90,
                "fitToWidth": 1, "fitToHeight": 1, "fitToPage": true}
  header-footer {"op": "header-footer", "oddFooter": "&LConfidential&RPage &P of &N",
                "oddHeader": ..., "evenHeader/evenFooter": ..., "firstHeader/firstFooter": ...,
                "differentOddEven": true, "differentFirst": true}
                (&L/&C/&R sections; &P page, &N total, &D date, &F file, &A sheet)

CONDITIONAL FORMATTING (GH-324):
  cf            {"op": "cf", "range": "A1:A10", "rule": "cellIs:greaterThan:100",
                "bold": true, "bg": "#FFC7CE", "fg": "#9C0006"}
                (rule DSL and flags as `xl cf add`; priorities auto-assigned in order)

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

Operations execute in order. Use "-" to read from stdin.
Use --dry-run to validate JSON without writing."""

  private val dryRunOpt =
    Opts.flag("dry-run", "Validate batch JSON and show summary without writing").orFalse

  val batchCmd: Opts[CliCommand] =
    Opts.subcommand("batch", batchHelp) {
      (batchArg, dryRunOpt).mapN(CliCommand.Batch.apply)
    }

  // --- Recalc command (GH-352) ---
  private val recalcHelp = """Recalculate all formulas and rewrite cached values.

Runs one whole-workbook recalculation and writes the result, so every formula
cell carries a cached value (<v>) that cached-value readers see (pandas,
openpyxl data_only=True, previewers, Excel before a manual recalc).

Formula errors (e.g. circular references) are data conditions, not tool
failures: affected cells are left uncached, the file is still written, the
errors are listed in the summary, and the exit code is 0. Pass the global
--strict flag (before the verb) to exit 1 instead when the summary carries
formula errors, iterative non-convergence, or --tables seed warnings.

Data-table records (<f t="dataTable">) keep their PINNED caches by default —
xl never evaluates TABLE(...) implicitly. Pass --tables to also replay each
table's what-if substitution and refresh its interior caches; that is what an
autoNoTable book needs, since Excel never recomputes tables on open.

USAGE:
  xl -f in.xlsx -o out.xlsx recalc
  xl -f in.xlsx -i recalc
  xl -f in.xlsx -o out.xlsx recalc --tables   # Also seed data-table interiors"""

  private val recalcTablesOpt: Opts[Boolean] =
    Opts
      .flag("tables", "Also seed data-table interior caches (default: pinned caches)")
      .orFalse

  val recalcCmd: Opts[CliCommand] =
    Opts.subcommand("recalc", recalcHelp) {
      recalcTablesOpt.map(CliCommand.Recalc.apply)
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

  // --- Import markdown command (GH-159) ---
  private val mdPathArg = Opts.argument[String]("md-file")
  private val mdStartOpt =
    Opts.option[String]("start", "Top-left cell for the imported table (default: A1)").orNone

  val importMdCmd: Opts[CliCommand] =
    Opts.subcommand("import-md", importMdHelp) {
      (mdPathArg, mdStartOpt, skipHeaderOpt, newSheetImportOpt, noTypeInferenceOpt)
        .mapN(CliCommand.ImportMarkdown.apply)
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

  // --- Freeze/Unfreeze commands ---

  private val freezeCmd: Opts[CliCommand] =
    Opts.subcommand(
      "freeze",
      "Freeze panes at cell reference (rows above and columns left are locked)"
    ) {
      refArg.map(CliCommand.Freeze.apply)
    }

  private val unfreezeCmd: Opts[CliCommand] =
    Opts.subcommand("unfreeze", "Remove freeze panes") {
      Opts(CliCommand.Unfreeze)
    }

  // --- Sheet appearance & print setup commands (GH-358) ---

  /** Parse an on|off (also true|false) option value. */
  private def onOffOpt(name: String, help: String): Opts[Option[Boolean]] =
    Opts
      .option[String](name, help)
      .mapValidated {
        case "on" | "true" => cats.data.Validated.valid(true)
        case "off" | "false" => cats.data.Validated.valid(false)
        case other =>
          cats.data.Validated.invalidNel(s"--$name expects on or off, got: $other")
      }
      .orNone

  private val sheetViewCmd: Opts[CliCommand] =
    Opts.subcommand(
      "sheet-view",
      "Set sheet view options: gridlines, zoom, tab selection (requires -o)"
    ) {
      val viewGridlines = onOffOpt("gridlines", "Show cell gridlines: on or off")
      val viewZoom = Opts.option[Int]("zoom", "Zoom percentage (10-400)").orNone
      val viewTabSelected = onOffOpt("tab-selected", "Select this sheet's tab: on or off")
      (viewGridlines, viewZoom, viewTabSelected).mapN(CliCommand.SheetViewOp.apply)
    }

  private val tabColorCmd: Opts[CliCommand] =
    Opts.subcommand(
      "tab-color",
      "Set the sheet tab color: named, #hex, rgb(r,g,b), or theme:accent1[:tint] (requires -o). " +
        "--clear removes a modeled color (a color preserved from the source XML is not stripped)."
    ) {
      val colorArg = Opts.argument[String]("color").orNone
      val clearFlag = Opts.flag("clear", "Clear the modeled tab color").orFalse
      (colorArg, clearFlag).mapN(CliCommand.TabColorOp.apply)
    }

  private val autoFilterCmd: Opts[CliCommand] =
    Opts.subcommand(
      "autofilter",
      "Set the sheet-level autoFilter range (filter dropdowns on the header row) or remove it " +
        "(requires -o). --clear strips the autoFilter even when it was preserved from the " +
        "source file; existing filter criteria ride along when only the range changes."
    ) {
      val rangeArg = Opts.argument[String]("range").orNone
      val clearFlag = Opts.flag("clear", "Remove the sheet's autoFilter").orFalse
      (rangeArg, clearFlag).mapN(CliCommand.AutoFilterOp.apply)
    }

  private val pageSetupCmd: Opts[CliCommand] =
    Opts.subcommand(
      "page-setup",
      "Set print page setup: orientation, scale, fit-to-page (requires -o)"
    ) {
      val orientationOpt =
        Opts.option[String]("orientation", "Page orientation: portrait or landscape").orNone
      val scaleOpt = Opts.option[Int]("scale", "Print scale percent (10-400)").orNone
      val fitToWidthOpt = Opts.option[Int]("fit-to-width", "Fit printout to N pages wide").orNone
      val fitToHeightOpt = Opts.option[Int]("fit-to-height", "Fit printout to N pages tall").orNone
      val fitToPageOpt = onOffOpt(
        "fit-to-page",
        "Force the sheetPr fitToPage flag: on or off. Omit to derive from --fit-to-width/height " +
          "and preserve whatever the source file carries (off actively strips a preserved flag)"
      )
      (orientationOpt, scaleOpt, fitToWidthOpt, fitToHeightOpt, fitToPageOpt)
        .mapN(CliCommand.PageSetupOp.apply)
    }

  private val headerFooterCmd: Opts[CliCommand] =
    Opts.subcommand(
      "header-footer",
      "Set print header/footer text with Excel codes: &L/&C/&R sections, &P page, &N total, " +
        "&D date, &F file, &A sheet (requires -o)"
    ) {
      val oddHeaderOpt = Opts.option[String]("odd-header", "Header for odd/all pages").orNone
      val oddFooterOpt = Opts.option[String]("odd-footer", "Footer for odd/all pages").orNone
      val evenHeaderOpt =
        Opts.option[String]("even-header", "Header for even pages (sets different-odd-even)").orNone
      val evenFooterOpt =
        Opts.option[String]("even-footer", "Footer for even pages (sets different-odd-even)").orNone
      val firstHeaderOpt =
        Opts
          .option[String]("first-header", "Header for the first page (sets different-first)")
          .orNone
      val firstFooterOpt =
        Opts
          .option[String]("first-footer", "Footer for the first page (sets different-first)")
          .orNone
      val diffOddEvenFlag =
        Opts.flag("different-odd-even", "Use even-page text on even pages").orFalse
      val diffFirstFlag = Opts.flag("different-first", "Use first-page text on page 1").orFalse
      (
        oddHeaderOpt,
        oddFooterOpt,
        evenHeaderOpt,
        evenFooterOpt,
        firstHeaderOpt,
        firstFooterOpt,
        diffOddEvenFlag,
        diffFirstFlag
      ).mapN(CliCommand.HeaderFooterOp.apply)
    }

  // --- Copy command ---

  private val copyCmd: Opts[CliCommand] =
    Opts.subcommand("copy", "Copy range to another location (with formula adjustment)") {
      val copySrcArg = Opts.argument[String]("source")
      val copyTgtArg = Opts.argument[String]("target")
      val valuesOnlyOpt =
        Opts.flag("values-only", "Copy values only (no formula adjustment)").orFalse
      (copySrcArg, copyTgtArg, valuesOnlyOpt).mapN(CliCommand.Copy.apply)
    }

  // --- Conditional formatting command (GH-324) ---

  private val cfCmd: Opts[CliCommand] =
    Opts.subcommand("cf", "Conditional formatting: add, list") {
      val addSub = Opts.subcommand(
        "add",
        """Add a conditional-formatting rule to a range (requires -o).

RULE DSL (--rule):
  cellIs:<op>:<value>       op: lessThan|lt, lessThanOrEqual|lte, equal|eq,
                            notEqual|ne, greaterThanOrEqual|gte, greaterThan|gt
  between:<lo>:<hi>         inclusive bounds (notBetween:<lo>:<hi> for the inverse)
  expression:<formula>      custom formula, e.g. expression:MOD(ROW(),2)=0
  colorScale:<c1>:<c2>[:<c3>]  2- or 3-point scale (mid at 50th percentile)
  dataBar:<color>
  top10:<n>[:percent]       bottom10:<n>[:percent] for bottom ranks
  text:<op>:<s>             op: contains, notContains, beginsWith, endsWith

FORMAT FLAGS (highlight rules; colorScale/dataBar carry inline colors instead):
  --bold --italic --underline --strike --bg <color> --fg <color>

Priorities are auto-assigned in add order (lower priority wins in Excel).

EXAMPLES:
  xl -f f.xlsx -s S1 -o o.xlsx cf add --range A1:A10 --rule 'cellIs:greaterThan:100' --bold --bg '#FFC7CE'
  xl -f f.xlsx -s S1 -o o.xlsx cf add --range B2:B20 --rule 'colorScale:red:white:green'
  xl -f f.xlsx -s S1 -o o.xlsx cf add --range C1:C50 --rule 'text:contains:overdue' --fg '#9C0006'"""
      ) {
        val cfRangeOpt =
          Opts.option[String]("range", "Target range, e.g. A1:A10 (qualified refs accepted)")
        val cfRuleOpt = Opts.option[String](
          "rule",
          "Rule DSL string, e.g. cellIs:greaterThan:100 (see cf add --help)"
        )
        val strikeOpt = Opts.flag("strike", "Strikethrough text").orFalse
        (cfRangeOpt, cfRuleOpt, boldOpt, italicOpt, underlineOpt, strikeOpt, bgOpt, fgOpt)
          .mapN(CliCommand.CfAdd.apply)
      }
      val listSub =
        Opts.subcommand("list", "List conditional-formatting rules on the sheet (read-only)") {
          Opts(CliCommand.CfList)
        }
      addSub orElse listSub
    }

  // --- Chart + image commands (GH-222) ---

  private val chartCmd: Opts[CliCommand] =
    Opts.subcommand("chart", "Chart operations: add") {
      Opts.subcommand("add", "Add a typed chart built from sheet data ranges") {
        val typeOpt =
          Opts.option[String]("type", "Chart type: column, bar, line, pie", "t")
        val groupingOpt = Opts
          .option[String](
            "grouping",
            "Bar grouping: clustered (default), stacked, percent-stacked (column/bar only)"
          )
          .orNone
        val dataOpt =
          Opts.option[String](
            "data",
            "Values range (e.g. B2:D10); column categories split per column"
          )
        val categoriesOpt =
          Opts.option[String]("categories", "Categories vector (e.g. A2:A10)").orNone
        val seriesNamesOpt = Opts
          .option[String]("series-names", "Comma-separated literal series names (positional)")
          .orNone
        val seriesColorsOpt = Opts
          .option[String](
            "series-colors",
            "Comma-separated series colors (positional), e.g. #307FE2,#005670; " +
              "unset series cycle the theme accents (bar/column/line only)"
          )
          .orNone
        val titleOpt = Opts.option[String]("title", "Chart title").orNone
        val legendOpt = Opts
          .option[String](
            "legend",
            "Legend position: right (default), left, top, bottom, top-right, none"
          )
          .orNone
        val atOpt = Opts.option[String](
          "at",
          "Placement: a range (chart stretches over it) or a single cell (default size)"
        )
        (
          typeOpt,
          groupingOpt,
          dataOpt,
          categoriesOpt,
          seriesNamesOpt,
          seriesColorsOpt,
          titleOpt,
          legendOpt,
          atOpt
        ).mapN(CliCommand.ChartAdd.apply)
      }
    }

  private val addImageCmd: Opts[CliCommand] =
    Opts.subcommand("add-image", "Embed an image (png/jpeg/gif/bmp/tiff/emf/wmf)") {
      val imageArg = Opts.argument[Path]("image-file")
      val atOpt = Opts.option[String](
        "at",
        "Placement: a single cell (natural size unless --size) or a range (image stretches)"
      )
      val sizeOpt =
        Opts.option[String]("size", "Pixel size WxH for a single-cell --at (e.g. 320x240)").orNone
      (imageArg, atOpt, sizeOpt).mapN(CliCommand.AddImage.apply)
    }

  // --- Structural editing commands (insert/delete rows & columns) ---

  private val insertRowsCmd: Opts[CliCommand] =
    Opts.subcommand("insert-rows", "Insert rows (shifts cells & rewrites formulas)") {
      val atArg = Opts.argument[Int]("at-row")
      val countArg = Opts.argument[Int]("count").withDefault(1)
      (atArg, countArg).mapN(CliCommand.InsertRows.apply)
    }

  private val deleteRowsCmd: Opts[CliCommand] =
    Opts.subcommand(
      "delete-rows",
      "Delete rows (shifts cells & rewrites formulas; #REF! on loss)"
    ) {
      val atArg = Opts.argument[Int]("at-row")
      val countArg = Opts.argument[Int]("count").withDefault(1)
      (atArg, countArg).mapN(CliCommand.DeleteRows.apply)
    }

  private val insertColsCmd: Opts[CliCommand] =
    Opts.subcommand("insert-cols", "Insert columns (shifts cells & rewrites formulas)") {
      val colArg = Opts.argument[String]("at-col")
      val countArg = Opts.argument[Int]("count").withDefault(1)
      (colArg, countArg).mapN(CliCommand.InsertColumns.apply)
    }

  private val deleteColsCmd: Opts[CliCommand] =
    Opts.subcommand(
      "delete-cols",
      "Delete columns (shifts cells & rewrites formulas; #REF! on loss)"
    ) {
      val colArg = Opts.argument[String]("at-col")
      val countArg = Opts.argument[Int]("count").withDefault(1)
      (colArg, countArg).mapN(CliCommand.DeleteColumns.apply)
    }

  // ==========================================================================
  // Command execution
  // ==========================================================================

  private def run(
    filePath: Path,
    sheetNameOpt: Option[String],
    outputOpt: Option[Path],
    displayOpt: Option[Path],
    backendOpt: Option[XmlBackend],
    maxSizeOpt: Option[Long],
    stream: Boolean,
    cmd: CliCommand
  ): IO[ExitCode] =
    runResult(
      filePath,
      sheetNameOpt,
      outputOpt,
      displayOpt,
      backendOpt,
      maxSizeOpt,
      stream,
      cmd
    ).flatMap(printRunResult)

  /**
   * Execute and render a command without printing, so in-place writes can commit first.
   *
   * GH-496: a [[StrictFailure]] is not a crash — the write completed and its summary is printed
   * verbatim (counts, failing refs, convergence verdict, plus the strict reason) — only the exit
   * code becomes 1. With `-i` that non-success code means the temp file is discarded and the input
   * is left untouched, which is the atomic reading of "this book did not pass the gate" — so the
   * summary's `Saved:` line is rewritten to say exactly that. With `-o` the output file genuinely
   * is written and `Saved:` stands.
   */
  private def runResult(
    filePath: Path,
    sheetNameOpt: Option[String],
    outputOpt: Option[Path],
    displayOpt: Option[Path],
    backendOpt: Option[XmlBackend],
    maxSizeOpt: Option[Long],
    stream: Boolean,
    cmd: CliCommand,
    policy: WritePolicy = WritePolicy.default
  ): IO[(ExitCode, String)] =
    execute(filePath, sheetNameOpt, outputOpt, backendOpt, maxSizeOpt, stream, cmd, policy).attempt
      .map {
        case Right(output) =>
          (ExitCode.Success, renderWithTarget(output, outputOpt, displayOpt))
        case Left(strict: StrictFailure) =>
          (
            ExitCode(1),
            unsayTheSave(
              renderWithTarget(strict.summary, outputOpt, displayOpt),
              outputOpt,
              displayOpt
            )
          )
        case Left(err) =>
          (ExitCode.Error, renderErrorMessage(err, outputOpt, displayOpt))
      }

  /**
   * GH-496: an in-place run whose exit code is non-success never commits its temp file, so the
   * summary a write command already built ("...\nSaved: <path>") describes a save that did not
   * happen. Replace that line with what is true. Only `-i` is affected: with `-o` the output file
   * IS written before the gate runs, so its `Saved:` line stays.
   */
  private def unsayTheSave(
    text: String,
    outputOpt: Option[Path],
    displayOpt: Option[Path]
  ): String =
    (outputOpt, displayOpt) match
      case (Some(write), Some(display)) if write != display =>
        text.linesIterator
          .map { line =>
            if line.startsWith("Saved: ") || line.startsWith("Saved (streaming): ") then
              s"NOT saved (--strict failure): $display left untouched"
            else line
          }
          .mkString("\n")
      case _ => text

  /**
   * In-place writes (-i) target a temp file that is atomically moved onto the input after the
   * command succeeds; user-facing messages must name the user-visible path, not the temp file that
   * no longer exists by the time it is read (GH-464).
   */
  private def renderWithTarget(
    text: String,
    outputOpt: Option[Path],
    displayOpt: Option[Path]
  ): String =
    (outputOpt, displayOpt) match
      case (Some(write), Some(display)) if write != display =>
        text.replace(write.toString, display.toString)
      case _ => text

  /**
   * GH-483: failure messages get the same temp→target rewrite as successes — an exception raised
   * mid-write (e.g. disk full) embeds the already-deleted .xl-inplace- temp path otherwise. Guards
   * `getMessage` being null (falls back to the exception's toString).
   */
  private[cli] def renderErrorMessage(
    err: Throwable,
    outputOpt: Option[Path],
    displayOpt: Option[Path]
  ): String =
    val message = Option(err.getMessage).getOrElse(err.toString)
    Format.errorSimple(renderWithTarget(message, outputOpt, displayOpt))

  private def printRunResult(result: (ExitCode, String)): IO[ExitCode] =
    IO.println(result._2).as(result._1)

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
        renderRasterizerTable(
          batikAvail = batikAvail,
          cairoAvail = cairoAvail,
          rsvgAvail = rsvgAvail,
          resvgAvail = resvgAvail,
          imageMagickAvail = imageMagickAvail,
          imageMagickDiag = imageMagickDiag,
          nativeImage = NativeImage.inNativeImage
        )
    }

  /**
   * Pure formatting core of `xl rasterizers` - separated from the availability probes so every
   * environment (JVM, native binary, nothing installed) is unit-testable.
   *
   * Returns (formatted output, true if at least one rasterizer works).
   *
   * Status terminology:
   *   - available: Works correctly
   *   - missing: Binary not in PATH
   *   - broken: Found but non-functional (e.g., delegate missing)
   *   - unavailable: Cannot be used in current environment (e.g., Batik on native-image)
   */
  private[cli] def renderRasterizerTable(
    batikAvail: Boolean,
    cairoAvail: Boolean,
    rsvgAvail: Boolean,
    resvgAvail: Boolean,
    imageMagickAvail: Boolean,
    imageMagickDiag: String,
    nativeImage: Boolean
  ): (String, Boolean) =
    val sb = new StringBuilder
    sb.append("SVG Rasterizer Status\n")
    sb.append("=" * 60 + "\n\n")
    sb.append(f"${"Backend"}%-14s | ${"Status"}%-11s | ${"Notes"}\n")
    sb.append("-" * 60 + "\n")

    // Batik - the default backend; "unavailable" when AWT not present (native image)
    val batikStatus = if batikAvail then "available" else "unavailable"
    val batikNote =
      if batikAvail then "Built-in default (requires AWT)"
      else if nativeImage then "Native binary: no AWT, by design"
      else "Requires AWT (not present here)"
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

    // ImageMagick (with delegate check) - explicit opt-in only, never tried automatically
    // "broken" = found but delegate missing, "missing" = not in PATH
    val imStatus =
      if imageMagickAvail then "available"
      else if imageMagickDiag.contains("missing") then "broken"
      else "missing"
    val imNote =
      val cleaned = imageMagickDiag.replaceAll("ImageMagick \\d+ \\((magick|convert)\\) ", "")
      val withOptIn =
        if imageMagickAvail then s"--rasterizer imagemagick only; $cleaned" else cleaned
      if withOptIn.length > 40 then withOptIn.take(37) + "..." else withOptIn
    sb.append(f"${"imagemagick"}%-14s | ${imStatus}%-11s | $imNote\n")

    sb.append("\n")

    val anyAvailable = batikAvail || cairoAvail || rsvgAvail || resvgAvail || imageMagickAvail
    if anyAvailable then
      sb.append("At least one rasterizer is available for PNG/JPEG/PDF export.\n")
    else
      sb.append("WARNING: No rasterizers available! PNG/JPEG/PDF export will fail.\n")
      if nativeImage then
        sb.append("This is the native binary: the bundled Batik backend needs AWT and can\n")
        sb.append("never work here - install one of the external tools below.\n")
      sb.append("\nInstall one of:\n")
      sb.append("  pip install cairosvg           # Python, most portable\n")
      sb.append("  apt install librsvg2-bin       # rsvg-convert, fast\n")
      sb.append(
        "  cargo install resvg            # or prebuilt: github.com/linebender/resvg/releases\n"
      )
      sb.append("  apt install imagemagick        # then pass --rasterizer imagemagick\n")

    (sb.toString, anyAvailable)

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

  /**
   * Run the diff command with diff-tool exit codes: 0 = identical, 1 = differences found, 2 = error
   * (unreadable file, sheet filter matching neither workbook, ...).
   */
  private[cli] def runDiff(
    fileA: Path,
    fileB: Path,
    sheetFilter: Option[String],
    maxSizeOpt: Option[Long],
    format: DiffFormat
  ): IO[ExitCode] =
    val excel = ExcelIO.instance[IO]
    val readerConfig = buildReaderConfig(maxSizeOpt)
    (for
      wbA <- excel.readWith(fileA, readerConfig)
      wbB <- excel.readWith(fileB, readerConfig)
      diff <- DiffCommands.computeDiff(wbA, wbB, sheetFilter) match
        case Right(d) => IO.pure(d)
        case Left(err) => IO.raiseError(new Exception(err))
      output = format match
        case DiffFormat.Markdown =>
          DiffCommands.renderMarkdown(diff, fileA.toString, fileB.toString)
        case DiffFormat.Json => DiffCommands.renderJson(diff)
    yield (output, diff.identical)).attempt.flatMap {
      case Right((output, identical)) =>
        IO.println(output).as(if identical then ExitCode.Success else ExitCode(1))
      case Left(err) =>
        IO.println(Format.errorSimple(err.getMessage)).as(ExitCode(2))
    }

  /**
   * Resolve lint's input file from the `-f` flag and the positional form — exactly one must be
   * given (GH-422). Errors use lint's exit-code 2 convention at the call site.
   */
  private[cli] def resolveLintFile(
    flagFile: Option[Path],
    positional: Option[Path]
  ): Either[String, Path] =
    (flagFile, positional) match
      case (Some(file), None) => Right(file)
      case (None, Some(file)) => Right(file)
      case (Some(f), Some(p)) =>
        Left(s"lint takes exactly one file — got both -f '$f' and positional '$p'")
      case (None, None) =>
        Left("lint requires a file: xl lint <file> (or xl -f <file> lint)")

  /**
   * Run the lint command with its exit-code convention: 0 = clean, 1 = findings, 2 = error
   * (unreadable file, missing/malformed core part). Opens the zip directly — NOT ExcelIO.read —
   * because a full parse would repair/normalize the very structure lint inspects (GH-397).
   */
  private[cli] def runLint(file: Path, format: LintFormat): IO[ExitCode] =
    IO.blocking(WorkbookLint.lint(file)).flatMap {
      case Right(findings) =>
        val output = format match
          case LintFormat.Text => LintCommands.renderText(file.toString, findings)
          case LintFormat.Json => LintCommands.renderJson(file.toString, findings)
        IO.println(output).as(if findings.isEmpty then ExitCode.Success else ExitCode(1))
      case Left(err) =>
        IO.println(Format.errorSimple(err.message)).as(ExitCode(2))
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
    cmd: CliCommand,
    policy: WritePolicy = WritePolicy.default
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
        // For write commands: stream flag uses the SAX/StAX workbook writer
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
          // GH-496: a streaming write never recalculates, so --strict could only ever report
          // "clean" — refuse rather than hand a CI lane a gate that cannot fail. --no-recalc
          // needs no guard: it is what the streaming path already does.
          if policy.strict then
            IO.raiseError(
              new Exception(
                "--strict is not supported with --stream (streaming writes never recalculate). Re-run without --stream."
              )
            )
          else executeStreamingWrite(filePath, sheetNameOpt, outputOpt, cmd)
        else
          val excel = ExcelIO.instance[IO]
          val readerConfig = buildReaderConfig(maxSizeOpt)
          for
            wb <- excel.readWith(filePath, readerConfig)
            sheet <- SheetResolver.resolveSheet(wb, sheetNameOpt)
            result <- executeCommand(wb, sheet, outputOpt, backendOpt, stream, cmd, policy)
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
          _,
          skipHidden
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
          headerRow,
          // GH-474: unsupported under --stream, but reported rather than silently dropped
          skipHidden
        )

    case CliCommand.Cell(refStr, noStyle) =>
      StreamingReadCommands.cell(filePath, sheetNameOpt, refStr, noStyle)

    case _ =>
      IO.raiseError(
        new Exception(
          "--stream not supported for this command. Supported: search, stats, bounds, view (markdown/csv/json only), cell"
        )
      )

  /** Validate batch JSON and show summary without writing. */
  private def batchDryRun(source: String): IO[String] =
    BatchParser.readBatchInput(source).flatMap { input =>
      BatchParser.parseBatchOperations(input).flatMap { result =>
        IO(result.warnings.foreach(System.err.println)) *>
          IO.pure {
            val summary = BatchParser.formatSummary(result.ops)
            s"Dry run - ${result.ops.size} operations parsed:\n$summary"
          }
      }
    }

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

    case CliCommand.Put(refStr, values, csvSplit, detect) =>
      if csvSplit then
        IO.raiseError(
          new Exception(
            "--csv auto-split is not supported with --stream. Omit --stream to use --csv."
          )
        )
      else
        outputOpt match
          case None =>
            IO.raiseError(new Exception("--output is required for put command"))
          case Some(outputPath) =>
            StreamingWriteCommands.put(
              filePath,
              outputPath,
              sheetNameOpt,
              refStr,
              values,
              detect
            )

    case CliCommand.PutFormula(refStr, formulas) =>
      outputOpt match
        case None =>
          IO.raiseError(new Exception("--output is required for putf command"))
        case Some(outputPath) =>
          StreamingWriteCommands.putFormula(filePath, outputPath, sheetNameOpt, refStr, formulas)

    case CliCommand.Batch(source, dryRun) if dryRun =>
      batchDryRun(source)

    case CliCommand.Batch(source, _) =>
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

  private[cli] def executeCommand(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    outputOpt: Option[Path],
    backendOpt: Option[XmlBackend],
    stream: Boolean,
    cmd: CliCommand,
    policy: WritePolicy = WritePolicy.default
  ): IO[String] = cmd match
    // Workbook commands (these are now handled in execute() before reaching here)
    case CliCommand.Sheets(action) =>
      action match
        case SheetsAction.List(_) => WorkbookCommands.sheets(wb)
        case SheetsAction.Hide(name, veryHide) =>
          requireOutput("sheets hide", outputOpt, backendOpt, stream)(
            SheetCommands.hideSheet(wb, name, veryHide, _, _, _)
          )
        case SheetsAction.Show(name) =>
          requireOutput("sheets show", outputOpt, backendOpt, stream)(
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
          rasterizer,
          skipHidden
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
        rasterizer,
        skipHidden
      )

    case CliCommand.Cell(refStr, noStyle) =>
      ReadCommands.cell(wb, sheetOpt, refStr, noStyle)

    case CliCommand.Search(pattern, limit, sheetsFilter) =>
      ReadCommands.search(wb, sheetOpt, pattern, limit, sheetsFilter)

    case CliCommand.Stats(refStr) =>
      ReadCommands.stats(wb, sheetOpt, refStr)

    case CliCommand.Filter(where, columns, limit, format, header) =>
      FilterCommands.filter(wb, sheetOpt, where, columns, limit, format, header)

    case CliCommand.Eval(formulaStr, overrides) =>
      ReadCommands.eval(wb, sheetOpt, formulaStr, overrides)

    case CliCommand.EvalArray(formulaStr, targetRef, overrides) =>
      ReadCommands.evalArray(wb, sheetOpt, formulaStr, targetRef, overrides)

    // Write commands (require output)
    case CliCommand.Put(refStr, values, csvSplit, detect) =>
      requireOutput("put", outputOpt, backendOpt, stream)(
        WriteCommands.put(wb, sheetOpt, refStr, values, _, _, _, csvSplit, detect, policy)
      )

    case CliCommand.PutFormula(refStr, formulas) =>
      requireOutput("putf", outputOpt, backendOpt, stream)(
        WriteCommands.putFormula(wb, sheetOpt, refStr, formulas, _, _, _, policy)
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
      requireOutput("style", outputOpt, backendOpt, stream) { (outputPath, config, streamWrite) =>
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
      requireOutput("row", outputOpt, backendOpt, stream)(
        WriteCommands.row(wb, sheetOpt, rowNum, height, hide, show, _, _, _)
      )

    case CliCommand.ColOp(colStr, width, hide, show, autoFit) =>
      requireOutput("col", outputOpt, backendOpt, stream)(
        WriteCommands.col(wb, sheetOpt, colStr, width, hide, show, autoFit, _, _, _)
      )

    // Row/column outline grouping (GH-421)
    case CliCommand.GroupRows(rows, level, collapsed) =>
      requireOutput("group-rows", outputOpt, backendOpt, stream)(
        WriteCommands.groupRows(wb, sheetOpt, rows, level, collapsed, _, _, _)
      )

    case CliCommand.GroupCols(cols, level, collapsed) =>
      requireOutput("group-cols", outputOpt, backendOpt, stream)(
        WriteCommands.groupCols(wb, sheetOpt, cols, level, collapsed, _, _, _)
      )

    case CliCommand.UngroupRows(rows) =>
      requireOutput("ungroup-rows", outputOpt, backendOpt, stream)(
        WriteCommands.ungroupRows(wb, sheetOpt, rows, _, _, _)
      )

    case CliCommand.UngroupCols(cols) =>
      requireOutput("ungroup-cols", outputOpt, backendOpt, stream)(
        WriteCommands.ungroupCols(wb, sheetOpt, cols, _, _, _)
      )

    case CliCommand.Batch(source, dryRun) if dryRun =>
      batchDryRun(source)

    case CliCommand.Batch(source, _) =>
      requireOutput("batch", outputOpt, backendOpt, stream)(
        WriteCommands.batch(wb, sheetOpt, source, _, _, _, policy)
      )

    case CliCommand.Recalc(tables) =>
      // A recalculation asked not to recalculate is a contradiction, not a no-op: say so rather
      // than writing a file the caller will read as freshened (GH-468).
      if policy.noRecalc then
        IO.raiseError(
          new Exception(
            "recalc cannot be combined with --no-recalc/--preserve-caches (it exists to rewrite caches). Drop the flag, or drop the recalc."
          )
        )
      else
        requireOutput("recalc", outputOpt, backendOpt, stream)(
          WriteCommands.recalc(wb, _, _, _, tables, policy)
        )

    case CliCommand.Import(csvPath, startRefOpt, delim, skipHeader, enc, newSheetOpt, noInfer) =>
      requireOutput("import", outputOpt, backendOpt, stream) {
        (outputPath, writerConfig, streamWrite) =>
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

    case CliCommand.ImportMarkdown(mdPath, startRefOpt, skipHeader, newSheetOpt, noInfer) =>
      requireOutput("import-md", outputOpt, backendOpt, stream) {
        (outputPath, writerConfig, streamWrite) =>
          ImportCommands.importMarkdown(
            wb,
            sheetOpt,
            mdPath,
            startRefOpt,
            skipHeader,
            newSheetOpt,
            noInfer,
            outputPath,
            writerConfig,
            streamWrite
          )
      }

    // Sheet management commands
    case CliCommand.AddSheet(name, afterOpt, beforeOpt) =>
      requireOutput("add-sheet", outputOpt, backendOpt, stream)(
        SheetCommands.addSheet(wb, name, afterOpt, beforeOpt, _, _, _)
      )

    case CliCommand.RemoveSheet(name) =>
      requireOutput("remove-sheet", outputOpt, backendOpt, stream)(
        SheetCommands.removeSheet(wb, name, _, _, _)
      )

    case CliCommand.RenameSheet(oldName, newName) =>
      requireOutput("rename-sheet", outputOpt, backendOpt, stream)(
        SheetCommands.renameSheet(wb, oldName, newName, _, _, _)
      )

    case CliCommand.Name(action) =>
      action match
        case NameAction.Add(nm, refersTo) =>
          requireOutput("name add", outputOpt, backendOpt, stream)(
            SheetCommands.nameAdd(wb, nm, refersTo, _, _, _)
          )
        case NameAction.Remove(nm) =>
          requireOutput("name rm", outputOpt, backendOpt, stream)(
            SheetCommands.nameRemove(wb, nm, _, _, _)
          )

    case CliCommand.MoveSheet(name, toIndexOpt, afterOpt, beforeOpt) =>
      requireOutput("move-sheet", outputOpt, backendOpt, stream)(
        SheetCommands.moveSheet(wb, name, toIndexOpt, afterOpt, beforeOpt, _, _, _)
      )

    case CliCommand.CopySheet(sourceName, targetName) =>
      requireOutput("copy-sheet", outputOpt, backendOpt, stream)(
        SheetCommands.copySheet(wb, sourceName, targetName, _, _, _)
      )

    // Cell commands
    case CliCommand.Merge(rangeStr) =>
      requireOutput("merge", outputOpt, backendOpt, stream)(
        CellCommands.merge(wb, sheetOpt, rangeStr, _, _, _)
      )

    case CliCommand.Unmerge(rangeStr) =>
      requireOutput("unmerge", outputOpt, backendOpt, stream)(
        CellCommands.unmerge(wb, sheetOpt, rangeStr, _, _, _)
      )

    case CliCommand.AddComment(refStr, text, author) =>
      requireOutput("comment", outputOpt, backendOpt, stream)(
        CommentCommands.addComment(wb, sheetOpt, refStr, text, author, _, _, _)
      )

    case CliCommand.RemoveComment(refStr) =>
      requireOutput("remove-comment", outputOpt, backendOpt, stream)(
        CommentCommands.removeComment(wb, sheetOpt, refStr, _, _, _)
      )

    case CliCommand.Clear(rangeStr, all, styles, comments) =>
      requireOutput("clear", outputOpt, backendOpt, stream)(
        CellCommands.clear(wb, sheetOpt, rangeStr, all, styles, comments, _, _, _)
      )

    case CliCommand.Fill(source, target, direction) =>
      requireOutput("fill", outputOpt, backendOpt, stream)(
        WriteCommands.fill(wb, sheetOpt, source, target, direction, _, _, _, policy)
      )

    case CliCommand.AutoFit(columnsOpt) =>
      requireOutput("autofit", outputOpt, backendOpt, stream)(
        WriteCommands.autoFit(wb, sheetOpt, columnsOpt, _, _, _)
      )

    case CliCommand.Sort(rangeStr, sortKeys, hasHeader) =>
      requireOutput("sort", outputOpt, backendOpt, stream)(
        WriteCommands.sort(wb, sheetOpt, rangeStr, sortKeys, hasHeader, _, _, _)
      )

    case CliCommand.Freeze(refStr) =>
      requireOutput("freeze", outputOpt, backendOpt, stream)(
        WriteCommands.freeze(wb, sheetOpt, refStr, _, _, _)
      )

    case CliCommand.Unfreeze =>
      requireOutput("unfreeze", outputOpt, backendOpt, stream)(
        WriteCommands.unfreeze(wb, sheetOpt, _, _, _)
      )

    // Sheet appearance & print setup (GH-358)
    case CliCommand.SheetViewOp(gridlines, zoom, tabSelected) =>
      requireOutput("sheet-view", outputOpt, backendOpt, stream)(
        WriteCommands.sheetView(wb, sheetOpt, gridlines, zoom, tabSelected, _, _, _)
      )

    case CliCommand.TabColorOp(color, clear) =>
      requireOutput("tab-color", outputOpt, backendOpt, stream)(
        WriteCommands.tabColor(wb, sheetOpt, color, clear, _, _, _)
      )

    // Sheet-level autoFilter authoring (GH-432)
    case CliCommand.AutoFilterOp(range, clear) =>
      requireOutput("autofilter", outputOpt, backendOpt, stream)(
        WriteCommands.autoFilter(wb, sheetOpt, range, clear, _, _, _)
      )

    case CliCommand.PageSetupOp(orientation, scale, fitToWidth, fitToHeight, fitToPage) =>
      requireOutput("page-setup", outputOpt, backendOpt, stream)(
        WriteCommands.pageSetup(
          wb,
          sheetOpt,
          orientation,
          scale,
          fitToWidth,
          fitToHeight,
          fitToPage,
          _,
          _,
          _
        )
      )

    case CliCommand.HeaderFooterOp(
          oddHeader,
          oddFooter,
          evenHeader,
          evenFooter,
          firstHeader,
          firstFooter,
          differentOddEven,
          differentFirst
        ) =>
      requireOutput("header-footer", outputOpt, backendOpt, stream)(
        WriteCommands.headerFooter(
          wb,
          sheetOpt,
          oddHeader,
          oddFooter,
          evenHeader,
          evenFooter,
          firstHeader,
          firstFooter,
          differentOddEven,
          differentFirst,
          _,
          _,
          _
        )
      )

    // Conditional formatting (GH-324)
    case CliCommand.CfAdd(range, rule, bold, italic, underline, strike, bg, fg) =>
      requireOutput("cf add", outputOpt, backendOpt, stream)(
        WriteCommands.cfAdd(
          wb,
          sheetOpt,
          range,
          rule,
          bold,
          italic,
          underline,
          strike,
          bg,
          fg,
          _,
          _,
          _
        )
      )

    case CliCommand.CfList =>
      WriteCommands.cfList(wb, sheetOpt)

    case CliCommand.Copy(source, target, valuesOnly) =>
      requireOutput("copy", outputOpt, backendOpt, stream)(
        WriteCommands.copyRange(wb, sheetOpt, source, target, valuesOnly, _, _, _, policy)
      )

    case CliCommand.ChartAdd(
          typeStr,
          grouping,
          data,
          categories,
          seriesNames,
          seriesColors,
          title,
          legend,
          at
        ) =>
      requireOutput("chart add", outputOpt, backendOpt, stream)(
        ChartCommands.chartAdd(
          wb,
          sheetOpt,
          typeStr,
          grouping,
          data,
          categories,
          seriesNames,
          seriesColors,
          title,
          legend,
          at,
          _,
          _,
          _
        )
      )

    case CliCommand.AddImage(imagePath, at, size) =>
      requireOutput("add-image", outputOpt, backendOpt, stream)(
        ChartCommands.addImage(wb, sheetOpt, imagePath, at, size, _, _, _)
      )

    case CliCommand.InsertRows(at, count) =>
      requireOutput("insert-rows", outputOpt, backendOpt, stream)(
        WriteCommands.insertRows(wb, sheetOpt, at, count, _, _, _, policy)
      )

    case CliCommand.DeleteRows(at, count) =>
      requireOutput("delete-rows", outputOpt, backendOpt, stream)(
        WriteCommands.deleteRows(wb, sheetOpt, at, count, _, _, _, policy)
      )

    case CliCommand.InsertColumns(col, count) =>
      requireOutput("insert-cols", outputOpt, backendOpt, stream)(
        WriteCommands.insertColumns(wb, sheetOpt, col, count, _, _, _, policy)
      )

    case CliCommand.DeleteColumns(col, count) =>
      requireOutput("delete-cols", outputOpt, backendOpt, stream)(
        WriteCommands.deleteColumns(wb, sheetOpt, col, count, _, _, _, policy)
      )

    // Diff has its own runner (two input files, custom exit codes) — never reaches here
    case CliCommand.Diff(_, _) =>
      IO.raiseError(new Exception("Internal: diff is dispatched via runDiff"))

    // Lint has its own runner (raw-zip inspection, custom exit codes) — never reaches here
    case CliCommand.Lint(_) =>
      IO.raiseError(new Exception("Internal: lint is dispatched via runLint"))

  // ==========================================================================
  // Helpers
  // ==========================================================================

  /** Usage message for write commands invoked without -o/-i — names the flag (GH-422). */
  private def missingOutputError(commandName: String): String =
    s"$commandName requires -o <out.xlsx> (or -i to modify in place)"

  private[cli] def requireOutput(
    commandName: String,
    outputOpt: Option[Path],
    backendOpt: Option[XmlBackend],
    stream: Boolean = false
  )(f: (Path, WriterConfig, Boolean) => IO[String]): IO[String] =
    val config = backendOpt.fold(WriterConfig.default)(b => WriterConfig(backend = b))
    outputOpt.fold(
      IO.raiseError[String](new Exception(missingOutputError(commandName)))
    )(path => f(path, config, stream))

  /**
   * Dispatch `run` with effective output path from --output / --in-place flags.
   *
   * The callback receives (writePath, displayPath) and returns the exit code plus rendered output
   * without printing it. The paths differ only for `-i`, where writes go to a temp file that is
   * moved onto the input before the successful output is printed (GH-464).
   *
   * Cases:
   *   - `-o` only: writes directly to the output path
   *   - `-i` only: writes to a sibling temp file then atomically moves onto input. If the command
   *     exits with a non-success code OR throws, the temp is deleted and the original is untouched
   *   - Neither: passes `None` through (for read-only subcommands that don't need output)
   *   - Both: errors with "mutually exclusive"
   */
  private[cli] def runWithOutput(
    outOpt: Option[Path],
    inPlace: Boolean,
    file: Path
  )(execute: (Option[Path], Option[Path]) => IO[(ExitCode, String)]): IO[ExitCode] =
    (outOpt, inPlace) match
      case (Some(_), true) =>
        IO.println(
          Format.errorSimple("--in-place (-i) and --output (-o) are mutually exclusive")
        ).as(ExitCode.Error)
      case (Some(out), false) => execute(Some(out), Some(out)).flatMap(printRunResult)
      case (None, false) => execute(None, None).flatMap(printRunResult)
      case (None, true) =>
        val tmpDir = Option(file.getParent).getOrElse(java.nio.file.Paths.get("."))
        val tmpResource = cats.effect.Resource.make(
          IO.blocking(java.nio.file.Files.createTempFile(tmpDir, ".xl-inplace-", ".xlsx"))
        )(tmp => IO.blocking(java.nio.file.Files.deleteIfExists(tmp)).void)

        tmpResource.use { tmp =>
          execute(Some(tmp), Some(file)).flatMap {
            case result @ (ExitCode.Success, _) =>
              // Atomic move: same directory guarantees same filesystem on POSIX/NTFS
              IO.blocking(
                java.nio.file.Files.move(
                  tmp,
                  file,
                  java.nio.file.StandardCopyOption.REPLACE_EXISTING
                )
              ) *> printRunResult(result)
            case result =>
              // Non-success exit: leave original file alone; Resource cleans up temp
              printRunResult(result)
          }
        }

  /** Require output path or raise user-friendly error, providing path to action */
  private def requireOutputAction(outputOpt: Option[Path], commandName: String)(
    f: Path => IO[String]
  ): IO[String] =
    outputOpt match
      case Some(path) => f(path)
      case None =>
        IO.raiseError(new Exception(missingOutputError(commandName)))

  /** Build WriterConfig from CLI backend option */
  private def buildWriterConfig(backendOpt: Option[XmlBackend]): WriterConfig =
    backendOpt.fold(WriterConfig.default)(b => WriterConfig(backend = b))

  /** Build ReaderConfig from CLI maxSize option (in MB). 0 means unlimited. */
  private def buildReaderConfig(maxSizeOpt: Option[Long]): ReaderConfig =
    maxSizeOpt match
      case Some(0) => ReaderConfig.permissive
      case Some(mb) => ReaderConfig.default.copy(maxUncompressedSize = mb * 1024 * 1024)
      case None => ReaderConfig.default
