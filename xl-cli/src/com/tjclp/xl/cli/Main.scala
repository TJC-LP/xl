package com.tjclp.xl.cli

import java.nio.file.{
  AccessDeniedException,
  AtomicMoveNotSupportedException,
  FileAlreadyExistsException,
  Files,
  NoSuchFileException,
  Path,
  StandardCopyOption
}

import scala.concurrent.duration.*

import cats.effect.{ExitCode, IO, IOApp, Ref, Resource}
import cats.effect.unsafe.IORuntimeConfig
import cats.implicits.*
import cats.syntax.parallel.*
import com.monovore.decline.*

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
  ImportCommands,
  InspectCommands,
  LintCommands,
  ReadCommands,
  SheetCommands,
  StreamingReadCommands,
  StreamingWriteCommands,
  WorkbookCommands,
  WriteCommands
}
import com.tjclp.xl.cli.read.{ReadQuery, Reads, SheetSource}
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
import com.tjclp.xl.cli.contract.{
  CliError,
  CliException,
  CliSignal,
  Diagnostics,
  ErrorCode,
  ExitCodes,
  FunctionDoc,
  Location,
  Outcome,
  OutputMode,
  Payload,
  Render,
  Schema,
  Warning,
  WarningCode
}
import com.tjclp.xl.cli.batch.OpRegistry
import com.tjclp.xl.cli.helpers.{BatchParser, Resolve}
import com.tjclp.xl.cli.output.Format

/** Read version from generated resource, fallback to dev */
private[cli] object BuildInfo:
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
 *   - `-s, --sheet` — Sheet name (ADR-017 §2.5: a qualified ref wins, a single-sheet book
 *     auto-selects, otherwise a sheet verb is `SHEET_REQUIRED`)
 *   - `-o, --output` — Output file for mutations (required for put/putf)
 *   - `--no-recalc` / `--preserve-caches` — write verbs only (GH-468): apply the edit and
 *     recalculate nothing. Non-structural verbs keep every cached formula value in the file; the
 *     structural verbs leave the caches their edit invalidated uncached (see [[WritePolicy]])
 *   - `--strict` — write verbs only (GH-496): exit 1 when the write's recalculation reports formula
 *     errors, non-convergence, or data-table seed warnings
 *   - `--json` — every verb (ADR-017 §2.4): print the result, success or failure, as one JSON
 *     envelope `{ok, exitCode, verb, version, data, warnings, error}` on stdout; `--format json`
 *     payloads ride inside it unchanged
 *
 * Global flags may go anywhere on the command line — [[contract.Argv.hoist]] moves them in front of
 * the verb before decline parses. (`view --eval --strict` is a separate, subcommand-scoped flag of
 * the same name and stays where it is.)
 *
 * This object is the process shell: option and verb definitions plus the handlers that read files
 * and print. The wiring between them is [[Cli.program]], and argv handling (help, version, parse
 * errors) is [[Cli.run]] — both take the [[CliIO]] to print through, so the contract suite can run
 * the identical tree in-process; the binary passes [[CliIO.system]].
 *
 * Channels and exit codes (ADR-017 §2.3): results go to stdout, diagnostics to stderr, and stdout
 * is empty on every failure. Exit 0 ok; 1 completed with findings or a failed gate (`diff` differs,
 * `lint` findings, `--strict`) — never a failure; 2 the command line is wrong; 3 the operation
 * could not complete. The code follows from the [[contract.CliError]]'s `code` through
 * [[contract.ExitCodes.forCode]], so no handler picks a number.
 */
object Main extends IOApp:

  /**
   * The name the staging tests know a run's result by: an [[contract.Outcome]], built here from the
   * three fields `runStagedOutput`'s commit rule reads (exit code, rendered output, completeness)
   * plus the run's warnings. Production code builds outcomes through [[contract.Outcome]]'s own
   * constructors, which carry the verb and the error as well.
   */
  private[cli] type CommandOutcome = Outcome

  private[cli] object CommandOutcome:
    /**
     * Total: the error is derived from the exit code so `ok ⇔ error.isEmpty ⇔ exitCode == 0` holds
     * for these outcomes too — 1 is a gate (`RECALC_GATE`, the output being its report), 2 usage,
     * anything else `INTERNAL`; the message is the output's first line.
     */
    def apply(
      exitCode: ExitCode,
      output: String,
      outputComplete: Boolean,
      warnings: Vector[Warning] = Vector.empty
    ): Outcome =
      val payload = Option.when(output.nonEmpty)(Payload.text(output))
      val error = Option.when(exitCode != ExitCodes.ok) {
        val code =
          if exitCode == ExitCodes.signal then ErrorCode.RECALC_GATE
          else if exitCode == ExitCodes.usage then ErrorCode.USAGE
          else ErrorCode.INTERNAL
        CliError(code, output.linesIterator.nextOption().getOrElse(s"exit ${exitCode.code}"))
      }
      Outcome(
        verb = "",
        payload = payload,
        warnings = warnings,
        error = error,
        exitCode = error.fold(ExitCodes.ok)(_.exitCode),
        outputComplete = outputComplete
      )

  /**
   * GH-519: SIGTERM/SIGINT must terminate the process even mid-recalculation. The evaluator is a
   * single non-yielding compute step, so fiber cancellation is only observed once the whole
   * computation finishes — under the default `shutdownHookTimeout = Duration.Inf` a TERM'd `xl`
   * kept computing at full CPU for the entire remaining recalc and then discarded the result. A
   * finite timeout lets the JVM (and the native image, which installs exit handlers by default on
   * GraalVM 25+) halt promptly after the cancellation attempt; the torn-`-o`-output window this
   * leaves is bounded by staging every `-o`/`-i` mutation beside its destination and replacing the
   * destination only after a complete write.
   *
   * The CPU-starvation checker is disabled outright: xl is a batch compute process, so "your
   * compute pool is busy" is the expected steady state, and on small containers the checker drowned
   * real diagnostics in warnings. Public (not protected) so tests can pin the config.
   */
  override def runtimeConfig: IORuntimeConfig =
    super.runtimeConfig.copy(
      shutdownHookTimeout = 2.seconds,
      cpuStarvationCheckInitialDelay = Duration.Inf
    )

  /**
   * The complete verb tree over the process streams. Tests parse it directly (`Command("xl",
   * "test")(Main.main).parse(args)`); the binary goes through [[run]].
   */
  def main: Opts[IO[ExitCode]] = Cli.program(CliIO.system)

  def run(args: List[String]): IO[ExitCode] = Cli.run(args, CliIO.system)

  // ==========================================================================
  // Global options
  // ==========================================================================

  private[cli] val fileOpt =
    Opts.option[Path]("file", "Excel file to operate on (required)", "f")

  private[cli] val sheetOpt =
    Opts
      .option[String]("sheet", "Sheet to select (required for sheet-level operations)", "s")
      .orNone

  private[cli] val outputOpt =
    Opts.option[Path]("output", "Output file (required)", "o")

  private[cli] val backendOpt: Opts[Option[XmlBackend]] =
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

  private[cli] val maxSizeOpt: Opts[Option[Long]] =
    Opts
      .option[Long](
        "max-size",
        "Max uncompressed size in MB for in-memory load (default: 100, 0 = unlimited). Use for large files when --stream is not supported."
      )
      .orNone

  private[cli] val streamOpt: Opts[Boolean] =
    Opts
      .flag(
        "stream",
        "Use O(1) memory streaming for large files (100k+ rows). Supports: search, stats, bounds, view (markdown/csv/json). 7-8x faster than in-memory."
      )
      .orFalse

  private[cli] val inPlaceOpt: Opts[Boolean] =
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
        "Exit 1 on formula evaluation errors, non-convergence, or data-table seed warnings, including formulas authored by put/putf/fill/copy (default: advisory, exit 0). With -o the output is written; with -i a strict failure leaves the input unchanged."
      )
      .orFalse

  /** Cross-cutting write posture (GH-468/GH-496), parsed before the verb like -f/-o/--stream. */
  private[cli] val writePolicyOpt: Opts[WritePolicy] =
    (noRecalcOpt, strictWriteOpt).mapN(WritePolicy.apply)

  /**
   * Global `--json` (ADR-017 §2.4): the result — success or failure — as one JSON envelope on
   * stdout. Orthogonal to a verb's own `--format`: `view --format json --json` wraps the bare
   * payload as `data`, and a pass-through verb with no `--format` at all uses its JSON payload
   * format under `--json` ([[CliCommand.viewFormat]]), so `data` is structured unless the user
   * asked for a text format.
   */
  private[cli] val jsonOpt: Opts[OutputMode] =
    Opts
      .flag("json", "Wrap every result in the JSON envelope; errors too")
      .orFalse
      .map(json => if json then OutputMode.Json else OutputMode.Text)

  // ==========================================================================
  // Command definitions
  // ==========================================================================

  private val rangeArg = Opts.argument[String]("range")

  /** `view [range]`: absent, the sheet's used range (W2.4). */
  private val viewRangeArg = Opts.argument[String]("range").orNone
  private val offsetOpt = Opts
    .option[Int]("offset", "Rows to skip from the top of the range (default: 0)")
    .withDefault(0)
  private val maxColsOpt = Opts
    .option[Int]("max-cols", "Maximum columns to display, from the left (default: 0 = all)")
    .withDefault(0)
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

  /**
   * `view --format`, as given: no baked default, so the runner can pick JSON under `--json` when
   * the user chose nothing ([[CliCommand.viewFormat]]); text mode defaults to markdown.
   */
  private val formatOpt: Opts[Option[ViewFormat]] = Opts
    .option[String](
      "format",
      "Output format: markdown (default; json under --json), html, svg, json, csv, png, jpeg, webp, pdf"
    )
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
    .orNone
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

  private val viewHelp = """View a range in multiple formats (table, JSON, image, PDF).

USAGE:
  xl -f file.xlsx -s Sheet1 view A1:D10
  xl -f file.xlsx view "Sheet1!A1:D10"    # Qualified ref (no -s needed)
  xl -f file.xlsx -s Sheet1 view          # No range: the sheet's used range

FORMATS:
  markdown (default), json, csv, html, svg, png, jpeg, webp, pdf

OUTPUT FLAGS:
  --format <fmt>      Output format
  --limit <n>         Max rows to display (default: 50; 0 = no limit).
                      When output is clipped, markdown appends a "… showing X of Y rows"
                      trailer; json adds "truncated"/"totalRows" fields (also with --stream);
                      csv/svg note on stderr; html notes on stderr and appends an HTML
                      comment; raster formats append the notice to the "Exported:" line.
  --offset <n>        Rows to skip from the top of the range (default: 0); pages with --limit
  --max-cols <n>      Max columns to display, from the left (default: 0 = all); json adds
                      "totalCols" when clipped
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

  private val diffFormatOpt: Opts[Option[DiffFormat]] =
    Opts
      .option[String]("format", "Output format: markdown (default; json under --json), json")
      .mapValidated { s =>
        s.toLowerCase match
          case "markdown" | "md" => cats.data.Validated.valid(DiffFormat.Markdown)
          case "json" => cats.data.Validated.valid(DiffFormat.Json)
          case other =>
            cats.data.Validated.invalidNel(s"Unknown format: $other. Use markdown or json")
      }
      .orNone

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
  - formulas / defined names referencing external workbook [N] with no
    N-th <externalReference> entry (the cross-workbook sheet-transplant
    class — Excel repairs the file by removing every such formula)
  - defined names Excel refuses: same-scope names colliding under its
    case/width/kana-insensitive comparison, over-long names, whitespace
    or control characters (Excel repairs by removing the named range)
  - xl/calcChain.xml entries naming a cell that holds no formula or a
    sheet id the workbook does not declare (Excel repairs the file on
    open; drop the part with its Override and Relationship, or rebuild it)
  - post-2007 functions (IFS, XLOOKUP, MAXIFS, ...) stored without Excel's
    _xlfn. prefix in <f>, a CF <formula>, a DV <formula1>/<formula2> or a
    <definedName> (the openpyxl class: a silent #NAME? on the first
    recalculation; xl heals a slot only when it regenerates it)

USAGE:
  xl lint report.xlsx
  xl -f report.xlsx lint                     # Equivalent flag form
  xl lint report.xlsx --format json          # Stable machine-readable schema

FINDING CATEGORIES:
  child-order | unresolved-rel-id | wrong-rel-type | missing-part |
  missing-content-type | ref-out-of-bounds | data-table-torn |
  data-table-unseeded | formula-leading-equals | external-ref-dangling |
  defined-name-invalid | calc-chain-stale | xlfn-missing

EXIT CODES:
  0 = no findings (package structure is clean)
  1 = findings reported
  2 = error (unreadable file, missing/malformed core part, ...)

Docs: xl lint is read-only; it never repairs or rewrites the file.

EXAMPLES:
  xl lint deliverable.xlsx && echo "safe to send"
  xl lint deliverable.xlsx --format json | jq '.findings'"""

  private val lintFormatOpt: Opts[Option[LintFormat]] =
    Opts
      .option[String]("format", "Output format: text (default; json under --json), json")
      .mapValidated { s =>
        s.toLowerCase match
          case "text" => cats.data.Validated.valid(LintFormat.Text)
          case "json" => cats.data.Validated.valid(LintFormat.Json)
          case other =>
            cats.data.Validated.invalidNel(s"Unknown format: $other. Use text or json")
      }
      .orNone

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

  /**
   * `xl schema` (ADR-017 §2.13): the CLI contract from the binary itself — the verb table as text,
   * or with `--json` the whole contract (exit and error codes, globals, verbs, the batch JSON
   * Schema, the function registry, the envelope's JSON Schema) as one document.
   */
  val schemaCmd: Opts[Unit] =
    Opts.subcommand(
      "schema",
      "Print the CLI contract: every verb with what it needs and how it exits; " +
        "--json adds exit/error/warning codes, globals, the batch op schema, the functions " +
        "and the envelope schema"
    ) {
      Opts.unit
    }

  /** The `schema` verb wired to its runner; [[Cli.program]] places it after `functions`. */
  private[cli] def schemaOpts(io: CliIO): Opts[IO[ExitCode]] =
    (jsonOpt, schemaCmd).mapN((mode, _) => runSchema(io, mode))

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
        viewRangeArg,
        formulasOpt,
        evalOpt,
        strictOpt,
        limitOpt,
        offsetOpt,
        maxColsOpt,
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
  - --stream scans the used range in O(1) memory (only the matching rows are kept)
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
  private val filterFormatOpt: Opts[Option[FilterFormat]] =
    Opts
      .option[String]("format", "Output format: markdown (default; json under --json), csv, json")
      .mapValidated { s =>
        s.toLowerCase match
          case "markdown" | "md" => cats.data.Validated.valid(FilterFormat.Markdown)
          case "csv" => cats.data.Validated.valid(FilterFormat.Csv)
          case "json" => cats.data.Validated.valid(FilterFormat.Json)
          case other =>
            cats.data.Validated.invalidNel(s"Unknown format: $other. Use markdown, csv, or json")
      }
      .orNone
  private val filterHeaderOpt =
    Opts
      .flag("header", "Treat the first used row as column names (excluded from matching)")
      .orFalse

  val filterCmd: Opts[CliCommand] =
    Opts.subcommand("filter", filterHelp) {
      (whereOpt, filterColumnsOpt, filterLimitOpt, filterFormatOpt, filterHeaderOpt)
        .mapN(CliCommand.Filter.apply)
    }

  // --- Inspect (ADR-017 §2.10): describe, audit, deps ---

  private val describeHelp = """Orient in a workbook: sheets, defined names, date system.

Metadata only by default — instant for any file size, works under --stream — listing every sheet
with its visibility state and dimension, every defined name (hidden ones flagged) and the date
system. --full loads the book and adds per-sheet counts: cells, formulas (and how many are
uncached), merges, comments, hyperlinks, freeze pane, tab color, autoFilter, tables, charts,
pictures, conditional formats, data validations, hidden rows/columns, plus the calcPr settings.

USAGE:
  xl -f model.xlsx describe
  xl -f model.xlsx --stream describe          # same card, O(1) memory
  xl -f model.xlsx describe --full
  xl -f model.xlsx --json describe --full     # {sheets: [...], definedNames: [...], date1904, calcPr}
"""

  private val fullOpt: Opts[Boolean] =
    Opts.flag("full", "Load the workbook and add per-sheet counts and calcPr").orFalse

  val describeCmd: Opts[CliCommand] =
    Opts.subcommand("describe", describeHelp) {
      fullOpt.map(CliCommand.Describe.apply)
    }

  private val auditHelp = """Find every reason a number can be wrong, in one pass.

Findings: cached error values (#DIV/0!, #REF!, ...), uncached formulas, formulas this evaluator
cannot parse, circular references, and readers of names that do not resolve. Notes (reported,
never findings): volatile TODAY/NOW/RAND/RANDBETWEEN cells, dynamic INDIRECT/OFFSET readers,
external-workbook references, and the file's calcPr. Text mode prints one section per non-empty
bucket; -s restricts the cell buckets to one sheet.

Exit 0 whether or not there are findings; --fail-on-findings exits 1 (AUDIT_FINDINGS) on a dirty
book with the report kept, so a CI lane can gate on it.

USAGE:
  xl -f model.xlsx audit
  xl -f model.xlsx -s Summary audit
  xl -f model.xlsx --json audit --fail-on-findings   # data.clean, data.findings, one array per bucket
"""

  private val failOnFindingsOpt: Opts[Boolean] =
    Opts.flag("fail-on-findings", "Exit 1 (AUDIT_FINDINGS) when the audit has findings").orFalse

  val auditCmd: Opts[CliCommand] =
    Opts.subcommand("audit", auditHelp) {
      failOnFindingsOpt.map(CliCommand.Audit.apply)
    }

  private val depsHelp = """Trace one cell's precedents and dependents, hop by hop.

Precedents are the cells the formula reads (single refs exactly, ranges as their occupied cells);
dependents are the formulas that read the cell, by name or through a range that contains it.
Each node carries its depth, formula and value. The ref follows the sheet rule: a qualified ref
('Q1 Data'!B4) names the sheet, else -s, else the only sheet of a single-sheet book.

USAGE:
  xl -f model.xlsx deps Summary!B4                          # both directions, one hop
  xl -f model.xlsx -s Data deps B4 --direction precedents --depth 3
  xl -f model.xlsx --json deps Summary!B4 --direction dependents --depth all
"""

  private val directionOpt: Opts[Direction] =
    Opts
      .option[String]("direction", "precedents, dependents or both (default: both)")
      .withDefault("both")
      .mapValidated {
        case "precedents" => cats.data.Validated.valid(Direction.Precedents)
        case "dependents" => cats.data.Validated.valid(Direction.Dependents)
        case "both" => cats.data.Validated.valid(Direction.Both)
        case other =>
          cats.data.Validated.invalidNel(
            s"Unknown direction: $other. Use precedents, dependents or both"
          )
      }

  // The parser is the only place the sentinels live: 'all' and 0 are Depth.All, an absent flag is
  // one hop, anything else is that many hops.
  private val depthOpt: Opts[Depth] =
    Opts
      .option[String]("depth", "Hops to follow: a number, or 'all' (0 = all; default: 1)")
      .mapValidated {
        case "all" => cats.data.Validated.valid(Depth.All)
        case text =>
          text.toIntOption.filter(_ >= 0) match
            case Some(0) => cats.data.Validated.valid(Depth.All)
            case Some(n) => cats.data.Validated.valid(Depth.Hops(n))
            case None =>
              cats.data.Validated.invalidNel(
                s"Invalid --depth: $text. Use a positive number, 0 or 'all'"
              )
      }
      .withDefault(Depth.Hops(1))

  val depsCmd: Opts[CliCommand] =
    Opts.subcommand("deps", depsHelp) {
      (refArg, directionOpt, depthOpt).mapN(CliCommand.Deps.apply)
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
  private[cli] val batchArg = Opts.argument[String]("operations").withDefault("-")

  /**
   * The `batch` help: the op table comes from [[OpRegistry.helpText]] — every op the registry
   * knows, its example and aliases — so the help can never list fewer ops than the parser accepts.
   */
  private[cli] val batchHelp: String =
    val head = """Apply multiple operations atomically from JSON.

USAGE:
  xl -f in.xlsx -s Sheet1 -o out.xlsx batch ops.json
  echo '[...]' | xl -f in.xlsx -s Sheet1 -o out.xlsx batch -
  xl batch --dry-run ops.json      # validate and summarise without a workbook
  xl batch --schema                # the document's JSON Schema (every op, field and alias)
"""
    val tail = """
STYLE PROPERTIES (style op):
  Font:      bold, italic, underline, fg, fontSize, fontName
  Fill:      bg (background color, e.g., "#FFFF00" or "yellow")
  Align:     align (left/center/right), valign (top/middle/bottom), wrap
  Format:    numFormat (general/number/currency/percent/date/text or an Excel format code)
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
Use --dry-run to validate JSON without writing; --schema prints the JSON Schema."""
    head + "\n" + OpRegistry.helpText + "\n" + tail

  private val dryRunHelp = "Validate batch JSON and show summary without writing"
  private val batchSchemaHelp =
    "Print the JSON Schema of the batch document (every op, field and alias) and exit"

  private val dryRunOpt = Opts.flag("dry-run", dryRunHelp).orFalse
  private val batchSchemaFlag = Opts.flag("schema", batchSchemaHelp).orFalse

  // decline runs every same-named subcommand over the same arguments and fails the parse if any of
  // them rejects one (Parser.Accumulator.OrElse.parseSub), so the full `batch` and the standalone
  // forms below accept the same two flags.
  val batchCmd: Opts[CliCommand] =
    Opts.subcommand("batch", batchHelp) {
      (batchArg, dryRunOpt, batchSchemaFlag).mapN(CliCommand.Batch.apply)
    }

  /**
   * The `batch` forms that need neither `-f` nor `-o`: `--dry-run <source>` validates the document
   * and `--schema` prints its JSON Schema — `OpRegistry.jsonSchema`, the `batchOps` of
   * `xl schema --json` (ADR-017 §2.13).
   */
  enum BatchStandalone derives CanEqual:
    case DryRun(source: String)
    case PrintSchema

  private[cli] val batchStandaloneArgs: Opts[BatchStandalone] =
    (batchArg, Opts.flag("dry-run", dryRunHelp))
      .mapN((source, _) => BatchStandalone.DryRun(source): BatchStandalone)
      .orElse(Opts.flag("schema", batchSchemaHelp).as(BatchStandalone.PrintSchema: BatchStandalone))

  /** The batch document's JSON Schema as a payload: the same JSON `xl schema --json` embeds. */
  private[cli] def batchSchemaPayload: Payload =
    Payload.Json(OpRegistry.jsonSchema(BuildInfo.version))

  private[cli] def batchStandaloneOutcome(
    form: BatchStandalone,
    io: CliIO,
    mode: OutputMode
  ): IO[Outcome] = form match
    case BatchStandalone.DryRun(source) => batchDryRunOutcome(source, io, mode)
    case BatchStandalone.PrintSchema => IO.pure(Outcome.ok("batch", batchSchemaPayload))

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
  xl -f in.xlsx -o out.xlsx recalc --tables       # Also seed data-table interiors
  xl -f in.xlsx -o out.xlsx recalc --parallel 4   # Independent regions on 4 threads"""

  private val recalcTablesOpt: Opts[Boolean] =
    Opts
      .flag("tables", "Also seed data-table interior caches (default: pinned caches)")
      .orFalse

  // GH-520: results are element-for-element identical to the sequential pass; books declaring
  // iterative calculation ignore the flag (cyclic fixpoints are sequential) with a stderr note.
  private val recalcParallelOpt: Opts[Option[Int]] =
    Opts
      .option[Int]("parallel", "Evaluate independent formula regions on N threads (GH-520)")
      .validate("--parallel must be at least 1")(_ >= 1)
      .orNone

  val recalcCmd: Opts[CliCommand] =
    Opts.subcommand("recalc", recalcHelp) {
      (recalcTablesOpt, recalcParallelOpt).mapN(CliCommand.Recalc.apply)
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

  private[cli] val freezeCmd: Opts[CliCommand] =
    Opts.subcommand(
      "freeze",
      "Freeze panes at cell reference (rows above and columns left are locked)"
    ) {
      refArg.map(CliCommand.Freeze.apply)
    }

  private[cli] val unfreezeCmd: Opts[CliCommand] =
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

  private[cli] val sheetViewCmd: Opts[CliCommand] =
    Opts.subcommand(
      "sheet-view",
      "Set sheet view options: gridlines, zoom, tab selection (requires -o)"
    ) {
      val viewGridlines = onOffOpt("gridlines", "Show cell gridlines: on or off")
      val viewZoom = Opts.option[Int]("zoom", "Zoom percentage (10-400)").orNone
      val viewTabSelected = onOffOpt("tab-selected", "Select this sheet's tab: on or off")
      (viewGridlines, viewZoom, viewTabSelected).mapN(CliCommand.SheetViewOp.apply)
    }

  private[cli] val tabColorCmd: Opts[CliCommand] =
    Opts.subcommand(
      "tab-color",
      "Set the sheet tab color: named, #hex, rgb(r,g,b), or theme:accent1[:tint] (requires -o). " +
        "--clear removes a modeled color (a color preserved from the source XML is not stripped)."
    ) {
      val colorArg = Opts.argument[String]("color").orNone
      val clearFlag = Opts.flag("clear", "Clear the modeled tab color").orFalse
      (colorArg, clearFlag).mapN(CliCommand.TabColorOp.apply)
    }

  private[cli] val autoFilterCmd: Opts[CliCommand] =
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

  private[cli] val pageSetupCmd: Opts[CliCommand] =
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

  private[cli] val headerFooterCmd: Opts[CliCommand] =
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

  private[cli] val copyCmd: Opts[CliCommand] =
    Opts.subcommand("copy", "Copy range to another location (with formula adjustment)") {
      val copySrcArg = Opts.argument[String]("source")
      val copyTgtArg = Opts.argument[String]("target")
      val valuesOnlyOpt =
        Opts.flag("values-only", "Copy values only (no formula adjustment)").orFalse
      (copySrcArg, copyTgtArg, valuesOnlyOpt).mapN(CliCommand.Copy.apply)
    }

  // --- Conditional formatting command (GH-324) ---

  private[cli] val cfCmd: Opts[CliCommand] =
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

  private[cli] val chartCmd: Opts[CliCommand] =
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

  private[cli] val addImageCmd: Opts[CliCommand] =
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

  private[cli] val insertRowsCmd: Opts[CliCommand] =
    Opts.subcommand("insert-rows", "Insert rows (shifts cells & rewrites formulas)") {
      val atArg = Opts.argument[Int]("at-row")
      val countArg = Opts.argument[Int]("count").withDefault(1)
      (atArg, countArg).mapN(CliCommand.InsertRows.apply)
    }

  private[cli] val deleteRowsCmd: Opts[CliCommand] =
    Opts.subcommand(
      "delete-rows",
      "Delete rows (shifts cells & rewrites formulas; #REF! on loss)"
    ) {
      val atArg = Opts.argument[Int]("at-row")
      val countArg = Opts.argument[Int]("count").withDefault(1)
      (atArg, countArg).mapN(CliCommand.DeleteRows.apply)
    }

  private[cli] val insertColsCmd: Opts[CliCommand] =
    Opts.subcommand("insert-cols", "Insert columns (shifts cells & rewrites formulas)") {
      val colArg = Opts.argument[String]("at-col")
      val countArg = Opts.argument[Int]("count").withDefault(1)
      (colArg, countArg).mapN(CliCommand.InsertColumns.apply)
    }

  private[cli] val deleteColsCmd: Opts[CliCommand] =
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

  private[cli] def run(
    filePath: Path,
    sheetNameOpt: Option[String],
    outputOpt: Option[Path],
    displayOpt: Option[Path],
    backendOpt: Option[XmlBackend],
    maxSizeOpt: Option[Long],
    stream: Boolean,
    cmd: CliCommand,
    io: CliIO,
    mode: OutputMode = OutputMode.Text
  ): IO[ExitCode] =
    runResult(
      filePath,
      sheetNameOpt,
      outputOpt,
      displayOpt,
      backendOpt,
      maxSizeOpt,
      stream,
      cmd,
      io = io,
      mode = mode
    ).flatMap(emit(_, mode, io))

  /**
   * Execute a command into an [[contract.Outcome]] without printing, so in-place writes can commit
   * first; [[emit]] renders it in the run's [[contract.OutputMode]].
   *
   * GH-496: a [[StrictFailure]] is not a crash — the write completed and its summary is the
   * payload, verbatim (counts, failing refs, convergence verdict, plus the strict reason); the
   * outcome is a `RECALC_GATE` signal (exit 1) whose one-line reason is the summary's
   * `STRICT FAILURE` line. With `-i` that non-success code means the temp file is discarded and the
   * input is left untouched, which is the atomic reading of "this book did not pass the gate" — so
   * the summary's `Saved:` line is rewritten to say exactly that. With `-o` the completed temp is
   * still committed and `Saved:` stands.
   *
   * Any other failure yields an outcome with NO payload and the exit code its [[contract.CliError]]
   * selects: 2 for a wrong command line, 3 for an operation that could not complete. An un-migrated
   * `new Exception(msg)` classifies as `INTERNAL` (exit 3) with the same message. Reader and view
   * warnings collected during the run ride on the outcome.
   */
  private[cli] def runResult(
    filePath: Path,
    sheetNameOpt: Option[String],
    outputOpt: Option[Path],
    displayOpt: Option[Path],
    backendOpt: Option[XmlBackend],
    maxSizeOpt: Option[Long],
    stream: Boolean,
    cmd: CliCommand,
    policy: WritePolicy = WritePolicy.default,
    strictFailureDiscardsOutput: Boolean = false,
    io: CliIO,
    mode: OutputMode = OutputMode.Text
  ): IO[Outcome] =
    Ref.of[IO, Vector[Warning]](Vector.empty).flatMap { warnings =>
      execute(
        filePath,
        sheetNameOpt,
        outputOpt,
        backendOpt,
        maxSizeOpt,
        stream,
        cmd,
        policy,
        io,
        warnings,
        mode
      ).attempt.flatMap { attempt =>
        warnings.get.map { collected =>
          attempt match
            case Right(payload) =>
              Outcome.ok(cmd.verb, relocate(payload, outputOpt, displayOpt), collected)
            case Left(strict: StrictFailure) =>
              val rendered = renderWithTarget(strict.summary, outputOpt, displayOpt)
              val summary =
                if strictFailureDiscardsOutput then unsayTheSave(rendered, outputOpt, displayOpt)
                else rendered
              Outcome.signal(
                cmd.verb,
                Payload.text(summary),
                CliError(ErrorCode.RECALC_GATE, strictReason(summary)),
                collected
              )
            // ADR-017 §2.10: a read verb that completed with findings (`audit --fail-on-findings`)
            case Left(signal: CliSignal) =>
              Outcome.signal(
                cmd.verb,
                relocate(signal.payload, outputOpt, displayOpt),
                signal.error,
                collected
              )
            case Left(err) =>
              // GH-483: the temp→target rewrite applies to failure messages too
              val classified = CliError.fromThrowable(err)
              val relocated = classified.copy(
                message = renderWithTarget(classified.message, outputOpt, displayOpt)
              )
              Outcome.failed(cmd.verb, relocated, collected)
        }
      }
    }

  /**
   * The gate's one-line reason: the summary's `STRICT FAILURE (--strict): …` line, else its first.
   */
  private def strictReason(summary: String): String =
    val lines = summary.linesIterator.toVector
    lines.find(_.startsWith("STRICT FAILURE")).orElse(lines.headOption).getOrElse(summary)

  /**
   * The GH-464 temp→target rewrite on a prose payload; JSON payloads never name the staging file.
   */
  private def relocate(
    payload: Payload,
    outputOpt: Option[Path],
    displayOpt: Option[Path]
  ): Payload =
    payload match
      case Payload.Text(text, saved, written) =>
        Payload.Text(renderWithTarget(text, outputOpt, displayOpt), saved, written)
      case json: Payload.Json => json
      case raw: Payload.Raw => raw

  /**
   * GH-496: an in-place run whose exit code is non-success never commits its temp file, so the
   * summary a write command already built ("...\nSaved: <path>") describes a save that did not
   * happen. Replace that line with what is true. Only `-i` is affected: with `-o` the completed
   * output is committed even though the gate selects exit code 1, so its `Saved:` line stays.
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

  /**
   * Print an outcome in the run's mode ([[contract.Render]]): the result on stdout (nothing at all
   * when there is none), the diagnostics and warnings — or the one-line `Error:` under `--json` —
   * on stderr, then the exit code.
   */
  private[cli] def emit(outcome: Outcome, mode: OutputMode, io: CliIO): IO[ExitCode] =
    val rendered = Render(mode)(outcome, BuildInfo.version)
    val out = if rendered.stdout.isEmpty then IO.unit else io.out(rendered.stdout)
    val err = if rendered.stderr.isEmpty then IO.unit else io.err(rendered.stderr)
    (out *> err).as(outcome.exitCode)

  /** An `ExcelIO` whose reader warnings land in `warnings` as `READER_WARNING`s. */
  private def readerCollecting(warnings: Ref[IO, Vector[Warning]]): ExcelIO[IO] =
    ExcelIO.withWarnings[IO] { warning =>
      warnings.update(_ :+ Warning(WarningCode.READER_WARNING, warning.toString))
    }

  /**
   * THE sheet rule's steps 2–3 for the run's default sheet, before dispatch ([[Resolve.default]],
   * ADR-017 §2.5): `-s` by name; else, for a verb that takes a sheet, the only sheet of a
   * single-sheet book — announced through [[announceAutoSelect]]. AllSheets and no-sheet verbs get
   * `None` and behave as they always have.
   */
  private def defaultSheet(
    wb: Workbook,
    sheetNameOpt: Option[String],
    cmd: CliCommand,
    mode: OutputMode,
    warn: Warning => IO[Unit]
  ): IO[Option[Sheet]] =
    IO.fromEither(Resolve.default(wb, sheetNameOpt, cmd.takesSheet).left.map(CliException(_)))
      .flatTap(sheet => announceAutoSelect(sheet.map(_.name), sheetNameOpt, cmd, mode, warn))

  /**
   * Step 3 as the run reports it: under `--json` only, when no `-s` was given and the verb takes a
   * sheet, the only sheet's name rides in `warnings[]` as `SHEET_AUTOSELECTED`. Text mode prints
   * nothing extra — its stdout and stderr stay byte-identical to before the rule.
   */
  private def announceAutoSelect(
    only: Option[SheetName],
    sheetNameOpt: Option[String],
    cmd: CliCommand,
    mode: OutputMode,
    warn: Warning => IO[Unit]
  ): IO[Unit] =
    only match
      case Some(name) if mode == OutputMode.Json && sheetNameOpt.isEmpty && cmd.usesDefaultSheet =>
        warn(Resolve.autoSelected(name))
      case _ => IO.unit

  /**
   * The streaming twins resolve their sheet from `workbook.xml` themselves ([[Resolve.sheetName]]);
   * the run announces step 3 from the same metadata, read only when the announcement can apply and
   * never failing the run — an unreadable file is the verb's own `IO_READ`.
   */
  private def streamAutoSelect(
    excel: ExcelIO[IO],
    filePath: Path,
    sheetNameOpt: Option[String],
    cmd: CliCommand,
    mode: OutputMode,
    warn: Warning => IO[Unit]
  ): IO[Unit] =
    if mode == OutputMode.Json && sheetNameOpt.isEmpty && cmd.usesDefaultSheet then
      excel.readMetadata(filePath).attempt.flatMap {
        case Right(meta) => announceAutoSelect(Resolve.only(meta), sheetNameOpt, cmd, mode, warn)
        case Left(_) => IO.unit
      }
    else IO.unit

  /**
   * Classify a failure while reading the input at `path` as `IO_READ` (exit 3), keeping the message
   * verbatim; a `CliException` already raised below passes through unchanged. Every read of the
   * input — full workbook, metadata quick path, lint's raw zip — goes through this so one condition
   * (missing or unreadable file) has one code.
   */
  private def classifyRead[A](path: Path)(read: IO[A]): IO[A] =
    read.adaptError {
      case cli: CliException => cli
      case other =>
        CliException(
          CliError(
            ErrorCode.IO_READ,
            CliError.messageOf(other),
            location = Some(Location.file(path.toString))
          )
        )
    }

  /** Read the input through `excel` under [[classifyRead]]. */
  private def readWorkbook(excel: ExcelIO[IO], path: Path, config: ReaderConfig): IO[Workbook] =
    classifyRead(path)(excel.readWith(path, config))

  /**
   * Classify a failure while producing the output at `target` — allocating its staging file,
   * committing it, or `new`'s direct write — as `IO_WRITE` (exit 3) naming the target: `cannot
   * write <target>: <reason>`, with the hint every unwritable destination needs. A `CliException`
   * already raised below passes through. The reason is stable text, never a staging path: a missing
   * directory (`NoSuchFileException` from the temp allocation, `FileNotFoundException` from a
   * stream) is named as such, permission failures likewise, anything else by its message.
   */
  private def classifyWrite[A](target: Path)(write: IO[A]): IO[A] =
    write.adaptError {
      case cli: CliException => cli
      case other =>
        CliException(
          CliError(
            ErrorCode.IO_WRITE,
            s"cannot write $target: ${writeReason(target, other)}",
            hint = Some("check that the directory exists and is writable"),
            location = Some(Location.file(target.toString))
          )
        )
    }

  private def writeReason(target: Path, failure: Throwable): String =
    val directory = Option(target.toAbsolutePath.getParent).fold(".")(_.toString)
    failure match
      case _: NoSuchFileException => s"no such directory: $directory"
      case _: AccessDeniedException => "permission denied"
      case other =>
        strerror(CliError.messageOf(other)) match
          case "No such file or directory" => s"no such directory: $directory"
          case "Permission denied" => "permission denied"
          case reason => reason

  /**
   * The OS reason inside a write failure's message. The writer and `ExcelIO` each wrap the
   * `FileOutputStream` failure once — `Failed to write XLSX: IO error: Failed to write XLSX: <path>
   * (No such file or directory)` — so peel those prefixes and keep the parenthesised reason when
   * the text ends in one (`java.io` spells every stream failure `<path> (<reason>)`).
   */
  private def strerror(message: String): String =
    val prefixes = List("Failed to write XLSX: ", "IO error: ")
    @scala.annotation.tailrec
    def peel(text: String): String = prefixes.find(text.startsWith) match
      case Some(prefix) => peel(text.stripPrefix(prefix))
      case None => text
    val inner = peel(message)
    val open = inner.lastIndexOf(" (")
    if open >= 0 && inner.endsWith(")") then inner.substring(open + 2, inner.length - 1)
    else inner

  private[cli] def runInfo(io: CliIO, mode: OutputMode = OutputMode.Text): IO[ExitCode] =
    val payload = mode match
      case OutputMode.Text => Payload.text(formatFunctionList())
      // Typed rows (ADR-017 §2.13): every registry function plus LET, the special form
      case OutputMode.Json => Payload.Json(FunctionDoc.toJson(FunctionDoc.all))
    emit(Outcome.ok("functions", payload), mode, io)

  /** `xl schema`: the verb table as text; `--json`: the whole contract ([[Schema.json]]). */
  private[cli] def runSchema(io: CliIO, mode: OutputMode): IO[ExitCode] =
    val payload = mode match
      case OutputMode.Text => Payload.text(Schema.verbTable(BuildInfo.version))
      case OutputMode.Json => Payload.Json(Schema.json(BuildInfo.version))
    emit(Outcome.ok("schema", payload), mode, io)

  private[cli] def runRasterizers(io: CliIO, mode: OutputMode = OutputMode.Text): IO[ExitCode] =
    probeRasterizers().flatMap { rows =>
      val anyAvailable = rows.exists(_.status == "available")
      val payload = mode match
        case OutputMode.Text =>
          Payload.text(renderRasterizerText(rows, anyAvailable, NativeImage.inNativeImage))
        case OutputMode.Json =>
          Payload.Json(
            ujson.Obj(
              "backends" -> ujson.Arr.from(rows.map { row =>
                ujson.Obj(
                  "name" -> ujson.Str(row.backend),
                  "status" -> ujson.Str(row.status),
                  "note" -> ujson.Str(row.note)
                )
              }),
              "anyAvailable" -> ujson.Bool(anyAvailable)
            )
          )
      emit(Outcome.ok("rasterizers", payload), mode, io)
    }

  /**
   * One line of the `xl rasterizers` table. Status terminology:
   *   - available: Works correctly
   *   - missing: Binary not in PATH
   *   - broken: Found but non-functional (e.g., delegate missing)
   *   - unavailable: Cannot be used in current environment (e.g., Batik on native-image)
   */
  private[cli] final case class RasterizerRow(backend: String, status: String, note: String)
      derives CanEqual

  /** Check all rasterizers in parallel and describe each. */
  private def probeRasterizers(): IO[Vector[RasterizerRow]] =
    (
      BatikRasterizer.isAvailable,
      CairoSvg.isAvailable,
      RsvgConvert.isAvailable,
      Resvg.isAvailable,
      ImageMagick.isAvailable,
      ImageMagick.diagnostics
    ).parMapN {
      (batikAvail, cairoAvail, rsvgAvail, resvgAvail, imageMagickAvail, imageMagickDiag) =>
        rasterizerRows(
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
    val rows = rasterizerRows(
      batikAvail,
      cairoAvail,
      rsvgAvail,
      resvgAvail,
      imageMagickAvail,
      imageMagickDiag,
      nativeImage
    )
    val anyAvailable = batikAvail || cairoAvail || rsvgAvail || resvgAvail || imageMagickAvail
    (renderRasterizerText(rows, anyAvailable, nativeImage), anyAvailable)

  /** The rows of the table, in display order. */
  private def rasterizerRows(
    batikAvail: Boolean,
    cairoAvail: Boolean,
    rsvgAvail: Boolean,
    resvgAvail: Boolean,
    imageMagickAvail: Boolean,
    imageMagickDiag: String,
    nativeImage: Boolean
  ): Vector[RasterizerRow] =
    // Batik - the default backend; "unavailable" when AWT not present (native image)
    val batikStatus = if batikAvail then "available" else "unavailable"
    val batikNote =
      if batikAvail then "Built-in default (requires AWT)"
      else if nativeImage then "Native binary: no AWT, by design"
      else "Requires AWT (not present here)"

    // CairoSvg
    val cairoStatus = if cairoAvail then "available" else "missing"
    val cairoNote = if cairoAvail then "pip install cairosvg" else "Not in PATH"

    // rsvg-convert
    val rsvgStatus = if rsvgAvail then "available" else "missing"
    val rsvgNote = if rsvgAvail then "librsvg2-bin" else "Not in PATH"

    // resvg
    val resvgStatus = if resvgAvail then "available" else "missing"
    val resvgNote = if resvgAvail then "cargo install resvg" else "Not in PATH"

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

    Vector(
      RasterizerRow("batik", batikStatus, batikNote),
      RasterizerRow("cairosvg", cairoStatus, cairoNote),
      RasterizerRow("rsvg-convert", rsvgStatus, rsvgNote),
      RasterizerRow("resvg", resvgStatus, resvgNote),
      RasterizerRow("imagemagick", imStatus, imNote)
    )

  /** The table as printed, byte for byte what `xl rasterizers` has always shown. */
  private def renderRasterizerText(
    rows: Vector[RasterizerRow],
    anyAvailable: Boolean,
    nativeImage: Boolean
  ): String =
    val sb = new StringBuilder
    sb.append("SVG Rasterizer Status\n")
    sb.append("=" * 60 + "\n\n")
    sb.append(f"${"Backend"}%-14s | ${"Status"}%-11s | ${"Notes"}\n")
    sb.append("-" * 60 + "\n")
    rows.foreach { row =>
      sb.append(f"${row.backend}%-14s | ${row.status}%-11s | ${row.note}\n")
    }
    sb.append("\n")

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

    sb.toString

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

  private[cli] def runHeadless(
    filePathOpt: Option[Path],
    sheetNameOpt: Option[String],
    maxSizeOpt: Option[Long],
    cmd: CliCommand,
    io: CliIO,
    mode: OutputMode = OutputMode.Text
  ): IO[ExitCode] =
    val readerConfig = buildReaderConfig(maxSizeOpt)
    Ref.of[IO, Vector[Warning]](Vector.empty).flatMap { warnings =>
      val excel = readerCollecting(warnings)
      val workbookIO: IO[Workbook] = filePathOpt match
        case Some(filePath) => readWorkbook(excel, filePath, readerConfig)
        case None => IO.pure(Workbook(Vector.empty)) // Truly empty workbook for constant formulas

      (for
        wb <- workbookIO
        sheet <- defaultSheet(wb, sheetNameOpt, cmd, mode, w => warnings.update(_ :+ w))
        payload <- (cmd, mode) match
          case (CliCommand.Eval(formulaStr, overrides), OutputMode.Text) =>
            ReadCommands.eval(wb, sheet, formulaStr, overrides).map(Payload.text)
          case (CliCommand.Eval(formulaStr, overrides), OutputMode.Json) =>
            ReadCommands.evalData(wb, sheet, formulaStr, overrides).map(Payload.Raw(_))
          case (CliCommand.EvalArray(formulaStr, targetRef, overrides), OutputMode.Text) =>
            ReadCommands.evalArray(wb, sheet, formulaStr, targetRef, overrides).map(Payload.text)
          case (CliCommand.EvalArray(formulaStr, targetRef, overrides), OutputMode.Json) =>
            ReadCommands
              .evalArrayData(wb, sheet, formulaStr, targetRef, overrides)
              .map(Payload.Raw(_))
          case (other, _) =>
            IO.raiseError(
              CliException(CliError(ErrorCode.INTERNAL, s"Unexpected headless command: $other"))
            )
      yield payload).attempt.flatMap { attempt =>
        warnings.get.flatMap { collected =>
          val outcome = attempt match
            case Right(payload) => Outcome.ok(cmd.verb, payload, collected)
            case Left(err) => Outcome.failed(cmd.verb, CliError.fromThrowable(err), collected)
          emit(outcome, mode, io)
        }
      }
    }

  /**
   * Run the diff command with its exit codes: 0 = identical, 1 = differences found (a result, not a
   * failure — a `DIFFERENCES_FOUND` signal whose report is the payload), 3 = error (unreadable
   * file, sheet filter matching neither workbook, ...) reported on stderr.
   */
  private[cli] def runDiff(
    fileA: Path,
    fileB: Path,
    sheetFilter: Option[String],
    maxSizeOpt: Option[Long],
    format: DiffFormat,
    io: CliIO = CliIO.system,
    mode: OutputMode = OutputMode.Text
  ): IO[ExitCode] =
    val excel = ExcelIO.instance[IO]
    val readerConfig = buildReaderConfig(maxSizeOpt)
    (for
      wbA <- readWorkbook(excel, fileA, readerConfig)
      wbB <- readWorkbook(excel, fileB, readerConfig)
      diff <- DiffCommands.computeDiff(wbA, wbB, sheetFilter) match
        case Right(d) => IO.pure(d)
        // The only refusal: a -s filter naming a sheet neither workbook has — SHEET_NOT_FOUND
        // with the nearest names from both books, keeping the diff's own message
        case Left(err) =>
          val names = (wbA.sheets ++ wbB.sheets).map(_.name.value).distinct
          IO.raiseError(
            CliException(
              sheetFilter.fold(CliError(ErrorCode.INTERNAL, err))(filter =>
                Resolve.sheetNotFound(names, filter).copy(message = err)
              )
            )
          )
      output = format match
        case DiffFormat.Markdown =>
          DiffCommands.renderMarkdown(diff, fileA.toString, fileB.toString)
        case DiffFormat.Json => DiffCommands.renderJson(diff)
    yield (output, diff.identical)).attempt.flatMap {
      case Right((output, identical)) =>
        // The JSON report rides as text (Payload.Raw): its numbers are never re-parsed
        val payload = (format, mode) match
          case (DiffFormat.Json, OutputMode.Json) => Payload.Raw(output)
          case _ => Payload.text(output)
        val outcome =
          if identical then Outcome.ok("diff", payload)
          else
            Outcome.signal(
              "diff",
              payload,
              CliError(ErrorCode.DIFFERENCES_FOUND, s"Differences found: $fileA vs $fileB")
            )
        emit(outcome, mode, io)
      case Left(err) =>
        emit(Outcome.failed("diff", CliError.fromThrowable(err)), mode, io)
    }

  /**
   * Resolve lint's input file from the `-f` flag and the positional form — exactly one must be
   * given (GH-422). A `Left` is a usage error (exit 2, stderr) at the call site.
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
   * Run the lint command with its exit codes: 0 = clean, 1 = findings (a result, not a failure), 3 =
   * error reported on stderr — an unreadable input (missing file, not a zip, missing/malformed core
   * part) as `IO_READ`, the same code every other verb gives that condition. Opens the zip directly
   * — NOT ExcelIO.read — because a full parse would repair/normalize the very structure lint
   * inspects (GH-397).
   */
  private[cli] def runLint(
    file: Path,
    format: LintFormat,
    io: CliIO = CliIO.system,
    mode: OutputMode = OutputMode.Text
  ): IO[ExitCode] =
    IO.blocking(WorkbookLint.lint(file)).flatMap {
      case Right(findings) =>
        val output = format match
          case LintFormat.Text => LintCommands.renderText(file.toString, findings)
          case LintFormat.Json => LintCommands.renderJson(file.toString, findings)
        val payload = (format, mode) match
          case (LintFormat.Json, OutputMode.Json) => Payload.Raw(output)
          case _ => Payload.text(output)
        val outcome =
          if findings.isEmpty then Outcome.ok("lint", payload)
          else
            Outcome.signal(
              "lint",
              payload,
              CliError(ErrorCode.LINT_FINDINGS, s"$file: ${findings.size} finding(s)")
            )
        emit(outcome, mode, io)
      case Left(err) =>
        val at = Some(Location.file(file.toString))
        val error = err match
          case XLError.IOError(_) | XLError.ParseError(_, _) =>
            CliError(ErrorCode.IO_READ, err.message, location = at, cause = Some(err))
          case other => CliError.fromXLError(other, at)
        emit(Outcome.failed("lint", error), mode, io)
    }

  private[cli] def runStandalone(
    outPath: Path,
    sheetName: String,
    sheets: List[String],
    backendOpt: Option[XmlBackend],
    io: CliIO,
    mode: OutputMode = OutputMode.Text
  ): IO[ExitCode] =
    val config = backendOpt.fold(WriterConfig.default)(b => WriterConfig(backend = b))
    (for
      // --sheet takes precedence over --sheet-name; if neither, default to "Sheet1"
      // A name the validator refuses is INVALID_SHEET_NAME with its own text
      names <- (if sheets.isEmpty then List(sheetName) else sheets).traverse(n =>
        IO.fromEither(Resolve.validSheetName(n).left.map(CliException(_)))
      )
      wb = Workbook(names.map(Sheet(_)).toVector)
      _ <- classifyWrite(outPath)(ExcelIO.instance[IO].writeWith(wb, outPath, config))
    yield
      val sheetList = names.map(_.value).mkString(", ")
      s"Created ${outPath.toAbsolutePath} with ${names.size} sheet(s): $sheetList"
    ).attempt.flatMap {
      case Right(output) =>
        val payload = Payload.Text(output, Some(outPath.toAbsolutePath.toString), written = true)
        emit(Outcome.ok("new", payload), mode, io)
      case Left(err) =>
        emit(Outcome.failed("new", CliError.fromThrowable(err)), mode, io)
    }

  /** A `--stream` refusal: usage (exit 2), naming the in-memory alternative as the hint. */
  private def unsupportedInStream(message: String, alternative: String): CliException =
    CliException(CliError(ErrorCode.UNSUPPORTED_IN_STREAM, message, hint = Some(alternative)))

  /**
   * Run one command to its [[contract.Payload]]. Under `--json` the typed verbs (`sheets`, `names`,
   * `bounds`, `batch --dry-run`) build their data directly and `view`/`filter` with their own
   * `--format json` pass it through ([[bridge]]); everything else is the prose the text mode
   * prints.
   */
  private def execute(
    filePath: Path,
    sheetNameOpt: Option[String],
    outputOpt: Option[Path],
    backendOpt: Option[XmlBackend],
    maxSizeOpt: Option[Long],
    stream: Boolean,
    cmd: CliCommand,
    policy: WritePolicy,
    io: CliIO,
    warnings: Ref[IO, Vector[Warning]],
    mode: OutputMode
  ): IO[Payload] =
    val excel = readerCollecting(warnings)
    val readerConfig = buildReaderConfig(maxSizeOpt)
    val warn: Warning => IO[Unit] = warning => warnings.update(_ :+ warning)
    // Handle metadata-only commands (instant for any file size)
    cmd match
      // Sheets list: quick mode (metadata only, instant) by default, --stats loads the workbook
      case CliCommand.Sheets(SheetsAction.List(stats)) =>
        (stats, mode) match
          case (true, OutputMode.Text) =>
            readWorkbook(excel, filePath, readerConfig)
              .flatMap(WorkbookCommands.sheets)
              .map(Payload.text)
          case (true, OutputMode.Json) =>
            readWorkbook(excel, filePath, readerConfig)
              .map(wb => Payload.Json(WorkbookCommands.sheetsData(wb)))
          case (false, OutputMode.Text) =>
            classifyRead(filePath)(WorkbookCommands.sheetsQuick(filePath)).map(Payload.text)
          case (false, OutputMode.Json) =>
            classifyRead(filePath)(WorkbookCommands.sheetsQuickData(filePath)).map(Payload.Json(_))
      case CliCommand.Sheets(SheetsAction.Hide(name, veryHide)) =>
        // Hide requires loading workbook and writing output
        requireOutputAction(outputOpt, "sheets hide") { outputPath =>
          val config = buildWriterConfig(backendOpt)
          readWorkbook(excel, filePath, readerConfig).flatMap { wb =>
            SheetCommands.hideSheet(wb, name, veryHide, outputPath, config, stream)
          }
        }.map(Payload.text)
      case CliCommand.Sheets(SheetsAction.Show(name)) =>
        // Show requires loading workbook and writing output
        requireOutputAction(outputOpt, "sheets show") { outputPath =>
          val config = buildWriterConfig(backendOpt)
          readWorkbook(excel, filePath, readerConfig).flatMap { wb =>
            SheetCommands.showSheet(wb, name, outputPath, config, stream)
          }
        }.map(Payload.text)

      // Names: always lightweight (defined names don't require cell data)
      case CliCommand.Names =>
        mode match
          case OutputMode.Text =>
            classifyRead(filePath)(WorkbookCommands.namesLight(filePath)).map(Payload.text)
          case OutputMode.Json =>
            classifyRead(filePath)(WorkbookCommands.namesData(filePath)).map(Payload.Json(_))

      // Bounds: dimension-first by default (--scan for full scan). The metadata read is the input
      // read, so a missing or unreadable file is IO_READ here as on every other verb.
      case CliCommand.Bounds(scan) =>
        classifyRead(filePath)(excel.readMetadata(filePath)).flatMap { meta =>
          announceAutoSelect(Resolve.only(meta), sheetNameOpt, cmd, mode, warn) *> {
            mode match
              case OutputMode.Text =>
                val text =
                  if scan then StreamingReadCommands.boundsScan(filePath, sheetNameOpt)
                  else StreamingReadCommands.boundsFromMetadata(meta, filePath, sheetNameOpt)
                text.map(Payload.text)
              case OutputMode.Json =>
                StreamingReadCommands
                  .boundsData(meta, filePath, sheetNameOpt, scan)
                  .map(Payload.Json(_))
          }
        }

      // A dry run validates the batch JSON and reads no workbook, whatever else is on the line
      case CliCommand.Batch(source, true, _) =>
        batchDryRunPayload(source, io, mode, warn)

      // --schema prints the document's JSON Schema; no workbook is read, nothing is written
      case CliCommand.Batch(_, _, true) =>
        IO.pure(batchSchemaPayload)

      // Describe (ADR-017 §2.10): metadata only — instant, streaming-safe — unless --full asks for
      // the loaded book's counts. A workbook verb: -s selects nothing, but an explicit -s must
      // still name a real sheet (SHEET_NOT_FOUND with candidates otherwise). The other workbook
      // verbs (sheets, names) are parsed without --sheet at all (Cli.scala), so only describe
      // can see the flag.
      case CliCommand.Describe(full) =>
        if !full then
          classifyRead(filePath)(excel.readMetadata(filePath)).flatMap { meta =>
            InspectCommands
              .checkSheetFlag(sheetNameOpt, "describe")(IO.pure(meta))
              .as(InspectCommands.describeLight(meta, mode))
          }
        else if stream then
          IO.raiseError(
            unsupportedInStream(
              "describe --full is not supported with --stream (the counts need the whole workbook)",
              "omit --full for the metadata-only card, or omit --stream; use --max-size <MB> for a large file"
            )
          )
        else
          InspectCommands.checkSheetFlag(sheetNameOpt, "describe")(
            classifyRead(filePath)(excel.readMetadata(filePath))
          ) *> readWorkbook(excel, filePath, readerConfig)
            .map(wb => InspectCommands.describe(wb, mode))

      // Audit and deps (ADR-017 §2.10) analyze the loaded workbook: never under --stream
      case CliCommand.Audit(failOnFindings) =>
        if stream then
          IO.raiseError(
            unsupportedInStream(
              "audit is not supported with --stream (the analysis needs the whole workbook)",
              "omit --stream; use --max-size <MB> to load a large file in memory"
            )
          )
        else
          for
            wb <- readWorkbook(excel, filePath, readerConfig)
            sheet <- defaultSheet(wb, sheetNameOpt, cmd, mode, warn)
            payload <- InspectCommands.audit(wb, sheet, failOnFindings, mode)
          yield payload

      case CliCommand.Deps(refStr, direction, depth) =>
        if stream then
          IO.raiseError(
            unsupportedInStream(
              "deps is not supported with --stream (the graph needs the whole workbook)",
              "omit --stream; use --max-size <MB> to load a large file in memory"
            )
          )
        else
          for
            wb <- readWorkbook(excel, filePath, readerConfig)
            sheet <- defaultSheet(wb, sheetNameOpt, cmd, mode, warn)
            payload <- InspectCommands.deps(wb, sheet, refStr, direction, depth, mode)
          yield payload

      // Other commands: regular execution path
      case _ =>
        // For write commands: stream flag uses the SAX/StAX workbook writer
        // For read commands: stream flag enables O(1) input memory (true streaming)
        // W2.4: the record-based reads are one function of a query and a source; the source
        // is the strategy (streaming reader or loaded workbook)
        val readQuery = ReadQuery.of(cmd, mode)

        // Check for streaming write commands (true O(1) memory transform)
        val isStreamingWriteCmd = cmd match
          case _: CliCommand.Style => true
          case _: CliCommand.Put => true
          case _: CliCommand.PutFormula => true
          case _: CliCommand.Batch => true
          case _ => false

        (stream, readQuery) match
          case (true, Some(query)) =>
            streamAutoSelect(excel, filePath, sheetNameOpt, cmd, mode, warn) *>
              Reads.run(query, SheetSource.streaming(filePath, excel), sheetNameOpt, mode, warn)
          case _ if stream && isStreamingWriteCmd =>
            // GH-496: a streaming write never recalculates, so --strict could only ever report
            // "clean" — refuse rather than hand a CI lane a gate that cannot fail. --no-recalc
            // needs no guard: it is what the streaming path already does.
            if policy.strict then
              IO.raiseError(
                unsupportedInStream(
                  "--strict is not supported with --stream (streaming writes never recalculate). Re-run without --stream.",
                  "omit --stream: the in-memory write recalculates the edit's dependency cone and can gate"
                )
              )
            else
              streamAutoSelect(excel, filePath, sheetNameOpt, cmd, mode, warn) *>
                executeStreamingWrite(filePath, sheetNameOpt, outputOpt, cmd, io, warn)
                  .map(Payload.text)
          case _ =>
            // --stream accepted but with no O(1) path for this verb: the workbook is loaded in
            // memory and only the write goes through the streaming backend — say so
            val backendOnly = Warning(
              WarningCode.STREAM_BACKEND_ONLY,
              s"--stream has no O(1) path for ${cmd.verb}: the workbook was loaded in memory " +
                "(a write still goes through the streaming writer)"
            )
            for
              _ <- warn(backendOnly).whenA(stream)
              wb <- readWorkbook(excel, filePath, readerConfig)
              sheet <- defaultSheet(wb, sheetNameOpt, cmd, mode, warn)
              payload <- readQuery match
                case Some(query) =>
                  Reads.run(query, SheetSource.inMemory(wb), sheet.map(_.name.value), mode, warn)
                case None =>
                  executeCommand(
                    wb,
                    sheet,
                    outputOpt,
                    backendOpt,
                    stream,
                    cmd,
                    policy,
                    io,
                    warn,
                    mode
                  ).map(Payload.text)
            yield payload

  /**
   * The verbs that run in O(1) memory under `--stream`: the record-based reads
   * ([[com.tjclp.xl.cli.read.Reads]] over the streaming source: search, stats, view, cell, filter),
   * `bounds` and the metadata fast paths of `runResult` (`describe` without `--full`, `sheets`
   * without `--stats`), and the writes [[executeStreamingWrite]] dispatches (put, putf, style,
   * batch). Every other write verb accepts the flag, loads the workbook in memory and only writes
   * through the streaming writer. `Schema.verbs`' `needs.streaming` column is pinned to this set by
   * SchemaSpec; extend both when a dispatch arm is added.
   */
  private[cli] val streamingVerbs: Set[String] = Set(
    "search",
    "stats",
    "bounds",
    "view",
    "cell",
    "filter",
    "describe",
    "sheets",
    "put",
    "putf",
    "style",
    "batch"
  )

  /** Read and parse the batch source once, for either rendering of a dry run. */
  private def parseBatchDryRun(source: String, io: CliIO): IO[BatchParser.ParseResult] =
    BatchParser.readBatchInput(source, io.stdin).flatMap(BatchParser.parseBatchOperations)

  /** The dry run's text: the count and one summary line per op. */
  private def batchDryRunText(result: BatchParser.ParseResult): String =
    s"Dry run - ${result.ops.size} operations parsed:\n${BatchParser.formatSummary(result.ops)}"

  /**
   * Validate batch JSON and show the summary without writing. The parse warnings go through the
   * run's warning sink like every other warning (ADR-017 §2.3): `Warning[CODE]: …` on stderr in
   * text mode, the envelope's `warnings[]` under `--json`.
   */
  private[cli] def batchDryRun(
    source: String,
    io: CliIO,
    warn: Warning => IO[Unit] = Diagnostics.warn(_, CliIO.system)
  ): IO[String] =
    parseBatchDryRun(source, io).flatMap { result =>
      result.warnings.traverse_(warn).as(batchDryRunText(result))
    }

  /**
   * The dry run as data (`batch --dry-run --json`): `{ops: [{index, op, summary}]}` — `index` the
   * op's 1-based position in the array, the same number every `Object N` message and
   * `location.opIndex` carry (ADR-017 §2.6); `op` the summary's leading keyword (`PUT`, `MERGE`,
   * ...); `summary` the line `formatSummary` prints. The parse warnings ride in the envelope's
   * `warnings[]`, never inside `data`.
   */
  private def batchDryRunData(result: BatchParser.ParseResult): ujson.Value =
    val ops = result.scoped.map { scoped =>
      val summary = BatchParser.formatSummary(Vector(scoped.op)).trim
      ujson.Obj(
        "index" -> ujson.Num(scoped.index),
        "op" -> ujson.Str(summary.takeWhile(_ != ' ')),
        "summary" -> ujson.Str(summary)
      )
    }
    ujson.Obj("ops" -> ujson.Arr.from(ops))

  /** The parsed dry run in the run's mode. */
  private def dryRunPayload(result: BatchParser.ParseResult, mode: OutputMode): Payload =
    mode match
      case OutputMode.Text => Payload.text(batchDryRunText(result))
      case OutputMode.Json => Payload.Json(batchDryRunData(result))

  /** A dry run inside a run: its parse warnings go to the run's sink. */
  private def batchDryRunPayload(
    source: String,
    io: CliIO,
    mode: OutputMode,
    warn: Warning => IO[Unit]
  ): IO[Payload] =
    parseBatchDryRun(source, io).flatMap { result =>
      result.warnings.traverse_(warn).as(dryRunPayload(result, mode))
    }

  /**
   * The standalone `batch --dry-run` (no -f, no -o): a bad source (invalid JSON, missing file) is a
   * failure classified like every other verb's, never an escaped exception; the parse warnings ride
   * on the outcome.
   */
  private[cli] def batchDryRunOutcome(source: String, io: CliIO, mode: OutputMode): IO[Outcome] =
    parseBatchDryRun(source, io).attempt.map {
      case Right(result) => Outcome.ok("batch", dryRunPayload(result, mode), result.warnings)
      case Left(failure) => Outcome.failed("batch", CliError.fromThrowable(failure))
    }

  /**
   * Execute streaming write command (O(1) memory transform). `warn` is the run's warning sink; the
   * streaming batch sends its parse warnings through it like the in-memory batch does.
   */
  private def executeStreamingWrite(
    filePath: Path,
    sheetNameOpt: Option[String],
    outputOpt: Option[Path],
    cmd: CliCommand,
    io: CliIO,
    warn: Warning => IO[Unit] = Diagnostics.warn(_, CliIO.system)
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
          IO.raiseError(outputRequired("--output is required for style command"))
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
          unsupportedInStream(
            "--csv auto-split is not supported with --stream. Omit --stream to use --csv.",
            "omit --stream to use --csv"
          )
        )
      else
        outputOpt match
          case None =>
            IO.raiseError(outputRequired("--output is required for put command"))
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
          IO.raiseError(outputRequired("--output is required for putf command"))
        case Some(outputPath) =>
          StreamingWriteCommands.putFormula(filePath, outputPath, sheetNameOpt, refStr, formulas)

    // `--dry-run` never reaches here: `execute` answers it before any dispatch
    case CliCommand.Batch(source, _, _) =>
      outputOpt match
        case None =>
          IO.raiseError(outputRequired("--output is required for batch command"))
        case Some(outputPath) =>
          StreamingWriteCommands.batch(filePath, outputPath, sheetNameOpt, source, io.stdin, warn)

    case _ =>
      IO.raiseError(
        unsupportedInStream(
          "--stream for write commands only supports: put, putf, style, batch",
          "omit --stream; use --max-size <MB> to load a large file in memory"
        )
      )

  /**
   * @param warn
   *   where a handler's out-of-band notices go (`view`'s csv/html/svg truncation and hidden-line
   *   notices); the runner collects them for the run's stderr or envelope, and the default prints
   *   `Warning[CODE]: …` to `io` directly
   */
  private[cli] def executeCommand(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    outputOpt: Option[Path],
    backendOpt: Option[XmlBackend],
    stream: Boolean,
    cmd: CliCommand,
    policy: WritePolicy = WritePolicy.default,
    io: CliIO = CliIO.system,
    warn: Warning => IO[Unit] = Diagnostics.warn(_, CliIO.system),
    mode: OutputMode = OutputMode.Text
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

    // W2.4: the record-based reads run over the loaded workbook as a source; their payload text
    // is what text mode prints (the runner reaches them through `execute`, which keeps the
    // typed payloads — this arm exists for the verbs the tests drive directly)
    case _: CliCommand.View | _: CliCommand.Cell | _: CliCommand.Search | _: CliCommand.Stats |
        _: CliCommand.Filter =>
      ReadQuery.of(cmd, mode) match
        case Some(query) =>
          Reads
            .run(query, SheetSource.inMemory(wb), sheetOpt.map(_.name.value), mode, warn)
            .map {
              case Payload.Text(text, _, _) => text
              case Payload.Raw(json) => json
              case Payload.Json(value) => ujson.write(value, indent = 2)
            }
        case None =>
          IO.raiseError(
            CliException(CliError(ErrorCode.INTERNAL, s"no read query for ${cmd.verb}"))
          )

    case CliCommand.Eval(formulaStr, overrides) =>
      ReadCommands.eval(wb, sheetOpt, formulaStr, overrides)

    case CliCommand.EvalArray(formulaStr, targetRef, overrides) =>
      ReadCommands.evalArray(wb, sheetOpt, formulaStr, targetRef, overrides)

    // Write commands (require output)
    case CliCommand.Put(refStr, values, csvSplit, detect) =>
      requireOutput("put", outputOpt, backendOpt, stream)(
        WriteCommands.put(wb, sheetOpt, refStr, values, _, _, _, csvSplit, detect, policy, warn)
      )

    case CliCommand.PutFormula(refStr, formulas) =>
      requireOutput("putf", outputOpt, backendOpt, stream)(
        WriteCommands.putFormula(wb, sheetOpt, refStr, formulas, _, _, _, policy, warn)
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

    // `--dry-run` never reaches here: `execute` answers it before any dispatch
    case CliCommand.Batch(source, _, _) =>
      requireOutput("batch", outputOpt, backendOpt, stream)(
        WriteCommands.batch(wb, sheetOpt, source, _, _, _, policy, io.stdin, warn)
      )

    case CliCommand.Recalc(tables, parallel) =>
      // A recalculation asked not to recalculate is a contradiction, not a no-op: say so rather
      // than writing a file the caller will read as freshened (GH-468).
      if policy.noRecalc then
        IO.raiseError(
          CliException(
            CliError.usage(
              "recalc cannot be combined with --no-recalc/--preserve-caches (it exists to rewrite caches). Drop the flag, or drop the recalc.",
              None
            )
          )
        )
      else
        requireOutput("recalc", outputOpt, backendOpt, stream)(
          WriteCommands.recalc(wb, _, _, _, tables, policy, parallel, warn)
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
        WriteCommands.fill(wb, sheetOpt, source, target, direction, _, _, _, policy, warn)
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
        WriteCommands.copyRange(wb, sheetOpt, source, target, valuesOnly, _, _, _, policy, warn)
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
        WriteCommands.insertRows(wb, sheetOpt, at, count, _, _, _, policy, warn)
      )

    case CliCommand.DeleteRows(at, count) =>
      requireOutput("delete-rows", outputOpt, backendOpt, stream)(
        WriteCommands.deleteRows(wb, sheetOpt, at, count, _, _, _, policy, warn)
      )

    case CliCommand.InsertColumns(col, count) =>
      requireOutput("insert-cols", outputOpt, backendOpt, stream)(
        WriteCommands.insertColumns(wb, sheetOpt, col, count, _, _, _, policy, warn)
      )

    case CliCommand.DeleteColumns(col, count) =>
      requireOutput("delete-cols", outputOpt, backendOpt, stream)(
        WriteCommands.deleteColumns(wb, sheetOpt, col, count, _, _, _, policy, warn)
      )

    // The inspection verbs are dispatched in execute() (typed payloads) — never reach here
    case CliCommand.Describe(_) =>
      IO.raiseError(new Exception("Internal: describe is dispatched in execute"))
    case CliCommand.Audit(_) =>
      IO.raiseError(new Exception("Internal: audit is dispatched in execute"))
    case CliCommand.Deps(_, _, _) =>
      IO.raiseError(new Exception("Internal: deps is dispatched in execute"))

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

  /**
   * A write verb without `-o`/`-i`: `OUTPUT_REQUIRED`, exit 2 (still raised after the read until
   * Wave 2).
   */
  private def outputRequired(message: String): CliException =
    CliException(CliError(ErrorCode.OUTPUT_REQUIRED, message))

  private[cli] def requireOutput(
    commandName: String,
    outputOpt: Option[Path],
    backendOpt: Option[XmlBackend],
    stream: Boolean = false
  )(f: (Path, WriterConfig, Boolean) => IO[String]): IO[String] =
    val config = backendOpt.fold(WriterConfig.default)(b => WriterConfig(backend = b))
    outputOpt.fold(
      IO.raiseError[String](outputRequired(missingOutputError(commandName)))
    )(path => f(path, config, stream))

  /**
   * Dispatch `run` with effective output path from --output / --in-place flags.
   *
   * The callback receives (writePath, displayPath) and returns the exit code plus rendered output
   * without printing it. For either write mode, writes go to a sibling temp file that is moved onto
   * the destination before successful output is printed (GH-464, GH-521). Keeping the temp on the
   * same filesystem permits an atomic replacement on supporting providers.
   *
   * Cases:
   *   - `-o` only: writes to a sibling temp and commits every complete output, including the file
   *     intentionally produced by a strict-validation exit
   *   - `-i` only: writes to a sibling temp file then atomically moves onto input. If the command
   *     exits with a non-success code OR throws, the temp is deleted and the original is untouched
   *   - Neither: passes `None` through (for read-only subcommands that don't need output)
   *   - Both: a usage error ("mutually exclusive", exit 2, stderr); nothing is read or written
   */
  private[cli] def runWithOutput(
    outOpt: Option[Path],
    inPlace: Boolean,
    file: Path,
    io: CliIO = CliIO.system,
    mode: OutputMode = OutputMode.Text,
    verb: String = ""
  )(execute: (Option[Path], Option[Path]) => IO[Outcome]): IO[ExitCode] =
    (outOpt, inPlace) match
      case (Some(_), true) =>
        val error =
          CliError.usage("--in-place (-i) and --output (-o) are mutually exclusive", None)
        emit(Outcome.failed(verb, error), mode, io)
      case (Some(out), false) =>
        runStagedOutput(out, ".xl-output-", io, mode, verb)(outcome => outcome.outputComplete)(
          execute
        )
      case (None, false) => execute(None, None).flatMap(emit(_, mode, io))
      case (None, true) =>
        runStagedOutput(file, ".xl-inplace-", io, mode, verb)(outcome =>
          outcome.outputComplete && outcome.exitCode == ExitCode.Success
        )(execute)

  /**
   * Allocate a sibling staging path and register it with the JVM before handing it to a writer. A
   * destination that cannot take a file (missing directory, no permission) fails here, as
   * `IO_WRITE` naming the target ([[classifyWrite]]).
   *
   * The Resource finalizer is the normal cleanup path. `deleteOnExit` is the hard-shutdown
   * backstop: the CLI runtime deliberately stops waiting after two seconds, so a canceled writer
   * may not get enough time to run its finalizer before the JVM exits.
   */
  private def stagedOutput(target: Path, prefix: String): Resource[IO, Path] =
    val directory = Option(target.getParent).getOrElse(Path.of("."))
    val filename = Option(target.getFileName).fold("output.tmp")(_.toString)
    val dot = filename.lastIndexOf('.')
    val suffix = if dot >= 0 then filename.substring(dot) else ".tmp"
    val acquire =
      IO.blocking(Files.createTempFile(directory, prefix, suffix)).flatMap { tmp =>
        IO.blocking(tmp.toFile.deleteOnExit()).attempt.flatMap {
          case Right(_) => IO.pure(tmp)
          case Left(error) =>
            IO.blocking(Files.deleteIfExists(tmp)).attempt *> IO.raiseError[Path](error)
        }
      }
    Resource.make(classifyWrite(target)(acquire))(tmp =>
      IO.blocking(Files.deleteIfExists(tmp)).void
    )

  /**
   * Run one command against a staging path and publish only an explicitly complete output. What the
   * payload then says about the write (`saved`, `written`) is what this step actually did — a
   * committed run names its target, a discarded one (an `-i` strict failure) says nothing was.
   *
   * Nothing escapes: a staging file that cannot be allocated or a commit that fails is an
   * `IO_WRITE` failure ([[classifyWrite]]) rendered like any other — no payload, exit 3, the run's
   * warnings kept — never a stack trace.
   */
  private def runStagedOutput(
    target: Path,
    prefix: String,
    io: CliIO,
    mode: OutputMode,
    verb: String
  )(
    shouldCommit: Outcome => Boolean
  )(
    execute: (Option[Path], Option[Path]) => IO[Outcome]
  ): IO[ExitCode] =
    stagedOutput(target, prefix)
      .use { tmp =>
        execute(Some(tmp), Some(target)).flatMap { outcome =>
          val commit: IO[Boolean] =
            if shouldCommit(outcome) then
              // A read-only command may accept the global -o flag but never touch its staging
              // file. An XLSX is necessarily non-empty, so do not replace a target with the
              // untouched file.
              classifyWrite(target)(
                IO.blocking(Files.size(tmp) > 0L)
                  .ifM(replaceAtomically(tmp, target).as(true), IO.pure(false))
              )
            else IO.pure(false)
          commit.attempt.flatMap {
            case Right(committed) =>
              emit(outcome.committed(Option.when(committed)(target.toString)), mode, io)
            case Left(failure) =>
              emit(
                Outcome.failed(verb, CliError.fromThrowable(failure), outcome.warnings),
                mode,
                io
              )
          }
        }
      }
      .handleErrorWith(failure =>
        emit(Outcome.failed(verb, CliError.fromThrowable(failure)), mode, io)
      )

  /** Replace `target` atomically when supported, with the JDK-prescribed total fallback. */
  private[cli] def replaceAtomically(source: Path, target: Path): IO[Unit] =
    val atomic =
      IO.blocking(
        Files.move(
          source,
          target,
          StandardCopyOption.ATOMIC_MOVE,
          StandardCopyOption.REPLACE_EXISTING
        )
      ).void
    val fallback =
      IO.blocking(Files.move(source, target, StandardCopyOption.REPLACE_EXISTING)).void
    atomicMoveOrFallback(atomic, fallback)

  /** Isolated for deterministic coverage of providers that reject `ATOMIC_MOVE`. */
  private[cli] def atomicMoveOrFallback(atomic: IO[Unit], fallback: IO[Unit]): IO[Unit] =
    atomic.handleErrorWith {
      case _: AtomicMoveNotSupportedException => fallback
      case _: FileAlreadyExistsException => fallback
      case error => IO.raiseError(error)
    }

  /** Require output path or raise user-friendly error, providing path to action */
  private def requireOutputAction(outputOpt: Option[Path], commandName: String)(
    f: Path => IO[String]
  ): IO[String] =
    outputOpt match
      case Some(path) => f(path)
      case None =>
        IO.raiseError(outputRequired(missingOutputError(commandName)))

  /** Build WriterConfig from CLI backend option */
  private def buildWriterConfig(backendOpt: Option[XmlBackend]): WriterConfig =
    backendOpt.fold(WriterConfig.default)(b => WriterConfig(backend = b))

  /** Build ReaderConfig from CLI maxSize option (in MB). 0 means unlimited. */
  private def buildReaderConfig(maxSizeOpt: Option[Long]): ReaderConfig =
    maxSizeOpt match
      case Some(0) => ReaderConfig.permissive
      case Some(mb) => ReaderConfig.default.copy(maxUncompressedSize = mb * 1024 * 1024)
      case None => ReaderConfig.default
