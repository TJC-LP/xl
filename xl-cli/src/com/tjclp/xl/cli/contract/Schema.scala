package com.tjclp.xl.cli.contract

import java.nio.charset.StandardCharsets

import scala.util.{Try, Using}

import com.tjclp.xl.cli.batch.OpRegistry
import com.tjclp.xl.cli.read.Capability

/** One global flag as `xl schema` documents it; `takesValue` comes from [[Argv.globals]]. */
final case class GlobalDoc(name: String, short: Option[String], doc: String) derives CanEqual

/** One row of the exit-code table (ADR-017 §2.3). */
final case class ExitCodeDoc(code: Int, meaning: String) derives CanEqual

/**
 * What a verb requires: `-f` an input workbook; `sheet` — it works on ONE sheet, so THE sheet rule
 * applies (a qualified ref, else `-s`, else the only sheet of a single-sheet book); `output` — a
 * write, so `-o` or `-i`; `streaming` — it runs in O(1) memory under `--stream`
 * ([[StreamSupport.O1]]; the other write verbs accept the flag, load the workbook in memory and
 * only write through the streaming writer; `Main.streamingVerbs` is the set, pinned by SchemaSpec).
 */
final case class Needs(file: Boolean, sheet: Boolean, output: Boolean, streaming: Boolean)
    derives CanEqual

/**
 * What `--stream` does to a verb (GH-638), the one table the runtime refuses from and `xl schema`
 * publishes as each verb's `stream`.
 */
enum StreamSupport derives CanEqual:
  /** Runs in O(1) memory: the streaming reader or writer, or a metadata-only path. */
  case O1

  /**
   * Accepted: the workbook is loaded in memory and only the write goes through the streaming
   * writer, under a `STREAM_BACKEND_ONLY` warning.
   */
  case BackendOnly

  /** Refused before any read (`UNSUPPORTED_IN_STREAM`, exit 2), for the reason given. */
  case Refused(reason: String)

  /** The name `xl schema --json` publishes: `o1`, `backend`, `refused`. */
  def name: String = this match
    case O1 => "o1"
    case BackendOnly => "backend"
    case Refused(_) => "refused"

/**
 * One verb of the contract (ADR-017 §2.13). `path` is the subcommand path (`["sheets", "hide"]`),
 * joined by a space it is the envelope's `verb`; `exit` lists every exit code the verb can end
 * with; `batchTwin` names the batch op that makes the same edit; `since` is the release the verb is
 * documented from (never earlier than its introduction).
 */
final case class VerbDoc(
  path: Vector[String],
  summary: String,
  needs: Needs,
  exit: Vector[Int],
  batchTwin: Option[String],
  since: String
) derives CanEqual:

  /** The envelope's `verb`: the path joined by a space. */
  def verb: String = path.mkString(" ")

  /** What `--stream` does to this verb ([[Schema.streamSupport]]). */
  def stream: StreamSupport = Schema.streamSupport(this)

  /**
   * The verb's own flags `--stream` refuses once parsed ([[Schema.refusedWith]]); empty otherwise.
   */
  def refusedWith: Vector[String] = Schema.refusedWith(this)

/**
 * The machine-readable contract `xl schema --json` publishes (ADR-017 §2.13): the exit-code table,
 * every error and warning code, the global flags, the verb table, the batch document's JSON Schema,
 * the function registry and the envelope's JSON Schema — one document, so an agent reads the whole
 * vocabulary from the binary it is talking to instead of from prose that drifts.
 *
 * The verb table is a literal here, pinned to the parser by `SchemaSpec` (its heads must equal the
 * subcommand names decline renders); Wave 2 derives it from the command registry that also builds
 * the parser.
 */
object Schema:

  // ---------------------------------------------------------------------------------------------
  // Exit codes and globals
  // ---------------------------------------------------------------------------------------------

  val exitCodes: Vector[ExitCodeDoc] = Vector(
    ExitCodeDoc(ExitCodes.ok.code, "ok"),
    ExitCodeDoc(
      ExitCodes.signal.code,
      "completed with findings or a failed gate (diff differs, lint findings, audit " +
        "--fail-on-findings, --strict) — never a failure; with -o the file is written, with -i " +
        "the input is left untouched"
    ),
    ExitCodeDoc(
      ExitCodes.usage.code,
      "usage — the command line is wrong; nothing read, nothing written"
    ),
    ExitCodeDoc(
      ExitCodes.failed.code,
      "failed — the operation could not complete; nothing written"
    )
  )

  /** The global flags, in `xl --help` order; every name and short is a key of [[Argv.globals]]. */
  val globals: Vector[GlobalDoc] = Vector(
    GlobalDoc("--json", None, "Wrap every result, success or failure, in the JSON envelope"),
    GlobalDoc("--file", Some("-f"), "Excel file to operate on"),
    GlobalDoc(
      "--sheet",
      Some("-s"),
      "Sheet to select; a sheet-qualified ref wins over it and a single-sheet book needs neither"
    ),
    GlobalDoc(
      "--max-size",
      None,
      "Max uncompressed size in MB for an in-memory load (default 100, 0 = unlimited); lifts the " +
        "security limit only — the heap (native image: 8 GB unless -Xmx<size> is passed; put it " +
        "before -f to be safe) still bounds what fits: a load estimated not to fit is refused with " +
        "RESOURCE_LIMIT, one that may not fit proceeds under a MEMORY_PRESSURE warning; below 0 " +
        "is a usage error"
    ),
    GlobalDoc("--output", Some("-o"), "Output file for a write"),
    GlobalDoc("--in-place", Some("-i"), "Edit the input file in place (instead of -o)"),
    GlobalDoc("--backend", None, "XML writer backend: scalaxml (default) or saxstax (faster)"),
    GlobalDoc(
      "--stream",
      None,
      "O(1)-memory streaming for large files: search, stats, bounds, view, cell, filter, " +
        "describe, sheets, names, lint; put, putf, style and the streamable batch ops (other " +
        "write verbs accept the flag but load the workbook; each verb's `stream` says which — " +
        "o1, backend, refused)"
    ),
    GlobalDoc(
      "--no-recalc",
      None,
      "Write verbs: apply the edit and recalculate nothing; structural edits leave the formulas " +
        "they invalidated uncached"
    ),
    GlobalDoc("--preserve-caches", None, "Alias for --no-recalc"),
    GlobalDoc(
      "--strict",
      None,
      "Write verbs: exit 1 when the recalculation reports formula errors, non-convergence or " +
        "data-table seed warnings (after `view` it is view's own --eval gate)"
    )
  )

  // ---------------------------------------------------------------------------------------------
  // The verb table
  // ---------------------------------------------------------------------------------------------

  private val plain: Vector[Int] = Vector(0, 2, 3)
  private val gated: Vector[Int] = Vector(0, 1, 2, 3)

  private def info(name: String, summary: String, since: String): VerbDoc =
    VerbDoc(Vector(name), summary, Needs(false, false, false, false), plain, None, since)

  private def read(
    path: String,
    summary: String,
    sheet: Boolean,
    streaming: Boolean,
    since: String,
    exit: Vector[Int]
  ): VerbDoc =
    VerbDoc(
      path.split(' ').toVector,
      summary,
      Needs(file = true, sheet = sheet, output = false, streaming = streaming),
      exit,
      None,
      since
    )

  private def write(
    path: String,
    summary: String,
    sheet: Boolean,
    twin: Option[String],
    since: String,
    streaming: Boolean,
    exit: Vector[Int]
  ): VerbDoc =
    VerbDoc(
      path.split(' ').toVector,
      summary,
      Needs(file = true, sheet = sheet, output = true, streaming = streaming),
      exit,
      twin,
      since
    )

  /** Every verb, in `xl --help` order (the same order as [[Argv.verbs]]). */
  val verbs: Vector[VerbDoc] = Vector(
    info("rasterizers", "List available SVG-to-raster backends and their status", "0.6.1"),
    info("functions", "List the supported Excel functions (--json: typed rows)", "0.4.2"),
    info(
      "schema",
      "Print the CLI contract: verbs, globals, exit and error codes, batch ops, functions (--json)",
      "0.20.0"
    ),
    info("new", "Create a blank xlsx file (--sheet <name> repeatable)", "0.1.0"),
    read(
      "diff",
      "Compare two workbooks (-g <file2>) and report cell, style and structure differences",
      sheet = false,
      streaming = false,
      "0.11.3",
      gated
    ),
    read(
      "lint",
      "Validate the raw package against the Excel-repair classes: child order, r:id resolution, " +
        "content-type coverage, over-max refs, data-table integrity, <f> canon, external refs, " +
        "defined names, calc chain (read-only)",
      sheet = false,
      streaming = true,
      "0.15.0",
      gated
    ),
    VerbDoc(
      Vector("eval"),
      "Evaluate a formula without modifying the sheet (--with overrides; no -f for constants)",
      Needs(file = false, sheet = true, output = false, streaming = false),
      plain,
      None,
      "0.4.2"
    ),
    read(
      "evala",
      "Evaluate an array formula and display, or spill (--at), the result grid",
      sheet = true,
      streaming = false,
      "0.9.0",
      plain
    ),
    read(
      "sheets",
      "List sheets with visibility state and dimension (--stats loads the book for counts)",
      sheet = false,
      streaming = true,
      "0.1.0",
      plain
    ),
    write(
      "sheets hide",
      "Hide a sheet from the sheet tabs (--very for VBA-only)",
      sheet = false,
      None,
      "0.9.2",
      streaming = false,
      plain
    ),
    write(
      "sheets show",
      "Show a hidden sheet",
      sheet = false,
      None,
      "0.9.2",
      streaming = false,
      plain
    ),
    read(
      "names",
      "List defined names (named ranges)",
      sheet = false,
      streaming = true,
      "0.2.0",
      plain
    ),
    read(
      "bounds",
      "Show the used range (instant from the dimension element; --scan for an accurate scan)",
      sheet = true,
      streaming = true,
      "0.9.0",
      plain
    ),
    read(
      "view",
      "View a range as markdown, json, csv, html, svg, png, jpeg, webp or pdf (--eval, --formulas)",
      sheet = true,
      streaming = true,
      "0.1.0",
      gated
    ),
    read(
      "cell",
      "Get one cell's value, style, comment and direct dependencies",
      sheet = true,
      streaming = true,
      "0.1.0",
      plain
    ),
    read(
      "search",
      "Search cells by regex (all sheets unless -s; the scan stops at --limit, --total counts every match)",
      sheet = false,
      streaming = true,
      "0.1.0",
      plain
    ),
    read(
      "stats",
      "Statistics for the numeric values in a range",
      sheet = true,
      streaming = true,
      "0.2.0",
      plain
    ),
    read(
      "filter",
      "Filter rows of the used range with a --where predicate (read-only)",
      sheet = true,
      streaming = true,
      "0.11.3",
      plain
    ),
    read(
      "describe",
      "Orient in a workbook: sheets, defined names, date system (--full adds per-sheet counts)",
      sheet = false,
      streaming = true,
      "0.20.0",
      plain
    ),
    read(
      "audit",
      "Find every reason a number can be wrong, in one pass (--fail-on-findings exits 1)",
      sheet = false,
      streaming = false,
      "0.20.0",
      gated
    ),
    read(
      "deps",
      "Trace one cell's precedents and dependents, hop by hop (--direction, --depth)",
      sheet = true,
      streaming = false,
      "0.20.0",
      plain
    ),
    write(
      "batch",
      "Apply multiple operations atomically from JSON (--dry-run validates; --schema prints the op schema)",
      sheet = true,
      None,
      "0.1.0",
      streaming = true,
      gated
    ),
    write(
      "put",
      "Write value(s) to a cell or range with smart type detection (--no-detect, --csv)",
      sheet = true,
      Some("put"),
      "0.1.0",
      streaming = true,
      gated
    ),
    write(
      "putf",
      "Write formula(s) to a cell or range; one formula over a range drags with $ anchoring",
      sheet = true,
      Some("putf"),
      "0.1.0",
      streaming = true,
      gated
    ),
    write(
      "style",
      "Apply formatting to cells; styles merge by default (--replace to overwrite)",
      sheet = true,
      Some("style"),
      "0.2.0",
      streaming = true,
      plain
    ),
    write(
      "row",
      "Set row properties: height, hide/show",
      sheet = true,
      Some("rowheight"),
      "0.3.0",
      streaming = false,
      plain
    ),
    write(
      "col",
      "Set column properties: width, hide/show, auto-fit (ranges like A:F)",
      sheet = true,
      Some("colwidth"),
      "0.3.0",
      streaming = false,
      plain
    ),
    write(
      "group-rows",
      "Group rows into a collapsible outline (--level, --collapsed)",
      sheet = true,
      Some("group-rows"),
      "0.18.0",
      streaming = false,
      plain
    ),
    write(
      "group-cols",
      "Group columns into a collapsible outline (--level, --collapsed)",
      sheet = true,
      Some("group-cols"),
      "0.18.0",
      streaming = false,
      plain
    ),
    write(
      "ungroup-rows",
      "Remove outline grouping from rows",
      sheet = true,
      Some("ungroup-rows"),
      "0.18.0",
      streaming = false,
      plain
    ),
    write(
      "ungroup-cols",
      "Remove outline grouping from columns",
      sheet = true,
      Some("ungroup-cols"),
      "0.18.0",
      streaming = false,
      plain
    ),
    write(
      "autofit",
      "Auto-fit column widths from content (--columns A:F)",
      sheet = true,
      Some("autofit"),
      "0.9.6",
      streaming = false,
      plain
    ),
    write(
      "recalc",
      "Recalculate every formula and rewrite the cached values (--tables, --parallel)",
      sheet = false,
      None,
      "0.12.6",
      streaming = false,
      gated
    ),
    write(
      "import",
      "Import CSV data with type detection (--new-sheet, --delimiter, --skip-header)",
      sheet = true,
      None,
      "0.7.0",
      streaming = false,
      plain
    ),
    write(
      "import-md",
      "Import a GFM markdown table with type detection (--start, --new-sheet)",
      sheet = true,
      None,
      "0.11.3",
      streaming = false,
      plain
    ),
    write(
      "add-sheet",
      "Add a new empty sheet (--after, --before)",
      sheet = false,
      Some("add-sheet"),
      "0.2.0",
      streaming = false,
      plain
    ),
    write("remove-sheet", "Remove a sheet", sheet = false, None, "0.2.0", streaming = false, plain),
    write(
      "rename-sheet",
      "Rename a sheet and rewrite every formula, defined name and chart reference to it",
      sheet = false,
      Some("rename-sheet"),
      "0.2.0",
      streaming = false,
      plain
    ),
    write(
      "move-sheet",
      "Move a sheet to a new position (--to, --after, --before)",
      sheet = false,
      None,
      "0.2.0",
      streaming = false,
      plain
    ),
    write(
      "copy-sheet",
      "Copy a sheet to a new name",
      sheet = false,
      None,
      "0.2.0",
      streaming = false,
      plain
    ),
    write(
      "merge",
      "Merge cells in a range",
      sheet = true,
      Some("merge"),
      "0.2.0",
      streaming = false,
      plain
    ),
    write(
      "unmerge",
      "Unmerge cells in a range",
      sheet = true,
      Some("unmerge"),
      "0.2.0",
      streaming = false,
      plain
    ),
    write(
      "comment",
      "Add a comment to a cell (--author)",
      sheet = true,
      Some("comment"),
      "0.6.0",
      streaming = false,
      plain
    ),
    write(
      "remove-comment",
      "Remove a cell's comment",
      sheet = true,
      Some("remove-comment"),
      "0.9.6",
      streaming = false,
      plain
    ),
    write(
      "clear",
      "Clear cell contents, styles or comments in a range (--all, --styles, --comments)",
      sheet = true,
      Some("clear"),
      "0.6.0",
      streaming = false,
      plain
    ),
    write(
      "fill",
      "Fill cells with a source value or formula, Excel Ctrl+D/Ctrl+R (--right)",
      sheet = true,
      None,
      "0.6.0",
      streaming = false,
      gated
    ),
    write(
      "sort",
      "Sort rows in a range by one or more columns (--by, --then-by, --desc, --numeric, --header)",
      sheet = true,
      None,
      "0.6.0",
      streaming = false,
      plain
    ),
    write(
      "freeze",
      "Freeze panes at a cell (rows above and columns left are locked)",
      sheet = true,
      Some("freeze"),
      "0.10.0",
      streaming = false,
      plain
    ),
    write(
      "unfreeze",
      "Remove freeze panes",
      sheet = true,
      Some("unfreeze"),
      "0.10.0",
      streaming = false,
      plain
    ),
    write(
      "copy",
      "Copy a range to another location with formula adjustment (--values-only)",
      sheet = true,
      Some("copy"),
      "0.10.0",
      streaming = false,
      gated
    ),
    write(
      "name add",
      "Add or replace a workbook-scoped named range",
      sheet = false,
      None,
      "0.10.0",
      streaming = false,
      plain
    ),
    write(
      "name rm",
      "Remove a named range",
      sheet = false,
      None,
      "0.10.0",
      streaming = false,
      plain
    ),
    write(
      "insert-rows",
      "Insert rows; shifts cells and rewrites formulas",
      sheet = true,
      None,
      "0.10.0",
      streaming = false,
      gated
    ),
    write(
      "delete-rows",
      "Delete rows; shifts cells and rewrites formulas (#REF! on loss)",
      sheet = true,
      None,
      "0.10.0",
      streaming = false,
      gated
    ),
    write(
      "insert-cols",
      "Insert columns; shifts cells and rewrites formulas",
      sheet = true,
      None,
      "0.10.0",
      streaming = false,
      gated
    ),
    write(
      "delete-cols",
      "Delete columns; shifts cells and rewrites formulas (#REF! on loss)",
      sheet = true,
      None,
      "0.10.0",
      streaming = false,
      gated
    ),
    write(
      "chart add",
      "Add a typed chart built from sheet data ranges (--type, --data, --at)",
      sheet = true,
      Some("chart"),
      "0.12.0",
      streaming = false,
      plain
    ),
    write(
      "add-image",
      "Embed an image (png/jpeg/gif/bmp/tiff/emf/wmf) at a cell or over a range",
      sheet = true,
      None,
      "0.12.0",
      streaming = false,
      plain
    ),
    write(
      "sheet-view",
      "Set sheet view options: gridlines, zoom, tab selection",
      sheet = true,
      Some("sheet-view"),
      "0.13.0",
      streaming = false,
      plain
    ),
    write(
      "tab-color",
      "Set or clear the sheet tab color (named, #hex, rgb(), theme:accent1[:tint])",
      sheet = true,
      Some("tab-color"),
      "0.13.0",
      streaming = false,
      plain
    ),
    write(
      "autofilter",
      "Set or clear the sheet-level autoFilter range",
      sheet = true,
      Some("autofilter"),
      "0.18.0",
      streaming = false,
      plain
    ),
    write(
      "page-setup",
      "Set print page setup: orientation, scale, fit-to-page",
      sheet = true,
      Some("page-setup"),
      "0.13.0",
      streaming = false,
      plain
    ),
    write(
      "header-footer",
      "Set print header/footer text with Excel codes (&L/&C/&R, &P, &N, &D, &F, &A)",
      sheet = true,
      Some("header-footer"),
      "0.13.0",
      streaming = false,
      plain
    ),
    write(
      "cf add",
      "Add a conditional-formatting rule (--range, --rule DSL, format flags)",
      sheet = true,
      Some("cf"),
      "0.13.0",
      streaming = false,
      plain
    ),
    read(
      "cf list",
      "List conditional-formatting rules on the sheet (read-only)",
      sheet = true,
      streaming = false,
      "0.13.0",
      plain
    )
  )

  // ---------------------------------------------------------------------------------------------
  // The envelope's JSON Schema
  // ---------------------------------------------------------------------------------------------

  /**
   * `schema/envelope.schema.json` from the classpath — the published shape of the `--json`
   * envelope, the same file the contract suite validates every envelope against. Total: a build
   * without the resource yields a schema-shaped placeholder that names the gap.
   */
  lazy val envelope: ujson.Value =
    Option(getClass.getResourceAsStream("/schema/envelope.schema.json"))
      .flatMap { stream =>
        Using(stream)(s => new String(s.readAllBytes(), StandardCharsets.UTF_8))
          .flatMap(text => Try(ujson.read(text)))
          .toOption
      }
      .getOrElse(
        ujson.Obj(
          "$comment" -> ujson.Str("schema/envelope.schema.json is missing from this build")
        )
      )

  // ---------------------------------------------------------------------------------------------
  // Renderings
  // ---------------------------------------------------------------------------------------------

  private def optStr(value: Option[String]): ujson.Value =
    value.fold[ujson.Value](ujson.Null)(ujson.Str.apply)

  private def verbJson(verb: VerbDoc): ujson.Obj =
    ujson.Obj(
      "path" -> ujson.Arr.from(verb.path.map(ujson.Str.apply)),
      "summary" -> ujson.Str(verb.summary),
      "needs" -> ujson.Obj(
        "file" -> ujson.Bool(verb.needs.file),
        "sheet" -> ujson.Bool(verb.needs.sheet),
        "output" -> ujson.Bool(verb.needs.output),
        "streaming" -> ujson.Bool(verb.needs.streaming)
      ),
      "stream" -> ujson.Str(verb.stream.name),
      "refusedWith" -> ujson.Arr.from(verb.refusedWith.map(ujson.Str.apply)),
      "exit" -> ujson.Arr.from(verb.exit.map(e => ujson.Num(e))),
      "batchTwin" -> optStr(verb.batchTwin),
      "since" -> ujson.Str(verb.since)
    )

  // ---------------------------------------------------------------------------------------------
  // --stream per verb (GH-638)
  // ---------------------------------------------------------------------------------------------

  /**
   * The verbs that refuse `--stream`, with the reason the refusal names. Every verb not here runs
   * in O(1) memory (`needs.streaming`) or accepts the flag and loads the workbook (the other
   * writes).
   */
  private val streamRefusals: Map[String, String] = Map(
    "rasterizers" -> "it reads no workbook",
    "functions" -> "it reads no workbook",
    "schema" -> "it reads no workbook",
    "new" -> "it writes a new file and reads none",
    "diff" -> "the two workbooks are compared in memory",
    "eval" -> "formula evaluation needs the loaded workbook",
    "evala" -> "formula evaluation needs the loaded workbook",
    "audit" -> "the analysis needs the whole workbook",
    "deps" -> "the graph needs the whole workbook"
  )

  /**
   * What `--stream` does to a verb: O(1) when it claims it, refused when the table says, else
   * backend-only.
   */
  def streamSupport(verb: VerbDoc): StreamSupport =
    if verb.needs.streaming then StreamSupport.O1
    else streamRefusals.get(verb.verb).fold(StreamSupport.BackendOnly)(StreamSupport.Refused(_))

  /**
   * The flags of an O(1) verb that `--stream` refuses once the verb has parsed them — each an
   * `UNSUPPORTED_IN_STREAM` from the verb's own handler, with the in-memory alternative as the
   * hint. (`batch`'s ops are marked `x-streamable` in its own schema, `xl batch --schema`.)
   */
  private val streamRefusedFlags: Map[String, Vector[String]] = Map(
    "view" -> Vector("--eval", "--format html/svg/png/jpeg/webp/pdf"),
    "sheets" -> Vector("--stats"),
    "describe" -> Vector("--full"),
    "put" -> Vector("--csv", "--strict"),
    "putf" -> Vector("--strict"),
    "style" -> Vector("--strict"),
    "batch" -> Vector("--strict")
  )

  /**
   * The verb's own flags `--stream` refuses ([[streamRefusedFlags]]); empty for every other verb.
   */
  def refusedWith(verb: VerbDoc): Vector[String] =
    streamRefusedFlags.getOrElse(verb.verb, Vector.empty)

  /**
   * The `UNSUPPORTED_IN_STREAM` refusal (exit 2) for a verb head — `audit`, `sheets`, `name` —
   * every form of which refuses `--stream`; `None` for a head with a form that takes it (a verb's
   * own flag may still be refused after parsing: `describe --full`, `sheets --stats`,
   * `view --eval`). The text is the refusal every such verb has always given: `<verb> is not
   * supported with --stream (<reason>)`, and the in-memory alternative as the hint.
   */
  def streamRefusal(head: String): Option[CliError] =
    val forms = verbs.filter(_.path.headOption.contains(head))
    val reasons = forms.map(_.stream).collect { case StreamSupport.Refused(reason) => reason }
    Option.when(forms.nonEmpty && reasons.size == forms.size)(reasons).flatMap(_.headOption).map {
      reason =>
        CliError(
          ErrorCode.UNSUPPORTED_IN_STREAM,
          s"$head is not supported with --stream ($reason)",
          hint = Some(
            if forms.exists(_.needs.file) then
              "omit --stream; use --max-size <MB> to load a large file in memory"
            else "omit --stream"
          )
        )
    }

  /**
   * `{version, exitCodes, errorCodes, warningCodes, globals, verbs, capabilities, batchOps,
   * functions, envelope}` — what `xl schema --json` prints as `data`. `capabilities` is the
   * [[com.tjclp.xl.cli.read.Capability]] table: what a read verb can ask of a loaded workbook and
   * of the streaming reader (`[{name, doc, inMemory, streaming}]`).
   */
  def json(version: String): ujson.Obj =
    ujson.Obj(
      "version" -> ujson.Str(version),
      "exitCodes" -> ujson.Arr.from(
        exitCodes.map(e =>
          ujson.Obj("code" -> ujson.Num(e.code), "meaning" -> ujson.Str(e.meaning))
        )
      ),
      "errorCodes" -> ujson.Arr.from(
        ErrorCode.all.map(code =>
          ujson.Obj("code" -> ujson.Str(code), "exit" -> ujson.Num(ExitCodes.forCode(code).code))
        )
      ),
      "warningCodes" -> ujson.Arr.from(WarningCode.all.map(ujson.Str.apply)),
      "globals" -> ujson.Arr.from(globals.map { g =>
        ujson.Obj(
          "name" -> ujson.Str(g.name),
          "short" -> optStr(g.short),
          "takesValue" -> ujson.Bool(Argv.globals.getOrElse(g.name, false)),
          "doc" -> ujson.Str(g.doc)
        )
      }),
      "verbs" -> ujson.Arr.from(verbs.map(verbJson)),
      "capabilities" -> Capability.json,
      "batchOps" -> OpRegistry.jsonSchema(version),
      "functions" -> FunctionDoc.toJson(FunctionDoc.all),
      "envelope" -> envelope
    )

  private def needsCell(needs: Needs): String =
    val parts = Vector(
      Option.when(needs.file)("-f"),
      Option.when(needs.sheet)("-s"),
      Option.when(needs.output)("-o|-i"),
      Option.when(needs.streaming)("--stream")
    ).flatten
    if parts.isEmpty then "-" else parts.mkString(" ")

  /** The STREAM cell: the verb's answer to `--stream`, and the flags it still refuses under it. */
  private def streamCell(verb: VerbDoc): String =
    val refused = verb.refusedWith
    if refused.isEmpty then verb.stream.name
    else s"${verb.stream.name} (not ${refused.mkString(", ")})"

  /** What `xl schema` prints: the verb table as text, with the globals and exit codes above it. */
  def verbTable(version: String): String =
    val headers = Vector("VERB", "NEEDS", "STREAM", "EXIT", "BATCH TWIN", "SINCE", "SUMMARY")
    val rows = verbs.map { v =>
      Vector(
        v.verb,
        needsCell(v.needs),
        streamCell(v),
        v.exit.mkString(" "),
        v.batchTwin.getOrElse("-"),
        v.since,
        v.summary
      )
    }
    val widths = headers.indices.map { i =>
      (headers(i) +: rows.map(_(i))).foldLeft(0)((w, cell) => math.max(w, cell.length))
    }
    def line(cells: Vector[String]): String =
      cells.zipWithIndex
        .map { (cell, i) => if i == cells.size - 1 then cell else cell.padTo(widths(i), ' ') }
        .mkString("  ")
        .replaceAll("\\s+$", "")
    val globalsLine = globals
      .map(g => g.short.fold(g.name)(s => s"$s|${g.name}"))
      .mkString(" ")
    val exitLines = exitCodes.map(e => s"  ${e.code}  ${e.meaning}")
    (Vector(
      s"xl $version — ${verbs.size} verbs. `xl <verb> --help` shows a verb's options; " +
        "`xl schema --json` prints this contract as JSON.",
      s"Globals (accepted anywhere on the command line): $globalsLine",
      "NEEDS: -f an input workbook; -s ONE sheet (a qualified ref, else -s, else the only sheet " +
        "of a single-sheet book); -o|-i an output; --stream runs in O(1) memory.",
      "STREAM: o1 runs in O(1) memory under --stream (not …: the verb's own flags it refuses " +
        "there); backend loads the workbook and only writes through the streaming writer; " +
        "refused is UNSUPPORTED_IN_STREAM (exit 2) before any read.",
      "Exit codes:"
    ) ++ exitLines ++ Vector("", line(headers)) ++ rows.map(line)).mkString("\n")
