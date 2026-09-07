# ADR-017: Agent-first operation algebra and CLI contract

**Date**: 2026-09-07 (synthesis of the agent-first design panel; supersedes nothing)
**Status**: ✅ Accepted — Wave 1 in execution (`docs/plan/agent-first-refactor.md`); Waves 2–4 scheduled as GitHub-issue-ready items
**Index**: to be listed in `docs/design/decisions.md` as "ADR-017: Agent-first operation algebra and CLI contract" (the integrator adds the index line; this file is the record)

The mission this record serves: *make xl the best Excel library in the world for agents; scripting
and the CLI are equally important.* An agent touches xl through exactly two surfaces — the
stateless `xl` CLI (verbs and `batch` JSON) and scala-cli scripts that `import
com.tjclp.xl.scripting.{*, given}`. Every decision below is judged by what those two surfaces make
an agent see, type, parse, and get silently wrong.

---

## 1. Context

Twelve understanding lanes, a completeness critic, three independent designs (algebra-first,
contract-first, agent-journey-first) and three adversarial judges converged on the same
verified evidence. Line numbers are for the working tree at `0d53e76` on
`claude/xl-refactor-agents-c1amtj` (= `origin/main` `a9cdb73`, PRs #569–#572 merged, plus the
proxy-safe bootstrap and the GH-477 codec fix).

**Three write vocabularies with drifting semantics.** The CLI verb ADT `enum CliCommand` has 55
cases, every field a `String` (`xl-cli/src/com/tjclp/xl/cli/Command.scala:35-213`). The batch
ADT `enum BatchOp` has 35 cases behind 32 op names, every ref a `String`
(`xl-cli/src/com/tjclp/xl/cli/helpers/BatchParser.scala:64-152`); the only complete list of op
names in the tool is the unknown-op error text (`:603`). The `--stream` path is a third
implementation: a range `putf` under `--stream` writes the same formula into every cell — "NO
dragging in streaming mode" (`StreamingWriteCommands.scala:281`) — and twenty-one op shapes are
rejected with one generic sentence (`:806-820`). Only `Style/Put/PutFormula/Batch` stream at all
(`Main.scala:2147-2152`); every other verb with `--stream` silently swaps the writer after a full
load. Spreadsheet semantics agents want from scripts (sort, fill, copy, autofit, grouping,
appearance) live only in `WriteCommands.scala` (2,468 lines) and its helpers.

**Prose-only output.** `ujson` is used by `FilterCommands`, `LintCommands`, `DiffCommands` and
`BatchParser` only; `view --format json` is a hand-built `StringBuilder`
(`output/JsonRenderer.scala`). Four of 56 verbs emit JSON, each a different shape; no error is
ever JSON; there is no schema for batch input.

**Exit 1 means five things and errors go to stdout.** Every runtime failure is `ExitCode.Error`
(`Main.scala:1762-1767`); the `--strict` gate is `ExitCode(1)` (`:1757`); `diff` differs and
`lint` findings are `ExitCode(1)` (`:2015`, `:2047`); `ExitCode(2)` exists at exactly three sites
(`:209`, `:2017`, `:2049`); decline usage errors exit 1 with the full 65-subcommand help. Error
text is printed with `IO.println` — to **stdout** — at `:1987`, `:2017`, `:2049`, `:2075`,
`:2872-2874`. `Format.error(XLError)` (`output/Format.scala:122`) has zero callers, and
`docs/reference/cli.md:1401-1408` documents an `Error/Location/Details/Suggestion` format no code
emits.

**Three sheet rules, documented as a fourth.** In-memory sheet-level verbs *require* `-s` or a
qualified ref (`helpers/SheetResolver.scala:50-56`, reached by every unqualified ref through
`resolveRef :92-109`); streaming reads default to the *first* sheet
(`StreamingReadCommands.scala:130-136`); streaming writes *refuse* multi-sheet books without `-s`
(`StreamingWriteCommands.scala:1125-1131`); `search`/`sheets`/`names` operate on all sheets. The
skill says "Commands default to first sheet" (`plugin/skills/xl-cli/SKILL.md:214`), the reference
says "required for unqualified ranges" (`docs/reference/cli.md:40`), and CLAUDE.md says `cell`
auto-detects a single sheet — which no in-memory path does. Batch re-implements "requires --sheet"
at 17 sites in `BatchParser.scala` (`:1175` … `:2043`), and 17 of the 32 ops cannot name a sheet.

**Rename is formula-blind.** `Workbook.rename` remaps typed chart references and marks
`SourceContext`, nothing else (`xl-core/src/com/tjclp/xl/workbooks/Workbook.scala:181-198`); the
`rename-sheet` verb and batch op call it directly (`SheetCommands.scala:144`,
`BatchParser.scala:1860`), so `='Old'!A1` survives the rename and evaluates to `#REF!` in Excel
(GH-559). `Sheet.insertRows` is likewise formula-blind (`Sheet.scala:149-166`); the formula-aware
`StructuralEditor.*Checked` lives in xl-evaluator and is exported by neither the `formula`
package object nor `formulaExports`.

**Twenty `transparent inline` union entry points** decide at the call site whether a script gets
`Sheet` or `XLResult[Sheet]` — `Sheet.scala:445/476/642/723/1273`,
`Workbook.scala:41/105/143/222/251/280/588/608`, `extensions.scala:72/138/159/175/190/209/343`.
Only `Sheet.named` (`Sheet.scala:1290`) has an explicit runtime twin.

**Docs are hand-copied and drift.** `SKILL.md:742` says "26 operations"; `cli.md:70` says "108
supported functions" (115 + LET exist, `FunctionRegistry.all` is macro-collected);
`plugin/.claude-plugin/plugin.json:4` is `0.19.1` while `build.mill:17` is `0.19.3`; CLAUDE.md
shows `putf … --from B2` but `putfCmd` is `(refArg, formulasArg)` with no `--from`
(`Main.scala:1027-1031`).

**Other verified facts the decision rests on.**
- `CommandIOApp.run(List[String])` is `final` (javap on decline-effect 2.5.0): argv normalisation
  and a usage exit code require `Main extends IOApp`.
- The xl-agent grader shells out to `xl … view <range> --format json --eval` and decodes the *bare*
  `{sheet, range, rows}` payload (`xl-agent/src/com/tjclp/xl/agent/benchmark/Evaluator.scala:61-79`).
- The surgical writer copies any sheet not in `tracker.modifiedSheets` byte-for-byte from the
  source zip (`xl-ooxml/src/com/tjclp/xl/ooxml/XlsxWriter.scala:94`, `:700-720`); every existing
  mutation goes through `Workbook.put/update/rename/...`, which mark the tracker
  (`Workbook.scala:78-88`, `:195-198`), and `RecalcResult.cacheResults` re-installs sheets "via
  `Workbook.put` — not copy" for exactly this reason (`Recalc.scala:313-324`).
- GH-560 root cause: `Sheet.putSingle` merges a codec-inferred style through `mergeStyles`, which
  applies the codec `numFmt` only when the existing one is `General` (`Sheet.scala:484-514`).
  Correct for an *inferred* hint (a `LocalDate`), wrong for the *explicit* batch `format` key,
  which reaches it as `Formatted(cellValue, numFmt)` (`BatchParser.scala:1303-1307`) while the
  `putf` path replaces (`:1319-1324`).
- GH-558 root cause is the worksheet writer, not the model: `Sheet.shiftAxis` remaps
  `rowProperties` correctly; `OoxmlWorksheet.scala:500-510` re-emits every cell-free *source*
  `<row>` at its original index with its original attributes, and `:516-521` emits the shifted
  domain property at the new index.
- `DependentRecalculation.recalculateAfterEdit(wb, sheetName, modifiedRefs, clock): RecalcResult`
  is `private[xl]` (`DependentRecalculation.scala:38-43`) and `RecalcResult.cacheResults` is
  `private[eval]` (`Recalc.scala:236`): the single cache-finalisation path exists and must be
  widened, not forked. The `@targetName` overload lattice is WorkbookEvaluator 11, SheetEvaluator
  7, DataTableSeeder 5, and `ExcelRecalc` has 5 `writeRecalculated` overloads.
- The prelude excludes the low-priority `default` given by name on two export hops because the
  LowPriority trick "does not survive export forwarding" (`xl/src/com/tjclp/xl/scripting.scala:31-44`);
  extension methods with default arguments must not be reached through a wildcard export
  (`exports.scala:62-65`, `SheetEvaluator.scala:376-378`) — `recalculateDependents(modifiedRefs,
  workbook = None, clock = Clock.system)` at `DependentRecalculation.scala:166-170` is why
  `DependentRecalculation` is exported as an object only.
- Test net: 742 `test(`/`property(` sites in xl-cli; only `MainSpec:1375`, `InPlaceSpec:304`,
  `LintCommandSpec:210` and `BatchRecalcSpec:558/569` parse real argv; zero xl-cli tests assert
  `Error:` on stdout; exit-code pins are `MainSpec:1447,1475,1477,1536,1537` (strict → 1),
  `DiffCommandSpec:265` (1) `:278` (2), `LintCommandSpec:122,134,152` (1) `:144,244,251` (2),
  `InPlaceSpec:132-136` (1) `:220` (`ExitCode.Error`), and `InPlaceSpec:105-109,265-276` inject
  `ExitCode.Error` as plumbing (unaffected by any table change).
- GH-477 is already fixed on this branch (`0d53e76`): `Cell.effectiveValue`,
  `Cell.isUncachedFormula`, `CellReader.readStrict`, `Sheet.readTypedStrict`, 25 tests, one probe.

---

## 2. Decision

Adopt the **contract-first spine** (design B) as the Wave 1 skeleton, graft the **agent
capabilities** of design C (describe/audit/deps, `RecalcOptions`, `SheetRenamer`, runtime twins,
the GH-558 writer fix, the `SourceContext` write-back invariant), and stage the **algebra** of
design A (`enum Edit`, `FormulaSupport`, laws, `--print-batch`, registry-derived docs) as the
Wave 2 destination — with the corrections the judges verified (§4). Nothing an agent already
parses changes shape; everything new is additive behind new flags, new verbs, new fields, and new
accepted inputs, except one deliberate, isolated, golden-pinned change: the exit-code table and
the error channel.

### 2.1 Modules and responsibilities (target)

| Module | Owns after the refactor | Never contains |
|---|---|---|
| `xl-core` | domain model; `XLError` with stable codes; `text.Suggest`; runtime twins (`putAt/styleAt/mergeAt/commentAt`, `Workbook.named`); bounded navigation; **Wave 2:** `com.tjclp.xl.ops` — `enum Edit`, `FormatHint`, `Scope`, `StyleOverlay`, `FormulaSupport` (trait + refusing `textOnly`), `EditSchema`, pure interpreter, laws | JSON libraries, IO, the evaluator |
| `xl-ooxml` | OOXML mapping; row/column property emission that is authoritative for modelled attributes (GH-558); **Wave 3:** SST compaction (#567) | CLI vocabulary |
| `xl-evaluator` | `RecalcOptions`/`IterativeMode` and the public `recalculateAfterEdit`/`recalculateUncached` wrappers; `FormulaOps`; `SheetRenamer`; `QualifiedGraph`; `WorkbookAudit`; `WorkbookSummary`; **Wave 2:** `EvalFormulaSupport`, `Recalc.afterEdits` (the promotion of `WriteCommands.scala:996-1229`) | IO, JSON |
| `xl-cats-effect` | `Excel[F]`/`ExcelIO`; **Wave 3:** `SheetSource`, additive sync-facade completions (`readSheet`, `readMetadata`, `modifyR`, `writeChecked`) | recalculating `Excel.write` by default (rejected, §4) |
| `xl` (aggregate) | the scripting prelude: one import, explicit `given FormulaSupport` (Wave 2), prelude probes for every script-visible addition | a second `Excel` object |
| `xl-cli` | the **contract layer**: `CliIO`, `Cli`, `Argv`, `CliError`/`ErrorCode`/`ExitCodes`/`Diagnostics`/`Warning`, `Outcome`/`Payload`/`Render`, `Resolve`, `OpSpec`/`OpRegistry`/`ScopedOp`, `FunctionDoc`/`Schema`; the golden corpus; **Wave 2:** `CommandSpec` registry, verb lowering, `--print-batch` | spreadsheet semantics that scripts cannot reach (moves to core/evaluator in Wave 2, verb bodies become forwarders) |

### 2.2 The CLI shell: `CliIO`, `Cli`, `Argv`

```scala
package com.tjclp.xl.cli

/** The only place bytes leave the process. */
final case class CliIO(
  out: String => IO[Unit],   // stdout: results only
  err: String => IO[Unit],   // stderr: diagnostics (errors, warnings)
  stdin: IO[String]          // `batch -`, later `putf <ref> -`
)
object CliIO:
  val system: CliIO
  /** Capturing sink for tests: the sink and an action yielding (stdout, stderr). */
  def capturing(stdin: String = ""): IO[(CliIO, IO[(String, String)])]

object Main extends IOApp:                       // was CommandIOApp; its run is final
  override def runtimeConfig: IORuntimeConfig    // unchanged (GH-519)
  def main: Opts[IO[ExitCode]]                   // KEPT — MainSpec/InPlaceSpec/LintCommandSpec parse Command("xl","test")(Main.main)
  def run(args: List[String]): IO[ExitCode] = Cli.run(args, CliIO.system)

object Cli:
  def program(io: CliIO): Opts[IO[ExitCode]]
  /**
   * hoist globals → parse → --help (stderr, 0 — decline's CommandIOApp channel, preserved; the
   * help goldens pin it) | usage error (stderr, 2) | run.
   */
  def run(args: List[String], io: CliIO): IO[ExitCode]

object Argv:
  /** Global flag names and whether each consumes a value. --strict is hoisted unless the verb is `view`. */
  val globals: Map[String, Boolean]
  /** Pure: move recognised globals (with their values) in front of the verb; stop at `--`; never reorder other tokens. */
  def hoist(args: List[String]): List[String]
  def verbOf(args: List[String]): Option[String]
```

Every `val *Cmd`, `Main.executeCommand`, `Main.runWithOutput`, `Main.CommandOutcome`,
`Main.requireOutput`, `Main.replaceAtomically`, `Main.atomicMoveOrFallback`,
`Main.renderErrorMessage`, `Main.runDiff`, `Main.runLint`, `Main.resolveLintFile` keep their
current names and signatures (specs reference them); they may forward.

### 2.3 Errors, exit codes, diagnostics

```scala
package com.tjclp.xl.cli.contract

final case class Location(file: Option[String], sheet: Option[String], ref: Option[String], opIndex: Option[Int])

final case class CliError(
  code: String,                         // XLError.code for domain failures; ErrorCode.* for CLI-only ones
  message: String,
  hint: Option[String] = None,
  candidates: Vector[String] = Vector.empty,
  location: Option[Location] = None,
  cause: Option[XLError] = None
):
  def exitCode: ExitCode = ExitCodes.forCode(code)

/** Carrier through IO. getMessage == error.message, so every `getMessage.contains` assertion holds. */
final class CliException(val error: CliError) extends Exception(error.message) with NoStackTrace

object CliError:
  def fromXLError(e: XLError, at: Option[Location]): CliError   // code = e.code, hint = e.hint, candidates = e.candidates
  def fromThrowable(t: Throwable): CliError                     // CliException → its error; StrictFailure → RECALC_GATE; XLException → fromXLError; NoSuchFileException → IO_READ; else INTERNAL
  def usage(message: String, hint: Option[String]): CliError

object ErrorCode:                       // CLI-only codes; domain codes come from XLError.code
  val USAGE, UNKNOWN_VERB, OUTPUT_REQUIRED, UNSUPPORTED_IN_STREAM, BATCH_JSON_INVALID,
      BATCH_OP_UNKNOWN, BATCH_OP_INVALID, BATCH_OP_FAILED, RASTERIZER_UNAVAILABLE,
      IO_READ, IO_WRITE, RECALC_GATE, DIFFERENCES_FOUND, LINT_FINDINGS, AUDIT_FINDINGS, INTERNAL: String
  val all: Vector[String]               // every CLI code + every XLError case code (from XLError.codes), unique, SCREAMING_SNAKE

object ExitCodes:
  val ok: ExitCode = ExitCode.Success   // 0
  val signal: ExitCode = ExitCode(1)    // completed with findings or a failed gate
  val usage: ExitCode = ExitCode(2)     // the command line is wrong; nothing read, nothing written
  val failed: ExitCode = ExitCode(3)    // the operation failed; nothing written
  def forCode(code: String): ExitCode   // RECALC_GATE/DIFFERENCES_FOUND/LINT_FINDINGS/AUDIT_FINDINGS → 1; USAGE/UNKNOWN_VERB/OUTPUT_REQUIRED/UNSUPPORTED_IN_STREAM/BATCH_JSON_INVALID/BATCH_OP_UNKNOWN/BATCH_OP_INVALID → 2; else 3

final case class Warning(code: String, message: String, location: Option[Location] = None)
object WarningCode:
  val READER_WARNING, TRUNCATED, HIDDEN_OMITTED, UNKNOWN_PROPERTY, FORMAT_HINT_IGNORED,
      STREAM_BACKEND_ONLY, RECALC_ERRORS, SHEET_AUTOSELECTED, FLAG_IGNORED: String

object Diagnostics:
  def render(err: CliError): String            // text form, see below
  def report(err: CliError, io: CliIO): IO[Unit]
```

**Exit-code table** (printed by `xl --help`, `xl schema --json`, cli.md, SKILL.md):

| exit | meaning | examples | file written? |
|---|---|---|---|
| 0 | ok | | as requested |
| 1 | completed with findings / a gate — **never a failure** | `diff` differs, `lint` findings, `audit --fail-on-findings`, `--strict` gate | `-o`: yes (today's GH-496 rule); `-i`: no |
| 2 | usage — the command line is wrong | unknown verb, flag not accepted, `-o` missing, `-i` with `-o`, `--stream` on an unsupported verb, batch JSON not an array, unknown or malformed batch op | no (nothing read) |
| 3 | failed — the operation could not complete | sheet not found, invalid ref, formula parse error, count mismatch, batch op failed at apply time, security limit, I/O | no |

Today: 1 = crash/usage/strict/diff/lint, 2 = diff/lint runtime error only. Every `1` that means
"gate or findings" is kept (the CI-lane pins at `MainSpec`, `InPlaceSpec:132-136`,
`LintCommandSpec:122,134,152`, `DiffCommandSpec:265` stay green); usage moves to 2 and runtime
failure to 3. `OUTPUT_REQUIRED` is raised inside `executeCommand` today (`Main.scala:2838-2847`),
i.e. after the workbook is read; it exits 2 from Wave 1 on and moves *before* the read with the
Wave 2 registry (`Needs` validation).

**Channels.** Results go to stdout; diagnostics go to stderr. In text mode the first stderr line
is today's exact text `Error: <message>` (nothing is lost for humans or for
`renderErrorMessage`), followed by indented `code: <CODE>`, `did you mean: <a>, <b>` (when
`candidates` is non-empty; ranked by `Suggest.closest`), and `hint: <text>`. Warnings are
`Warning[<CODE>]: <message>` lines on stderr. In `--json` mode the envelope (§2.4) — success *or*
failure — is the only thing on stdout, and stderr carries the one-line `Error: <message>` so a
human tailing a log still sees it. Usage errors detected before dispatch also produce the envelope
when `--json` is among the raw arguments.

### 2.4 Outcome, Payload, Render — the envelope

```scala
package com.tjclp.xl.cli.contract

enum Payload:
  /** Legacy bridge for prose verbs: text as today, plus the two facts every write knows. */
  case Text(text: String, saved: Option[String], written: Boolean)
  /** Verbs that already emit JSON (view/filter/diff/lint) pass their payload through unchanged; typed verbs build ujson directly. */
  case Json(value: ujson.Value)

final case class Outcome(
  verb: String,
  payload: Option[Payload],
  warnings: Vector[Warning],
  error: Option[CliError],
  exitCode: ExitCode,
  outputComplete: Boolean          // carried over from CommandOutcome: runStagedOutput's commit rule is unchanged
)

final case class Rendered(stdout: String, stderr: String)
object Render:
  def text(o: Outcome): Rendered                    // today's bytes on stdout; diagnostics on stderr
  def json(o: Outcome, version: String): Rendered   // the envelope on stdout; one-line Error: on stderr
```

The envelope has exactly seven keys, identical for every verb, success or failure:

```json
{ "ok": true,  "exitCode": 0, "verb": "view", "version": "0.20.0",
  "data": { "sheet": "Data", "range": "A1:C3", "rows": [ … ] },
  "warnings": [ { "code": "TRUNCATED", "message": "Showing 50 of 120 rows" } ],
  "error": null }

{ "ok": false, "exitCode": 3, "verb": "put", "version": "0.20.0",
  "data": null, "warnings": [],
  "error": { "code": "SHEET_NOT_FOUND", "message": "Sheet not found: 'Sales'",
             "hint": "use -s <name> or a qualified ref like 'Data'!A1",
             "candidates": ["Sales ", "Data"],
             "location": { "file": "in.xlsx", "sheet": "Sales", "ref": null, "opIndex": null } } }
```

`--json` is a global flag; `--format` chooses the payload. **`--format json` keeps printing the bare
payload forever** (`{sheet, range, rows}` for `view`; the existing shapes for `filter`, `diff`,
`lint`) — the xl-agent grader and field scripts piping to `jq` depend on it. Under `--json`,
`view`/`filter`/`diff`/`lint` default to that JSON payload when no `--format` is given, so
`view --json` yields `data` equal to what `view --format json` prints bare; an explicit text
`--format` under `--json` rides inside as `data.text`. Prose verbs yield
`data: {"text": "...", "saved": "out.xlsx", "written": true}` until they are typed (Wave 2's
`Written` payload with per-edit before/after). Truncation is reported inside the verb payload
(`view` already emits `truncated`/`totalRows`) and as a `TRUNCATED` warning; the envelope itself
does not grow an eighth key. `schema --json` publishes the envelope's JSON Schema.

### 2.5 `Resolve` — ONE sheet rule

```scala
package com.tjclp.xl.cli.helpers

object Resolve:
  enum Target derives CanEqual:
    case Cell(ref: ARef)
    case Range(range: CellRange)
  final case class Resolved(sheet: Sheet, target: Target, viaQualifiedRef: Boolean)

  /** THE rule: 1) a qualified ref wins; 2) else -s; 3) else the only sheet of a single-sheet book; 4) else SHEET_REQUIRED. */
  def target(wb: Workbook, sheetFlag: Option[String], refStr: String, verb: String): Either[CliError, Resolved]
  def sheet(wb: Workbook, sheetFlag: Option[String], verb: String): Either[CliError, Sheet]
  /** The same rule over metadata, for the streaming paths (a name, no Sheet in memory). */
  def sheetName(meta: LightMetadata, sheetFlag: Option[String], qualified: Option[SheetName], verb: String): Either[CliError, SheetName]
  /** The same rule for one batch op: qualified ref > op `sheet` key > -s > single sheet; a qualified ref that disagrees with `sheet` is BATCH_OP_INVALID. */
  def forOp(wb: Workbook, defaultSheet: Option[SheetName], opSheet: Option[String], qualified: Option[SheetName], index: Int, op: String): Either[CliError, SheetName]
```

The rule, stated once and enforced everywhere (verbs, batch ops, `--stream` reads and writes):

1. A sheet-qualified ref (`'Q1 Report'!A1:D9`) names the sheet.
2. Otherwise `-s/--sheet` (for a batch op, its `sheet` key comes before `-s`).
3. Otherwise, if the workbook has exactly one sheet, that sheet (a `SHEET_AUTOSELECTED` warning is
   emitted in `--json` mode only).
4. Otherwise `SHEET_REQUIRED` — `XLError.SheetRequired(context, available)`, exit 3 (the file has
   been read), `candidates` = the sheet names.

`AllSheets` verbs ignore steps 3–4: `search` (without `-s`), `sheets`, `names`, `diff` (without
`-s`), `lint`, `describe`, `audit`. Verbs that need no sheet at all: `recalc`, `new`, `functions`,
`rasterizers`, `schema`; `eval`/`evala` take an optional file and an optional sheet. The streaming
first-sheet default and the streaming "Multiple sheets found" refusal are both replaced by the rule
(single-sheet auto-select stays; multi-sheet without a sheet is `SHEET_REQUIRED`). Qualified refs
become legal under `--stream` writes. `SheetResolver.{resolveSheet, requireSheet, findSheet,
resolveRef}` keep their signatures and forward, preserving the two message texts specs pin
(`"Sheet not found: X. Available: …"`, `"… requires --sheet or qualified ref (e.g., Sheet1!A1).
Available sheets: …"`).

### 2.6 Batch: `OpSpec`, `OpRegistry`, `ScopedOp`

```scala
package com.tjclp.xl.cli.batch

enum FieldKind derives CanEqual:
  case Str, Num, Int, Bool, Ref, Range, Sheet, Color, NumFmt, Formula, Value, Values, Formulas
  case Enum(values: Vector[String])
  case Arr(of: FieldKind)

final case class Field(name: String, kind: FieldKind, required: Boolean, aliases: Vector[String] = Vector.empty, doc: String = "") derives CanEqual

final case class OpSpec(
  name: String,                    // canonical kebab op name ("put", "remove-comment", …)
  aliases: Vector[String],
  fields: Vector[Field],           // camelCase = today's JSON keys; kebab spellings accepted as aliases
  oneOf: Vector[Set[String]],      // e.g. put: {value} | {values}
  sheetScoped: Boolean,            // accepts the `sheet` key
  cellMutating: Boolean,           // mirrors WriteCommands.isCellMutating
  structural: Boolean,
  needsFormula: Boolean,
  streamable: Boolean,             // mirrors buildStreamingBatchPatches' supported set (dragging putf included; the streaming putf VERB is the one that does not drag)
  cliVerb: Option[String],
  since: String,
  doc: String,
  example: ujson.Obj
) derives CanEqual

object OpRegistry:
  val all: Vector[OpSpec]                          // 32 today; the seed of Wave 2's EditSchema
  def find(name: String): Option[OpSpec]           // exact, alias, kebab/camel-insensitive
  def jsonSchema(version: String): ujson.Value     // draft 2020-12, oneOf per op, enums inlined, x-aliases/x-streamable/x-cliVerb/x-example
  def helpText: String                             // replaces the hand-written 17-op list in batchHelp

/** Every op gets a sheet without touching the 35 BatchOp cases the specs construct. */
final case class ScopedOp(op: BatchParser.BatchOp, sheet: Option[String], index: Int)
// BatchParser.ParseResult(ops, warnings) gains `scoped: Vector[ScopedOp] = Vector.empty`; ops == scoped.map(_.op)
// BatchParser.applyScoped(wb, defaultSheetOpt, scoped, recalcDependents): IO[Workbook]  — applyBatchOperations keeps its signature and delegates
```

Rules: unknown op → `BATCH_OP_UNKNOWN` (exit 2) with `did you mean`; unknown property → a
`UNKNOWN_PROPERTY` warning (stderr in text mode, `warnings[]` in `--json`) — the default does
**not** flip to an error in 0.20; `--strict-input` (Wave 2) promotes it. Aliases are honoured
silently and never trigger the warning (`format`↔`numFormat`, `from`↔`anchor`, `target`↔`url`,
`align`↔`halign`, kebab↔camel). Every apply-time failure is `BATCH_OP_FAILED` with
`location.opIndex` (1-based, as today's `Object N` text). A `rename-sheet` of the batch's default
sheet updates the default for the ops that follow. Under `--stream`, ops whose spec is not
`streamable`, or whose resolved sheet differs from the streamed worksheet, are refused **by index
before the zip is opened** (`UNSUPPORTED_IN_STREAM`, exit 2). GH-560: an explicit `format` on
`put` replaces the target's `numFmt` (the `putf` path at `BatchParser.scala:1319-1324`); the
inferred hint keeps the General-only rule.

### 2.7 `XLError`: additive extensions and new cases (xl-core)

```scala
package com.tjclp.xl.error

enum XLError derives CanEqual:               // CanEqual is additive
  // …the 30 existing cases, unchanged…
  case EditFailed(index: Int, op: String, cause: XLError)                  // message: s"op $index ($op): ${cause.message}"
  case UnsupportedCapability(op: String, capability: String, hint: String) // message: s"$op requires $capability: $hint"
  case SheetRequired(context: String, available: Vector[String])          // message: s"$context requires a sheet: pass -s <name> or qualify the ref (Sheet!A1). Available: …"

object XLError:
  extension (error: XLError)
    def message: String                        // existing
    def code: String                           // SCREAMING_SNAKE of the case's simple name; EditFailed → the cause's code; Other → OTHER
    def hint: Option[String]                   // per-case one-liners (SheetNotFound → "list sheets with `xl -f <file> sheets`", ValueCountMismatch → "provide exactly N values, or 1 to fill", …)
    def candidates: Vector[String]             // SheetRequired → available; EditFailed → cause.candidates; else empty
    def root: XLError                          // unwraps EditFailed
    def opIndex: Option[Int]                   // EditFailed → Some(index)
  val codes: Vector[String]                    // one code per case, unique; feeds ErrorCode.all and schema --json

package com.tjclp.xl.text
object Suggest:
  /** Case-insensitive Levenshtein ranking; distance ≤ max(2, len/3); pure. */
  def closest(input: String, candidates: Iterable[String], limit: Int = 3): Vector[String]
```

Adding fields to existing cases is source-breaking for downstream matches, so the new information
lives in new cases and extension accessors. Adding cases makes downstream exhaustive matches warn
(the repo has no `-Xfatal-warnings`); this is called out in CHANGELOG. `CliError.fromXLError` is a
projection (`code = e.code`) — one vocabulary, no hand-written mapping table.

### 2.8 Recalculation options and the public after-edit seam (xl-evaluator)

```scala
package com.tjclp.xl.formula.eval

enum IterativeMode derives CanEqual:
  case Off
  case FromCalcPr                      // honour <calcPr iterate>, Off when absent — what `xl recalc` does today
  case Force(calc: IterativeCalc)

final case class RecalcOptions(
  clock: Clock = Clock.system,
  rng: Rng = Rng.system,
  iterative: IterativeMode = IterativeMode.FromCalcPr,
  parallelism: Int = 1,
  seedTables: Boolean = false
) derives CanEqual
object RecalcOptions:
  val default: RecalcOptions = RecalcOptions()

// Additive members of WorkbookEvaluator's existing `extension (wb: Workbook)` block (no default arguments — wildcard-export landmine):
def recalculate(options: RecalcOptions): RecalcResult
/** Public form of DependentRecalculation.recalculateAfterEdit (private[xl]); wraps, never re-derives the cone. */
def recalculateAfterEdit(sheet: SheetName, modified: Set[ARef], options: RecalcOptions): RecalcResult
/** Computes ONLY formula cells with no cache; cached cells stay byte-identical (GH-468 doctrine). */
def recalculateUncached(options: RecalcOptions): RecalcResult

// Additive member of RecalcResult (after toEither):
def summary: String                     // byte-identical to WriteCommands.formatRecalcSummary's text, so CLI and scripts print the same
```

The 23-overload `@targetName` lattice stays as forwarders in 0.20 and is deprecated in 0.21 once
the skill has moved. `Recalc.afterEdits` (the promotion of `recalcHonoringCalcPr`, `changedRefs`,
`dirtyCone`, `applyConeCaches`, `scopeToCone`, `strictGate` from `WriteCommands.scala:996-1229`)
is Wave 2 and builds on these seams.

### 2.9 `FormulaOps`, `SheetRenamer` (xl-evaluator)

```scala
object FormulaShifter:                 // additive
  def renameSheet[A](expr: TExpr[A], from: SheetName, to: SheetName): TExpr[A]
  /** referencesSheet ∪ SheetNameRef — referencesSheet deliberately returns false for SheetNameRef and must keep doing so for StructuralEditor. */
  def mentionsSheet(expr: TExpr[?], sheet: String): Boolean

package com.tjclp.xl.formula.printer
object FormulaOps:                     // string in, string out — the six CLI parse→shift→print copies collapse onto it in Wave 2
  def renameSheet(text: String, from: SheetName, to: SheetName): XLResult[String]
  def mentionsSheet(text: String, sheet: SheetName): Boolean
  def shift(text: String, colDelta: Int, rowDelta: Int): XLResult[String]   // pins today's clamp at the sheet edge (FormulaShifter.scala:207-213)

package com.tjclp.xl.formula.eval
object SheetRenamer:
  /** Workbook.rename, then rewrite every reference to `from` — cell formulas on every sheet (caches preserved), defined names, CF and DV formulas. Refuses (Left) before mutating when a text that mentions the sheet cannot be parsed. Writes changed sheets back with Workbook.put. */
  def rename(wb: Workbook, from: SheetName, to: SheetName): XLResult[Workbook]
  def references(wb: Workbook, sheet: SheetName): Vector[DependencyGraph.QualifiedRef]
```

`Workbook.rename` keeps its 0.19 meaning (tab, chart refs, tracker) and its Scaladoc says it is
formula-blind, like `Sheet.insertRows` vs `StructuralEditor`. The `rename-sheet` verb and batch op
call `SheetRenamer.rename`.

### 2.10 `QualifiedGraph`, `WorkbookAudit`, `WorkbookSummary` (xl-evaluator — shared by CLI and prelude)

```scala
package com.tjclp.xl.formula.graph
final case class QualifiedGraph(
  dependencies: Map[DependencyGraph.QualifiedRef, Set[DependencyGraph.QualifiedRef]],
  dependents:   Map[DependencyGraph.QualifiedRef, Set[DependencyGraph.QualifiedRef]]
):
  def precedents(q: DependencyGraph.QualifiedRef, depth: Int): Vector[Vector[DependencyGraph.QualifiedRef]]  // layer k = exactly k hops; depth 0 = unbounded
  def dependents(q: DependencyGraph.QualifiedRef, depth: Int): Vector[Vector[DependencyGraph.QualifiedRef]]
  def sccs: Vector[DependencyGraph.Scc]
object QualifiedGraph:
  def of(wb: Workbook): QualifiedGraph   // fromWorkbookFormulaGraph (private[formula]) + fromWorkbookDependencyIndex (private[xl]); bounded/symbolic, never expands A:A

package com.tjclp.xl.formula.eval
final case class WorkbookAudit(
  errorCells: Vector[(DependencyGraph.QualifiedRef, CellError)],
  uncachedFormulas: Vector[DependencyGraph.QualifiedRef],
  unparseable: Vector[(DependencyGraph.QualifiedRef, String)],
  volatile: Vector[DependencyGraph.QualifiedRef],        // TODAY/NOW/RAND/RANDBETWEEN by name (no FunctionFlags.volatile exists yet)
  dynamic: Vector[DependencyGraph.QualifiedRef],         // DependencyGraph.dynamicCells
  cycles: Vector[DependencyGraph.Scc],
  externalRefs: Vector[DependencyGraph.QualifiedRef],
  unresolvedReaders: Vector[DependencyGraph.QualifiedRef], // DependencyGraph.unresolvedReaders
  calcPr: Option[CalcPr]
) derives CanEqual:
  def isClean: Boolean
object WorkbookAudit:
  def of(wb: Workbook): WorkbookAudit

final case class SheetSummary(name: SheetName, index: Int, state: Option[String], dimension: Option[CellRange],
  cellCount: Int, formulaCount: Int, uncachedFormulas: Int, mergedRanges: Int, comments: Int, hyperlinks: Int,
  freeze: Option[FreezePane], tabColor: Option[Color], autoFilter: Option[CellRange], tables: Vector[String],
  charts: Int, pictures: Int, conditionalFormats: Int, dataValidations: Int, hiddenRows: Int, hiddenCols: Int) derives CanEqual
final case class WorkbookSummary(sheets: Vector[SheetSummary], definedNames: Vector[DefinedName], date1904: Boolean, calcPr: Option[CalcPr]) derives CanEqual
object WorkbookSummary:
  def of(wb: Workbook): WorkbookSummary   // xl-evaluator, NOT the aggregate: xl-cli cannot see `xl` (xl-cli/package.mill:8)
```

CLI verbs: `describe [--full]` (metadata-only from `LightMetadata` without `--full`, so it works
under `--stream`; `--full` loads the book and prints `WorkbookSummary`), `audit [-s S]
[--fail-on-findings]` (exit 1 = `AUDIT_FINDINGS` when asked), `deps <ref> [--direction
precedents|dependents|both] [--depth N|all]`. `cell` uses `QualifiedGraph.of(wb)` instead of the
O(workbook) reverse fold at `ReadCommands.scala:259-265`. Prelude: `wb.describe: WorkbookSummary`,
`wb.audit: WorkbookAudit`, `QualifiedGraph.of(wb)`.

### 2.11 Runtime twins and bounded navigation (xl-core)

```scala
// Sheet — explicit runtime forms beside the transparent-inline unions; all parse via RefType.parseToXLError and reject qualified refs
def putAt[A: CellWriter](ref: String, value: A): XLResult[Sheet]
def putAt[A: CellWriter](ref: String, value: A, style: CellStyle): XLResult[Sheet]
def styleAt(ref: String, style: CellStyle): XLResult[Sheet]        // cell or range
def mergeAt(range: String): XLResult[Sheet]
def commentAt(ref: String, comment: Comment): XLResult[Sheet]
// Workbook companion — twins of the transparent-inline apply(String…)
def named(name: String): XLResult[Workbook]
def named(first: String, second: String, rest: String*): XLResult[Workbook]
// ARef companion extension block (non-inline, issue #252)
def tryShift(colOffset: Int, rowOffset: Int): Option[ARef]        // None instead of minting A0 / column -1 / beyond XFD1048576
def tryDown(n: Int): Option[ARef]
def tryRight(n: Int): Option[ARef]
def clampShift(colOffset: Int, rowOffset: Int): ARef
// CellRange
def rows: Iterator[CellRange]
def columns: Iterator[CellRange]
def row(i: Int): Option[CellRange]
def column(i: Int): Option[CellRange]
// CellStyle
def withUnderline(u: Underline): CellStyle                        // = withFont(font.withUnderline(u)) (#465)
```

The transparent forms are not deprecated in 0.20 (they are the literal-time safety net); their
Scaladoc names the twin, and the xl-scripting skill stops teaching them for computed strings.

### 2.12 Wave 2 destination: `enum Edit` in xl-core

```scala
package com.tjclp.xl.ops

/** Where a value's number format came from. Only Explicit may override a non-General numFmt (GH-560). */
enum FormatHint derives CanEqual:
  case Inferred(fmt: NumFmt)     // codec inference (LocalDate → Date): applied only when the existing numFmt is General
  case Explicit(fmt: NumFmt)     // the agent asked for it: replaces

final case class Loc(sheet: Option[SheetName], ref: ARef) derives CanEqual
final case class Area(sheet: Option[SheetName], range: CellRange) derives CanEqual
final case class Scope(defaultSheet: Option[SheetName]) derives CanEqual

enum Edit derives CanEqual:
  case Put(at: Loc, value: CellValue, format: Option[FormatHint])       // constructed through Edit.put[A: CellWriter](at, value), which lifts the codec hint as Inferred — a bare CellValue never loses inference by accident
  case PutValues(at: Area, values: Vector[CellValue], format: Option[FormatHint])
  case PutFormula(at: Loc, formula: String, format: Option[FormatHint])
  case DragFormula(at: Area, formula: String, anchor: ARef, format: Option[FormatHint])
  case Fill(source: Area, target: CellRange, direction: Edit.FillDir)
  case Copy(source: Area, target: Loc, valuesOnly: Boolean)
  case Sort(at: Area, keys: Vector[Edit.SortKeySpec], hasHeader: Boolean)
  case Clear(at: Area, what: ClearWhat)
  case Style(at: Area, overlay: StyleOverlay, mode: StyleMode)
  case Merge(at: Area); case Unmerge(at: Area)
  case RenameSheet(from: SheetName, to: SheetName)
  case InsertRows(sheet: Option[SheetName], at: Row, count: Int)       // … and the remaining ~35 cases mirroring the 32 ops + the 15 verbs without a twin
  // enums that would collide with xl-cli's FillDirection/SortDirection/SortMode are nested under Edit (FillDir, SortDir, SortKeySpec) and are NOT exported through `api`

/** Formula-aware work the core interpreter delegates. Implemented by xl-evaluator's EvalFormulaSupport. */
trait FormulaSupport:
  def validate(formula: String): XLResult[Unit]
  def shift(formula: String, colDelta: Int, rowDelta: Int): XLResult[String]
  def insertRows(wb: Workbook, sheet: SheetName, at: Row, count: Int): XLResult[Workbook]   // + deleteRows/insertCols/deleteCols
  def renameSheet(wb: Workbook, from: SheetName, to: SheetName): XLResult[Workbook]
object FormulaSupport:
  /** REFUSES: shift/structural/rename are Left(UnsupportedCapability(op, "formula support", "import the evaluator")); validate accepts non-empty text. No `given default` anywhere in xl-core. */
  val textOnly: FormulaSupport

object Edit:
  def validate(edit: Edit)(using FormulaSupport): XLResult[Unit]
  def plan(wb: Workbook, edits: Vector[Edit], scope: Scope)(using FormulaSupport): XLResult[Vector[Planned]]
  def applyAll(wb: Workbook, edits: Vector[Edit], scope: Scope)(using FormulaSupport): XLResult[Applied]  // fail-fast, all-or-nothing; every sheet written back through Workbook.put/update
  def lower(edit: Edit, existing: Sheet): Option[Patch]                                                   // the Sheet-local subset lowers to the law-tested Patch kernel
```

The prelude supplies `given FormulaSupport = EvalFormulaSupport` explicitly (no LowPriority
trick — `scripting.scala:31-44` shows it does not survive export forwarding); the CLI passes it
explicitly. `StyleOverlay` carries `Option` fields (so un-bold is expressible) with border colour
as per-side data (`BorderSideOverlay(style, color)`), never "colour applies to every side being
set" — that rule broke the monoid-action law. `Patch` stays; `Patch.toEdits(p, sheet)` and
`Edit.lower` are connected by coherence laws (§6). `EditSchema` (JSON-free) replaces `OpRegistry`
as the single metadata table; batch JSON becomes the codec of `Vector[Edit]`; every mutating verb
lowers to `Vector[Edit]` and can print it with `--print-batch`.

### 2.13 Docs generated from code, gated in CI

`xl schema --json` → `{version, exitCodes[], errorCodes[], warningCodes[], globals[], verbs[{path,
summary, needs, exit, batchTwin, since}], batchOps: <OpRegistry.jsonSchema>, functions:
<FunctionDoc list>}`; `xl batch --schema` prints `batchOps` alone; `xl functions --json` prints
`functions` alone (115 registry functions + `LET` flagged `specialForm: true`). Wave 1's verb list
is a table asserted equal to the subcommand names decline renders in `xl --help`; Wave 2 derives
it from the `CommandSpec` registry that also builds `Main.main`.
`docs/reference/generated/{cli-verbs,batch-ops,functions,exit-codes,error-codes}.md` are rendered
by `DocsGenSpec` and compared with the checked-in files (`XL_UPDATE_DOCS=1` rewrites; a drift is a
test failure, so CI is the gate). `cli.md`, `SKILL.md` and `CLAUDE.md` link to the generated files
instead of copying counts. `plugin/.claude-plugin/plugin.json` must equal `build.mill`'s version
(shell check in `ci.yml`; `/release-prep` keeps it).

---

## 3. Invariants

Stated once, tested forever (Judge 3 §5.1, adopted verbatim):

1. **Every interpreter writes sheets back through `Workbook.put`/`update`** so `modifiedSheets` is
   right, with a property test that goes through a real file — a `SheetRenamer`, an `Edit`
   interpreter, or a recalc that assembles the result with `copy(sheets = …)` renames or edits in
   memory while the surgical writer copies the stale source part.
2. **`--stream` honours an op with identical semantics or refuses it by index before any byte is
   written** (`UNSUPPORTED_IN_STREAM`). It never recalculates and never degrades (no "same formula
   in every cell" instead of a drag).
3. **`--format json` is the payload format and `--json` the envelope, forever.**
4. **A hint on a value is `Inferred | Explicit`, and only `Explicit` overrides a non-General
   numFmt** (GH-560). Inferred hints keep `Sheet.putSingle`'s General-only rule on every surface.
5. **Exit `1` is reserved for "completed with findings/gate", never for failure.**

Two further rules the panel added: no new `transparent inline` union entry points and no default
arguments on extension methods reached through wildcard exports; every prelude-visible addition
gets a probe in `xl/test/src/xlprelude/ScriptingPreludeTest.scala`.

---

## 4. Alternatives considered

**Algebra-first in Wave 1 (design A).** Correct destination; rejected as the *first* move.
Verified flaws: (a) `given default: FormulaSupport = textOnly` in a low-priority trait "wins
without ambiguity" — the repo's own prelude says the opposite and excludes `default` by name on two
hops (`scripting.scala:31-44`), so `summon[FormulaSupport]` in a script would be ambiguous; and
`textOnly` *degrading* (`RenameSheet` renames without rewriting plus a warning) re-creates the
"same op, different semantics by context" class keyed on the classpath — the seam is kept, but
`textOnly` **refuses**. (b) `Edit.Put(at, value: CellValue, format)` "keeps Date inference" —
false: `Sheet.put(ref, value: CellValue)` (`Sheet.scala:395-399`) bypasses `putSingle`; the codec
hint exists only for typed `A` (`:418`, `:496-514`). Hence `FormatHint = Inferred | Explicit`
lifted at the edge by `Edit.put[A: CellWriter]`. (c) `export ops.{FillDirection, SortDirection,
SortMode, …}` from `api` collides with xl-cli's same-package enums (`Command.scala:216,237,242`,
used in `Main.scala` without import) because `api.scala:154` re-exports `api.*` at package level
and every xl-cli file imports `com.tjclp.xl.{*, given}` — a compile break; the enums are nested
under `Edit`. (d) The `StyleOverlay` border rule ("colour applies to every side being set")
contradicts the monoid-action law `(a |+| b).applyTo(s) == b.applyTo(a.applyTo(s))`. (e)
Wave 1 sizing ~2× low; landing the interpreter beside the verbs leaves four implementations of
sort/fill/copy alive for a release. Kept from A: `EditSchema`/`EditSpec` as the shape of
`OpSpec`, `XLError.code/hint/root/opIndex` + `EditFailed` + `Suggest` in xl-core, streaming
refusal by index, `--print-batch`, the seven laws, `Scope` rule 4 (rename updates the default).

**Agent-journey-first (design C).** Best value-per-hour ordering and the only correct GH-558
diagnosis; rejected as the skeleton. Verified flaws: (a) `WorkbookSummary` in the `xl` aggregate
is unreachable from xl-cli (`xl-cli/package.mill:8` depends on xl-cats-effect and xl-evaluator
only) — two `describe` implementations by construction; it lives in xl-evaluator. (b) A prelude
`object Excel` shadowing `io.Excel` whose `write` recalculates and returns `RecalcResult` flips the
most-called script API (46 call sites in probes/docs/examples) with no deprecation window, breaks
`ExcelRecalc`'s five `extension (excel: Excel.type)` overloads, and its "recompute uncached cells
*and their dependents* while never touching cached cells" is self-contradictory; kept instead:
additive `Excel.writeChecked`/`readSheet`/`readMetadata`/`modifyR` (Wave 3) and the skill teaching
`writeRecalculated`. (c) `--format json` implying the envelope breaks the xl-agent grader
(`Evaluator.scala:61-79`). (d) A `Console[IO]` harness captures nothing — `Main` has zero
`Console[` uses and 12 `System.err.println` sites; `CliIO` is the workable version. (e) The JSON
error envelope on stderr with empty stdout leaves an agent piping `| jq` with nothing.
(f) Unconditional `Argv.hoist` of `--strict` breaks `view --eval --strict` (two flags share the
name: `Main.scala:295-301` vs `:320-321`) — hence the `view` exception in `Argv.globals`.

**Contract-first (design B) — adopted, with corrections.** Rejected parts: `Registry.flagsOf`
parsing decline's rendered help as an argv-grammar input (the collision set is hard-coded;
`schema --json`'s per-flag docs wait for typed descriptors in Wave 2); a second Levenshtein in
xl-cli (`Suggest` goes to xl-core now); deferring GH-559 and GH-558 to Wave 2 (both are isolated,
high-value, and scheduled in Wave 1); W2.6's core-side GH-558 diagnosis (wrong module); the
streaming `view --format json` shape change in Wave 1 (deferred to the `SheetSource` wave — one
"Changed" entry per release is the ceiling).

**Cross-cutting rejections.** Making `--format json` envelope-shaped (A, C); flipping unknown
batch keys to exit 2 or `"bold": false` to un-bold in 0.20 (A) — warnings in `warnings[]` plus
`--strict-input` and `StyleOverlay` in Wave 2 are the non-breaking paths; JSON errors on stderr
(C); recalculating `Excel.write` by default (C).

---

## 5. Consequences

**Compatibility and deprecations.**
- Text stdout stays byte-identical for every verb (the golden corpus proves it); the JSON payloads
  of `view/filter/diff/lint --format json` are untouched. New behaviour is reached through
  `--json`, new verbs (`describe`, `audit`, `deps`, `schema`), new flags (`batch --schema`,
  `audit --fail-on-findings`), new fields, and newly accepted inputs (globals after the verb,
  single-sheet auto-select, `sheet` on every batch op, qualified refs under `--stream`).
- The breaking changes in 0.20.0, each led with **Breaking:** in CHANGELOG: the exit table (usage
  1→2, failures 1/2→3); errors and warnings on stderr instead of stdout; streaming reads on a
  multi-sheet book without a sheet exit 3 `SHEET_REQUIRED` instead of silently using the first
  sheet; typed reads (`readTyped`/`readTypedOpt`/`readTypedOr`) return a formula cell's cached
  value (GH-477; `readTypedStrict` keeps the 0.19 rule); and `XLError` gains cases, so exhaustive
  downstream matches warn. `--help` stays on stderr with exit 0 (decline's channel; the help
  goldens pin it).
- Published library API (0.19.3 on Maven Central): otherwise additive — `XLError` extensions,
  `derives CanEqual`, twins, navigation, `RecalcOptions` and three extension methods,
  `FormulaOps`/`SheetRenamer`/`QualifiedGraph`/`WorkbookAudit`/`WorkbookSummary` exported from
  `formulaExports`. Nothing is removed; the `@targetName` lattices and `Workbook.rename` keep their
  meaning; deprecations (0.21) are announced, not shipped, in this wave.
- The 754 xl-cli handler tests keep passing because handler signatures are frozen:
  `WriteCommands.put(wb, sheetOpt, refStr, values, outputPath, config, stream, csvSplit, detect,
  policy)` and the rest of `WriteCommands`/`SheetCommands`/`CellCommands`/`CommentCommands`/
  `ChartCommands`/`ReadCommands`/`FilterCommands`/`DiffCommands`/`LintCommands`/`ImportCommands`
  are untouched or forwarded; `BatchParser.parseBatchJson(json): Either[Exception, ParseResult]`,
  `applyBatchOperations(wb, defaultSheetOpt, ops, recalcDependents)`, `BatchOp` cases,
  `StyleProps`, `ParsedValue` are kept (parse errors become `CliException`, a subclass of
  `Exception` with the same message); `SheetResolver.*` forward with verbatim messages;
  `Main.main`, `Main.putCmd`, `Main.recalcCmd`, `Main.executeCommand`, `Main.runWithOutput`,
  `Main.CommandOutcome` survive. The only planned adjustments are the exit-code pins
  `DiffCommandSpec:278` (2→3), `LintCommandSpec:144` (2→3), `InPlaceSpec:220` (`ExitCode.Error`→2)
  and streaming specs that pinned the first-sheet default or the qualified-ref refusal.

**Performance.** `QualifiedGraph.of` is bounded/symbolic and replaces an O(workbook) reverse fold
per `cell` call; `recalculateAfterEdit` is the already-shipped cone; `describe` without `--full` is
metadata-only; `Argv.hoist` is O(argv). No hot path in xl-core or xl-ooxml changes in Wave 1
except the GH-558 row emission, which is byte-identity-guarded for untouched books.

**Docs generation.** Verb, op, function, exit-code and error-code tables are generated and
CI-diffed; counts stop appearing in prose. Test counts are still recorded by hand in three places
(CLAUDE.md, STATUS.md, testing-guide.md) and taken from the Mill run.

**Process.** Every cluster is TDD in an isolated worktree, ≤2 concurrent, runs Mill as
`MILL_VERSION=1.1.5-jvm ./mill`, formats with `reformatAll __.sources`, keeps handler signatures,
adds probes for script-visible API, and leaves CHANGELOG/roadmap/decisions.md/test-count
bookkeeping to the integrator. The four gates before the PR: `./mill __.compile`, `./mill
__.test`, `scripts/test-examples.sh`, `scripts/verify-skill-snippets.sh --local`.

---

## 6. Laws and invariants to test

- **Golden corpus**: for every pinned case, `CliHarness.run(args*)` reproduces `(exit, stdout,
  stderr)` after normalisation (`<DIR>`, `<TMP>`, `<VERSION>`); a golden diff is a contract change
  and needs a CHANGELOG line.
- **Envelope**: for every verb and both outcomes, the stdout of `--json` parses, has exactly the
  seven keys, `ok == (error == null)`, `exitCode == ExitCodes.forCode(error.code)` on failure,
  `version == BuildInfo.version`; `view --json`'s `data` equals `view --format json`'s payload
  byte-for-byte after `ujson.read`.
- **Exit table**: every code in `ErrorCode.all` maps to exactly one of {1, 2, 3}; every `1` is a
  findings/gate code; the strict/diff/lint pins are unchanged.
- **Sheet rule**: property over `Generators.genSheetName`: a qualified ref wins regardless of `-s`;
  single-sheet books auto-select for every sheet-level verb, batch op and streaming path;
  two-sheet books without a sheet are `SHEET_REQUIRED` with candidates = the names.
- **Argv**: `hoist(hoist(a)) == hoist(a)`; `hoist` never reorders non-global tokens; `--` stops
  hoisting; `view … --strict` is not hoisted; `xl view A1:B2 -f f -s S` and `xl -f f -s S view
  A1:B2` produce identical `(exit, stdout, stderr)`.
- **Registry parity**: `OpRegistry.all.map(_.name).toSet` == the op names the 32-arm parser
  accepts; `cellMutating` agrees with `WriteCommands.isCellMutating` for every `BatchOp`;
  `streamable` agrees with `buildStreamingBatchPatches`' supported set; every `example` validates
  against `jsonSchema` and round-trips through `parseBatchJson`.
- **Rename**: `renameSheet(renameSheet(e, a, b), b, a) == e` for generated expressions; cached
  values are preserved; sheets that never mention the old name are byte-identical on disk;
  read-file → rename → write → re-read shows the rewritten text (the `SourceContext` invariant,
  through a real file, not in memory).
- **Recalc**: `recalculate(RecalcOptions())` ≡ `recalculate()` on an acyclic book;
  `recalculateAfterEdit(sheet, refs, opts)` ≡ `DependentRecalculation.recalculateAfterEdit(wb,
  sheet, refs, opts.clock)`; `recalculateUncached` leaves cached cells byte-identical;
  `RecalcResult.summary` equals the CLI's summary text on a shared fixture.
- **GH-558**: write → read → `insertRowsChecked` → `deleteRowsChecked` → write yields the original
  row set; every `r=` is emitted once; an untouched book round-trips byte-identically.
- **GH-560**: an explicit `format` replaces a `Custom` numFmt; an inferred date hint does not
  override a `Currency` numFmt.
- **Twins**: `Sheet.named(s).flatMap(_.putAt(r, v))` equals the literal form for valid strings and
  is `Left(InvalidReference)` for a qualified `r`; `tryShift(dc, dr).flatMap(_.tryShift(-dc, -dr))
  == Some(ref)` when in bounds and `None` past A1/XFD1048576.
- **Wave 2 algebra laws** (acceptance for `enum Edit`): identity (`applyAll(wb, Vector.empty) ==
  Right(Applied(wb, …))`); fold (`applyAll(a ++ b) == applyAll(a).flatMap(applyAll(b))`);
  idempotence for the declared class; lowering coherence (`lower(e) == Some(p)` ⇒ applying `e`
  equals `Patch.applyPatch(_, p)`); desugar coherence (`Patch.toEdits(p, sheet)` applied equals
  `Patch.applyPatch`); determinism; all-or-nothing; the 30 `PatchSpec` laws unchanged; every Wave 1
  golden unchanged.

---

## 7. Open questions

1. **`--strict` promotion of warnings.** Should `--strict` (write verbs) also promote
   `UNKNOWN_PROPERTY`/`FORMAT_HINT_IGNORED` to exit 1, or is a separate `--strict-input` (Wave 2)
   cleaner? Recommendation: separate flag; `--strict` stays a recalculation gate.
2. **`audit` under the global `--strict`.** Wave 1 uses a verb-local `--fail-on-findings` because
   `--strict` is only parsed in the write group; the Wave 2 registry can unify them.
3. **`putf <ref> @file` / `-` input channels.** Deferred from Wave 1 (their only home is
   `Main.executeCommand`, which the Wave 1 pairing keeps out of the same group as
   `resolve-and-argv`); land with the registry in Wave 2 or as a standalone issue.
4. **Streaming `view --format json` typed rows.** Never documented as stable; deferred to the
   `SheetSource` wave to keep one "Changed" entry per release.
5. **`FunctionFlags.volatile`.** `WorkbookAudit` uses a name set in Wave 1; adding the flag
   touches `FunctionSpecs*` (Zinc/macro gotcha) and is a Wave 2 chore.
6. **Drag edge policy.** `FormulaShifter.shift` clamps at A1; `Left(OutOfBounds)` would change
   `putf` drag at the sheet edge. Pinned as-is until a maintainer decides.
7. **`recalc` as a batch directive vs an op** in the Wave 2 codec — decide after field use.
8. **`Excel`/`Excel[F]` homonym** (sync object and F-polymorphic trait share a name): rename or
   alias in 1.0.
