# Agent-first refactor — wave plan

Design record: `docs/design/agent-first-architecture.md` (ADR-017). Branch:
`claude/xl-refactor-agents-c1amtj`, rebased on `origin/main` `a9cdb73` (PRs #569–#572 merged) and
carrying `93ab913` (proxy-safe bootstrap) and `0d53e76` (GH-477: `Cell.effectiveValue`,
`Cell.isUncachedFormula`, `CellReader.readStrict`, `Sheet.readTypedStrict`,
`CellCodecCachedFormulaSpec` with 25 tests, one prelude probe — **GH-477 is done; nothing below
plans it again**). Test count on the branch before those commits: 5,625; the branch adds 26.

## Rules for every cluster

- **TDD.** The "Tests first" list is written first and fails; then the implementation; then the
  module suite. A cluster that changes an existing test explains why in the commit body.
- **Mill invocation.** Always `MILL_VERSION=1.1.5-jvm ./mill -j 2 <target>` with a 600000 ms
  timeout: `<module>.compile`, `<module>.test`, and `<module>.test.testOnly <Spec>` while
  iterating; `__.compile` once before the commit. Never `__.test` in a worktree (the integrator runs
  it at integration). After editing any `FunctionSpecs*` trait (no Wave 1 cluster does) run
  `./mill clean xl-evaluator.compile` first — the Zinc/macro gotcha in CLAUDE.md.
- **Format pair before every commit.**
  `MILL_VERSION=1.1.5-jvm ./mill mill.scalalib.scalafmt.ScalafmtModule/reformatAll __.sources`
  then `…/checkFormatAll __.sources` (CI checks all sources; `__.reformat` skips tests).
- **Owned by the integrator, never edited by an implementer:** `CHANGELOG.md`,
  `docs/plan/roadmap.md`, `docs/design/decisions.md`, `docs/STATUS.md`, the test-count lines in
  `CLAUDE.md` and `docs/reference/testing-guide.md`. Each cluster's "Integrator notes" carries the
  CHANGELOG bullet verbatim and the bookkeeping the integrator applies.
- **Handler signatures are frozen.** `WriteCommands.*`, `SheetCommands.*`, `CellCommands.*`,
  `CommentCommands.*`, `ChartCommands.*`, `ReadCommands.*`, `StreamingReadCommands.*`,
  `StreamingWriteCommands.*`, `FilterCommands.*`, `DiffCommands.*`, `LintCommands.*`,
  `ImportCommands.*`, `BatchParser.{parseBatchJson, parseBatchOperations, applyBatchOperations,
  readBatchInput, formatSummary}`, `BatchOp` cases, `StyleProps`, `ParsedValue`,
  `SheetResolver.*`, `Main.{main, putCmd, recalcCmd, executeCommand, runWithOutput,
  CommandOutcome, requireOutput, replaceAtomically, atomicMoveOrFallback, renderErrorMessage,
  runDiff, runLint, resolveLintFile}` keep their names and positional shapes. If a body moves,
  leave a one-line forwarder. Never edit `WriteCommands.scala:60-1260` (write paths, cone, strict
  gate), `Recalc.scala` bodies, `WorkbookEvaluator.recalculate*` bodies, `DataTableSeeder`,
  `Evaluator.scala`, `DependencyGraph.scala`, `FunctionSpecs*`.
- **Prelude probes.** Anything a script can reach gets a test in
  `xl/test/src/xlprelude/ScriptingPreludeTest.scala` (package `xlprelude`, outside `com.tjclp.xl`,
  `import com.tjclp.xl.scripting.{*, given}` only). No default arguments on extension methods; no
  new `transparent inline` unions; new enums/case classes `derives CanEqual`.
- **Additive JSON.** Never remove or rename a key in an existing payload; `--format json` stays
  the bare payload; the envelope has exactly `ok, exitCode, verb, version, data, warnings, error`.
- **Commits.** Conventional subject (`feat(cli): …`, `fix(core): …`, `docs: …`), body says why and
  lists the new tests with counts, `Refs #n`, the session's trailers. Stage paths explicitly —
  never `git add -A` or `git add .`.
- **Definition of done (implementer):** tests first and green; `<module>.test` green for every
  touched module; `__.compile` green; format pair clean; probes added; docs in the cluster's list
  updated; commit body complete. **Definition of done (integrator, before the PR):** the four
  gates with `set -o pipefail` — `./mill __.compile`, `./mill __.test`, `scripts/test-examples.sh`,
  `scripts/verify-skill-snippets.sh --local` — plus the golden runner, CHANGELOG entries, roadmap
  section, ADR-017 index line, test counts from the Mill run.

## Conflict and ordering note

Six parallel groups, at most two clusters each; a group starts when every `DependsOn` of both
members is integrated. Members of one group never touch the same file (source **or** prose). Each
group has at most one cluster that touches xl-core or xl-evaluator; its partner is xl-cli-only,
with one exception: group 2 pairs `errors-exit-codes` (xl-core `XLError.scala` + `text/Suggest.scala`,
otherwise xl-cli) with `structural-row-props` (xl-ooxml only) because every xl-cli-only cluster
depends on the error types — xl-ooxml recompiles its dependants but not xl-core or xl-evaluator.

| Group | Clusters | Shared-file rule |
|---|---|---|
| 1 | `harness-golden` (alone, in progress) | owns `Main.scala`, `Cli.scala`, `CliIO.scala`, `xl-cli/package.mill` |
| 2 | `errors-exit-codes` ∥ `structural-row-props` | errors owns `Main.scala`, `SheetResolver.scala`, `Format.scala`, `Session.scala`, `contract/*`, `XLError.scala`, `text/Suggest.scala`, `DiffCommandSpec`/`LintCommandSpec`/`InPlaceSpec`, goldens, `cli.md`, xl-cli `SKILL.md`; row-props owns `OoxmlWorksheet.scala`, its new spec, `StructuralCommandSpec` |
| 3 | `envelope-json` ∥ `recalc-options-renamer` | envelope owns `Main.scala`, `WorkbookCommands.scala`, `ReadCommands.scala`, `contract/*`, goldens, `cli.md`, xl-cli `SKILL.md`; renamer owns xl-evaluator files, `exports.scala`, `SheetCommands.scala`, `BatchParser.applyRenameSheet` (one method), the `WriteCommands.scala:940-949` comment, probes, `scripting.md`, xl-scripting `SKILL.md`, `LIMITATIONS.md` |
| 4 | `batch-opspec-scope` ∥ `describe-audit-deps` | batch owns `BatchParser.scala`, `batch/*`, `ValueParser.scala`, `WriteCommands.scala:950` (visibility) and `:1268-1301` (one call site), `StreamingWriteCommands.batch`; describe owns `Main.scala`, `Command.scala`, `InspectCommands.scala`, `ReadCommands.cell`, `exports.scala`, evaluator files, probes, `cli.md`, xl-cli `SKILL.md`. Batch leaves its `cli.md`/`SKILL.md` deltas to `docs-from-code` |
| 5 | `resolve-and-argv` ∥ `twins-and-navigation` | resolve owns `Resolve.scala`, `SheetResolver.scala`, `BatchParser.scala` (17 sites), `StreamingReadCommands.scala`, `StreamingWriteCommands.scala`, `Argv.scala`, `Cli.scala`, goldens, `cli.md`, xl-cli `SKILL.md`; twins owns xl-core files, probes, `scripting.md`, xl-scripting `SKILL.md` |
| 6 | `docs-from-code` (alone) | owns `Main.scala`, `contract/{FunctionDoc,Schema}.scala`, `DocsGenSpec`, `docs/reference/generated/*`, `cli.md`, `SKILL.md`, `CLAUDE.md`, `plugin.json`, `ci.yml` |

`Main.scala` is edited by exactly one cluster per group (1, 2, 3, 4-describe, 6); `Cli.scala` by
harness-golden and resolve-and-argv only; `BatchParser.scala` by renamer (one method), batch,
resolve in that order; `exports.scala` by renamer then describe. If `harness-golden` placed the
option-group chain in `Cli.scala` rather than `Main.scala`, the integrator serialises group 5
(`describe-audit-deps` is already integrated by then, so only `resolve-and-argv` needs a rebase).
`twins-and-navigation` and `structural-row-props` are independent and may be pulled into any
earlier group whose partner is xl-cli-only if a slot frees. If the session runs short, cut in this
order: `docs-from-code` (integrator does the pointer edits by hand), `twins-and-navigation`,
`describe-audit-deps`; never `errors-exit-codes`, `envelope-json`, `batch-opspec-scope`,
`resolve-and-argv`.

Nominal hours below already include the judges' 1.5–2× correction: ≈38.5 agent-hours, ≈25 wall
hours at two concurrent implementers plus integration.

---

## Wave 1 — this session

### Cluster harness-golden

**Status**: done — `7510c5d`, `3d13f14` (in-process harness, 39 goldens, `Locale.US` pin, concurrent-run test)
**Modules**: xl-cli (src + test)
**Files**: new `xl-cli/src/com/tjclp/xl/cli/CliIO.scala`; new `xl-cli/src/com/tjclp/xl/cli/Cli.scala`; `xl-cli/src/com/tjclp/xl/cli/Main.scala`; `xl-cli/package.mill`; new `xl-cli/test/src/com/tjclp/xl/cli/contract/CliHarness.scala`; new `xl-cli/test/src/com/tjclp/xl/cli/contract/GoldenSpec.scala`; new `xl-cli/test/src/com/tjclp/xl/cli/contract/CliHarnessSpec.scala`; new `xl-cli/test/src/com/tjclp/xl/cli/contract/TestFixtures.scala`; new `xl-cli/test/resources/golden/*.golden` (~30); `docs/reference/testing-guide.md`
**DependsOn**: —
**ParallelGroup**: 1
**HoursEstimate**: 4
**ClosesIssues**: —

**Implementer brief** (design B's C1, as written; the safety net every later cluster stands on —
no behaviour change in this cluster):

1. Add `CliIO(out: String => IO[Unit], err: String => IO[Unit], stdin: IO[String])` with
   `CliIO.system` (`IO.println`, `IO.blocking(System.err.println(_))`,
   `IO.blocking(scala.io.Source.stdin.mkString)`) and
   `CliIO.capturing(stdin: String = ""): IO[(CliIO, IO[(String, String)])]` (two
   `Ref[IO, StringBuilder]`s; the second element yields `(stdout, stderr)`).
2. Convert `Main` from `CommandIOApp` to `IOApp` — `CommandIOApp.run(List[String])` is `final`.
   Reproduce decline-effect's behaviour line for line: `--help` → help text on stdout, exit 0;
   parse failure → help on stderr, exit **1 in this cluster** (`errors-exit-codes` makes it 2);
   `--version` → `BuildInfo.version` on stdout, exit 0. Keep `override def runtimeConfig`
   (GH-519), keep `def main: Opts[IO[ExitCode]]` (specs parse `Command("xl","test")(Main.main)`),
   keep every `val *Cmd`. `Main.run(args) = Cli.run(args, CliIO.system)`;
   `Cli.run(args, io) = Command("xl", header, helpFlag = true)(program(io)).parse(args, sys.env)`
   matched into the three outcomes above. No argv normalisation yet.
3. Thread `io` through `Cli.program`: `printRunResult`, `runInfo`, `runRasterizers`, `runDiff`,
   `runLint`, `runHeadless`, `runStandalone`, `batchDryRun`, `runWithOutput`'s mutually-exclusive
   error and the `Format.errorSimple` sites in the option groups write through `io.out`/`io.err`.
   Errors still go to `io.out` in this cluster (they move in `errors-exit-codes`). The 12 handler-
   level `System.err.println` sites (batch warnings, `StyleBuilder` numFmt hints) stay where they
   are and are captured by the harness bracket below until Wave 2.
4. `CliHarness.run(args: String*)(stdin: String = ""): IO[CliRun]` with
   `final case class CliRun(exit: Int, stdout: String, stderr: String)`: calls `Cli.run(args.toList,
   io)` with a capturing sink and brackets `System.setErr` under one **process-global**
   `Semaphore[IO](1)` (MUnit runs suites in parallel; `NumFmtParsingSpec:113`,
   `ViewHiddenSpec:74`, `ViewTruncationSpec:70` also `setErr`).
5. Golden corpus: one file per case under `xl-cli/test/resources/golden/<case>.golden` with
   sections `## args`, `## stdin` (optional), `## exit`, `## stdout`, `## stderr`. Fixtures are
   built in-test from `Sheet`/`Workbook` builders into a per-run temp dir; normalisers replace the
   temp dir with `<DIR>`, temp file names with `<TMP>`, `BuildInfo.version` with `<VERSION>`.
   `XL_UPDATE_GOLDEN=1` rewrites; otherwise a diff fails with a unified diff in the message.
   `TestFixtures.simpleBook()` (sheets `Data` — text, number, date, a cached formula, an uncached
   formula — and `Summary`, empty) and `TestFixtures.singleSheetBook()`.
6. Pin ~30 cases of today's contract exactly as it is: `--help`; `put --help`; `view --help`;
   `batch --help`; `sheets`; `names`; `bounds`; `view Data!A1:C3`; `view --format json`; `view
   --format csv`; `cell Data!A1`; `search Hello`; `stats Data!B1:B3`; `functions`; `put` success
   (`-o`); `put` without `-o` (text + exit); `put` with an unknown sheet (candidates text + exit);
   `putf` with a bad formula; `batch -` happy path via stdin; `batch --dry-run -`; `batch -` with an
   unknown op; `diff` identical (0), differing (1), `--format json`; `lint` clean (0), missing file
   (2 today); `eval "=1+1"`; `-i recalc` (GH-464 message); `--stream view --format json`; flags
   after the verb (`view A1:B2 -s Data` → today's decline error, pinned so later clusters show a
   diff); `-s` with `names` (today a parse error, pinned); `--stream view` on a two-sheet book
   without `-s` (today: first sheet — pinned so `resolve-and-argv`'s change is a reviewed diff).
7. `xl-cli/package.mill` test object: `override def moduleDeps = super.moduleDeps ++
   Seq(build.`xl-ooxml`.test, build.`xl-core`.test)` (the pattern `xl-evaluator/package.mill:19`
   uses) so the 20-workbook fixture corpus and `Generators` are on the CLI test classpath.

**Tests first**: `GoldenSpec` (parametrised over the corpus); `CliHarnessSpec` — "stdout and
stderr are captured separately", "stdin reaches `batch -`", "exit code propagates", "two
concurrent harness runs do not interleave stderr".
**Existing tests that may need adjusting**: none expected. `MainSpec:1375`, `InPlaceSpec:304`,
`LintCommandSpec:210` keep parsing `Main.main`; `MainSpec:1456`/`InPlaceSpec:310` keep capturing
`System.out` because `CliIO.system.out` is `IO.println`.
**Docs/skill files to touch**: `docs/reference/testing-guide.md` — replace the "Golden File Tests
(Future - P11)" section (`:167-186`) with "CLI contract goldens" (layout, normalisers,
`XL_UPDATE_GOLDEN=1`, the review discipline: a golden diff is a contract change and needs a
CHANGELOG line).
**Integrator notes**: CHANGELOG `### Added` — "CLI contract golden corpus and in-process
`CliHarness` (`xl-cli/test/resources/golden`); `Main` is a plain `IOApp`, behaviour unchanged."
Record the new test count from `xl-cli.test`. Confirm at merge whether the option-group chain
stayed in `Main.scala` (default assumption for group 5).

---

### Cluster errors-exit-codes

**Status**: done — `7ab44f2`, `9f98be4` (typed `CliError`, exit table 0/1/2/3, diagnostics on stderr, `Session.scala` removed)
**Modules**: xl-core (two files), xl-cli
**Files**: `xl-core/src/com/tjclp/xl/error/XLError.scala`; new `xl-core/src/com/tjclp/xl/text/Suggest.scala`; new `xl-core/test/src/com/tjclp/xl/error/XLErrorCodesSpec.scala`; new `xl-core/test/src/com/tjclp/xl/text/SuggestSpec.scala`; new `xl-cli/src/com/tjclp/xl/cli/contract/{CliError,ErrorCode,ExitCodes,Diagnostics,Warning}.scala`; `xl-cli/src/com/tjclp/xl/cli/Main.scala`; `xl-cli/src/com/tjclp/xl/cli/helpers/SheetResolver.scala`; `xl-cli/src/com/tjclp/xl/cli/output/Format.scala`; `xl-cli/src/com/tjclp/xl/cli/Session.scala` (delete); new `xl-cli/test/src/com/tjclp/xl/cli/contract/{CliErrorSpec,DiagnosticsSpec,ExitCodesSpec}.scala`; `xl-cli/test/src/com/tjclp/xl/cli/{DiffCommandSpec,LintCommandSpec,InPlaceSpec}.scala`; goldens; `docs/reference/cli.md`; `plugin/skills/xl-cli/SKILL.md`
**DependsOn**: harness-golden
**ParallelGroup**: 2
**HoursEstimate**: 4
**ClosesIssues**: — (retires "exit 1 means five things" and "errors on stdout")

**Implementer brief**:

1. xl-core, `XLError.scala`: add `derives CanEqual` to `enum XLError`. Add three cases after
   `Other`: `EditFailed(index: Int, op: String, cause: XLError)`,
   `UnsupportedCapability(op: String, capability: String, hint: String)`,
   `SheetRequired(context: String, available: Vector[String])`; extend the `message` match
   (`s"op $index ($op): ${cause.message}"`, `s"$op requires $capability: $hint"`,
   `s"$context requires a sheet: pass -s <name> or qualify the ref (Sheet!A1). Available: ${available.mkString(", ")}"`).
   Add extensions in `object XLError`: `code: String` (SCREAMING_SNAKE of `productPrefix`:
   `SheetNotFound` → `SHEET_NOT_FOUND`, `InvalidCellRef` → `INVALID_CELL_REF`, `IOError` →
   `IO_ERROR`; `EditFailed` → `cause.code`; `Other` → `OTHER`), `hint: Option[String]` (a per-case
   table: `SheetNotFound` → "list sheets with `xl -f <file> sheets`"; `SheetRequired` → "use -s
   <name> or a qualified ref like 'Name'!A1"; `ValueCountMismatch` → "provide exactly N values, or
   1 to fill the range"; `FormulaError` → "check the formula with `xl eval`"; `SecurityError` →
   "re-run with --max-size 0 or --stream"; `OutOfBounds` → "valid cells are A1..XFD1048576"; others
   `None`), `candidates: Vector[String]` (`SheetRequired` → `available`; `EditFailed` →
   `cause.candidates`; else empty), `root: XLError`, `opIndex: Option[Int]`. Add
   `val codes: Vector[String]` built from one sample instance per case (a private table) so the
   set is enumerable for `schema --json`. Grep the repo for exhaustive `XLError` matches
   (`case XLError\.`) and add arms where the compiler warns.
2. xl-core, `text/Suggest.scala`: `def closest(input: String, candidates: Iterable[String], limit:
   Int = 3): Vector[String]` — case-insensitive Levenshtein, keep distance ≤ `max(2, input.length /
   3)`, stable order by (distance, candidate). Leave `FormulaParser.levenshteinDistance`
   (`FormulaParser.scala:1304`) alone this wave (Wave 2 delegates).
3. xl-cli `contract/`: implement ADR-017 §2.3 exactly — `Location`, `CliError(code: String, message,
   hint, candidates, location, cause: Option[XLError])` with `exitCode = ExitCodes.forCode(code)`,
   `CliException(error) extends Exception(error.message) with NoStackTrace`,
   `CliError.fromXLError(e, at)` (`code = e.code`, `hint = e.hint`, `candidates = e.candidates`;
   for `SheetNotFound` the caller adds `Suggest.closest(name, wb.sheetNames.map(_.value))` when it
   has the workbook), `CliError.fromThrowable(t)` (`CliException` → its error; `StrictFailure` →
   `RECALC_GATE` carrying the summary as the message; `XLException` → `fromXLError`;
   `java.nio.file.NoSuchFileException` → `IO_READ`; anything else → `INTERNAL` with `getMessage`,
   `null`-safe via `Option(t.getMessage).getOrElse(t.toString)`), `CliError.usage(message, hint)`;
   `object ErrorCode` with the CLI-only constants and `val all: Vector[String] = cliCodes ++
   XLError.codes` (assert unique); `object ExitCodes` (`ok`, `signal = ExitCode(1)`, `usage =
   ExitCode(2)`, `failed = ExitCode(3)`, `forCode`); `Warning(code, message, location)`,
   `object WarningCode`; `Diagnostics.render(err)` = `"Error: <message>"` first line (today's exact
   text), then `"  code: <CODE>"`, `"  did you mean: a, b"` when candidates non-empty, `"  hint:
   <text>"` when present; `Diagnostics.report(err, io)` writes to `io.err`.
4. `Main.runResult`: `Left(e: CliException)` → `CommandOutcome(e.error.exitCode, "", outputComplete
   = false)` with `Diagnostics.report` to stderr; `Left(strict: StrictFailure)` unchanged (exit 1,
   summary on stdout, `unsayTheSave` under `-i`); `Left(other)` → `fromThrowable` → exit 3 via
   stderr. `runHeadless`/`runStandalone` errors → stderr, exit 3. `runDiff`/`runLint`: findings keep
   1; the runtime-error arms (`:2017`, `:2049`) become 3; `resolveLintFile`'s usage error stays 2
   but moves to stderr; the `-i`+`-o` conflict (`:2872-2874`) → `USAGE`, exit 2, stderr; decline
   parse failures in `Cli.run` → exit 2 (still the decline help text; the one-line usage arrives
   with `resolve-and-argv`). Stdout is empty on every failure.
5. Migrate raise sites to `CliException`: `SheetResolver` (4 methods; keep both message texts
   verbatim; `resolveSheet`/`findSheet` → `SHEET_NOT_FOUND` with `candidates = Suggest.closest`;
   `requireSheet` → `XLError.SheetRequired(context, names)` projected), `requireOutput` /
   `requireOutputAction` / `missingOutputError` (→ `OUTPUT_REQUIRED`, exit 2 — note it still fires
   after the read until Wave 2), the streaming refusals at `Main.scala:2159-2164`, `:2209-2214`,
   `:2233-2238`, `:2282` (→ `UNSUPPORTED_IN_STREAM` with the in-memory alternative in `hint`),
   `StreamingWriteCommands.scala:1125-1131` and `:817-820` (→ `UNSUPPORTED_IN_STREAM` /
   `SHEET_REQUIRED`). Leave `WriteCommands` alone; an un-migrated `new Exception(msg)` still yields
   a well-formed `INTERNAL` diagnostic with the same message.
6. Wire `ExcelIO.withWarnings` (`ExcelIO.scala:1396`) in `Main.execute` and `runHeadless`: collect
   `XlsxReader.Warning` values into a per-run `Ref[IO, Vector[Warning]]` and print
   `Warning[READER_WARNING]: <case>` to stderr after the result (the envelope picks them up in
   `envelope-json`).
7. Delete `Session.scala` (zero references) and the `Format` renderers with zero callers
   (`batchSuccess`, `openSuccess`, `createSuccess`, `selectSuccess`, `error`, `errorDetails`;
   grep each before deleting; keep `errorSimple`, `saveSuffix`, `putSuccess`, `evalSuccess`,
   `evalArraySuccess`, `cellInfo`). Add the four-line exit-code table to the `xl --help` header.

**Tests first**: `XLErrorCodesSpec` — every case has a non-empty code matching
`^[A-Z][A-Z0-9_]*$`, codes are unique, `EditFailed(3, "put", SheetNotFound("x")).code ==
"SHEET_NOT_FOUND"`, `XLError.Other("a") == XLError.Other("a")` compiles without a cast;
`SuggestSpec` — `closest("Sale", Seq("Sales","Data")) == Vector("Sales")`, ranking by distance,
empty when nothing is close, case-insensitive; `CliErrorSpec` — every `XLError.codes` entry appears
in `ErrorCode.all`, `fromThrowable` on a `null`-message exception is safe, `StrictFailure` →
`RECALC_GATE` → exit 1; `ExitCodesSpec` — every code maps to exactly one of 1/2/3 and every 1 is
a findings/gate code; `DiagnosticsSpec` — text shape with and without candidates; harness cases:
"sheet-not-found writes nothing to stdout, `Error:` + `code:` on stderr, exit 3"; "missing -o exits
2"; "`-i` with `-o` exits 2"; "diff differs 1 / unreadable file 3"; "lint findings 1 / usage 2 /
corrupt zip 3"; "strict gate still 1 and `-o` still writes"; "reader warning surfaces on stderr".
**Existing tests that may need adjusting**: `DiffCommandSpec:278` (`ExitCode(2)` → `ExitCode(3)`:
an unreadable file is a failure, not usage); `LintCommandSpec:144` (same); `InPlaceSpec:220`
(`ExitCode.Error` → `ExitCode(2)`: `-i`+`-o` is usage). `InPlaceSpec:105-109, 265-276` inject
`ExitCode.Error` and assert the plumbing — unchanged. `MainSpec:1447,1475,1477,1536,1537`,
`InPlaceSpec:132-136`, `LintCommandSpec:122,134,152,244,251`, `DiffCommandSpec:265` unchanged.
Regenerate goldens with `XL_UPDATE_GOLDEN=1` and review: only exit codes and the stdout→stderr move
may differ.
**Docs/skill files to touch**: `docs/reference/cli.md:1401-1408` ("Error Format" that nothing
emits) replaced by the real stderr format and the exit table; `plugin/skills/xl-cli/SKILL.md`
`### Strict exit codes for pipelines` (`:685`) gains the table.
**Integrator notes**: CHANGELOG `### Changed` (breaking, call it out) — "CLI exit codes: 0 ok,
1 completed with findings or a failed `--strict` gate (unchanged), 2 usage (was 1), 3 operation
failed (was 1; `diff`/`lint` runtime errors were 2). Errors and warnings go to stderr with a stable
`code:` line; stdout carries results only." `### Added` — "Stable `XLError.code`/`hint`/
`candidates`, `XLError derives CanEqual`, new cases `EditFailed`, `UnsupportedCapability`,
`SheetRequired`; `com.tjclp.xl.text.Suggest`." `### Removed` — "dead `Session.scala` and unused
`Format` renderers." Note the downstream exhaustive-match warning in the `### Changed` entry.

---

### Cluster structural-row-props

**Status**: done — `86afeb7` (domain `rowProperties` authoritative on both writers; `PropertyOnlyRowShiftSpec`)
**Modules**: xl-ooxml (+ one xl-cli test)
**Files**: `xl-ooxml/src/com/tjclp/xl/ooxml/worksheet/OoxmlWorksheet.scala` (`:496-525`); new `xl-ooxml/test/src/com/tjclp/xl/ooxml/PropertyOnlyRowShiftSpec.scala`; `xl-cli/test/src/com/tjclp/xl/cli/StructuralCommandSpec.scala` (+1 case)
**DependsOn**: —
**ParallelGroup**: 2
**HoursEstimate**: 2
**ClosesIssues**: #558

**Implementer brief** (design C's cluster I; root cause verified by all three judges):

1. Root cause: `Sheet.shiftAxis` already remaps `rowProperties` correctly. The writer duplicates:
   `preservedRowsWithoutDomainCells` (`OoxmlWorksheet.scala:500-510`) re-emits every cell-free
   *source* `<row>` at its ORIGINAL index with its original `ht`/`customHeight`/`hidden`/
   `outlineLevel`/`collapsed`/`s`/`customFormat`, then `emptyRowsFromDomain` (`:516-521`) emits
   the shifted domain property at the NEW index — a property-only row appears twice after
   `insert-rows`/`delete-rows`.
2. Fix: the domain `rowProperties` is authoritative for every attribute `XlsxReader` models. For a
   preserved cell-free source row: strip the modelled attributes, keep only unmodelled ones
   (`spans`, `thickBot`, `thickTop`, `ph`, …), then apply the domain props for that index if any;
   drop the row when nothing is left (no cells, no attributes). Every row index is emitted at most
   once. Check `<col>` emission (`WorksheetHelpers`) for the symmetric bug and apply the same rule
   if present; #558 reports rows only.
3. Byte-identity guard: for a book read and written with no edit, the re-applied domain props must
   reproduce the source row exactly — `OoxmlRoundTripSpec`, `FixturePreservationSpec`,
   `DeterminismSpec` and `FixtureCorpusSpec` are the gate; run them.

**Tests first**: `PropertyOnlyRowShiftSpec` — build the #558 fixture (`A1,A2,A4,A6` text, `A7
=A2&A4`; row 3 `ht=3`, row 5 `hidden outlineLevel=1`, row 8 `ht=5.15`) through a real write→read
so `sourceContext` is populated (the CLI path); `StructuralEditor.insertRowsChecked(wb, S, 1, 1)`
then write: the XML has `<row r="4" ht="3">`, `<row r="6" hidden outlineLevel="1">`, `<row r="9"
ht="5.15">` and no `r="3"/"5"/"8"` property rows; `deleteRowsChecked(…, 3, 1)` → 3, 4, 7 only;
law: write→read→insert→delete→write yields the original row set; an unmodelled attribute
(`thickBot="1"`) on a cell-free source row survives an unrelated edit; an untouched book
round-trips byte-identically. xl-cli `StructuralCommandSpec`: `WriteCommands.insertRows` on the
fixture, raw-zip assertion that each `r=` appears once.
**Existing tests that may need adjusting**: none expected; if `RowColumnOperationsSpec` pinned the
duplicate, fix the pin with a GH-558 comment.
**Docs/skill files to touch**: none.
**Integrator notes**: CHANGELOG `### Fixed` — "Structural edits no longer duplicate
property-only rows (height, hidden, outline level) at both the old and the new index (#558)."

---

### Cluster envelope-json

**Status**: done — `2f879a8`, `7640737`, `5b60dff` (seven-key envelope, typed `data` for the read verbs, warnings sink, review rework closed the `--stream`/`--eval` stderr leaks)
**Modules**: xl-cli
**Files**: new `xl-cli/src/com/tjclp/xl/cli/contract/{Outcome,Payload,Render}.scala`; `xl-cli/src/com/tjclp/xl/cli/Main.scala`; `xl-cli/src/com/tjclp/xl/cli/commands/WorkbookCommands.scala`; `xl-cli/src/com/tjclp/xl/cli/commands/ReadCommands.scala`; new `xl-cli/test/src/com/tjclp/xl/cli/contract/{RenderSpec,EnvelopeSpec}.scala`; new `xl-cli/test/resources/schema/envelope.schema.json`; goldens; `docs/reference/cli.md`; `plugin/skills/xl-cli/SKILL.md`
**DependsOn**: errors-exit-codes
**ParallelGroup**: 3
**HoursEstimate**: 4
**ClosesIssues**: — (retires "machine-readable output for 4 of 56 verbs")

**Implementer brief**:

1. Implement ADR-017 §2.4: `Payload.Text(text, saved: Option[String], written: Boolean)`,
   `Payload.Json(value: ujson.Value)`, `Outcome(verb, payload, warnings, error, exitCode,
   outputComplete)`, `Rendered(stdout, stderr)`, `Render.text`, `Render.json(o, version)`. Keep
   `private[cli] final case class CommandOutcome(exitCode, output, outputComplete)` as a forwarder
   built from `Outcome` (`InPlaceSpec:41-42` constructs it); `runStagedOutput`'s commit rule reads
   `outputComplete` as today.
2. Global `--json` flag (`Opts.flag("json", "Wrap every result in the JSON envelope; errors too").orFalse`)
   added to every option group in the chain (the Wave 2 registry collapses them) and threaded into
   `runResult`/`run*` as `OutputMode.Text | Json`. Also scan the raw args for `--json` in `Cli.run`
   so decline usage errors produce an envelope too.
3. `Render.text` reproduces today's stdout bytes for every verb (the goldens prove it).
   `Render.json` emits the seven-key envelope: `verb` = the subcommand path joined by a space
   (`"sheets hide"`), `version = BuildInfo.version`, `data` from the payload — `Text` →
   `{"text": …, "saved": …, "written": …}` (`saved` = the display path after the GH-464 temp→target
   rewrite, `null` and `written:false` when nothing was committed, i.e. an `-i` strict failure);
   `Json(v)` → `v`. `view`/`filter`/`diff`/`lint` in `--json` mode reuse their existing JSON
   strings parsed with `ujson.read` — do not rebuild them. Typed `data` for `sheets`
   (`[{name, index, state, dimension}]`), `names` (`[{name, refersTo, scope, hidden}]`), `bounds`
   (`{sheet, range, dimension: bool}`), `eval`/`evala` (`{formula, result: {type, value,
   formatted}, overrides}`), `functions` (`[{name}]` until `docs-from-code` types it),
   `rasterizers`: add `sheetsData(wb)`, `sheetsQuickData(path)`, `namesData(path)` in
   `WorkbookCommands` and a typed accessor beside `ReadCommands.eval`/`evalArray`, keeping the
   `IO[String]` entry points. Under `--json`, `search`/`stats`/`cell`/`put` and every write verb
   wrap their prose as `Payload.Text` (typed forms are Wave 2).
4. Warnings: one `Ref[IO, Vector[Warning]]` per run in `Cli.program`, fed by the reader-warning
   hook from `errors-exit-codes` and by `ReadCommands.view`'s hidden-row / truncation notices where
   they already exist as strings (`HIDDEN_OMITTED`, `TRUNCATED`); text mode prints
   `Warning[CODE]: …` lines on stderr, JSON mode fills `warnings[]`.
5. Failure in `--json` mode: the `ok:false` envelope on stdout, one `Error: <message>` line on
   stderr, exit per the table. A strict-gate failure is `ok:false, exitCode:1, error.code:
   RECALC_GATE` with `data: {"text": <summary>, "saved": …, "written": …}` — the summary is data,
   not lost.
6. `batch --dry-run --json` → `data: {ops: [{index, op, summary}], warnings: [...]}` from the
   existing `formatSummary` lines.

**Tests first**: `RenderSpec` — envelope shape for each payload case; `ok` ⇔ `error == null`;
`exitCode` consistent with the error code; `version` present; exactly seven keys. `EnvelopeSpec`
via the harness: "`view --json` data equals `view --format json` payload"; "`put --json` success:
`ok:true`, `data.saved` = output path, `data.written:true`"; "`put --json` sheet-not-found:
`ok:false`, `error.code == SHEET_NOT_FOUND`, `exitCode:3`, `data:null`, stderr one line";
"`sheets --json` typed"; "`eval --json` typed"; "`-i --strict recalc --json` on a cyclic book:
exit 1, `data.written:false`"; "`batch --dry-run --json`"; "`--json` with a decline parse error
still yields an envelope"; "every envelope validates against `envelope.schema.json`".
**Existing tests that may need adjusting**: `InPlaceSpec` compiles through the `CommandOutcome`
forwarder. Goldens: additions only (new `--json` cases).
**Docs/skill files to touch**: `docs/reference/cli.md` `## Output Format` (`:1387`) gains an
"Output contract" subsection (envelope, `--json` vs `--format`, channels);
`plugin/skills/xl-cli/SKILL.md` `## Quick Reference` gains "Always pass `--json` when a program
reads the result".
**Integrator notes**: CHANGELOG `### Added` — "Global `--json`: every verb, success or failure,
prints one envelope `{ok, exitCode, verb, version, data, warnings, error}` on stdout; `--format
json` payloads are unchanged." Record the new test count.

---

### Cluster recalc-options-renamer

**Status**: done — `fdaf89b`, `850198c` (rework round: chart remap kept, `recalculateUncached` never withdraws a cache, 3-D ranges refuse)
**Modules**: xl-evaluator, xl-cli (two call sites + one comment), xl (probes)
**Files**: new `xl-evaluator/src/com/tjclp/xl/formula/eval/RecalcOptions.scala`; `xl-evaluator/src/com/tjclp/xl/formula/eval/WorkbookEvaluator.scala` (additive members of the existing `extension (wb: Workbook)` block only); `xl-evaluator/src/com/tjclp/xl/formula/eval/Recalc.scala` (one additive method after `toEither`); `xl-evaluator/src/com/tjclp/xl/formula/printer/FormulaShifter.scala` (additive); new `xl-evaluator/src/com/tjclp/xl/formula/printer/FormulaOps.scala`; new `xl-evaluator/src/com/tjclp/xl/formula/eval/SheetRenamer.scala`; `xl-evaluator/src/com/tjclp/xl/exports.scala` (additive); new `xl-evaluator/test/src/com/tjclp/xl/formula/{RecalcOptionsSpec,SheetRenamerSpec,FormulaOpsSpec}.scala`; `xl-cli/src/com/tjclp/xl/cli/commands/SheetCommands.scala` (`renameSheet` body); `xl-cli/src/com/tjclp/xl/cli/helpers/BatchParser.scala` (`applyRenameSheet` body only, `:1852-1861`); `xl-cli/src/com/tjclp/xl/cli/commands/WriteCommands.scala` (the `isCellMutating` doc comment at `:940-949` only); new `xl-cli/test/src/com/tjclp/xl/cli/RenameSheetSpec.scala`; `xl/test/src/xlprelude/ScriptingPreludeTest.scala`; `docs/reference/scripting.md`; `plugin/skills/xl-scripting/SKILL.md`; `docs/LIMITATIONS.md`
**DependsOn**: — (the only evaluator cluster in its group)
**ParallelGroup**: 3
**HoursEstimate**: 5
**ClosesIssues**: #559

**Implementer brief**:

1. `RecalcOptions.scala`: `enum IterativeMode derives CanEqual { case Off; case FromCalcPr; case
   Force(calc: IterativeCalc) }`; `final case class RecalcOptions(clock: Clock = Clock.system, rng:
   Rng = Rng.system, iterative: IterativeMode = IterativeMode.FromCalcPr, parallelism: Int = 1,
   seedTables: Boolean = false) derives CanEqual`; `object RecalcOptions { val default =
   RecalcOptions() }`.
2. `WorkbookEvaluator.scala`, appended inside the existing `extension (wb: Workbook)` block (no
   default arguments; no existing body changes): `def recalculate(options: RecalcOptions):
   RecalcResult` — `FromCalcPr` → `wb.metadata.calcPr.filter(_.iterativeCalculation).map(IterativeCalc.fromCalcPr)`
   exactly as `WriteCommands.recalcHonoringCalcPr` (`:996-1003`), then `parallelism > 1` →
   `recalculateParallel(options.clock, parallelism)` else `recalculate(clock, rng[, iterative])`;
   `seedTables` → `seedDataTablesReport` after. `def recalculateAfterEdit(sheet: SheetName,
   modified: Set[ARef], options: RecalcOptions): RecalcResult` = `DependentRecalculation.recalculateAfterEdit(wb, sheet, modified, options.clock)`
   (`private[xl]`, reachable from `com.tjclp.xl.formula.eval`) — when the book declares iteration
   fall back to `recalculate(options)` exactly as `writeAfterRefresh` does; copy that four-line
   decision, do not move it. `def recalculateUncached(options: RecalcOptions): RecalcResult` —
   evaluate ONLY cells whose value is `Formula(_, None, _)`, in dependency order, reading their
   inputs' caches as they are; cached cells stay byte-identical (GH-468 doctrine); finalise
   through `RecalcResult.cacheResults` (`private[eval]`, reachable).
3. `Recalc.scala`: after `toEither`, `def summary: String` reproducing
   `WriteCommands.formatRecalcSummary`'s text byte-for-byte (copy the text; the CLI delegation is
   Wave 2). Do not touch `object RecalcResult.cacheResults`.
4. `FormulaShifter.scala` (additive): `def renameSheet[A](expr: TExpr[A], from: SheetName, to:
   SheetName): TExpr[A]` rewriting `SheetRef`, `SheetPolyRef`, `SheetRange`,
   `RangeLocation.CrossSheet` and `SheetNameRef` whose sheet matches `from` case-insensitively
   (`ExternalRange` untouched); structure it like `shiftStructuralInternal`. `def
   mentionsSheet(expr: TExpr[?], sheet: String): Boolean` = `referencesSheet(expr, sheet)` OR any
   `SheetNameRef` matches — `referencesSheet` (`:279-323`) deliberately returns `false` for
   `SheetNameRef` and `StructuralEditor` depends on that; do not change it.
5. `FormulaOps.scala`: `renameSheet(text, from, to): XLResult[String]` (parse → `renameSheet` →
   `FormulaPrinter.print`/`printFileForm` matching the caller's convention; `Left(FormulaError(text,
   reason))` when the parser rejects text that `mentionsSheet`), `mentionsSheet(text, sheet):
   Boolean` (the `definedNameRefusal` qualifier regex at `StructuralEditor.scala:140-143`, applied
   outside string literals — lift it into `FormulaOps` and make `StructuralEditor` call the new
   public form so there is one copy), `shift(text, colDelta, rowDelta): XLResult[String]` (parse →
   `FormulaShifter.shift` → print; keeps today's clamp).
6. `SheetRenamer.rename(wb, from, to): XLResult[Workbook]`: (a) refuse-before-mutate — collect
   every cell formula, defined-name segment (split on top-level commas like
   `shiftDefinedNameText`), CF formula (`CfRule.CellIs.formula1/formula2`,
   `CfRule.Expression.formula`, `Cfvo.Formula` inside `ColorScale`/`DataBar`) and DV formula
   (`DvKind.List/Custom/Bounded`) that `mentionsSheet(from)` and fails to parse → `Left`; (b)
   `wb.rename(from, to)`; (c) for every sheet, rewrite each `CellValue.Formula(text, cached,
   kind)` that mentions `from` with `FormulaOps.renameSheet`, **preserving `cached` and `kind`**
   (a rename changes no value); rewrite CF/DV texts (mirror `rewriteCfFormulas`/`rewriteDvFormulas`
   shapes; `Preserved` payloads and text-family rules untouched); (d) rewrite defined names via
   `metadata.copy(definedNames = …)` and mark metadata modified; (e) write every changed sheet back
   with `wb.put(sheet)` — never `copy(sheets = …)` — so `SourceContext.modifiedSheets` is right
   (ADR-017 invariant 1). `references(wb, sheet): Vector[QualifiedRef]` lists the cells whose text
   mentions the sheet (for `deps`/`describe`).
7. `exports.scala` (additive): `export formula.eval.{RecalcOptions, IterativeMode, SheetRenamer,
   StructuralEditor}`, `export formula.printer.{FormulaShifter, FormulaOps}`, `type QualifiedRef =
   formula.graph.DependencyGraph.QualifiedRef; val QualifiedRef = formula.graph.DependencyGraph.QualifiedRef`.
   Do not collapse `formula/package.scala` onto `formulaExports` this wave, and do not
   wildcard-export `DependentRecalculation.*` (its Sheet-level extension has default args).
8. xl-cli: `SheetCommands.renameSheet` body calls `SheetRenamer.rename` (same error mapping, summary
   gains `"; N formula(s) rewritten"` when N > 0); `BatchParser.applyRenameSheet` body calls
   `SheetRenamer.rename`; fix the `isCellMutating` comment that claims a rename "rewrites
   referencing formulas" (now true — say so).

**Tests first**: `RecalcOptionsSpec` — `recalculate(RecalcOptions())` ≡ `recalculate()` on an
acyclic book; on an `iterate`-declared book ≡ `recalculate(Clock.system,
IterativeCalc.fromCalcPr(cp))`; `IterativeMode.Off` reports the cycle as an error; `parallelism =
4` ≡ sequential; `recalculateUncached` leaves cached cells byte-identical and fills uncached ones;
`recalculateAfterEdit(sheet, Set(A1), opts)` ≡ `DependentRecalculation.recalculateAfterEdit(wb,
sheet, Set(A1), opts.clock)`; `summary` equals the three summary strings `BatchRecalcSpec`/
`StrictWriteIntegritySpec` assert on. `SheetRenamerSpec` — the #559 repro (`Sheet2!A1 =
=Sheet1!A1*2`, `=SUM(Sheet1!A1:A1)` → `Data!…`, caches preserved); rename to a name needing quotes
(`Q1 Data` → `'Q1 Data'!A1`); defined name `Sheet1!$A$1` and a comma-union name rewritten; CF
`Expression("Sheet1!A1>0")`, `Cfvo.Formula`, DV list `Sheet1!$A$1:$A$3` rewritten; `="Sheet1"`
string literal untouched; `SheetNameRef` rewritten; case-insensitive match; `ExternalRange`
identity; property `renameSheet(renameSheet(e, a, b), b, a) == e` over generated expressions;
refusal when an unparsable name mentions the sheet (workbook untouched); `Workbook.rename`
unchanged (tab only); **the SourceContext law: read a fixture file → rename → write → re-read shows
the rewritten text on a sheet other than the renamed one, and a sheet that never mentioned the old
name is byte-identical in the zip.** `FormulaOpsSpec` — `shift("=A1+$B$1", 1, 1) == Right("=B2+$B$1")`;
`mentionsSheet` ignores string literals. xl-cli `RenameSheetSpec` — `SheetCommands.renameSheet` and
batch `rename-sheet` write rewritten formulas; summary counts. Probes: `RecalcOptions()`,
`wb.recalculate(RecalcOptions.default)`, `wb.recalculateAfterEdit(...)`, `wb.recalculateUncached(...)`,
`result.summary`, `SheetRenamer.rename`, `FormulaOps.renameSheet`, `StructuralEditor.insertRowsChecked`,
`FormulaShifter.shift`, `QualifiedRef(...)` all resolve through the prelude.
**Existing tests that may need adjusting**: none expected; `BatchRecalcSpec`'s classification of
`rename-sheet` as non-recalculating stays valid (values unchanged).
**Docs/skill files to touch**: `docs/reference/scripting.md` `## Formulas: build, recalculate,
inspect` (one primitive, `RecalcOptions`, `SheetRenamer`); `plugin/skills/xl-scripting/SKILL.md`
recipe + Gotchas (rename now rewrites); `docs/LIMITATIONS.md:332` (rename limitation → typed
formulas, CF, DV and names are rewritten; `Preserved` payloads still are not). Leave `cli.md`'s
`rename-sheet` line and the xl-cli `SKILL.md` gotcha to `docs-from-code`.
**Integrator notes**: CHANGELOG `### Fixed` — "Renaming a sheet (`rename-sheet`, batch
`rename-sheet`, `SheetRenamer.rename`) rewrites every reference to it: cell formulas on every
sheet, defined names, conditional-format and data-validation formulas; cached values preserved
(#559)." `### Added` — "`RecalcOptions`/`IterativeMode`, `wb.recalculate(options)`,
`wb.recalculateAfterEdit(sheet, refs, options)`, `wb.recalculateUncached(options)`,
`RecalcResult.summary`; `FormulaOps`; `StructuralEditor`, `FormulaShifter`, `QualifiedRef`
exported from the formula surface." Apply the `cli.md`/xl-cli `SKILL.md` deltas with
`docs-from-code`.

---

### Cluster batch-opspec-scope

**Status**: done — `f058ddc`, `49ba36c`, `d60ad13` (32 `OpSpec`s, `sheet` on every op, `--stream` refusal by index, #560; rework: streaming formats keep the cell style)
**Modules**: xl-cli
**Files**: new `xl-cli/src/com/tjclp/xl/cli/batch/{OpSpec,OpRegistry}.scala`; `xl-cli/src/com/tjclp/xl/cli/helpers/BatchParser.scala`; `xl-cli/src/com/tjclp/xl/cli/helpers/ValueParser.scala` (`:59-66`); `xl-cli/src/com/tjclp/xl/cli/commands/WriteCommands.scala` (`:950` visibility and the one call site in `batch` at `:1268-1301`); `xl-cli/src/com/tjclp/xl/cli/commands/StreamingWriteCommands.scala` (`batch`, `:434-463`); new `xl-cli/test/src/com/tjclp/xl/cli/batch/{OpRegistrySpec,BatchSheetScopeSpec,BatchPutFormatSpec,BatchStreamRefusalSpec}.scala`; `xl-cli/test/src/com/tjclp/xl/cli/BatchPutSpec.scala` (additions)
**DependsOn**: errors-exit-codes, recalc-options-renamer
**ParallelGroup**: 4
**HoursEstimate**: 4
**ClosesIssues**: #560

**Implementer brief**:

1. `OpSpec.scala`/`OpRegistry.scala` per ADR-017 §2.6 (`FieldKind`, `Field`, `OpSpec`,
   `OpRegistry.all/find/jsonSchema/helpText`, `ScopedOp`). One `OpSpec` per op name (32):
   `put` (`ref|range`, `value` xor `values`, `format` alias `numFormat`, `detect`, `sheet`),
   `putf` (`ref|range`, `value|formula` xor `values`, `from` alias `anchor`, `format`, `sheet`),
   `style` (the `StyleProps` fields; `numFormat` alias `format`; `align` alias `halign`;
   `replace`), `merge`, `unmerge`, `colwidth`, `rowheight`, `comment`, `remove-comment`,
   `hyperlink` (`target` alias `url`), `clear`, `col-hide`, `col-show`, `row-hide`, `row-show`,
   `autofit`, `add-sheet`, `rename-sheet`, `freeze`, `unfreeze`, `copy`, `sheet-view`,
   `tab-color`, `page-setup`, `header-footer`, `cf`, `chart`, `autofilter`, `group-rows`,
   `group-cols`, `ungroup-rows`, `ungroup-cols`. `cellMutating` mirrors `WriteCommands.isCellMutating`
   (make it `private[cli]` — one-line visibility change — so `OpRegistrySpec` can assert parity);
   `streamable` mirrors `buildStreamingBatchPatches` (`put`, `putf` single, explicit `values` and
   dragging `from` — the streaming writer already shifts, `StreamingWriteSpec` pins it; the
   "NO dragging" divergence is the streaming `putf` *verb*, not the batch op — `style`, `merge`,
   `unmerge`, `colwidth`, `rowheight`, `col-hide/show`, `row-hide/show`; not the 21 ops at
   `:806-816`);
   `sheetScoped = true` for every op except `add-sheet`/`rename-sheet`; `example` a valid object.
2. `BatchParser.parseBatchJson`: replace the 20 hand-listed `known*Props` sets with
   `OpRegistry.find(op).fields` (+ aliases, camel and kebab spellings) — the 32 arms stay; the
   unknown-property warning (`collectUnknownPropsWarning`) keeps its text and never fires on an
   alias; unknown op → `CliException(CliError(BATCH_OP_UNKNOWN, "Object N: Unknown operation '<x>'.
   Valid operations: …", candidates = Suggest.closest(x, OpRegistry.all.map(_.name)), location =
   Location(opIndex = Some(N))))` — the existing text is kept as a prefix (a spec matches on it) and
   `Did you mean: …` is appended; shape errors → `BATCH_OP_INVALID` with `opIndex`; a non-array
   document → `BATCH_JSON_INVALID`. Every op accepts an optional `"sheet"` key: the parser produces
   `ScopedOp(op, sheet, index)`; `ParseResult` gains `scoped: Vector[ScopedOp] = Vector.empty` with
   `ops == scoped.map(_.op)`; `parseBatchJson`/`parseBatchOperations` keep their signatures.
3. `BatchParser.applyScoped(wb, defaultSheetOpt: Option[Sheet], scoped: Vector[ScopedOp],
   recalcDependents: Boolean): IO[Workbook]` — the sheet for an op is `Resolve`-shaped: qualified
   ref > op `sheet` > default; a qualified ref that disagrees with `sheet` is `BATCH_OP_INVALID`
   (until `resolve-and-argv` lands, implement this locally in `BatchParser` and let that cluster
   forward it to `Resolve.forOp`); a `rename-sheet` of the current default updates the default for
   later ops; every apply-time failure is wrapped as `CliException(CliError(BATCH_OP_FAILED,
   s"Object ${i+1} (${op}): ${cause.getMessage}", location = Location(opIndex = Some(i+1))))` — the
   cause text stays a suffix so existing `contains` assertions hold. `applyBatchOperations(wb,
   sheetOpt, ops, recalcDependents)` keeps its signature and delegates with
   `ops.zipWithIndex.map((op, i) => ScopedOp(op, None, i + 1))`. `WriteCommands.batch`
   (`:1268-1301`): one call-site change — pass `parseResult.scoped` to `applyScoped` instead of
   `ops` to `applyBatchOperations`; `isCellMutating` is evaluated over `scoped.map(_.op)`. Nothing
   else in `WriteCommands` changes.
4. GH-560: in `updateSheetWithFormat` (`:1296-1310`) an explicit `format` must REPLACE the cell's
   numFmt — put the value, then apply the `putf` path (`applyNumFmt`, `:1319-1324`:
   `existing.withNumFmt(numFmt)`) instead of wrapping in `Formatted(cellValue, numFmt)`. The
   inferred hint (no `format` key) keeps `Sheet.put`'s General-only rule. Verify with the issue's
   fixture (`0.0%_);\(0.0%\)` on A1 → `#,##0_);\(#,##0\)`). `page-setup` accepts `fitToHeight: 0`
   (GH-463's batch half; the CLI flag half lands with the registry).
5. `--stream`: in `StreamingWriteCommands.batch`, after parsing and BEFORE `resolveSheetPath`,
   partition `scoped` by `OpRegistry.find(op).streamable` and by `sheet`/qualified ref equal to the
   streamed worksheet; any rejects → `CliException(CliError(UNSUPPORTED_IN_STREAM, "…ops [3 putf,
   7 clear] are not supported with --stream; drop --stream to apply them in memory", location
   opIndex of the first))` before any byte is written (ADR-017 invariant 2). The generic throw at
   `:817-820` becomes unreachable and is deleted.
6. `ValueParser.dataTableFormulaError` (`:59-66`) stops advertising the non-existent `data-table`
   batch op; point at `sheet.dataTable(...)` in scripts only.

**Tests first**: `OpRegistrySpec` — "every `BatchOp` case corresponds to a registered op name
(32)"; "`cellMutating` agrees with `WriteCommands.isCellMutating` for every op"; "`streamable`
agrees with the streaming supported set (dragging `putf` shifts identically under `--stream`)"; "every example
validates against `jsonSchema` and round-trips through `parseBatchJson`"; "names + aliases
unique". `BatchSheetScopeSpec` — each op with `"sheet":"Other"` lands on `Other`; a qualified ref
beats `sheet`; a disagreeing pair is `BATCH_OP_INVALID` naming the index; a missing sheet is
`BATCH_OP_FAILED` with candidates; `rename-sheet` of the default sheet retargets later ops;
`applyBatchOperations` still works with plain ops. `BatchPutFormatSpec` — the #560 fixture; a
`put` of a `LocalDate`-shaped string without `format` onto a `Currency` cell keeps `Currency`.
`BatchStreamRefusalSpec` — `--stream` with a `fill`/`clear`/dragging `putf` op refuses by index
and writes nothing; `--stream` with only streamable ops still works; `--stream` op with a `sheet`
different from the streamed one refuses. `BatchPutSpec` additions: unknown op suggests the nearest;
alias keys (`numFormat`/`format`, `num-format`, `anchor`, `url`, `halign`) raise no warning;
unknown key still warns with the known-property list.
**Existing tests that may need adjusting**: any test asserting the exact unknown-op message must
accept the appended `Did you mean:` line; `MainSpec` "provides clear error for invalid JSON"
(still `Left`, message non-empty). The 79 `BatchOp.*(...)` constructions and the
`NumFmtParsingSpec:71` pattern compile unchanged (wrapper design).
**Docs/skill files to touch**: none directly — hand the batch table, `--schema`, the `sheet` key,
aliases and the GH-560 note to `docs-from-code` (listed below).
**Integrator notes**: CHANGELOG `### Added` — "Every batch op accepts `sheet`; unknown ops
suggest the nearest name; apply-time errors carry the op index; `--stream` refuses unsupported
ops by index before writing anything; `OpRegistry` is the single table behind batch parsing."
`### Fixed` — "batch `put` with an explicit `format` replaces an existing custom number format
(#560); the phantom `data-table` op is no longer advertised." Pass the doc deltas to
`docs-from-code`.

---

### Cluster describe-audit-deps

**Status**: pending
**Modules**: xl-evaluator, xl-cli, xl (probes)
**Files**: new `xl-evaluator/src/com/tjclp/xl/formula/graph/QualifiedGraph.scala`; new `xl-evaluator/src/com/tjclp/xl/formula/eval/WorkbookAudit.scala`; new `xl-evaluator/src/com/tjclp/xl/formula/eval/WorkbookSummary.scala`; `xl-evaluator/src/com/tjclp/xl/exports.scala` (additive); new `xl-evaluator/test/src/com/tjclp/xl/formula/{QualifiedGraphSpec,WorkbookAuditSpec,WorkbookSummarySpec}.scala`; new `xl-cli/src/com/tjclp/xl/cli/commands/InspectCommands.scala`; `xl-cli/src/com/tjclp/xl/cli/Command.scala` (+3 cases); `xl-cli/src/com/tjclp/xl/cli/Main.scala` (+3 subcommands in the read-only group, execute arms, `describe` under `--stream`); `xl-cli/src/com/tjclp/xl/cli/commands/ReadCommands.scala` (`cell`, `:253-266`); new `xl-cli/test/src/com/tjclp/xl/cli/InspectCommandsSpec.scala`; `xl/test/src/xlprelude/ScriptingPreludeTest.scala`; `docs/reference/cli.md`; `plugin/skills/xl-cli/SKILL.md`
**DependsOn**: envelope-json, recalc-options-renamer, harness-golden
**ParallelGroup**: 4
**HoursEstimate**: 5
**ClosesIssues**: — (J1 "orient" and J5 "trace a broken number" become one call each)

**Implementer brief**:

1. `QualifiedGraph.of(wb)` = `DependencyGraph.fromWorkbookFormulaGraph(wb)` (`private[formula]`,
   reachable from `formula.graph`) packaged with `fromWorkbookDependencyIndex(wb)` (`private[xl]`)
   for `dependents`; `precedents/dependents(q, depth)` BFS by layer (`depth = 0` = unbounded;
   layer k = refs exactly k hops away, deduplicated against earlier layers); `sccs` from
   `qualifiedCyclicNodes`/`Scc`. Bounded and symbolic — an `A:A` reader never expands to a million
   refs (assert a node-count bound in the spec). Do not edit `DependencyGraph.scala`.
2. `WorkbookAudit.of(wb)` per ADR-017 §2.10 from existing functions: cached Excel error values
   (`CellValue.Formula(_, Some(CellValue.Error(e)), _)` and bare `Error` cells), uncached formulas,
   unparseable formulas (`FormulaParser.parse` `Left` with `ParseError.formatWithContext`),
   volatile (function names in `Set("TODAY","NOW","RAND","RANDBETWEEN")` — there is no
   `FunctionFlags.volatile`; adding one is a Wave 2 chore), dynamic (`DependencyGraph.dynamicCells`),
   cycles (`QualifiedGraph.sccs.filter(_.cyclic)`), external refs (`TExpr.ExternalRef/ExternalRange`
   present), unresolved readers (`DependencyGraph.unresolvedReaders`), `calcPr`; `isClean` = every
   bucket except `volatile`, `dynamic`, `externalRefs`, `calcPr` is empty.
3. `WorkbookSummary.of(wb)` per ADR-017 §2.10 (`SheetSummary` per sheet; `state` as
   `WorkbookCommands.sheets` reports it; counts are folds over `Sheet` fields; `hiddenRows/
   hiddenCols` from `rowProperties`/`columnProperties`).
4. `exports.scala` (additive, after the renamer's lines): `export formula.graph.QualifiedGraph`,
   `export formula.eval.{WorkbookAudit, WorkbookSummary}`, plus `extension (wb: Workbook) { def
   describe: WorkbookSummary; def audit: WorkbookAudit }` in a new `object WorkbookInspect` in
   `WorkbookSummary.scala`, exported as `WorkbookInspect.*` (no default args).
5. xl-cli: `CliCommand.Describe(full: Boolean)`, `Audit(failOnFindings: Boolean)`,
   `Deps(ref: String, direction: String, depth: Option[Int])` (`direction` ∈
   `precedents|dependents|both`, default `both`; `depth` `None` = 1, `0` = all — accept the literal
   `all`). `InspectCommands.describe(wb, full): IO[Outcome]`, `describeLight(meta: LightMetadata):
   IO[Outcome]` (the `--stream` and no-`--full` path: sheets, state, dimension, defined names with
   `hidden`, `date1904` — from `WorkbookMetadataReader`), `audit(wb, sheetOpt, failOnFindings)`
   (exit 1 `AUDIT_FINDINGS` when asked and `!isClean`; text mode prints one section per non-empty
   bucket), `deps(wb, sheetOpt, ref, direction, depth)` (nodes `{ref: "'Sheet'!A1", depth,
   formula?, value?}` using `SheetName.quoteForFormula`; resolves the ref through
   `SheetResolver.resolveRef` — it inherits the one rule when `resolve-and-argv` lands). Add the
   three `Opts.subcommand`s to the read-only group (`sheetReadOnlySubcmds`), the `execute` arms
   (`describe` without `--full` is metadata-only and works under `--stream`), and JSON payloads
   through `Payload.Json`. `ReadCommands.cell`: replace the `fromWorkbook` + reverse fold
   (`:253-266`) with `QualifiedGraph.of(wb)` — text output unchanged (pin it).

**Tests first**: `QualifiedGraphSpec` — cross-sheet chain A→B→C: `precedents(C, 2)` is two
layers, `dependents(A, 0)` = {B, C}; an `A:A` reader stays under a node bound; `sccs` finds a
two-cycle. `WorkbookAuditSpec` — a fixture with one cached `#DIV/0!`, one uncached formula, one
`TODAY()`, one `INDIRECT`, one 2-cycle, one `'[1]Ext'!A1`, one `UNSUPPORTED(1)` populates exactly
one entry per bucket; a clean book is `isClean`. `WorkbookSummarySpec` — counts on
`TestFixtures`-shaped books and on three fixtures of the 20-workbook corpus (never errors).
`InspectCommandsSpec` via the harness: `describe --json` on a two-sheet fixture (hidden name
flagged); `--stream describe` returns metadata only; `describe --full` adds counts; `audit --json`
buckets; `audit --fail-on-findings` exits 1 on a dirty book and 0 on a clean one; `deps Sheet2!A1
--json` lists `Sheet1!A1` as a precedent with `value`; `cell` text output unchanged. Probes:
`wb.describe`, `wb.audit`, `QualifiedGraph.of(wb).precedents(...)`.
**Existing tests that may need adjusting**: none (new verbs); pin `cell`'s current text in the
new spec before rewriting it.
**Docs/skill files to touch**: `docs/reference/cli.md` `## Commands` table + three new sections
under `## Command Details`; `plugin/skills/xl-cli/SKILL.md` `## Essential Patterns` gains a
task→verb table (orient → `describe`; broken number → `audit`, `deps`).
**Integrator notes**: CHANGELOG `### Added` — "`describe [--full]`, `audit
[--fail-on-findings]`, `deps <ref> [--direction] [--depth]`; `WorkbookSummary`, `WorkbookAudit`,
`QualifiedGraph` in xl-evaluator with `wb.describe`/`wb.audit` in the prelude; `cell` uses the
bounded graph." Record the test count.

---

### Cluster resolve-and-argv

**Status**: pending
**Modules**: xl-cli
**Files**: new `xl-cli/src/com/tjclp/xl/cli/helpers/Resolve.scala`; `xl-cli/src/com/tjclp/xl/cli/helpers/SheetResolver.scala`; `xl-cli/src/com/tjclp/xl/cli/helpers/BatchParser.scala` (the 17 `requires --sheet` sites and the op-scope helper from `batch-opspec-scope`); `xl-cli/src/com/tjclp/xl/cli/commands/StreamingReadCommands.scala` (`resolveSheetName`, `cell`'s first-sheet default at `:130-136`); `xl-cli/src/com/tjclp/xl/cli/commands/StreamingWriteCommands.scala` (`resolveSheetPath`); new `xl-cli/src/com/tjclp/xl/cli/contract/Argv.scala`; `xl-cli/src/com/tjclp/xl/cli/Cli.scala`; new `xl-cli/test/src/com/tjclp/xl/cli/contract/{ResolveSpec,ArgvSpec}.scala`; `xl-cli/test/src/com/tjclp/xl/cli/{StreamingReadSpec,StreamingWriteSpec}.scala`; goldens; `docs/reference/cli.md`; `plugin/skills/xl-cli/SKILL.md`
**DependsOn**: errors-exit-codes, batch-opspec-scope, harness-golden
**ParallelGroup**: 5
**HoursEstimate**: 4
**ClosesIssues**: — (retires the three sheet policies, the 17 hand-written `requires --sheet`, and the flags-before-the-verb trap)

**Implementer brief**:

1. `Resolve.scala` per ADR-017 §2.5: `Target`, `Resolved`, `target(wb, sheetFlag, refStr, verb)`,
   `sheet(wb, sheetFlag, verb)`, `sheetName(meta, sheetFlag, qualified, verb)`, `forOp(wb,
   defaultSheet, opSheet, qualified, index, op)`. THE rule: qualified ref > `-s` (> op `sheet` for
   batch) > the only sheet of a single-sheet book (emit `Warning(SHEET_AUTOSELECTED)` in `--json`
   mode only) > `CliError.fromXLError(XLError.SheetRequired(context, names))` (exit 3). Message
   texts: keep the two legacy strings verbatim as `message` for `SHEET_REQUIRED`/`SHEET_NOT_FOUND`.
2. `SheetResolver.{resolveSheet, requireSheet, findSheet, resolveRef}` become one-line forwarders
   (`IO.fromEither(... .left.map(CliException(_)))`). The 17 `BatchParser` sites route through
   `Resolve.forOp`; `StreamingReadCommands.resolveSheetName` and the `cell` first-sheet default
   route through `Resolve.sheetName(meta, …)` (`excel.readMetadata`); `StreamingWriteCommands.resolveSheetPath`
   takes its *name* from `Resolve.sheetName` (a qualified ref now selects the sheet under `--stream`
   writes) and keeps only the zip lookup. The streaming first-sheet default and the "Multiple
   sheets found" refusal are replaced by the rule: single-sheet books auto-select, multi-sheet books
   without a sheet are `SHEET_REQUIRED`.
3. `Argv.scala`: `val globals: Map[String, Boolean]` (`-f/--file` true, `-s/--sheet` true,
   `-o/--output` true, `-i/--in-place` false, `--backend` true, `--max-size` true, `--stream`
   false, `--no-recalc` false, `--preserve-caches` false, `--json` false, `--strict` false);
   `hoist(args)`: walk tokens; a recognised global (and its value, also the `--flag=value` form) is
   moved in front of the first non-flag token (the verb); `--strict` is hoisted unless
   `verbOf(args) == Some("view")` (the two `--strict` flags at `Main.scala:295-301` and
   `:320-321`); everything after `--` is untouched; non-global tokens keep their relative order.
   `verbOf(args)`: the first token that is not a global, a global's value, or `--help/--version`.
   Wire `Cli.run(args, io)`: `hoist` → if `verbOf` is defined and not one of the 56 verb names (a
   literal `Vector` in `Argv`, asserted against the subcommand names in `xl --help` by `ArgvSpec`)
   → `UNKNOWN_VERB` with `Suggest.closest`, one-line usage, exit 2; decline parse failure → stderr
   one-line usage `usage: xl [-f FILE] [-s SHEET] [-o OUT | -i] [--json] <verb> …` + the first
   decline error + `Run 'xl <verb> --help'`, exit 2 — never the 65-subcommand dump. Verify
   decline's own `--` handling through the harness (`put A1 -- -5`); implement it in `hoist` only if
   decline does not stop there.

**Tests first**: `ResolveSpec` — the four steps as unit tests; property over
`Generators.genSheetName`: a qualified ref wins regardless of `-s`; `forOp` conflict rule.
`ArgvSpec` — ScalaCheck: `hoist(hoist(a)) == hoist(a)`; non-global order preserved; `--`
respected; `view … --strict` not hoisted; `recalc --strict` hoisted; `--file=x` form; the verb
list equals the subcommand names decline renders. Harness: "single-sheet book: `view A1:B2`,
`cell A1`, `stats A1:A3`, `put A1 5`, `--stream view A1:B2` all work without `-s`"; "two-sheet
book: the same commands exit 3 `SHEET_REQUIRED` with candidates, in memory and under `--stream`";
"`--stream put 'Data'!A1 5` works"; "`--stream view` with `-s X` and a different qualified sheet
uses the qualified one"; "`xl view A1:B2 -f f -s Data` ≡ `xl -f f -s Data view A1:B2`"; "`xl -f f
recalc --strict -o out` gates"; "`xl -f f viwe A1` exits 2 with `did you mean: view` and fewer
than 10 lines"; "`put A1 -- -5` stores -5"; "batch op on `Sheet2!A1` after `rename-sheet` of the
default resolves by the new identity".
**Existing tests that may need adjusting**: `StreamingWriteSpec` cases pinning "qualified refs
rejected under `--stream`" (now accepted) and "Multiple sheets found" (now `SHEET_REQUIRED`);
`StreamingReadSpec` cases relying on the first-sheet default on a multi-sheet fixture (now
`SHEET_REQUIRED`) — add a GH reference in each; the two goldens pinned in `harness-golden` for the
decline-error shapes and the streaming first-sheet default now change (the reviewed diff).
**Docs/skill files to touch**: `docs/reference/cli.md:40` and `plugin/skills/xl-cli/SKILL.md:214`
state the one rule (both currently contradict the code); delete the flags-before-the-verb gotcha
(`cli.md:74-75`, `SKILL.md:646-647`) and document the `view --strict` exception; delete the
`'"'"'` quoting recipe at `SKILL.md:293` in favour of `sheet` on the op (the `@file`/`-` channel is
Wave 2). Fix the CLAUDE.md-derived "sheet auto-detected if unambiguous" wording in `cli.md:55` to
"single-sheet books auto-select for every verb".
**Integrator notes**: CHANGELOG `### Added` — "Global flags are accepted anywhere on the command
line (`--strict` after `view` stays view's own); unknown verbs suggest the nearest name; one sheet
rule for every verb, batch op and streaming path: qualified ref, then `-s`, then the only sheet of a
single-sheet book, otherwise `SHEET_REQUIRED` with candidates; qualified refs work under `--stream`
writes." `### Changed` — "Streaming reads on multi-sheet books without a sheet exit 3
`SHEET_REQUIRED` instead of silently using the first sheet." Fix the `putf … --from B2` example in
CLAUDE.md with `docs-from-code`.

---

### Cluster twins-and-navigation

**Status**: done — `90b9170`, `3dc2d4e` (corner-form-only contract documented after review)
**Modules**: xl-core, xl (probes)
**Files**: `xl-core/src/com/tjclp/xl/sheets/Sheet.scala` (additive: `putAt` ×2, `styleAt`, `mergeAt`, `commentAt`; Scaladoc of the transparent forms names the twin); `xl-core/src/com/tjclp/xl/workbooks/Workbook.scala` (companion: `named` ×2); `xl-core/src/com/tjclp/xl/addressing/ARef.scala` (`tryShift`, `tryDown`, `tryRight`, `clampShift`); `xl-core/src/com/tjclp/xl/addressing/CellRange.scala` (`rows`, `columns`, `row`, `column`); `xl-core/src/com/tjclp/xl/styles/cellstyle.scala` (`withUnderline`); new `xl-core/test/src/com/tjclp/xl/sheets/RuntimeTwinsSpec.scala`; new `xl-core/test/src/com/tjclp/xl/addressing/BoundedNavigationSpec.scala`; `xl-core/test/src/com/tjclp/xl/StyleSpec.scala` (+1); `xl/test/src/xlprelude/ScriptingPreludeTest.scala`; `docs/reference/scripting.md`; `plugin/skills/xl-scripting/SKILL.md`
**DependsOn**: —
**ParallelGroup**: 5
**HoursEstimate**: 3
**ClosesIssues**: #465 (the `withUnderline` half; `collapseRows/Cols` is Wave 2)

**Implementer brief**:

1. `Sheet`: `def putAt[A: CellWriter](ref: String, value: A): XLResult[Sheet]`, `def putAt[A:
   CellWriter](ref: String, value: A, style: CellStyle): XLResult[Sheet]`, `def styleAt(ref: String,
   style: CellStyle): XLResult[Sheet]` (a cell → `withCellStyle`, a range → `withRangeStyle`), `def
   mergeAt(range: String): XLResult[Sheet]`, `def commentAt(ref: String, comment: Comment):
   XLResult[Sheet]`. All parse through `RefType.parseToXLError` (`RefType.scala:147`); a
   `QualifiedCell`/`QualifiedRange` is `Left(XLError.InvalidReference("sheet-qualified refs are not
   accepted by Sheet.putAt/styleAt/mergeAt/commentAt; use wb.update(sheet, …)"))`; a range where a
   cell is required is `Left(InvalidCellRef(ref, "expected a single cell"))`. The `put`/`style`/
   `merge`/`comment` transparent forms are unchanged; their Scaladoc gains "for a name computed at
   runtime prefer `putAt`" in the style of `Sheet.named`'s doc (`Sheet.scala:1268-1289`).
2. `object Workbook`: `def named(name: String): XLResult[Workbook]` and `def named(first: String,
   second: String, rest: String*): XLResult[Workbook]` — `SheetName(_)` each, `InvalidSheetName`
   on failure, `DuplicateSheet` on repeats; semantically the dynamic branch of the transparent
   `apply` (`Workbook.scala:588-612`).
3. `object ARef` extension block (non-inline, see the issue #252 note at `ARef.scala:56-59`): `def
   tryShift(colOffset: Int, rowOffset: Int): Option[ARef]` (`None` when the result would leave
   column 0..16383 or row 0..1048575), `def tryDown(n: Int): Option[ARef]`, `def tryRight(n: Int):
   Option[ARef]`, `def clampShift(colOffset: Int, rowOffset: Int): ARef`. `shift` is unchanged.
4. `CellRange`: `def rows: Iterator[CellRange]`, `def columns: Iterator[CellRange]`, `def row(i:
   Int): Option[CellRange]` (0-based within the range), `def column(i: Int): Option[CellRange]`.
5. `CellStyle.withUnderline(u: Underline): CellStyle = withFont(font.withUnderline(u))`
   (`Font.withUnderline(u: Underline)` exists at `Font.scala:21`).

**Tests first**: `RuntimeTwinsSpec` — `Sheet.named(s).flatMap(_.putAt(r, 1))` equals the literal
form for valid strings; each twin is `Left(InvalidReference)` for a qualified ref and
`Left(InvalidCellRef/InvalidRange)` for garbage; `styleAt` on a range styles every cell;
`Workbook.named("A","B")` has two sheets, `named("A","A")` is `Left(DuplicateSheet)`.
`BoundedNavigationSpec` — `tryDown(1)` at row 1048576 is `None`; `tryShift(-1, 0)` at column A is
`None`; property `tryShift(dc, dr).flatMap(_.tryShift(-dc, -dr)) == Some(ref)` when in bounds;
`clampShift` is idempotent; `rows.size == height`, `row(i)` is `None` out of range, `columns`
mirrors. `StyleSpec` — `withUnderline(Underline.Double).font.underline == Underline.Double`.
Probes: `Sheet.named(s).flatMap(_.putAt(r, 1))`, `styleAt`, `mergeAt`, `commentAt`,
`Workbook.named`, `ref"A1".tryDown(1)`, `ref"A1:B3".rows.size == 3`, `CellStyle.default.withUnderline(Underline.Single)`
all resolve through the prelude (ARef extensions resolve via the companion implicit scope; the
probe is what proves it).
**Existing tests that may need adjusting**: none (additive).
**Docs/skill files to touch**: `docs/reference/scripting.md` `## Compile-time vs runtime refs`
(the two rules: literals are total, computed strings use `*At`/`named`; bounded navigation);
`plugin/skills/xl-scripting/SKILL.md` `## Essential Patterns` and `## Gotchas` (stop teaching the
transparent forms for computed strings).
**Integrator notes**: CHANGELOG `### Added` — "Runtime twins `Sheet.putAt/styleAt/mergeAt/
commentAt` and `Workbook.named` (explicit `XLResult`), bounded navigation
`ARef.tryShift/tryDown/tryRight/clampShift`, `CellRange.rows/columns/row/column`,
`CellStyle.withUnderline` (#465)."

---

### Cluster docs-from-code

**Status**: pending
**Modules**: xl-cli, docs, plugin, CI
**Files**: `xl-cli/src/com/tjclp/xl/cli/Main.scala` (`functions --json`; `schema [--json]` verb; `batch --schema` flag on `batchCmd`; `batchHelp` body from `OpRegistry.helpText`); new `xl-cli/src/com/tjclp/xl/cli/contract/{FunctionDoc,Schema}.scala`; new `xl-cli/test/src/com/tjclp/xl/cli/contract/{DocsGenSpec,FunctionDocSpec,SchemaSpec}.scala`; new `docs/reference/generated/{cli-verbs,batch-ops,functions,exit-codes,error-codes}.md`; `docs/reference/cli.md`; `plugin/skills/xl-cli/SKILL.md`; `plugin/skills/xl-cli/reference/FORMULAS.md`; `CLAUDE.md` (counts and the `putf --from` example only; test-count lines stay with the integrator); `plugin/.claude-plugin/plugin.json`; `.github/workflows/ci.yml`; goldens
**DependsOn**: envelope-json, batch-opspec-scope, resolve-and-argv, describe-audit-deps
**ParallelGroup**: 6
**HoursEstimate**: 3.5
**ClosesIssues**: — (retires the 108/112/115, 17/26/27/32 and 0.19.1/0.19.3 drift)

**Implementer brief**:

1. `FunctionDoc(name: String, minArgs: Int, maxArgs: Option[Int], args: List[String], returnsDate:
   Boolean, returnsTime: Boolean, dynamicDeps: Boolean, specialForm: Boolean)` from
   `FunctionRegistry.all` (macro-collected; evaluate once in a `lazy val`); `args` from the arg
   spec's `describeParts`, `maxArgs = None` for variadic specs; append `LET` with `specialForm =
   true` so the count reads 115 + LET. `xl functions` text unchanged; `xl functions --json` →
   `Payload.Json`.
2. `Schema.json(version): ujson.Value` = `{version, exitCodes: [{code, meaning}], errorCodes:
   ErrorCode.all, warningCodes, globals: [{name, short, takesValue, doc}] from Argv.globals, verbs:
   [{path, summary, needs: {file, sheet, output, streaming}, exit, batchTwin, since}], batchOps:
   OpRegistry.jsonSchema(version), functions: [FunctionDoc…]}`. The verb table is a literal
   `Vector[VerbDoc]` in `Schema.scala` — Wave 2 derives it from the `CommandSpec` registry;
   `SchemaSpec` asserts its `path.head`s equal the subcommand names decline renders in
   `Command("xl","test")(Main.main).showHelp` (65 declarations → 56 distinct names + 8 nested), and
   that every `batchTwin` names an `OpRegistry` op. `xl schema --json` prints it; `xl schema` prints
   the verb table. `xl batch --schema` prints `batchOps` alone (flag on `batchCmd`, no `-f`
   required, same shape as `--dry-run`); `batchHelp` (`Main.scala:1163-1224`) lists all 32 ops from
   `OpRegistry.helpText` instead of 17 by hand.
3. `DocsGenSpec` renders the five generated markdown files from `Schema.json` and compares them
   with the checked-in files; `XL_UPDATE_DOCS=1` rewrites. Same discipline as goldens: a diff is a
   contract change.
4. `docs/reference/cli.md`: replace the hand-written command summary table (`:165-240`) and the
   batch-op table (`:1104` region) with links/includes of the generated files; fix "108 supported
   functions" (`:70`); add the `sheet` key, aliases, `--schema`, the GH-560 note, the `rename-sheet`
   rewrite line (handed over by `batch-opspec-scope` and `recalc-options-renamer`); keep the prose.
   `plugin/skills/xl-cli/SKILL.md`: cut to a routing layer — install, the mental model (globals
   anywhere; the sheet rule; `-o`/`-i`; `--json` always; `xl <verb> --help`; batch is the primary
   write surface; the exit table), a task→verb table, current-version gotchas only (delete the
   `On ≤0.12.x` clauses in `## Field Gotchas`), `"26 operations"` (`:742`) → "run `xl batch --schema`",
   the `rename-sheet` gotcha removed, links to the generated reference; add `Requires xl >= 0.20.0`
   and an `xl --version` check. `reference/FORMULAS.md` regenerated from `functions --json`.
   `CLAUDE.md`: fix the `putf … --from B2` example (the verb drags from `range.start`; there is no
   `--from`), point the "All 32 batch operations" and "115 Functions" lists at the generated files.
5. `plugin/.claude-plugin/plugin.json` `version` → the value in `build.mill:17` (`0.19.3` today;
   `/release-prep` bumps both); add a step to `.github/workflows/ci.yml` that fails when the two
   differ (`grep`/`jq`, no build needed).

**Tests first**: `DocsGenSpec` (drift fails); `FunctionDocSpec` — "115 registry functions + LET",
"every FunctionDoc has args", "`functions --json` validates"; `SchemaSpec` — verb names equal
decline's, every error code in the table maps to an exit code, `batchOps` has 32 `oneOf` entries;
harness: "`schema --json`", "`batch --schema` is valid JSON Schema", "`batch --help` lists 32 ops",
"`functions --json`".
**Existing tests that may need adjusting**: the `batch --help` golden changes (expected diff).
**Docs/skill files to touch**: this cluster is docs (listed above).
**Integrator notes**: CHANGELOG `### Added` — "`xl schema [--json]`, `xl batch --schema`, `xl
functions --json`; `docs/reference/generated/*` rendered from the binary and CI-gated;
`plugin.json` version checked against `build.mill`." Then the integrator: consolidate the Wave 1
CHANGELOG entries under `## [Unreleased]` (Keep-a-Changelog form), add "Unreleased — wave 27:
agent-first contract (0.20.0)" to `docs/plan/roadmap.md` listing the ten clusters and the Wave 2–4
items below with issue links once filed, add the ADR-017 index line to `docs/design/decisions.md`,
and set the test counts in `CLAUDE.md:103/409/414`, `docs/STATUS.md:212`,
`docs/reference/testing-guide.md:225` from the `__.test` run (expected: 5,625 + 26 from `0d53e76`
+ every cluster's additions; never estimate).

---

## Wave 2 — the algebra underneath the contract (GitHub-issue-ready)

Each item is one issue: title, modules, scope, acceptance. All are additive to the Wave 1
contract; the Wave 1 goldens are the acceptance test that nothing an agent sees changed unless the
item says "Changed".

**W2.1 `enum Edit` algebra in xl-core with a refusing `FormulaSupport`** · xl-core, xl-evaluator,
xl · Scope: `com.tjclp.xl.ops` per ADR-017 §2.12 — `Edit` (one case per batch op plus the 15
verbs without a twin), `Loc`/`Area`/`ColSpan`/`RowSpan` parsers over `RefType.parseToXLError`,
`Scope` with rule 4 (rename updates the default), `FormatHint = Inferred | Explicit`,
`StyleOverlay` with `Option` fields and per-side border colour, `FormulaSupport` trait +
`textOnly` that **refuses** shift/structural/rename with `UnsupportedCapability` (no `given
default`), `EvalFormulaSupport` in xl-evaluator wrapping `StructuralEditor.*Checked`,
`FormulaOps.shift`, `SheetRenamer.rename`, `FormulaParser.parse`; `Edit.validate/plan/applyAll/
lower`; `Patch.toEdits(p, sheet)`; `EditSchema` (JSON-free) with the `OpSpec` field shape; sort/
fill/copy/autofit/group/appearance semantics **moved** (not copied) from `WriteCommands`/`CopyOps`/
`GroupingOps`/`AppearanceOps`/`ColumnAutoFit` into `Sheet` methods, the CLI originals becoming
forwarders in the same PR; the prelude exports `Edit`, `wb.edit(edits*)(using FormulaSupport)`,
`sheet.edit`, and `given FormulaSupport = EvalFormulaSupport` explicitly. Acceptance: the seven
laws (identity, fold, idempotence class, lowering coherence, desugar coherence, determinism,
all-or-nothing) as ScalaCheck properties over a `genEdit` covering every case; `PatchSpec`'s 30
laws unchanged; `EditSchema.all.size == number of cases`; no `FillDirection/SortDirection/SortMode`
exported from `api`; every Wave 1 golden unchanged; an `EditPreludeProbe` outside `com.tjclp.xl`.
Size honestly: four clusters (model+schema, interpreter+laws, formula support, prelude).

**W2.2 Batch on the algebra: twins, verb lowering, `--print-batch`** · xl-cli · Scope: `EditJson`
codec (`decode(encode(e)) == Right(e)`, all errors collected with op indexes, aliases from
`EditSchema`, `sheet`/qualified-ref conflict rule); `BatchParser` becomes the adapter (`BatchOp.toEdit`,
`Edit.toBatchOp`, `applyScoped` = `traverse(toEdit) → Edit.applyAll`; the `apply*` helpers
deleted); the 15 missing twins (`insert-rows`, `delete-rows`, `insert-cols`, `delete-cols`, `fill`,
`sort`, `import`, `import-md`, `remove-sheet`, `move-sheet`, `copy-sheet`, `image`, `name`,
`sheets-hide`, `sheets-show`) plus `add-sheet.before`, `name` with `scope` (#462); `Lowering.toEdits(cmd:
CliCommand, scope): XLResult[Vector[Edit]]` for every mutating verb (`putf <range> <one formula>`
lowers to `DragFormula(anchor = range.start)`); `--print-batch` on every mutating verb prints the
lowering and the replay command; `--strict-input` promotes `UNKNOWN_PROPERTY`/`FORMAT_HINT_IGNORED`
to exit 2 (the default stays a warning); `--dry-run` with `-f` runs `Edit.plan` (semantic dry-run).
Acceptance: `BatchParitySpec` — every JSON literal in the existing batch specs decodes to the same
`Edit` through both paths; `LoweringParitySpec` — every mutating `CliCommand` lowers and every
`Edit` case is produced by some lowering or is declared batch-only; `xl batch --schema` validates
every example; `WriteCommands`/`BatchParser` shrink by ≥ 1,500 lines. Closes #462, GH-463's CLI
half.

**W2.3 Registry-driven `Main` with `Needs` validation** · xl-cli · Scope: `CommandSpec(path,
summary, help, needs: CliCommand => Needs, exit, batchTwin, since, opts)`, `Registry.all` (one per
verb), one dispatch group `(globalOptsAll, Registry.verbOpts).mapN(dispatch)` replacing the eleven
`orElse` groups; `Needs(file, sheet, output, streaming, policy)` validated **before** the read
(`OUTPUT_REQUIRED` exits 2 without touching the file; `--stream` on a `BackendOnly` verb → a
`STREAM_BACKEND_ONLY` warning instead of a silent "Saved (streaming)"; on an `Unsupported` verb →
usage 2); typed `Flag` descriptors feed `schema --json` (replacing the Wave 1 verb table) and
per-verb `--help` headers; `Argv.hoist`'s verb-ownership from the registry (the `view --strict`
exception becomes data); `--strict` reaches `audit` (retiring `--fail-on-findings` as an alias);
`putf <ref> @file` / `putf <ref> -` / `put <ref> @file` input channels with one documented quoting
rule; `xl --version --json`; `XL_JSON=1` default. Acceptance: `RegistrySpec` — every subcommand
has a spec, every spec's `--help` parses, `needs` are consistent; every Wave 1 golden unchanged
except the two decline-error shapes and the `view -o x` message; `MainSpec`/`InPlaceSpec`/
`LintCommandSpec` argv tests green.

**W2.4 `CellRecord` and `SheetSource`: one projection, streaming as a strategy** · xl-cli,
xl-cats-effect · Scope: `CellRecord(ref, sheet, kind, value, formatted, formula: Option[FormulaInfo],
hidden, mergedInto, style)` with `toJson(legacyKeys = true)` byte-identical to today's
`view --format json`; `JsonRenderer`, `CsvRenderer`, `Markdown` and the streaming renderers consume
`Vector[Vector[CellRecord]]`; `SheetSource` (`inMemory(wb)` / `streaming(path, excel)`) with a
`capabilities` set published by `schema --json`; read verbs become `ReadQuery => SheetSource =>
IO[Outcome]`; typed `--json` payloads for `search`, `stats`, `cell`, `filter`; `view` default range
= used range with `--offset`/`--max-cols`; **Changed**: streaming `view --format json` becomes the
typed in-memory shape (never documented as stable). Acceptance: property over generated books —
for every read verb, in-memory and streaming produce equal payloads when the book uses only shared
capabilities; unsupported combinations are in-band `UNSUPPORTED_IN_STREAM`; `filter --stream`
works; the nine scalar textualisers and three JSON/CSV escapers are gone.

**W2.5 Streaming writes over the algebra** · xl-cli, xl-cats-effect · Scope: `StreamingWriteCommands`
consumes `Vector[Edit]`, declares support from `EditSchema.streamable`, refuses by index before
opening the zip (already the batch behaviour from Wave 1; now also the verb path: `--stream putf
<range> <one formula>` is refused instead of filling the same formula), parse-validates formulas
through `EvalFormulaSupport.validate`, honours `StyleMode.Merge` via `StylePatcher.getStyle`,
deletes the `*Sync` style clones (`StreamingWriteCommands.scala:968-1081`) in favour of
`StyleOverlay`, resolves sheets through `Resolve.sheetName`. Acceptance: property `stream(edits) ≡
memory(edits)` on the streamable subset for generated edits (cell XML equality); `StreamingWriteSpec`
green; no divergence sentence left in `docs/reference/cli.md`.

**W2.6 `Recalc.afterEdits` promotion** · xl-evaluator, xl-cli, xl · Scope: `RecalcPolicy(scope:
Cone | Whole | None, strict: Boolean)` and `Recalc.afterEdits(before: Workbook, applied: Applied,
policy: RecalcPolicy, options: RecalcOptions): RecalcOutcome` absorbing `recalcHonoringCalcPr`,
`changedRefs`, `dirtyCone`, `applyConeCaches`, `scopeToCone`, `scopedRecalc`, `strictGate`,
`formatRecalcSummary` from `WriteCommands.scala:996-1229`; name edits (`DefineName`/`RemoveName`)
treated as `Whole` (closing the cone-soundness gap for formulas reading a redefined name);
`WriteCommands.writeAfterRefresh` and `batch` become thin callers; `Excel.writeRecalculated` uses
the same policy; `RecalcResult.summary` is the only summary text; the `@targetName` lattices are
`@deprecated` bridges; `FormulaParser.levenshteinDistance` delegates to `Suggest`;
`SheetEvaluator.evalErrorToXLError` (`cell_<Long>` rendering) deleted in favour of
`EvalError.toXLError` with `'Sheet'!A1` names. Acceptance: `BatchRecalcSpec`,
`StrictWriteIntegritySpec`, `AfterEditRecalculationSpec` unchanged; `WriteCommands.scala` loses
≥ 350 lines; `xl recalc`, batch tails, `view --eval` and `writeRecalculated` are projections of one
function.

**W2.7 `_xlfn.` prefixes (#556) — only after PR #576 lands** · xl-evaluator, xl-ooxml · Scope:
the writer prefixes the post-2007 function set in `<f>` (a table on `FunctionFlags`, e.g.
`since2010`), the parser strips `_xlfn.`/`_xlws.` on read, `FormulaPrinter` round-trips, lint
`xlfn-missing`; also `FunctionFlags.volatile` for `WorkbookAudit` (touches `FunctionSpecs*` — run
`./mill clean xl-evaluator.compile`). Acceptance: the issue's repro opens in Excel without
`#NAME?`; parse∘print identity on the fixture corpus; `WorkbookAudit.volatile` uses the flag.

**W2.8 Scripting completions without a default flip** · xl-cats-effect, xl, xl-core · Scope:
additive `Excel.writeChecked(wb, path, options: RecalcOptions): RecalcResult` (recalculates
uncached cells via `recalculateUncached` and writes; `Excel.write` keeps 0.19 semantics),
`Excel.readSheet(path, sheet): Sheet` (`XLException` naming candidates), `Excel.readMetadata(path):
LightMetadata`, `Excel.modifyR(path)(f: Workbook => XLResult[Workbook]): Unit`, `orExit[A](r:
XLResult[A]): A`; the `ExcelRecalc` overloads stay; `Sheet.collapseRows/collapseCols` (#465's second
half); `CodecError.UncachedFormula` considered again with the exhaustive-match sites
(`EvalError.scala:138-142`, `SheetEvaluator.scala:695-697`) listed; the xl-scripting skill teaches
`writeChecked`/`writeRecalculated` and never `write` for freshly built models. Acceptance: probes
for every name; `scripts/test-examples.sh` and `verify-skill-snippets.sh --local` green; no
`Excel.write` call in docs/examples writes a formula without a cache.

**W2.9 `RowCodec` records** · xl-core, xl · Scope: `trait RowCodec[A]` with `RowCodec.derived` via
`Mirror.ProductOf`, `Option[T]` fields as empty cells, `RowCodecError.Field(row, column, field,
cause)`, `Sheet.readRows[A](range)`, `readRowsByHeader[A](headerRow)`, `headers(row)`,
`column(header, headerRow)`, `putRows`/`putRowsWithHeader`/`putTable`; reads through
`effectiveValue`. **Gate**: the prelude probe (`case class Order(...) derives RowCodec` through
`com.tjclp.xl.scripting`) is written and green before the cluster is scheduled — the inline/export
landmine is exactly what it tests; fallback is a non-inline `RowCodec.of[A](codecs*)`. Acceptance:
round-trip law `readRows(putRows(at, rows).range) == Right(rows)` over generators; ADR-008 amended
("primitives hand-written, row codecs derived").

**W2.10 Warnings never silent; `Written` payload** · xl-cli · Scope: write handlers gain `*Report`
twins returning `WriteReport(text, saved, changes: Vector[{sheet, target, before, after}], recalc,
warnings)` beside the `IO[String]` entry points; `Payload.Written`; the 12 handler-level
`System.err.println` sites replaced by the warnings sink; `renderWithTarget`/`unsayTheSave` deleted
(the report carries `saved`/`written` as data). Acceptance: `put --json` returns
`data.changes[0].before/after`; the `--strict` exit-1 payload carries `recalc.errors`; no
`System.err.println` remains under `xl-cli/src`; the `CliHarness` `setErr` bracket is removed.

**W2.11 Harness and goldens in CI; JAR smoke** · CI, xl-agent · Scope: `ci.yml` runs the golden
runner explicitly and a JAR smoke (`--help`, `sheets`, `view --json`, `schema --json`, `lint`) on
every PR with `timeout-minutes`; `release.yml` smokes the native binary the same way; the xl-agent
grader reads `sheets --json` instead of scraping the markdown table (`Evaluator.scala:82-95`) and
locks the skill to the binary version; `ContractSpec` enumerates every error/warning code against
the generated tables. Acceptance: a golden diff fails CI with the unified diff in the log; the
grader has no markdown parsing left.

## Wave 3 — self-describing, packaged, one vocabulary (GitHub-issue-ready)

**W3.1 Errors as one vocabulary end to end** · xl-core, xl-cats-effect, xl-cli · Scope:
`ExcelIO` raises `XLException(XLError)` instead of `new Exception(String)` at its ten wrap sites
(`ExcelIO.scala:109,120,241,251,260,294,449,464,494,512`) so `IO_READ`/`SECURITY_LIMIT`/`PARSE_ERROR`
codes come from the domain; `CodecError` folded into `XLError` with a bridge (`readTyped:
XLResult[Option[A]]` behind a deprecation); the six dead `XLError` cases (`InvalidColumn`,
`InvalidRow`, `NumberFormatError`, `StyleError`, `ColorError`, `UnsupportedType`) deprecated;
`Either[String, _]` parsers in `addressing` (`Column.parse`, `SheetName.apply`) gain `XLResult`
twins. Acceptance: `CliError.fromThrowable` has no `INTERNAL` fall-through for any exception xl
itself raises; probes for the twins.

**W3.2 Retire the `transparent inline` union entry points** · xl-core, xl · Scope: the 20 sites
(`Sheet.put/comment/apply(String…)`, `Workbook.apply/put/remove/update/upsert/delete(String…)`,
`extensions.style/cell/range/get/merge`) become `@deprecated` in 0.21 with messages naming the twin
(`putAt`, `styleAt`, `mergeAt`, `commentAt`, `Sheet.named`, `Workbook.named`, `wb(SheetName)`);
literal forms stay as `inline` macros returning the total type; removal in 1.0. Acceptance: the
skill, `docs/reference/scripting.md`, `examples/*.sc` and `verify-skill-snippets.sh` snippets
compile without deprecation warnings; `ElegantSyntaxSpec` updated.

**W3.3 SST compaction (#567)** · xl-ooxml · Scope: rebuild `sharedStrings.xml` to referenced
entries on write (remap `t="s"` indexes, recompute `count`/`uniqueCount`) behind
`WriterConfig.compactSharedStrings` (default on for dirty writes, off when byte-stability of
untouched sheets is requested), plus lint `shared-string-orphan`. Acceptance: the issue's repro
shows `uniqueCount` unchanged after replacing text; round-trip law; `DeterminismSpec` green.

**W3.4 Packaging and version single-sourcing** · build, scripts, CI · Scope: an `xl-cli-native`
module carries `jvmId = "graalvm-community:25.0.1"` so `__.compile`/`__.test` stop downloading
GraalVM; one `scripts/install.sh` feeding the Makefile, the tarball and the skill zip;
`scripts/sync-version.sh` writes `build.mill`, `plugin.json`, the `//> using dep` pins in
`scripting.scala`, `SKILL.md` and `docs/reference/scripting.md` from one value, with the CI check
from `docs-from-code` extended to every pin; cats-effect versions reconciled (xl-cli 3.6.3 vs
xl-cats-effect 3.5.7); `xl --version --json` reports distribution kind and rasterizer availability.
Acceptance: `/release-prep` has no hand edits left for versions; CI fails on any pin drift.

**W3.5 `whatif` and `explain`** · xl-evaluator, xl-cli, xl · Scope: a tracing `Evaluator` producing
an `EvalTrace` tree; `xl explain <ref|formula>`; `wb.whatIf(overrides, watch, options):
WhatIfResult` cone-scoped and cross-sheet; `xl whatif --set Sheet2!B1=5 --watch Summary!C10`; all
`--json`. Acceptance: `explain` on a cross-sheet chain shows every hop with values; `whatif` equals
a full recalculation on the watched cells for generated books.

**W3.6 Skill generated end to end** · plugin, CI · Scope: `plugin/skills/xl-cli/SKILL.md` and its
`reference/*.md` rendered from `xl schema --json` by `release.yml`; xl-agent prompt profiles
(`token|correct`), a trace classifier by gotcha signature; contract tests running `describe`,
`audit`, `deps` and every read verb over the 20-workbook foreign fixture corpus; byte-identity and
idempotence properties for every write verb over `Generators`. Acceptance: no hand-authored verb,
op, function or code list remains anywhere under `plugin/` or `docs/reference/`.

## Wave 4 — platform reach (GitHub-issue-ready)

**W4.1 `ops` on Scala Native / Scala.js** (ADR-016 execution meets ADR-017): the `Edit` algebra,
`EditSchema`, `FormulaOps` and `SheetRenamer` cross-compile with the library legs; the SN `xl`
binary passes the golden corpus. Acceptance: `./mill __.native.test` green on the algebra laws;
goldens pass on the SN binary.

**W4.2 Published `xl-ops-json` module**: the `Edit` codec as a tiny module depending on xl-core +
ujson so scripts can serialise edits (`Edit.toJson`, `EditJson.decodeBatch`) without the CLI;
xl-core stays JSON-free. Acceptance: a scala-cli script writes a batch file the CLI replays
byte-identically.

**W4.3 `xl-testkit` populated**: `Generators`, the law helpers (round-trip, monoid action, lowering
coherence), `CliHarness` and the golden runner published for downstream users. Acceptance: the
module has sources and a probe consumes it from a scala-cli script.

**W4.4 Rename the `Excel`/`Excel[F]` homonym** for 1.0 (sync facade vs F-polymorphic trait share a
name — a documented agent footgun): decide alias vs rename with a deprecation window.
