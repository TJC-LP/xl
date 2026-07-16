# Changelog

All notable changes to the XL project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Changed

- **Excel error VALUES are first-class evaluation results** (#344): `=1/0`
  now evaluates to `Right(CellValue.Error(#DIV/0!))` — catchable by
  IFERROR/ISERROR, cached by `recalculate()` (written as `t="e"` cells),
  rendered by the CLI as `#DIV/0!` — instead of failing the whole formula.
  Aggregates propagate error elements per Excel policy (SUM/MIN/MAX/AVERAGE/
  SUMPRODUCT propagate — including raw ranges and SUMIF-matched cells —
  while COUNT skips and COUNTA counts); AND/OR/NOT evaluate arguments
  eagerly and return the first error; scalar comparisons AND equality
  propagate (`=1=#REF!` no longer swallows to FALSE); numeric domain
  failures carry op-level Excel codes (`=2^1024` → `#NUM!`, SQRT/LOG/LN
  domains → `#NUM!`, `MOD(x,0)` → `#DIV/0!`, RANK → `#N/A`); mismatched
  broadcast dimensions pad with `#N/A` per Excel; literal `"TRUE"`/`"FALSE"`
  text coerces in condition positions. Host failures (parse errors, missing
  sheets, cycles) remain loud `Left`s — the boundary is law-tested, and
  `RecalcResult` distinguishes cached error values (`excelErrors`) from
  host failures (`errors`). Design record: `docs/design/error-propagation.md`.

### Fixed

- **xl-agent robustness** (#344 items 7–9): `pause_turn` stops auto-resume
  with bounded re-sends (usage summed across resumes, container id
  preserved past strategy request configuration — a real bug found by
  adversarial review); errored tasks now count in benchmark summary totals
  (reported distinctly, never as passes); the engine's last-resort trace
  writer never overwrites a skill-saved trace (fallback lands suffixed).

## [0.13.0] "Fixpoint" - 2026-07-16

The replication-campaign feature train (waves 13–16): professional models
now parse (`=+`, `%`, defined names), compute (coercion parity, opt-in
iterative recalculation for circular schedules), author (data validation,
calcPr, comments/CF via Patch, text rotation, tab colors, scrolled panes),
and finish (appearance/print-setup/cf CLI) — end to end, no XML surgery.

### Added

- **Sheet appearance & print-setup CLI commands** (#358, completing the
  issue): `sheet-view --gridlines on|off --zoom N [--tab-selected]`,
  `tab-color <color>|--clear` (hex, named, `rgb()`, and new
  `theme:<slot>[:<tint>]` syntax), `page-setup --orientation --scale
  --fit-to-width/--fit-to-height/--fit-to-page`, and `header-footer`
  (odd/even/first) — each with a batch-op twin. The full deliverable
  finish (gridlines off, zoom 85, tab colors, landscape + fit, confidential
  footer) is one batch file with zero XML patching.
- **Conditional-formatting CLI** (#324): `xl cf add --range A1:A10 --rule
  'cellIs:greaterThan:100' --bold --bg #FFC7CE` (+ `cf list` and a batch
  `cf` op) — colon rule DSL covering cellIs/between/expression/colorScale/
  dataBar/top10/text families with clean errors; priorities auto-stamp.
- **Batch `putf` format field** (#356): `{"op":"putf","ref":"C1",
  "value":"=A1*2","format":"#,##0.0"}` applies the numFmt with the formula
  in all three variants (single, `values[]`, `from`-dragging), on both the
  in-memory and streaming batch paths — no more second `style` pass.
- **Rasterizer diagnostics** (#359): failed PNG/PDF exports now name the
  probed backend chain, point at `xl rasterizers`, and say when you're on
  the native binary (where Batik is unavailable by design) with one-liner
  installs; `xl rasterizers` gains explicit native-image awareness.

### Changed

- **JSON formula cells carry a dedicated `formula` field** (#357):
  `view --format json` now always emits `{ref, type:"formula", formula,
  value, formatted}` — `value`/`formatted` hold the computed/cached value.
  Previously, under `--formulas`, `value` held the formula *text* and the
  computed value was absent; consumers keying on `value` there were reading
  the misfeature this fixes. CSV/markdown behavior is unchanged.

- **Defined-name resolution in formulas** (#384): `=IF(case=2,…)`,
  `=entry_mult*ltm_ebitda`, `=SUM(rev_range)` parse, evaluate, and
  round-trip byte-faithfully via a dedicated name AST node — workbook- and
  sheet-scoped names (sheet-scoped shadows global per OOXML), constants,
  ranges, and name→name chains with a cycle guard; unresolvable names are
  clean per-cell errors. Names contribute dependency edges, so
  `recalculate()` orders name-gated formula families correctly — on one
  real LBO gold this class was 926 of 1,571 recalc-probe rejections.
- **Opt-in bounded iterative recalculation** (#373, completing part b):
  `recalculate(IterativeCalc(maxIter, maxChange))` Jacobi-fixpoints
  declared cycles (seeded at 0, per-Excel keep-last-values on maxIter,
  dependents evaluate off converged values, clock pinned across
  iterations, per-cell containment inside the loop); default
  `recalculate()` still isolates cycles as errors.
  `IterativeCalc.fromCalcPr` bridges the workbook's authored calcPr.
  Circular LBO debt schedules now verify end-to-end.
- **MROUND** (#386): the 108th registry function —
  `multiple × ROUND(number/multiple)`, `MROUND(x, 0) = 0`, sign-mismatch
  errors — the standard term-loan sizing idiom.

### Fixed

- **Coercion parity with Excel** (#385): raw serial `Number` cells coerce
  in date positions (YEAR/MONTH/DAY/EOMONTH/EDATE/DATEDIF/NETWORKDAYS/
  WORKDAY/YEARFRAC — professional workbooks store dates exclusively as
  serials), and blank cells coerce to 0 in scalar numeric positions
  (`=A1+1`, `=-A1`, `=A1%`) — while range aggregates still skip blanks and
  error cells still propagate. Two real-model recalc-probe rejection
  classes eliminated.

- **Data-validation authoring** (#375): typed list dropdowns —
  `sheet.withDataValidation(range, DataValidation.list("\"Yes,No\""))` or a
  range-ref formula — with allowBlank/showDropdown (OOXML's inverted
  `showDropDown` handled), multi-range sqref, structural-edit range
  shifting, and read-side parse into the model (read→edit→write keeps
  validations on rewritten sheets); unmodeled validation kinds survive
  byte-faithfully as `Preserved`, mirroring conditional formatting.
- **calcPr authoring** (#373, part a): `WorkbookMetadata.calcPr` /
  `Workbook.withCalcPr(CalcPr(iterativeCalculation, maxIterations,
  maxChange))` emits `<calcPr iterate iterateCount iterateDelta/>` on
  scratch builds and overlays onto preserved calcPr (calcId/refMode
  survive) — from-scratch replicas of circular models open in Excel with
  iterative calc ON. Bounded iterative evaluation (part b) remains open.
- **Patch comment + conditional-format cases** (#379):
  `Patch.SetComment` / `Patch.SetConditionalFormat` with DSL operators
  (`ref.comment(...)`, `range.conditionalFormat(rule)`) — section builders
  composed as patches can now attach provenance comments and CF rules;
  auto-priority stamping routes through `Sheet.conditionalFormat`.
- **Text rotation** (#380): `Align.textRotation` (0–180 + 255 stacked, with
  `CellStyle.rotated(deg)` accepting Excel-UI negative degrees) — rotated
  sensitivity-grid captions author, serialize on both backends, stream
  through `StylePatcher`, and participate in style dedup correctly.
- **Runtime column handles** (#361): `Column.parse("D")` and a total
  `RefType.col` — column-oriented builders can fold over runtime letters
  (`Column.parse(c)` → `setColumnProperties`) with no compile-time literal.
- **Recalculating write** (#360): `Excel.writeRecalculated(wb, path)` (xl
  aggregator/scripting prelude) recalculates, writes the cached workbook,
  and returns the `RecalcResult` — the one-call answer to the
  "wrote blanks because nothing recalculated" footgun; formula errors are
  data (the file still writes; inspect `result.errors`).

- **Excel percent postfix operator** (#355): `=A1*10%`, `=10%`, `=(1+5%)^2`
  parse, evaluate (`10%` → exact `0.1`), broadcast over ranges (`=A1:C1%`),
  and print back byte-identically (never rewritten to `/100`); precedence
  matches Excel (`2^3%` ≡ `2^(3%)`), and deep `%` chains respect the
  parser's nesting guard.
- **Leading unary plus preserved through print** (#374): `=+A1`,
  `=+Model!B2`, `=++A1`, `=2^+2` round-trip byte-for-byte via a dedicated
  identity AST node — replicated formula text now matches professional-model
  sources exactly. This deliberately reverses #271's printer normalization
  (`=+A1` no longer prints as `=A1`); evaluation is unchanged (pure
  identity), and references under `+`/`%` still shift on drag and
  structural edits and contribute dependency edges.
- **Freeze panes read into the model** (#372): the reader populates
  `Sheet.freezePane` from `<pane>` (frozen and frozenSplit states), so
  read→modify→write keeps freeze state through model regeneration and
  typed dumps see it; `SheetView` gains `tabSelected`. Unmodified sheets
  still round-trip byte-identically.
- **Scrolled frozen panes** (#382): `FreezePane.At` carries an optional
  scroll target (`freezeAt(anchor, scrolledTo)`), representing
  `<pane ySplit="9" topLeftCell="F40"/>` — frozen at one row, scrolled to
  another; structural edits shift both refs.
- **Sheet tab color, model half** (#358): `Sheet.tabColor: Option[Color]`
  (rgb and theme+tint forms) with `withTabColor`, read + write on both
  writer backends with preserve-if-None semantics and schema-correct
  `sheetPr` child order. CLI exposure remains open in #358.

## [0.12.7] "Integrity" - 2026-07-16

File-integrity wave from the tjc-modeling byte-exact replication campaign
(#327, #328, #378, #381, #383, #387, #388): every workaround that campaign
shipped as post-write zip surgery is now unnecessary — packages are
self-consistent, schema-valid, and `recalculate()` is total.

### Fixed

- **Property-only rows survive scratch-build writes** (#381): rows carrying
  only `RowProperties` (height/hidden/outline via `setRowProperties`, no
  cells) now emit their `<row>` element on the default writer backend —
  spacer-row heights and collapsed outline groups no longer vanish. The
  streaming (SaxStax) backend already handled this; backends now agree,
  pinned by cross-backend parity tests.
- **Formula cells keep cached DateTime results** (#378): `=TODAY()`,
  `=EDATE(...)`, `=EOMONTH(...)` cells written after `recalculate()` now
  carry their cached value as an Excel serial in `<v>` (with `t="n"`) on all
  three writer backends, so `data_only` readers, pandas, and previewers see
  values instead of blanks. The cats-effect streaming writer also gained the
  missing `t="str"` mapping for cached text.
- **Comment rich runs are schema-valid** (#383): run properties in
  `xl/comments*.xml` (and freshly-authored shared-string/inline rich runs)
  emit `<rFont val=…/>` per CT_RPrElt instead of the styles.xml `<name>`
  shape — strict parsers (openpyxl) can open XL-authored comments again.
  The reader now parses `<rFont>` (real Excel comments no longer silently
  lose their font) while still accepting legacy `<name>` files.
- **Modified sheets keep their source part names** (#327): on source-backed
  writes, a modified sheet's output part, its rels, and the workbook-rels
  reconciliation all derive from one identity-resolved source path — writes
  against Excel-reordered files no longer fail with a duplicate-zip-entry
  `IOError` when an index-derived name collides with an unmodified
  neighbor's part.
- **Removing a sheet's last comment prunes its package footprint** (#328):
  a cell-edit write that drops a sheet's final comment no longer leaves the
  orphan comments/VML `[Content_Types].xml` overrides, dangling sheet rels,
  or `legacyDrawing` reference behind; other sheets' comments survive, and
  foreign part dialects (openpyxl naming) prune by resolved path.
- **Scratch workbooks with theme colors ship a theme part** (#387): writing
  a from-scratch workbook whose styles (or conditional-formatting dxfs)
  reference `Color.Theme` now emits a default Office `xl/theme/theme1.xml`
  plus its content-type override and workbook relationship — packages are
  self-consistent for strict consumers. RGB-only workbooks gain nothing;
  source-backed writes keep preserving the original theme.
- **Financial-function divergence is contained per-cell** (#388): `IRR()`
  Newton divergence (and the same overflow class in XIRR/XNPV/NPV and the
  TVM family) returns a per-cell evaluation error instead of throwing
  `ArithmeticException` out of `recalculate()`; a defense-in-depth backstop
  in the workbook evaluator guarantees recalculation is total even if a
  future function forgets to guard.

## [0.12.6] "Fieldwork" - 2026-07-15

Production-feedback hardening from agent-fleet use on real deal workbooks
(#349–#354, PRs #363–#368): external-workbook references parse and keep their
Excel caches, batch writes recalculate, output truncation is visible,
DOCTYPE-bearing workbooks read, the native binary reports real parse errors,
and linux-arm64 binaries ship.

### Added

- **`xl recalc` command** (#352): recalculate any workbook from the CLI
  (`xl -f model.xlsx -i recalc`), reporting per-cell formula errors without
  failing the write.
- **Truncation reporting** (#351): `view`/`search` now say when `--limit`
  (default 50) clips output — a `… showing N of M rows` trailer in markdown,
  a stderr notice for csv (stdout stays machine-parseable), and
  `truncated`/`totalRows` fields in json; `--limit 0` means unlimited, and
  `search` reports its true match total.
- **linux-arm64 native binaries** (#354): the release matrix builds on
  `ubuntu-24.04-arm` with arch-safe cache keys and a per-platform binary smoke
  test; the generated installer maps `Linux-aarch64` and fails loud when an
  asset is missing instead of installing an error page.

### Fixed

- **External-workbook references parse and pin their Excel caches** (#353):
  `[2]Book1!A1` / `'[3]Sheet Name'!B2` forms parse to a dedicated AST node
  with exact printer round-trip and anchor-aware shifting; they contribute no
  dependency edges, `recalculate()` preserves the Excel-written cached value
  of external-formula cells verbatim (dependents compute from those caches),
  uncached external cells report a clear per-cell error instead of
  `UnexpectedChar([`, and one such formula no longer poisons evaluation of
  the rest of the workbook. Resolving external workbooks / reading
  `externalLinks` cache parts remains future work.
- **Batch writes recalculate** (#352): cell-mutating `batch` runs end with one
  global recalculation, so `putf` cells carry cached `<v>` values (previously
  blank in `data_only` readers, pandas, and previewers); formula errors are
  surfaced in the batch summary without aborting the write. Also fixes a
  latent core bug: recalculated caches now survive surgical writes of
  disk-read workbooks (the modification tracker never saw cache-only changes,
  so the writer preserved stale worksheet XML verbatim); sheets whose caches
  are unchanged still ride byte-for-byte preservation.
- **Benign leading DOCTYPE no longer rejects core parts** (#350): a
  `<!DOCTYPE …>` prolog (with or without an internal subset) on
  workbook/sheet/styles/sharedStrings parts is stripped by a conservative
  scanner before parsing — the parser itself stays fully locked down
  (doctype disallowed, external entities off) — and core-part parse errors
  now carry line/column plus the offending construct. A hostile-but-valid
  fixture joins the corpus laws.
- **Native-image parse diagnostics** (#349): the JDK-internal Xerces/Xalan
  message bundles are registered for native-image, so the shipped binary
  reports real parse errors (line/column and reason) instead of
  `Could not load any resource bundle by …XMLMessages`; the release workflow
  now smoke-tests every native binary against a valid and a malformed
  workbook.

### Docs

- **Skill field-gotchas refresh** (#362): production-verified gotchas,
  warning triage, and 0.12.5 recalculation notes in the `xl-cli` and
  `xl-scripting` skills (updated again in this release for the fixes above).

## [0.12.5] "Memo" - 2026-07-13

Kills an exponential blowup in the formula evaluator on recursive multi-branch
models (the LBO debt-schedule shape) — hours of CPU down to milliseconds — and
moves whole-workbook recalculation onto a single workbook-level dependency
order.

### Fixed

- **Recursive uncached-reference evaluation is memoized per pass** (#346): an
  uncached formula reference (`CellValue.Formula(_, None)`) was recursively
  re-evaluated once per *path* through the dependency graph — exponential on
  multi-branch recursive chains (an LBO debt schedule at ~800 formulas ran for
  hours at 100% CPU; the repro in `RecalcPerfSpec` never finished at n=24
  periods). A pass-local memo (`Evaluator.EvalMemo`, keyed by sheet identity
  then ref) threads through recursive evaluation, `EvalContext`, and the
  aggregate/array uncached-cell readers, so each distinct cell evaluates at
  most once per top-level evaluation — Excel's own one-computation-per-cell
  semantics. Errors memoize too (a cycle no longer re-burns the depth guard
  per path). The n=24 repro drops from unbounded to milliseconds.
- **`Workbook.recalculate()` orders cells workbook-level, not per sheet**
  (#346): recalculation now builds one qualified (sheet!cell) dependency graph
  (`DependencyGraph.fromWorkbookBounded`, ranges bounded by each *target*
  sheet's used range), prunes cycles with a workbook-level Tarjan pass, and
  evaluates a single global topological order threading the partially
  evaluated workbook — every precedent, same-sheet or cross-sheet, is computed
  exactly once before its dependents instead of being re-derived on demand
  against the original uncached snapshot. Consequences: mutually-referencing
  sheets recalculate in linear time; cross-sheet aggregates over uncached
  formula cells read computed values regardless of sheet order; **cycles that
  span sheets are now detected and reported as circular** (participants
  circular, downstream blocked, acyclic remainder evaluates) instead of
  surfacing as depth-guard evaluation errors. The GH-274/GH-301 dynamic
  (INDIRECT/OFFSET) evaluate-last bucket and cache-strip semantics are
  preserved, with the closure now workbook-level.

## [0.12.4] "Carriage" - 2026-07-09

Wave 11: elementwise error semantics through array operations, and loud benchmark
failure diagnostics. Follow-up divergences tracked in #344.

### Fixed

- **Errors carry elementwise through array operations** (#337): comparisons and
  arithmetic over arrays containing error values produce the error *as that element*
  (`=10/A1:A3` with a zero spills `{10, #DIV/0!, 5}`) instead of failing the whole
  formula. Left operand's error wins for binary ops; per-element division by zero
  yields `#DIV/0!`; uncoercible text yields `#VALUE!`. New invariant, property-tested:
  broadcast operations return a formula error only for dimension mismatch.
- **Unused IF/IFS branch errors no longer poison the array path** (#339):
  `=SUM(IF(A1:A2>0,A1:A2,1/0))` on all-positive data evaluates — branch failures
  demote to error elements that unselected positions discard, per Excel CSE.
  Scalar-path branch laziness is unchanged and pinned.
- **Aggregates fail loudly on carried errors instead of silently mis-summing**:
  SUM/MIN/MAX/AVERAGE/SUMPRODUCT over expression arrays containing an error value
  return a formula error naming the Excel code (previously SUMPRODUCT coerced error
  elements to 0 — a silently wrong number). COUNT keeps Excel's ignore-errors
  behavior; the guards remain IFERROR-catchable.
- **Benchmark truncations are loud** (#340, internal xl-agent): the agent loop now
  captures `stop_reason` — `max_tokens`/`refusal`/`pause_turn` turns surface as case
  errors in reports, glyphs, and traces instead of reading as clean completions;
  default per-iteration output cap raised to 32,768 with a new `--max-tokens` flag;
  partial usage is recovered from completed turns; engine-level failures leave a
  metadata trace; all-errored tasks render red; stream lines carry the skill name.

### Changed

- New `EvalError → CellError` demotion table (`#DIV/0!`, `#REF!`, `#VALUE!`);
  circular references remain fatal formula errors. Documented divergences from
  Excel (pow-overflow `#VALUE!` vs `#NUM!`, AND/OR error aborts, scalar error
  refusals, aggregate error-value returns) are enumerated in #344.

## [0.12.3] "Parity" - 2026-07-09

Waves 10–11: evaluator correctness gaps found by live SpreadsheetBench dogfooding, plus the
xl-agent benchmark refresh that found them.

### Fixed

- **Excel comparison semantics** (#335): ordered comparisons (`<` `<=` `>` `>=`) no longer
  raise `TypeMismatch` on text operands. Comparisons now follow Excel's total order — text
  compares case-insensitively and lexicographically, cross-type rank is
  `number < text < logical` with `FALSE < TRUE`, dates compare by their Excel serial,
  and empty cells coerce to the other operand's zero value (`0` / `""` / `FALSE`) — uniformly
  in scalar and array/broadcast paths (`=SUMPRODUCT((B1:B2<"z")*1)` works).
- **Array-aware IF; evaluator crash family eliminated** (#333): `=MIN(IF(range>0,range,99))`
  and friends no longer escape the total-function boundary with a `ClassCastException`
  (a raw GraalVM stack trace in the native CLI). IF broadcasts array conditions elementwise
  per Excel CSE semantics; scalar IF keeps lazy branch selection.
- **AND/OR aggregate over array conditions; NOT and IFS broadcast** (#338): logical
  functions fold arrays per Excel (AND=forall, OR=exists) instead of silently collapsing to
  the top-left element; `=NOT(range>0)` spills elementwise; IFS array conditions follow CSE
  chaining. Cross-argument short-circuit and scalar laziness are preserved.
- **Benchmark failure diagnostics preserved** (#334, internal xl-agent): per-case errors now
  serialize into reports and partial conversation traces survive mid-run failures.

### Changed

- **Equality semantics** (`=`, `<>`, `SWITCH`) now equate `Empty` with `0`/`""`/`FALSE` and
  dates with their Excel serial (Excel-correct; previously type-strict).
- **Numeric-looking text no longer parses as a number under ordered comparisons**
  (`="2">100` is `TRUE`, matching Excel's type ranking).

### Internal (xl-agent benchmark harness, #332)

- Claude 5 model registry (`claude-sonnet-5` agent default, `claude-opus-4-8` grader),
  anthropic-java 2.11.1 → 2.48.0, code-execution tool `20260521`, typed container uploads.
- Prompt caching on code-execution requests (−52–59% cost per benchmark task) and
  cache-token capture end-to-end (traces price by the model that actually ran).
- Version-agnostic release-asset resolution (auto-download of the latest binary/skill now
  works); Skills API drift fixes (`files[]` form field, list pagination, flat-zip layouts).

## [0.12.2] "Interop" - 2026-06-11

Wave 9: the LibreOffice edit-corruption fix and the writer follow-ups it surfaced.

### Fixed (wave 9)

- **Editing LibreOffice-produced workbooks no longer corrupts them** (#320):
  `workbook.xml.rels` regenerates in the same pass as `workbook.xml` — sheet rIds stay
  consistent with their rels (LO maps rId1 to the theme; any edit previously made the first
  sheet resolve to `theme1.xml`), preserved workbook-level elements survive the regeneration,
  and multi-add/multi-rename allocates distinct sheetIds. Verified end-to-end: xl and
  openpyxl both read edited LO files correctly; the wave-7 CfPreservationSpec workaround is
  undone (full re-reads assert directly).
- **Comment content-type registration follows actual emitted paths** (#321 — proven already
  fixed by the GH-315 rework; falsifiability-tested regression guard committed).
- **No dangling [Content_Types] overrides for dropped writer-owned parts** (#322): deleted
  sheets' worksheet/comment/VML overrides are pruned from the preserved merge; exotic
  non-writer-owned overrides still survive.
- **New sheets join surgical SST accounting** (#323): `Workbook.put` of a fresh sheet on a
  source-backed workbook references the combined SST (`t="s"`) with exact counts instead of
  re-inlining its strings.

## [0.12.1] "Clean Sweep" - 2026-06-11

Wave 7: every remaining open issue closed in one wave — conditional formatting lands as
the headline feature alongside twelve fidelity, writer-internals, and polish fixes.

### Added (clean-sweep wave 7)

- **Conditional formatting** (#136): typed model for cellIs / expression / colorScale /
  dataBar / top10 / text-operator rules with dxf differential formats, `sheet.conditionalFormat`
  authoring with auto-priority (saturating), structural row/column edits shift rule ranges,
  unmodeled rule families ride namespace-self-contained `Preserved` payloads — x14/extLst
  blocks (LibreOffice dataBars) survive dirty writes byte-faithfully. Joins the generative
  round-trip law; openpyxl fixture in the corpus.
- **activeTab serialization** (#294): `Workbook.activeSheetIndex` round-trips through
  `bookViews/workbookView` both directions; equivalence un-ignores it.
- **fitToPage tri-state** (#284): `PageSetup.fitToPage: Option[Boolean]` — `Some(false)`
  actively strips a preserved flag; `None` keeps the 0.11.1 derived behavior.

### Fixed (wave 7)

- **openpyxl comments surface in the model** (#292): comment parts resolve via worksheet
  relationship type instead of path patterns — `xl/comments/comment1.xml` (subdirectory
  dialect) now populates `Sheet.comments`.
- **RichText SST keying** (#303): rich-text entries key by run structure, not plain text —
  a RichText whose plain text equals an existing plain string no longer collides; RichText
  cells now generate in the round-trip law.
- **Exact surgical SST counts** (#304): edits that remove or duplicate references to existing
  strings recount the modified sheets' contributions exactly.
- **Content_Types preservation** (#314): metadata-modified writes merge the preserved
  `[Content_Types].xml` (exotic overrides survive) instead of rebuilding from minimal.
- **Namespace-URI rel detection** (#316): preserved drawing fragments binding the
  relationships namespace to exotic prefixes are detected as rel-bearing.
- **Identity-keyed source mappings** (#315): drawing/comment path mappings and SST
  modified-sheet accounting key by stable sheet name instead of index — sheet deletion or
  reordering combined with drawing/comment edits in one write now regenerates correctly
  (the wave-6a skip-guard is gone); fresh comment parts allocate collision-free paths and
  register their actual emitted paths in [Content_Types].xml.
- **`Cell.comment` deprecated and wired** (#295): the never-serialized field now
  write-throughs into `Sheet.comments` on `put`; lenient sheet-ref parsing pinned as
  intentional (#281).

### Changed (wave 7)

- **Codec put paths 2.4x faster** (#297): style registration fused into the codec put
  pipeline — chained codec puts 2.44x, varargs 1.56x (smoke-mode JMH, `SheetPutBenchmark`).
- **Dead writer paths deleted** (#313): `regenerateAll`/`writeZip*` and unused async
  static-part helpers removed (−455 lines, zero callers).

## [0.12.0] "Visual" - 2026-06-11

The Visual release: embedded pictures and typed charts with hybrid preservation —
the drawing layer xl could previously only carry verbatim is now a first-class,
law-governed part of the model. Closes the original 0.12.0 theme (#221, #222).

### Added (Visual wave 6b)

- **Typed chart model** (#222): `com.tjclp.xl.charts.Chart` — bar, line, pie with series
  (name/categories/values as sheet-qualified ranges), title, legend; charts anchor through the
  6a drawing layer (`Drawing.ChartFrame`). Typed read of the supported subset with `Preserved`
  fallback for everything else (stacked/scatter/3D round-trip byte-faithfully — pinned by new
  `chart-stacked.xlsx`/`chart-scatter.xlsx` fixtures). Deterministic `xl/charts/chartN.xml`
  emission with series formulas quoted via `SheetName.quoteForFormula` and value caches filled
  from live cells. **Structural row/column edits rewrite chart series ranges** like formulas.
  Charts join the generative round-trip law (`genChart`, dedicated `ChartRoundTripSpec` laws).
- **CLI**: `xl chart add --type bar --data B2:D10 --categories A2:A10 --title "Revenue" --at
  F2:K15` and `xl add-image logo.png --at B2` (#222, #221 surface).

### Added (Visual wave 6a)

- **Embedded pictures** (#221): new pure `com.tjclp.xl.drawings` package — `Drawing` enum
  (`Picture` + `Preserved` for unmodeled content), `DrawingAnchor` (OneCell/TwoCell/Absolute
  over the `Emu` unit type), `ImageData` (7 formats with total magic detection and
  png/gif/jpeg/bmp dimension sniffing, content-addressed sha256). `Sheet.addImage` (4
  overloads incl. natural-size), `pictures`, `removeDrawing`; structural row/column edits
  remap picture anchors like comments. Full OOXML round-trip on both writer backends:
  deterministic drawing/media part emission, sha-deduplicated media, relationship-preserving
  hybrid regeneration (clean drawings ride byte-preservation — the wave-2 corpus and the #291
  namespace machinery hold by construction). Streaming envelope documented (in-memory writes
  only). Pictures join the generative round-trip law (`genPicture`).

## [0.11.3] "Robustness" - 2026-06-11

Waves 4+5 of the backlog burn-down: streaming/OOXML robustness (docProps, streaming SST,
1904 dates, fraction/conditional formats, reader parity, the drawing-corruption fix that
unblocks 0.12.0 Visual) and the CLI/evaluator batch (import-md, diff, filter; implicit
intersection and the totality sweep).

### Added (CLI & evaluator totality, wave 5)

- **`import-md`** (#159): GFM pipe-table import (file or stdin) with smart per-cell type
  detection (currency/percent/date via `FormattedParsers.detect`), `--start`/`--new-sheet`/
  `--skip-header`/`--no-type-inference`, alignment markers map to cell alignment.
- **`diff`** (#137): compare two workbooks — cells (value/formula/resolved-style), sheets,
  merges, comments, hyperlinks; markdown + stable JSON output; diff-tool exit codes
  (0 identical / 1 differs / 2 error).
- **`filter`** (#134, phase 1): `--where` row predicates (comparisons, AND/OR/NOT, LIKE,
  BETWEEN, IN, IS EMPTY; `--header` name resolution; `--columns`/`--limit`;
  markdown/csv/json) — total evaluation, type mismatch = no match, never an error.
- **Font-metric autofit** (#156): column auto-fit measures real AWT font metrics
  (family/size/bold) with the heuristic as headless fallback.
- **Batik-first rasterization** (#86): pure-JVM Batik is the documented default; subprocess
  backends (incl. ImageMagick) are explicit opt-ins; best-effort native-image configs added.
- **Scheduled law fuzzing** (#308): weekly random-seed runs of the generative law at 2000
  cases (`law-fuzz.yml`), seeds printed for replay via `-Dxl.roundtrip.seed`.

### Fixed (wave 5)

- **Default display strategy prefers cached formula values** (#282) — xl-core-only consumers
  see numbers, not formula text, for already-evaluated files.
- **SVG renderer** (#298): shared-edge borders resolve by weight (heavier wins, one line per
  edge); clipPath ids keyed by cell ref (no duplicate ids in hidden ranges).
- **Evaluator totality sweep** (#302, #306, #307, #301, #280): ArrayResult collapses to its
  top-left value in scalar argument AND operator positions (implicit intersection —
  `IFERROR(INDIRECT("A1"), x)` works); cross-type call-result coercion per cell-decoder
  conventions (`=UPPER(SUM(...))`, `=IF(SUM(...),1,2)`); integer-argument coercion
  total-with-error across the date-function family (toInt sweep); OFFSET joins INDIRECT's
  deferred-partition recalc semantics (dynamicDeps parity); internal diagnostics quote
  cell-ref-shaped sheet names.
- **CLI border merging delegates to core `Border.merge`** (#279) — CLI semantics can no
  longer drift from the library's.
- **Test sources join the scalafmt gate** (#296): CI and pre-commit run `checkFormatAll
  __.sources`; the 130-file formatting backlog cleared.

### Added (streaming & OOXML robustness, wave 4)

- **docProps emission** (#242): `docProps/core.xml` + `app.xml` written from `WorkbookMetadata`
  (deterministic — only modeled fields, no GUIDs or wall-clock stamps); reader parses them back;
  the model wins over preserved verbatim parts; metadata fields un-ignored in the generative
  equivalence.
- **Streaming SST + style registry** (#223): two-pass streaming writes emit a real
  `sharedStrings.xml` (first-occurrence order, per-reference counts) and `styles.xml` instead of
  inline strings everywhere — per the smart-streaming design; SST accumulator created per run
  (referential transparency: re-running the same compiled `IO` no longer reuses a dirty
  accumulator).
- **1904 date system** (#243): `workbookPr date1904` parsed into `WorkbookMetadata`, preserved on
  write, epoch-aware serial conversion (no phantom-leap-day in the 1904 system); streaming-edit
  envelope documented.
- **Fraction + conditional formats** (#243, #285): Excel fraction rendering (`# ?/?`, `# ??/??`,
  fixed denominators; continued-fraction nearest under the digit budget, total for BigDecimal
  beyond Double range) and `[>100]`-style condition-prefix section routing.

### Fixed (wave 4)

- **Date display gaps** (#283): General + date value renders the serial (Excel behavior);
  custom date codes route through sections (`;;;` hides dates).
- **Streaming reader parity** (#293, #305): inlineStr cells keep their `s=` style indices;
  `<f>` formula text trimmed like the DOM reader; `_xHHHH_` escapes decode per-run (bare `<t>`
  ignored when rich runs present, matching DOM).
- **StyleParser totality** (#278): malformed `<sz val>` (and sibling numeric attributes) no
  longer throw through Font's domain guard on read.
- **Drawing corruption on modify** (#291): namespace prefixes (`xmlns:r`) stay bound when
  regenerating worksheets with preserved drawings — bindings retained at capture AND hoisted at
  re-emission; modifying cells in chart/image sheets now produces valid workbooks (unblocks the
  0.12.0 Visual theme).

## [0.11.2] "Laws & Functions" - 2026-06-10

Waves 2+3 of the backlog burn-down: the test-infrastructure release that ships its laws
together with everything those laws caught, plus the evaluator-breadth batch (LET, RAND,
INDIRECT, YEARFRAC parity, format inheritance — registry 104 → 107).

### Added (evaluator breadth, wave 3)

- **LET** (#193): `LET(name1, value1, ..., calculation)` lexical bindings with let* semantics —
  new `TExpr.Let`/`BindingRef` AST nodes, parser-level lexical scope, total coercion of bindings
  in typed argument positions, range-valued bindings via parse-time substitution. Divergences
  from Excel documented in LIMITATIONS.md.
- **RAND / RANDBETWEEN** (#115): randomness as an explicit `Rng` capability mirroring `Clock` —
  `Rng.system` by default, `Rng.seeded(seed)` for deterministic scripts/tests, threaded through
  explicit `@targetName` overloads (`evaluateFormula(f, clock, rng)`, `recalculate(clock, rng)`).
- **INDIRECT** (#274): dynamic text-to-reference resolution riding the OFFSET `ArrayResult`
  mechanism — composes with aggregates (`=SUM(INDIRECT("A1:A"&B1))`), sheet-qualified and quoted
  names, full-column clamping; `recalculate()` defers INDIRECT-bearing cells and their dependents
  to an evaluate-last partition with cache stripping. R1C1 mode documented-unsupported.
  Registry: 104 → **107 functions**.
- **Format inheritance for formulas, opt-in** (#184): `sheet.putFormulaInheriting(ref, formula)`
  applies the referenced cells' number format to General-formatted targets (Excel's behavior on
  formula entry), via `FormulaFormatting.inferFormatFromReferences`.

### Fixed (wave 3)

- **YEARFRAC Excel parity** (#93): full US/NASD 30/360 rules for basis 0 (Feb end-of-month before
  day-31 rules, in Excel's order) and Excel's actual basis-1 algorithm (366 only when the span
  includes Feb 29; average-year-length for multi-year spans); start > end normalizes positive for
  all bases.
- **Round-trip defects found by the wave-2 generative law** — fixed and the corresponding
  generator exclusions removed, so the law now locks them: schema-invalid lowercase border/fill
  enum tokens → canonical ECMA-376 camelCase (#287); CR in cell text lost → `_x000D_` ST_Xstring
  escaping on every writer/reader path including streaming (#288); SST collapsed NFC-equivalent
  strings → exact-string dedup, original codepoints preserved (#289); comment author whitespace
  corrupted text on read → author canonicalized (trimmed) at write time (#290); surgical SST
  combine orphaned the SST on no-new-strings writes, stored NFC-normalized text, and undercounted
  references (#277).
- **Cross-sheet array stringification** (latent, found during #274): a reference that recursively
  evaluates to an array no longer collapses to `Text("ArrayResult(...)")` — it takes the top-left
  value per scalar-context convention.

### Added (test infrastructure, #240/#40/#47)

- **Generative round-trip law** (#240): `forAll(genWorkbook) { wb => read(write(wb)) ≈ wb }`
  over rich generated workbooks (styles, comments, hyperlinks, formulas, sheet views, page
  setup) with a documented workbook-equivalence (`WorkbookEquivalence`, fixed-seed CI runs,
  `XL_ROUNDTRIP_MIN_SUCCESS` override). Generators extended accordingly (`genCellStyle`,
  `genComment`, `genHyperlink`, `genRichSheet`/`genRichWorkbook`).
- **Real-file fixture corpus + parity laws** (#240): 10 synthetic fixtures generated by
  `scripts/generate-fixtures.py` (openpyxl + LibreOffice dialects, documented provenance,
  deterministic regeneration) under `xl-ooxml/test/resources/fixtures/`; streaming vs
  in-memory reader parity asserted over every fixture (`StreamingParitySpec`); chart/image
  preservation behavior pinned (`FixturePreservationSpec`) ahead of the 0.12.0 Visual work.
- **`SheetPutBenchmark`** (#40): JMH comparison of `Sheet.put` overloads (chained, varargs,
  pre-constructed `Cell`, `Patch` batch).
- **Renderer edge-case tests** (#47): empty range with column properties, HTML/SVG dimension
  parity, shared-edge border behavior (documented), zero-area and all-hidden ranges.

The new laws immediately surfaced 12 real defects — filed as #287–#298 (schema-invalid
border/fill enum casing, CR loss, SST NFC collapse, comment-author corruption, drawing
corruption on modify, openpyxl comment dialect, streaming parity gaps, activeTab,
vestigial `Cell.comment`, scalafmt test-source gate, put-path performance, SVG edge defects).

## [0.11.1] "Totality" - 2026-06-10

Wave 1 of the post-0.11.0 backlog burn-down: every open bug, fixed TDD-first under
adversarial review (each new test proven to fail without its fix).

### Added

- **Leading unary plus in formulas** (#271): `=+A1`, `=+SUM(A1:B2)`, chained `=++A1`, and
  exponent positions (`=2^+2`) now parse — the ubiquitous banker idiom from real bank-built
  models. Unary plus is identity: it parses to the same AST as the bare expression, so the
  printer normalizes (`=+A1` prints `=A1`) and the parse∘print round-trip law is untouched.
  Deeply chained `+` counts against the recursion budget (`NestingTooDeep`, never overflow).
- **Even/first-page headers & footers + fitToPage** (#266): `HeaderFooter` gains
  `evenHeader/evenFooter/firstHeader/firstFooter` and `differentOddEven/differentFirst`;
  setting `fitToWidth`/`fitToHeight` now also emits `<sheetPr><pageSetUpPr fitToPage="1"/>`
  (without which Excel ignores fit-to-page). Full write→read round-trip.

### Fixed

- **Sheet names that look like cell refs are now quoted** (#263): a sheet named `Q1`, `A1`,
  or `R1C1` previously produced unquoted, ambiguous formulas (`Q1!$A$1:$B$2`) in printed
  formulas, `toA1`, and `_xlnm` print names. One shared predicate (`SheetName.needsQuoting` /
  `SheetName.quoteForFormula`) now drives `RefType`, `FormulaPrinter`, and `PrintNames` —
  including `TRUE`/`FALSE` as sheet names, which the parser would otherwise re-read as booleans.
- **Trailing empty format sections preserved** (#262): `"0.0;;"` (the hide-zero idiom) now
  splits into three sections and empty sections render as empty string instead of falling
  back to General. `";;;"` (hide-everything) works.
- **Evaluating display strategy prefers cached formula values** (#275): after
  `wb.recalculate()`, cross-sheet formulas display their cached values in `excel""` /
  `displayCell` output instead of raw formula text.
- **Streaming `StylePatcher` totality** (#264): malformed `indent="-2"` (and other invalid
  numeric attributes) in styles.xml no longer throw through `Align`'s domain guard during
  streaming style patches — hardened to match the DOM-side parser exactly.
- **Streaming writer emits sheet metadata for fresh sheets** (#265): `DirectSaxEmitter` now
  writes `sheetPr`, `sheetViews` (freeze panes, gridline settings), `pageMargins`,
  `pageSetup`, `headerFooter`, and `hyperlinks` in schema order — previously all silently
  dropped on the SaxStax path; also fixes orphaned hyperlink relationships.

### Changed

- **`SheetEvaluator` is now var-free** (#48): the two mutable accumulation blocks were
  refactored to `foldLeft`, removing the last `@SuppressWarnings(Var)` exemptions in the
  evaluator's hot path. Behavior is identical (locked by the existing suite + new
  equivalence pins).
- **SST whitespace preservation through surgical writes pinned by regression spec** (#17):
  investigated with a falsifiability protocol (three hypothesized bug classes injected, each
  caught) — the 2025-11 report is not reproducible on the current write path; a permanent
  regression spec now guards it.

## [0.11.0] "Scripting" - 2026-06-10

The scripting release: one-import prelude, total-by-semantics APIs, whole-workbook
recalculation with per-cell error reporting, the xl-scripting Claude Code skill with
compile-verified snippets, and the render/style fidelity needed to reproduce professional
(banker-grade) templates exactly. The new gate machinery (external-consumer probes, examples
CI, adversarial reviews, property seeds) surfaced and fixed eight pre-existing bugs.

### Added (render & style fidelity, #254-#260)

- **Per-side border DSL + range outlines** (#257): `CellStyle.borderTop/borderBottom/borderLeft/
  borderRight` builders (with color overloads) merging into the existing border, and
  `range.outlined(style[, color])` — an edge-correct outline patch built on the new
  `Patch.MergeBorder` case (apply-time border overlay preserving font/fill/numFmt; note for
  exhaustive matchers: `Patch` gains a case).
- **Alignment indent DSL** (#260): `CellStyle.indent(n)` shortcut (model + OOXML round-trip
  already existed); style dedup and round-trip now covered by tests; styles parser hardened
  against negative `indent` attributes in malformed files.
- **Sheet view settings** (#258): `SheetView(showGridLines, zoomScale)` on
  `Sheet.viewSettings`, serialized into the same `sheetView` element as freeze panes; SVG
  renderer suppresses gridlines when a sheet disables them.
- **Print setup** (#259): `PageSetup` gains `headerFooter` (odd header/footer with `&P`/`&N`
  codes), `margins`, `printArea`, and `repeatRows` — emitted in schema order and round-tripped;
  print area/titles ride the defined-names pipeline as sheet-scoped `_xlnm` names. Follow-ups
  (even/first headers, `fitToPage` flag) tracked separately.

### Fixed (render & style fidelity, #254-#260)

- **Custom number formats: zero routed to the negative section** (#254): two-section codes
  (`pos;neg`) now format zero through the positive section per Excel's rule — `$0.0`, not
  `($0.0)`. One unified section-selection site covers display, HTML/SVG, CLI, and `TEXT()`.
- **SVG cell fonts silently overridden** (#255): the embedded `.cell-text` CSS rule's
  font-family/font-size outranked per-cell presentation attributes in every CSS-aware
  rasterizer (and the attribute value carried CSS-style quotes fontconfig can't match). Class
  rules no longer declare fonts; every text path emits explicit, unquoted font attributes.
- **SVG underline dropped** (#256): `Font.underline` now emits `text-decoration="underline"`
  on SVG text (HTML already had it), so underlined headers survive PNG/PDF export.
- **Parser ate keyword-prefixed names** (#268): `=SUM(Notes1!A1:A2)` parsed as
  `=SUM(NOT(es1!A1:A2))` — logical keywords (NOT/AND/OR) now require a word boundary; sheets
  literally named `NOT`/`AND`/`OR` parse as references.
- **Verbatim copy always failed** (#261): the read digest stopped at the ZIP central directory
  while the copy check hashed the whole file, so writing an untouched workbook to a new path
  errored loudly. Untouched copies are byte-identical again; tampering is still refused.
- **excel"" printed raw formulas through the prelude**: the evaluator's `evaluating` display
  strategy never survived wildcard export; the prelude now selects it explicitly, so scripts
  display computed, NumFmt-formatted values.

### Changed

- Reading a workbook now populates `Sheet.pageSetup` for any sheet with `<pageMargins>`
  (virtually all real files) — previously only sheets with `<pageSetup>`; modelable
  `_xlnm.Print_Area`/`_xlnm.Print_Titles` defined names are lifted into `pageSetup` instead of
  surfacing in `metadata.definedNames`.


### Added (scripting hardening, #252)

- **Scripting prelude** `com.tjclp.xl.scripting.{*, given}`: one import for scripts — core API,
  DSL, compile-time literals, formula evaluation, sync `Excel`, streaming `ExcelIO` + `RowData`,
  smart detection sugar (`String.toFormatted`), and the `.unsafe` boundary. The plain
  `com.tjclp.xl.{*, given}` import stays 100% pure.
- **Range fill**: `range := value` puts the value in every cell (Excel Ctrl+Enter semantics) —
  new `:=` overloads on `CellRange`, total `:=` on `RefType`.
- **Total ARef navigation**: `down`/`up`/`right`/`left` (default 1) for Either-free loops.
- **`Workbook.upsert(name, f)`**: total update-or-create counterpart of `update` (string
  literals compile-time validated; runtime strings return `XLResult`).
- **Typed read helpers**: `Sheet.readTypedOr(ref, default)` and `Sheet.readTypedOpt(ref)`.
- **Workbook-level evaluation**: `wb.evaluateFormula(formula, onSheet)` wires cross-sheet
  context automatically.
- **`Workbook.recalculate(clock)`**: total whole-workbook recalculation returning
  `RecalcResult` (cached workbook + per-sheet values + per-cell `CellEvalError`s); reference
  cycles are isolated (participants circular, downstream dependents blocked) while the acyclic
  remainder evaluates.
- **`FormattedParsers.detect`**: total smart value detection (currency/accounting/percent/ISO
  date/number/boolean/text) promoted from CLI-internal code; the CLI delegates.
- **xl-scripting skill** (`plugin/skills/xl-scripting/`): SKILL.md + API reference + 7 runnable
  recipes; packaged as `xl-scripting-skill-<version>.zip` on release with a version-pin gate.
- **Anti-rot CI**: `scripts/test-examples.sh` (compiles every `examples/*.sc` against the local
  build, runs a curated subset) and `scripts/verify-skill-snippets.sh` (compiles every complete
  snippet in the skill), wired into GitHub Actions (`examples` job, `skill-verify` workflow).

### Fixed (scripting hardening, #252)

- **Opaque-type extensions unusable outside the package**: export forwarders for extension
  methods and companion factories on opaque types (`ARef.toA1/shift/col/row`, `Column`/`Row`
  arithmetic, `SheetName.value`, style units) produced path-dependent proxy types or failed
  extension lookup for any consumer outside `com.tjclp.xl` (e.g. `Column.from0(0).toLetter` in
  `examples/demo.sc` did not compile against 0.10.0). Fixed via explicit alias + singleton-val
  pairs in `api`, de-inlined opaque extensions/factories (still static methods on primitives),
  and companion-scope resolution. Guarded by new external-consumer probes
  (`xl/test/src/xlprelude/`).
- **`withCachedFormulas` uncached entire sheets**: one failing formula dropped caching for the
  whole sheet, and cross-sheet formulas always failed (no workbook context), silently uncaching
  their sheet. Now per-cell and cross-sheet aware (implemented on `recalculate`).
- **Silent no-op `range := value`**: previously returned `Patch.empty` (documented wart); now
  fills the range. Code relying on the no-op behavior must filter ranges explicitly.
- **Scripts hung at exit after `toHtml`/`toSvg`**: render text measurement initialized
  non-headless AWT, whose non-daemon AWT-Shutdown thread kept the JVM alive after main
  completed. AWT now defaults to headless unless the embedder set `java.awt.headless`
  explicitly (font metrics don't need a display).

## [0.10.0] "Trust & Author" - 2026-06-09

A correctness-and-authoring release: fix defects that silently corrupted data or
falsified flagship guarantees, close the read↔write asymmetry, and broaden the
formula library. See `docs/archive/plan/v0.10.0-triage.md` for rationale.

#### Added

- **Structural editing** (#128, #129) — `insert-rows` / `delete-rows` /
  `insert-cols` / `delete-cols` CLI verbs and a `StructuralEditor` workbook API.
  Shifts cells, merges, row/column properties and freeze panes, then rewrites
  every affected formula (including cross-sheet) with correct `#REF!` generation
  and range shrinking. Pure `Workbook => Workbook`, byte-identical on re-run.
- **16 formula functions** (#76, #120, #122) — IFS, SWITCH, CHOOSE, LARGE, SMALL,
  RANK, PERCENTILE, QUARTILE, HLOOKUP, MAXIFS, MINIFS, OFFSET, and the dynamic-array
  functions SEQUENCE, SORT, UNIQUE, FILTER. **Registry 88 → 104.** OFFSET returns a
  range (spills standalone, collapses to a scalar when 1×1) and composes with
  aggregates — `SUM(OFFSET(...))` works.
- **XLOOKUP binary-search modes** (#55) — `search_mode` 2 (ascending) and -2
  (descending) now accepted with correct iteration direction.
- **Named ranges** — `DefinedName` serialization (previously read-only) plus a
  CLI `name add` / `name rm` verb.
- **Hyperlinks** — `Cell.hyperlink` is now serialized to `<hyperlinks>` and
  populated on read; added a `hyperlink` batch op.

#### Fixed

- **Silent data loss on modify** (C1) — inline worksheet elements
  (`dataValidations`, `hyperlinks`, `sheetProtection`, `autoFilter`, …) were
  excluded from the unknown-part catch-all but never captured, so any sheet edit
  dropped them. Now preserved through the modify path.
- **`=A1=B1` / `=IF(A1=B1,…)` "Unresolved PolyRef"** (C2) — bare cell-ref
  operands in equality are now resolved during parsing.
- **Case-insensitive scalar text equality** (C3) — `="A"="a"` now matches Excel
  and the array/XLOOKUP/criteria paths.
- **1900 leap-year** (C4) — pre-1900-03-01 dates and serial 60 now convert with
  Excel parity; Scaladoc corrected.
- **Deterministic serialization** (C5) — XML-illegal control characters are
  sanitized identically across backends, and numerics emit plain decimals (no
  `1.0E+10`) in `<v>`.

### Changed

- **Scala 3.7.4 -> 3.8.3** — Language upgrade with betterFors stabilization and VarHandle-based lazy vals
- **Mill 1.1.0-RC3 -> 1.1.5** — Stable release with Scala 3.8 support
- **WartRemover 3.4.1 -> 3.5.6** — Required for Scala 3.8.3 compiler plugin compatibility

### Fixed

- **Option2Iterable WartRemover violations** — Replaced `.toSeq` on `Option` with `.toList` in Comments, Table, StyleSerializer, and StylePatcher (newly caught by WartRemover 3.5.6)

---

## [0.9.7] - 2026-04-15

### Added

- **`formula` alias for `putf` batch ops + `--dry-run` flag** (#224)
  - Batch JSON `putf` operations now accept `"formula"` as an alias for `"value"`, matching natural user intent
  - New `--dry-run` flag validates batch JSON without reading or writing files

### Fixed

- **Resolve worksheet paths via workbook relationships** (#226, GH-225)
  - Streaming reader now resolves worksheet paths from `xl/workbook.xml.rels` instead of assuming `xl/worksheets/sheet{N}.xml` naming convention
  - Fixes failures on workbooks with non-standard worksheet paths (e.g., files produced by third-party tools)

---

## [0.9.6] - 2026-02-19

### Added

- **Batch completeness — 17 operations** (#219, GH-88)
  - 10 new batch ops: `comment`, `remove-comment`, `clear`, `col-hide`, `col-show`, `row-hide`, `row-show`, `autofit`, `add-sheet`, `rename-sheet`
  - Makes `batch` the primary interface for atomic multi-step workbook modifications
  - Streaming mode supports col/row visibility ops; other new ops require full workbook mode with clear error message
  - SAX-based `readExistingWorksheetMetadata` preserves col/row properties (widths, outline levels, styles, visibility) during streaming writes

### Fixed

- **Significant digit counting in `formatGeneral`** (#218)
  - Now counts significant digits (not characters) to match Excel's General format threshold
  - Fixes edge case where `countSignificantDigits` returned 0 for `"0"` input

- **Repeatable `--with` flag** (#218)
  - `eval`/`evala` commands now support multiple `--with` flags

- **Domain row/col hidden properties override original XML** (#219)
  - `applyDomainRowProps` used OR logic making it impossible to unhide rows once hidden; domain properties now always override preserved XML attributes

- **Preserve styles on untouched sheets during metadata writes** (#220, TJC-751)
  - `StyleIndex.fromWorkbookWithSource` only created style remappings for modified sheets, but metadata-only changes (add-sheet, reorder) regenerate all sheets — unmodified sheets silently lost all styling
  - Now generates remappings for all sheets unconditionally when full regeneration is triggered, with a performance guard skipping unmodified sheets in the normal surgical path

---

## [0.9.5] - 2026-02-08

### Fixed

- **move-sheet silent data corruption** (#207, #213)
  - `move-sheet` reordered sheet names but left cell data in original positions, causing cross-sheet formula references to point to wrong data
  - Surgical writer now regenerates all sheets when reorder flag is set

- **Multi-hop same-sheet formula evaluation** (#208, #214)
  - Formulas referencing cells containing uncached formulas (e.g., `=C1` where C1=`B1+50`, B1=`A1*2`) now recursively evaluate instead of failing with `TypeMismatch`
  - Mirrors the cross-sheet recursive evaluation pattern already in place

- **`cell` command scientific notation** (#211, #216)
  - `NumFmtFormatter.formatGeneral` now uses `toPlainString` instead of `toString`, preventing BigDecimal from producing scientific notation like `5.809440E-01`

- **`eval` requires `-s` for fully-qualified formulas** (#210, #216, #217)
  - `=SUM(Revenue!A1:A3)` no longer requires `--sheet` flag since all references are qualified
  - New `containsUnqualifiedCellReferences` method distinguishes qualified vs unqualified refs

- **Style log says "(merge)"** (#210, #216)
  - Changed to "(additive)" to avoid confusion with cell merging

- **`rasterizers` returns exit code 1** (#210, #216)
  - Informational command now always returns exit code 0

### Changed

- **Documentation updates** (#209, #212, #215)
  - Function count corrected from 81 to 82
  - Added shell quoting guide for sheet names with spaces
  - Added accounting format, custom format string, and rasterizer install examples
  - Clarified `putf` vs `put`, batch JSON `numFormat` key, HTML export limitations
  - Added `move-sheet` cross-sheet reference warning

---

## [0.9.4] - 2026-02-06

### Fixed

- **Batch `put` now supports `values` array** (TJC-740, #206)
  - `{"op":"put","ref":"A1:C1","values":["Q1","Q2","Q3"]}` works like CLI `put A1:C1 "Q1" "Q2" "Q3"`
  - Per-element smart detection for numbers, booleans, dates, currency, percent
  - Consistent with existing `putf` `values` support

- **Sort now moves comments with data rows** (TJC-741, #206)
  - `Sheet.comments` (rich comments with author) are remapped to follow data during sort
  - Cell styles were already correctly preserved (confirmed via new tests)

### Added

- **SVG date rendering regression tests** (TJC-742, #206)
  - Verified date formatting works correctly in SVG output for both `Number` with `NumFmt.Date` and `DateTime` values

### Changed

- **SKILL.md documentation improvements** (#206)
  - Added batch JSON style property mapping table (CLI flags → JSON property names)
  - Documented batch `put` with `values` array

---

## [0.9.3] - 2026-02-04

### Fixed

- **Surgical write comment file path mapping** (#205)
  - Excel numbers comment files sequentially (comments1.xml, comments2.xml...) across only sheets that have comments, NOT by sheet index
  - XlsxWriter was incorrectly using sheet indices, causing Content_Types.xml to reference non-existent files
  - Fixed by preserving original comment file paths in SourceContext during read

- **Theme color serialization** (#205)
  - ThemeSlot enum ordinals didn't match OOXML spec indices (Light1/Dark1 were swapped)
  - Caused black backgrounds to appear where light fills should be
  - Fixed by adding themeSlotToIndex() for correct mapping during write

---

## [0.9.2] - 2026-02-04

### Added

- **`--strict` flag for `--eval` error handling** (#204)
  - `xl view --eval --strict` fails with non-zero exit code on formula evaluation errors
  - Default behavior unchanged: warnings to stderr, shows original formula
  - Useful for CI/CD pipelines and scripts requiring reliable error detection

- **xl-agent benchmark module** (#194)
  - LLM skill comparison framework for Excel manipulation tasks
  - Parallel execution of benchmark work units
  - Conversation tracing for debugging agent behavior

- **Exponentiation operator** (#203)
  - Formula parser now supports `^` for exponentiation: `=2^10` → 1024

- **Sheet hide/show commands** (#201)
  - `xl sheets hide <name>` and `xl sheets show <name>`
  - Support for `veryHidden` sheets (only visible via code)

### Fixed

- **Issues triaged and closed**: #192, #196, #197, #178, #179, #180, #182

---

## [0.9.1] - 2026-02-02

### Added

- **Boolean coercion and SUMPRODUCT array expressions** (#199)
  - SUMPRODUCT now supports array conditions: `=SUMPRODUCT((A:A="X")*(B:B))`
  - Boolean-to-numeric coercion: TRUE→1, FALSE→0 in arithmetic contexts
  - Enables Excel-style conditional aggregation patterns

### Changed

- **Eliminate WartRemover warnings across codebase** (#200)
  - All `var`/`while` constructs refactored to tail-recursive functions
  - Replaced `.head`/`.tail` with `.headOption` or pattern matching
  - Silenced legitimate uses in macros with `@SuppressWarnings`

---

## [0.9.0] - 2026-01-30

### Added

- **Cross-sheet SUMIFS with full-column optimization** (#195)
  - SUMIFS, COUNTIFS, AVERAGEIFS now work correctly with cross-sheet references
  - Full-column references (A:A) optimized via bounds computation to avoid iterating 1M+ rows
  - Cell references as criteria now resolve correctly (e.g., `=SUMIFS(Sheet2!D:D, Sheet2!A:A, A1)`)
  - Fixed PolyRef resolution in criteria evaluation

- **Array formula support with TRANSPOSE and broadcasting** (#191)
  - TRANSPOSE function for array transposition
  - Array arithmetic with broadcasting (scalar, row, column operations)
  - `evala` command for array formula evaluation with optional spill to cells

### Fixed

- **Formula cell type detection** (#190)
  - ROW() and COLUMN() with zero arguments now correctly use formula position
  - Proper cell type inference for formula results

---

## [0.8.1] - 2026-01-28

### Fixed

- **Aggregate functions**: Evaluate uncached formulas in SUM, AVERAGE, COUNT, etc. (#188)
  - Formula cells with no cached value are now evaluated before aggregation
  - Fixes incorrect results when aggregating ranges containing unevaluated formulas

- **INDEX function**: 2-argument form now correctly treats second arg as column number for single-row arrays (#189)
  - `INDEX(D2:G2, 4)` now returns the 4th column (was incorrectly treating as row number)
  - Fixes header lookup patterns like `=INDEX($D$2:$G$2, MATCH(MAX(D3:G3), D3:G3, 0))`
  - Single-column arrays still correctly treat second arg as row number

---

## [0.8.0] - 2026-01-27

### Added

- **Streaming writes** with O(1) memory via SAX→StAX architecture (#183)
  - Write 1M+ rows with constant memory using `excel.writeStream(path, sheetName)`
  - Early-abort optimization for efficient large file generation
  - Compatible with fs2 Stream pipelines

- **Enhanced JSON batch syntax** with typed values and format hints (#185)
  - Native JSON types: numbers, booleans, null parsed directly
  - Smart detection: currency (`$1,234.56`), percent (`45.5%`), dates (`2025-01-15`)
  - Explicit format hints: `{"value": 0.455, "format": "percent"}`
  - Formula dragging with `from` parameter for fill-down patterns
  - Opt-out smart detection with `"detect": false`

### Changed

- **12-180x streaming performance improvement** via chunked batching and range-bounded reads (#176)
  - Streaming operations now use chunked batching for optimal throughput
  - Range-bounded streaming avoids scanning entire files

- **API rename**: `writeStreamTrue` → `writeStream` for clarity (#175)

### Fixed

- Plugin structure fixes for Claude Code marketplace compatibility

---

## [0.7.0] - 2026-01-23

### Added

- **Plugin marketplace support**: Repository now functions as a Claude Code plugin marketplace
  - Install via `/plugin marketplace add TJC-LP/xl` then `/plugin install xl-cli@xl`
  - Upload `xl-plugin-*.zip` to Claude.ai via Settings > Features
  - Extensible structure for future plugins (xl-scripting, etc.)

- **Self-documenting CLI**: Comprehensive `--help` for high-value commands
  - `view`: formats, output flags, raster options, examples
  - `style`: font, fill, alignment, borders, colors, examples
  - `import`: options, type inference, limitations
  - `put`: modes (single/fill/batch), negative numbers
  - `putf`: formula dragging, anchor modes, running totals
  - `sort`: options, behavior details

- **Batch styling operations** in JSON batch commands (#88)
  - Full style properties: bold, italic, underline, bg, fg, fontSize, fontName
  - Alignment: align, valign, wrap
  - Number formats: numFormat (general, number, currency, percent, date, text)
  - Borders: border, borderTop/Right/Bottom/Left, borderColor
  - Replace mode: `"replace": true` to overwrite instead of merge

### Changed

- **SKILL.md reduced 54%**: From 861 to 406 lines by moving reference content to CLI help
- **Auto-latest installation**: SKILL.md now auto-detects latest release from GitHub API
- **Skill location**: Moved from `.claude/skills/` to `plugin/skills/xl-cli/` (distributed via plugin marketplace)

### Fixed

- **Sheet rename preserves styles**: Renaming sheets no longer loses cell formatting
- **Batch JSON parsing**: Migrated to uPickle for more robust JSON handling (#67)

---

## [0.6.1] - 2026-01-22

### Added

- **`xl rasterizers` command**: List available SVG-to-raster backends with status (#158)
  - Shows availability, notes, and delegate info for each rasterizer
  - Exits with code 1 if no rasterizers available (useful for CI/automation)
  - Parallelized checks for faster execution

- **ImageMagick SVG delegate detection** (#160)
  - Detects when ImageMagick's SVG delegate (rsvg-convert) is missing
  - Falls back to v6 if v7 delegate is broken
  - Prevents confusing "delegate failed" errors in containerized environments

### Fixed

- **CSV `--no-header` type inference**: Numeric columns now correctly detected when using `--no-header` on a CSV that has a header row (#170)
  - Changed from unanimity-based to majority-based (80% threshold) type inference
  - Individual cells that can't parse gracefully fall back to Text

---

## [0.6.0] - 2026-01-21

### Added

- **7 new CLI commands** for data manipulation and formatting:
  - `csv` - Import CSV files with automatic type detection (#149)
  - `comment` - Add, view, and remove cell comments (#151)
  - `clear` - Clear cell ranges (contents, styles, and/or comments) (#152)
  - `fill` - Excel-style fill down/right with formula shifting (#154)
  - `sort` - Sort rows by one or more columns (#157)

- **CLI enhancements**:
  - Batch `put` with smart mode detection: single cell, fill pattern, or batch values (#150)
  - `--auto-fit` flag for automatic column width calculation (#155)
  - `&` concatenation operator in formulas (#152)

### Fixed

- **Cross-sheet formula evaluation**: References to formula cells on other sheets now correctly evaluate the formula instead of returning the cached value (#162)
- **Eager recalculation**: Dependent formulas now recalculate immediately when upstream cell values change via `put`, `putf`, or `fill` commands (#164)

---

## [0.5.0] - 2026-01-12

### Fixed

- **SVG text overflow**: Added clip-path constraints to prevent text from overflowing beyond cell boundaries in SVG/PNG exports (GH-146)
  - Text properly clips when adjacent cells have content
  - Maintains Excel-like overflow behavior into empty cells
  - All raster formats (PNG, JPEG, WebP, PDF) benefit from fix

### Documentation

- **CLI skill 100% parity**: Added missing documentation for `functions` command, `--backend` flag, `--value` flag, and performance guide
  - Achieves complete coverage of all CLI features
  - Better guidance for LLM agents using xl

### Changed

- **36-39% faster streaming writes** via SaxStax backend and DirectSaxEmitter (PR #145)
  - New `DirectSaxEmitter` bypasses intermediate XML construction
  - Optimized attribute handling and string building
  - Benchmarked against Apache POI with consistent improvements

### Added

- **34 new formula functions** bringing total to 81:
  - **Financial (TVM)**: PMT, FV, PV, RATE, NPER - time value of money calculations
  - **Statistical**: MEDIAN, STDEV, STDEVP, VAR, VARP - descriptive statistics
  - **Type checking**: ISNUMBER, ISTEXT, ISBLANK, ISERR, ISERROR - value type inspection
  - **Conditional aggregation**: AVERAGEIF, AVERAGEIFS - conditional averaging
  - **Count/Reference**: COUNTBLANK, ROW, COLUMN, ROWS, COLUMNS, ADDRESS - cell reference utilities
  - **Math**: SQRT, MOD, POWER, LOG, LN, EXP, FLOOR, CEILING, TRUNC, SIGN, INT - numeric operations

- **Variadic aggregate functions**: SUM, COUNT, AVERAGE, MIN, MAX, MEDIAN, STDEV, VAR now support Excel-compatible variadic syntax
  - `=SUM(1,2,3)` - individual values
  - `=SUM(A1:A5, B1:B5)` - multiple ranges
  - `=SUM(A1, 5, B1:B3)` - mixed ranges and values

- **Dynamic `xl functions` command**: Now shows all 81 functions from registry instead of hardcoded list
- **Formula system refactored**: Reorganized into modular traits for better maintainability

### Security

- **Formula injection guards** complete (WI-31)
- **ZIP bomb detection** complete (WI-30)
- **XXE prevention** verified

---

## [0.5.0-RC2] - 2025-12-27

### Changed

- **36-39% faster streaming writes** via SaxStax backend and DirectSaxEmitter (PR #145)
  - New `DirectSaxEmitter` bypasses intermediate XML construction
  - Optimized attribute handling and string building
  - Benchmarked against Apache POI with consistent improvements

### Documentation

- Cleaned up internal planning docs, consolidated roadmap
- Removed 14 obsolete planning documents

---

## [0.5.0-RC1] - 2025-12-25

### Changed

- **36-39% faster streaming writes** via SaxStax backend and DirectSaxEmitter (PR #145)
  - New `DirectSaxEmitter` bypasses intermediate XML construction
  - Optimized attribute handling and string building
  - Benchmarked against Apache POI with consistent improvements

### Documentation

- Cleaned up internal planning docs, consolidated roadmap
- Removed 14 obsolete planning documents

---

## [0.5.0-RC1] - 2025-12-25

### Added

- **34 new formula functions** bringing total to 81:
  - **Financial (TVM)**: PMT, FV, PV, RATE, NPER - time value of money calculations
  - **Statistical**: MEDIAN, STDEV, STDEVP, VAR, VARP - descriptive statistics
  - **Type checking**: ISNUMBER, ISTEXT, ISBLANK, ISERR, ISERROR - value type inspection
  - **Conditional aggregation**: AVERAGEIF, AVERAGEIFS - conditional averaging
  - **Count/Reference**: COUNTBLANK, ROW, COLUMN, ROWS, COLUMNS, ADDRESS - cell reference utilities
  - **Math**: SQRT, MOD, POWER, LOG, LN, EXP, FLOOR, CEILING, TRUNC, SIGN, INT - numeric operations

- **Variadic aggregate functions**: SUM, COUNT, AVERAGE, MIN, MAX, MEDIAN, STDEV, VAR now support Excel-compatible variadic syntax
  - `=SUM(1,2,3)` - individual values
  - `=SUM(A1:A5, B1:B5)` - multiple ranges
  - `=SUM(A1, 5, B1:B3)` - mixed ranges and values

- **Dynamic `xl functions` command**: Now shows all 81 functions from registry instead of hardcoded list

### Changed

- **Formula system refactored**: Reorganized into modular traits for better maintainability
  - Split `FunctionSpecs.scala` into focused modules (Aggregate, Financial, Lookup, etc.)
  - Extracted `TExpr` helpers into dedicated traits
  - Deduplicated range extraction and criteria parsing helpers

### Fixed

- **Formula parsing edge cases**: Fixed scientific notation, date arithmetic, cross-sheet dependencies
- **WartRemover compliance**: Added `@SuppressWarnings` annotations for intentional type casts in formula DSL

---

## [0.4.3] - 2025-12-19

### Added

- **Rasterizer fallback chain**: Multiple SVG-to-raster backends for native image support in environments where Batik/AWT is unavailable (like Claude.ai)
  - Batik (built-in), cairosvg, rsvg-convert, resvg, ImageMagick
  - Automatic fallback tries each rasterizer in order until one succeeds
  - New `--rasterizer <name>` flag to force a specific backend

- **Targeted range evaluation**: `--eval` now only evaluates formulas within the viewed range (plus their dependencies) instead of the entire sheet
  - Significantly faster for small ranges on large sheets
  - Complexity: O(R + D) where R = range cells, D = dependencies vs O(N) for all formulas

### Fixed

- **Cross-sheet formula evaluation**: Fixed `--eval` flag to work correctly with cross-sheet references like `='Sheet2'!A1`
- **CairoSvg JPEG handling**: Now explicitly rejects JPEG format to allow fallback chain to find appropriate rasterizer

### Breaking Changes

- **`--use-imagemagick` flag removed**: Replaced with `--rasterizer <name>` option for specifying rasterizer backend
  - Migration: `--use-imagemagick` → `--rasterizer imagemagick`

---

## [0.4.2] - 2025-12-19

### Added

- **PI() function**: Mathematical constant (3.14159...) now available in formulas
- **Headless formula evaluation**: `xl eval "=PI()*2"` works without a file for constant formulas
- **`xl functions` command**: Lists all 47 supported formula functions with categories and descriptions
- **Cross-sheet formula references**: `=Sheet1!A1` syntax now fully supported in evaluator

### Fixed

- **PolyRef type safety**: Eliminated unsafe `asInstanceOf` casts in formula evaluator by resolving polymorphic references at parse time
- **VLOOKUP text lookup**: Text lookups now correctly return text results
- **Nested formula evaluation**: Formula cells with cached values now properly extract values during evaluation
- **Wildcard pattern matching**: `Widget*` patterns in SUMIF/COUNTIF now match correctly (TJC-353)
- **Addition with cell references**: Fixed regression where `=A1+B1` failed to evaluate

---

## [0.4.0] - 2025-12-17

### Added

- **Pure JVM rasterization** (PR #91): Apache Batik-based SVG to PNG/JPEG conversion
  - No external ImageMagick dependency required
  - Falls back to ImageMagick in GraalVM native images (AWT unavailable)
  - `--use-imagemagick` flag for explicit ImageMagick usage

- **Indexed color support**: Proper handling of legacy Excel indexed colors
  - Maps indices 0-63 to RGB values per ECMA-376 specification
  - Preserves indexed colors during read/write round-trips

- **Per-side border styling**: New CLI flags for individual border control
  - `--border-top`, `--border-right`, `--border-bottom`, `--border-left`
  - Border merging is now per-side (matches Excel behavior)

- **Date calculation functions**: 6 new Excel-compatible date functions
  - `EOMONTH` - End of month N months from start date
  - `EDATE` - Same day N months from start date
  - `DATEDIF` - Difference between dates (Y/M/D/MD/YM/YD units)
  - `NETWORKDAYS` - Working days between dates (excludes weekends/holidays)
  - `WORKDAY` - Date N working days from start (excludes weekends/holidays)
  - `YEARFRAC` - Year fraction between dates (5 day-count conventions)

### Changed

- **Mill upgraded to 1.1.0-RC3**: Improved Scala 3.7.3 support and build performance
  - Tested across all CI platforms (Linux/macOS/Windows)

- **Style merging is now default**: `xl style` command merges with existing styles
  - Use `--replace` flag for previous replacement behavior
  - Preserves existing formatting when adding new properties

- **Formula evaluation warnings**: `--eval` flag now prints warnings to stderr when formulas fail to evaluate instead of silently continuing

### Breaking Changes

- **`xl style` command now merges by default**: Style commands now merge with existing cell styles instead of replacing them entirely. Use `--replace` flag for the previous behavior.
  - Before: `xl style A1 --bold` replaced entire cell style with just bold
  - After: `xl style A1 --bold` adds bold to existing style
  - Migration: Add `--replace` flag to scripts that rely on replacement behavior

### Fixed

- **Batik @SuppressWarnings scope**: Moved annotation from object level to method level for tighter suppression

- **DRY forkArgs**: Extracted shared JVM options to `BuildConfig.lazyValsJvmArgs` constant

- **Partial file cleanup**: Batik rasterizer now cleans up partial files on failure using idiomatic `Using`

---

## [0.3.0] - 2025-12-13

### Added

- **Financial Functions** (PR #77): 2 new date-aware financial functions
  - `XNPV` - Net present value with irregular dates
  - `XIRR` - Internal rate of return with irregular dates

- **Security Hardening** (PR #78): Protection against malicious files
  - ZIP bomb detection with configurable thresholds
  - Formula injection guards (escape `=`, `+`, `-`, `@` prefixes)
  - Configurable via `SecurityConfig` in read/write operations

### Fixed

- **Style components not rendering** (PR #80): When adding new styles (bold, fill color) to existing files, font/fill/border components weren't being added to `styles.xml`. New styles now correctly include their component definitions.

- **Column widths lost on save** (PR #80): Column properties set via API/CLI were overwritten by preserved XML on subsequent operations. Domain properties now take priority over preserved XML.

---

## [0.2.3] - 2025-12-07

### Changed

- **GraalVM upgraded to 25.0.1 LTS**: Native image builds now use JDK 25
  - 22% heap reduction from Compact Object Headers (JEP 519)
  - Up to 8x faster String::hashCode for constant keys
  - Better Vector API support for numeric operations

### Added

- **Formula validation**: `CellValue.formula()` smart constructor validates:
  - Non-empty expression (OOXML requirement)
  - Cached value is not a nested Formula (illegal state)
- **Render constants documented**: Magic numbers extracted to `RenderUtils`
  - `IndentPxPerLevel` (21px) - Excel indent spacing
  - `InterRunGapPx` (4px) - Rich text run spacing
- **Enhanced scaladocs**: Union return type pattern explained
  - Documents `transparent inline` type narrowing behavior
  - Examples for compile-time vs runtime validation paths

---

## [0.2.2] - 2025-12-06

### Fixed

- **WartRemover warnings eliminated**: All 82 compile-time warnings resolved
  - Source code: Refactored `isInstanceOf` to pattern matching in `PutLiteral.scala`
  - Source code: Added `@nowarn` for unreachable `PolyRef` case in `FormulaShifter.scala`
  - Test code: Added class-level `@SuppressWarnings` annotations following project policy

### Changed

- **Release workflow improved**: `/release-prep` command now creates annotated tags with CHANGELOG extraction
  - Ensures GitHub releases display proper release notes instead of commit messages
  - Tags are now annotated (`git tag -a`) with message extracted from CHANGELOG.md

---

## [0.2.1] - 2025-12-05

### Added

- **`xl` aggregate package**: Single artifact bundling all modules for simpler onboarding
  - Depend on `com.tjclp::xl:0.2.1` instead of 4 separate modules
  - Individual modules (`xl-core`, `xl-ooxml`, `xl-cats-effect`, `xl-evaluator`) remain available for minimal footprint
- **`/release-prep` slash command**: Streamlined version bump workflow for releases

### Changed

- Documentation updated with simplified single-dependency examples
- Examples now use aggregate package by default

---

## [0.2.0] - 2025-12-05

Major CLI enhancement release with 38 commits across 8 PRs.

### Added

- **CLI Sheet Management Commands** (PR #63)
  - `rename` - Rename sheets
  - `move` - Reorder sheets within workbook
  - `copy` - Duplicate sheets
  - `merge` - Combine multiple sheets
  - `add-sheet` - Add new sheet with optional positioning (`--before`, `--after`)
  - `remove-sheet` - Delete sheets
  - `stats` - Sheet statistics (cell count, ranges, formulas)
  - `names` - List defined names/named ranges
  - Qualified reference support: `Sheet1!A1:B10` syntax

- **Formula Dragging** (PR #64)
  - Excel-style formula dragging with `$` anchoring support
  - `putf` command supports `--from` flag for formula range application
  - Per-endpoint anchor modes: relative (`A1`), absolute (`$A$1`), mixed (`$A1`, `A$1`)
  - Example: `xl putf B2:B10 "=SUM($A$1:A1)" --from B2`

- **CLI Syntax Improvements** (PR #65)
  - Map syntax for batch `put` and `style` operations with compile-time validation
  - `--eval` flag for formula caching in repeated evaluations
  - Improved batch operation duplicate detection

- **Repeatable `--sheet` Flag** (PR #69)
  - Create multiple sheets in one command: `xl new out.xlsx --sheet Data --sheet Summary`
  - Backward compatible with existing `--sheet-name` option

- **Example Improvements** (PR #66)
  - Shebang support for direct script execution (`#!/usr/bin/env -S scala-cli shebang`)
  - Centralized dependency management via `project.scala`
  - Standardized naming conventions (underscores instead of hyphens)

### Fixed

- **AVERAGE in Nested Formulas** (PR #68)
  - Dedicated `TExpr.Average` case prevents crash in expressions like `=SUM(A1)+AVERAGE(B1:B2)`
  - Single-pass optimization for MIN/MAX/AVERAGE using foldLeft
  - Added iterator consumption regression tests

- **CLI Batch Performance** (PR #66)
  - O(N) batch operations using grouped sheet updates (was O(N²))
  - File descriptor leak prevention in `BatchParser`

- **Excel File Corruption** (PR #62)
  - Fixed corruption during surgical modification operations
  - Improved release body generation from tag messages

### Changed

- **CLI Architecture Refactor** (PR #66)
  - Split command handlers into focused modules for maintainability
  - Extracted helper modules: `SheetResolver`, `ValueParser`, `StyleBuilder`, `BatchParser`
  - Extracted enum types: `CliCommand`, `ViewFormat`
  - Extracted shared renderer utilities: `RendererCommon`

- **CLI Behavior**
  - Requires explicit `--sheet` flag for unqualified ranges (prevents ambiguity)
  - Clearer error messages for empty range evaluations

- **Documentation**
  - Updated xl-cli skill with full CLI capabilities
  - Clarified unsafe import usage in README

---

## [0.1.4] - 2025-12-04

### Fixed

- Multiple native binary build improvements
- Release asset naming consistency

---

## [0.1.3] - 2025-12-04

### Added

- **GraalVM Native Image Support**
  - Zero-dependency native binaries for macOS (Intel/ARM) and Linux
  - `xl` CLI runs without JDK installation
- **Windows Support**
  - Native Windows binary (`xl.exe`)
  - `mill.bat` for Windows build compatibility
- **Release Automation**
  - CLI tarball included in GitHub releases
  - Versioned skill package in releases

### Fixed

- Dynamic version from `PUBLISH_VERSION` environment variable
- Windows native image build shell configuration

---

## [0.1.2] - 2025-12-04

### Added

- **Maven Central Publishing**
  - Automated publishing via GitHub Actions
  - GPG signing with Mill's `SonatypeCentralPublishModule`

### Fixed

- Release workflow environment variable conventions
- GPG key import process

---

## [0.1.1] - 2025-12-04

### Fixed

- **OOXML Compatibility**
  - Bind `r` namespace in sheet elements for openpyxl compatibility
  - Prevent chained write corruption with atomic temp file strategy

---

## [0.1.0] - 2025-12-04

Initial public release of XL - the pure functional Excel library for Scala 3.

### Added

- **Core Domain Model** (`xl-core`)
  - Pure functional Excel types: `Cell`, `Sheet`, `Workbook`, `CellValue`
  - Opaque types with zero overhead: `Column`, `Row`, `ARef` (packed 64-bit), `SheetName`, `CellRange`
  - `Patch` and `StylePatch` monoids for composable modifications
  - `CellStyle` with font, fill, border, number format, alignment
  - Compile-time validated macros: `ref"A1"`, `fx"=SUM(A1:A10)"`, `money"$1,000"`
  - DSL operators: `:=` for assignment, `++` for patch composition

- **OOXML Read/Write** (`xl-ooxml`)
  - `XlsxReader` and `XlsxWriter` for .xlsx files
  - Shared strings table (SST) with intelligent caching
  - Style deduplication via `CellStyle.canonicalKey`
  - Canonical XML output for byte-identical, diffable files
  - Surgical modification with byte-perfect preservation of unmodified parts

- **Cats Effect Integration** (`xl-cats-effect`)
  - `Excel[F]` type class and `ExcelIO` convenience object
  - SAX-based streaming reads for O(1) memory consumption
  - `fs2` streaming write support for large workbooks
  - `ExcelR[F]` for explicit error handling without exceptions

- **Formula Evaluator** (`xl-evaluator`)
  - `TExpr[A]` GADT for typed formula AST
  - 30 Excel functions: SUM, COUNT, AVERAGE, MIN, MAX, IF, AND, OR, NOT, CONCATENATE, LEFT, RIGHT, LEN, UPPER, LOWER, TODAY, NOW, DATE, YEAR, MONTH, DAY, NPV, IRR, VLOOKUP, SUMIF, COUNTIF, SUMIFS, COUNTIFS, SUMPRODUCT, XLOOKUP
  - `DependencyGraph` with cycle detection (Tarjan's SCC algorithm)
  - Round-trip verified `FormulaParser` and `FormulaPrinter`

- **CLI** (`xl-cli`)
  - Commands: `view`, `cell`, `search`, `put`, `putf`, `batch`, `export`, `new`, `sheets`
  - Export formats: CSV, JSON, TSV, HTML, SVG, Markdown
  - LLM-optimized output with `--format` options

- **Additional Features**
  - Cell-level codecs for 9 types: String, Int, Long, Double, BigDecimal, Boolean, LocalDate, LocalDateTime, RichText
  - Auto-format inference: LocalDate → Date, LocalDateTime → DateTime, BigDecimal → Decimal
  - RichText DSL: `"Bold".bold.red + " normal " + "Italic".italic.blue`
  - HTML export with inline CSS preservation
  - Optics library (Lens, Optional) with zero external dependencies
  - Row/column properties: height, width, hidden, outline level

- **High-Fidelity Rendering**
  - SVG renderer with all 14 Excel border styles
  - HTML renderer with content-aware alignment
  - Theme color resolution from OOXML theme XML
  - Text overflow into adjacent empty cells

### Performance

- **Streaming reads**: 35% faster than Apache POI at 1k rows, competitive at 10k rows
- **O(1) memory**: SAX-based streaming maintains constant memory regardless of file size
- **Surgical writes**: 11x speedup for unmodified workbooks (verbatim copy optimization)
- **Partial modification**: 2-5x speedup when only some sheets are modified

### Testing

- 731+ tests across all modules
- Property-based law verification for Monoid, Lens, and round-trip laws
- ScalaCheck generators for all core types
