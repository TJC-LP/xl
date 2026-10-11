# Changelog

All notable changes to the XL project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

The 2026-09-25 Weaver dogfood of 0.23.1 filed five findings against the CLI — `view --eval` exiting
`INTERNAL` on one unevaluable cell, `--strict` still writing `-o` (#677), three gaps in the PNG
render (text overflow, row height, conditional formats — #497), `diff` blind to sheet structure,
and `deps` expanding a whole-column range cell by cell. Each was designed by independent panels,
implemented test-first in its own worktree and adversarially reviewed (mutation checks, headless
Chromium for the HTML render, the LibreOffice oracle for the evaluator) before integration.

Wave 31 then swept the Wave-30 follow-ups (#460, #669–#681) in six worktree clusters, each
test-first and adversarially reviewed, with the evaluator and display changes pinned against a
LibreOffice 25.8 recalculation: the formula grammar Excel writes and the parser refused (union and
intersection, array constants, `TRUE()`), flat operator chains of any length, lookup and financial
error-code parity, `LOOKUP`/`XMATCH`, Excel's 11-character General display, raster backends that
exit early, a data-table staleness note in `audit`, and the library gaps `describe`/`lint` left.

### Added

- **The recalc summary names its error-valued cells, with the evaluator's reason** (#692):
  `Recalculated 3 formulas (1 error value: Data!D1 #N/A — VLOOKUP exact match not found:
  VLOOKUP(42, A1:B3, 2, FALSE))` — the first ten of `RecalcResult.excelErrors`
  (`RecalcResult.MaxNamedErrorCells`) as `RecalcResult.errorCells: Vector[ErrorValueCell]`
  (`sheet`, `ref`, `error`, `reason`), the rest counted as `… and N more`. The reason is the
  diagnostic the evaluator raised the error with (a lookup's miss echoing its call, a division's
  operands); an error that arrived as a plain value (`NA()`, an error literal, a precedent's error
  read through a reference, a cycle member) has none. `xl recalc --json` publishes the same list as
  `errorCells` (`{ref, error, reason}`) beside `errorValuedCycles`; the envelope is otherwise
  unchanged.
- **Union and intersection reference operators, array constants, `TRUE()`/`FALSE()`** (#669): the
  parser now reads `=SUM((A1,A2))`, `=INDEX((A1:B2,A1:C2),1,1,2)`, `=AREAS((A1,B1))` (LibreOffice's
  `~` union too) and `=SUM((A1:B2 B1:C2))`; the aggregates fold every area, `INDEX` gains
  `area_num`, a union anywhere else is `#VALUE!` and an empty intersection `#NULL!`, and
  intersection dependency edges are the shared cells. Array constants `{1,2;3,4}` (numbers,
  negatives, text, booleans, error values) parse, print verbatim and evaluate as arrays, in plain
  cells too (`SUM({1,2}*2)` is 6). `TRUE()`/`FALSE()` are zero-argument calls, a bare `NOT` at the end
  or before an operator is a defined name, and `Sheet1!A1:Sheet1!B2` is the range `Sheet1!A1:B2`.
  Batch `putf` and the `putf` verb accept all of them. Structured references, 3-D references and
  `LAMBDA` calls remain parse errors (#669 stays open for them).
- **`LOOKUP` and `XMATCH`** (#670): `LOOKUP` in its vector and array forms; `XMATCH` with
  match_mode −1/0/1/2 and search_mode 1/−1/2/−2 (a 2-D array or an unknown mode is `#VALUE!`).
  Misses are the typed `#N/A` `IFNA`/`ISNA` see. The registry holds 124 functions.
- **`Excel.lint` / `Excel.lintStream`** (#674): the library lint runs the same parser-backed
  `formula-unparseable` rule `xl lint` does (the oracle moved to xl-evaluator's
  `UnparseableFormula`); exported with `Finding`/`LintCategory`/`LintSeverity` from the scripting
  prelude and `com.tjclp.xl.{*, given}`.
- **`Workbook.effectiveDefinedNames`** (#674): the workbook's names including the print area and
  print titles a read lifts into `Sheet.pageSetup`, in the order the writer emits them.
  `WorkbookSummary.definedNames` (so `wb.describe`) and `describe --full` read it.
- **`ExcelIO.writeStreamStyledWithAutoDetect`** (#675): the O(1) two-pass writer with a style table,
  so a CSV imported into a new workbook through the streaming branch shows its dates as dates.
  Streaming zip entries carry time 0, so identical input writes byte-identical archives.
- **Lint rule `autofilter-name-mismatch`** (hygiene, #460): a sheet-scoped `_xlnm._FilterDatabase`
  naming a range other than the sheet's `<autoFilter ref>`. Probed on Excel for Mac and LibreOffice:
  a stale name opens without repair, so it is hygiene, and a missing or leftover name is not a
  finding. `xl autofilter` and the batch op now move the name with the filter, as Excel does.
- **`xl audit` notes stale data tables** (#678): after `put` re-solves a model, a data table's
  interior keeps its old caches (Excel parity under `autoNoTable`). The audit re-evaluates up to 8
  interior cells per table through the `recalc --tables` path and reports a `dataTableStale` note
  naming them, with the lever. Tables whose what-if cone is volatile (directly or through defined
  names) are skipped, so two audits of one file agree.
- **Cycles that settle on error values are named** (#678): `SccReport.errorValued`,
  `RecalcResult.errorValuedCycles`, `recalc --json` `data.errorValuedCycles`, and the summary adds
  `, 1 cycle settled on error values`. `certified` keeps meaning "at a fixpoint with no host errors"
  (Excel keeps an error as a stable value).

- **Conditional formatting in the pictures** (#497): `xl view --format html|svg|png|jpeg|webp|pdf`
  paints a sheet's conditional formatting as Excel does, with no flag — cell-value rules (all eight
  operators), formula rules (evaluated as array formulas, as Excel evaluates them:
  `=SUM(($A$1:$A$10=A1)*1)>1`, `=OR(A1=$X$1:$X$5)`), the four text rules, top/bottom N and N%, 2- and 3-point colour
  scales, Excel 2007 data bars, and Excel's blanks, errors and dates-occurring rules. Rules compose
  in Excel's precedence: the lower priority wins each property, non-conflicting formats combine,
  and Stop If True stops lower rules, scales and bars included. Relative references anchor at the
  block's top-left cell and shift per cell. Rules see the values the picture draws — cached, or live
  under `--eval`, which evaluates every formula the window's rules read with the window (a
  top/bottom, scale or bar rule's whole range, the cells a formula rule reads, including local
  precedents reached through workbook or sheet-scoped name aliases), so a cell's paint
  never depends on the window asked for. Without `--eval` an uncached formula (as openpyxl writes
  them) is computed, and top/bottom, scale and bar statistics follow a chain filled down or across
  whatever its length (a formula rule, or a chain read from its far end, stops at the 100-level
  recursion guard and is reported: run `xl recalc` first, or view with `--eval`). The new informational `CF_NOT_RENDERED` warning names every rule not painted
  — a kind xl does not paint yet (icon sets, above/below average, duplicate/unique values, Excel
  2010+ data bars), a rule whose formula does not parse, or a numeric rule over a cell xl cannot
  compute, which is never painted from partial statistics nor affects another rule — and every
  rule painted only where it could be evaluated (`partly rendered`, naming the first cell it failed
  at). It never gates, `--strict` included. An engine failure degrades to an unpainted picture with
  that warning, never `INTERNAL`.
- **Conditional-formatting library API**: `sheet.conditionalFormatOverlay(range)` and
  `sheet.evaluateConditionalFormats(range, workbook, clock)` produce a `CfOverlay` (engine:
  xl-evaluator's `CfEvaluator`, reporting unpainted rules as `CfUnevaluated`); the new overloads
  `sheet.toSvg(range, overlay)` / `sheet.toHtml(range, overlay)` (plus full-option and
  `SvgRenderer`/`HtmlRenderer` object forms) paint it, while the existing `toSvg`/`toHtml` take no
  overlay and stay CF-blind (`CfOverlay.empty` renders byte-identically to them; their output still
  changes in this release through the text-spill and row-autofit fixes under Fixed). Also new: `Dxf.orElse` (a monoid), `Dxf.applyTo` (a homomorphism
  onto `CellStyle`), the shared `CfTextOp.formula`, and xl-ooxml's render-only `CfRenderLift.lift`.
- **Excel array lifting for scalar functions**: a scalar function handed an array evaluates once
  per element, as in Excel 365 — `=ABS(C2:C4+D2:D4)` is `{2;2;3}` under `evala`,
  `=SUMPRODUCT(--ISNUMBER(r))` counts, `=SUMPRODUCT(1/COUNTIF(r,r))` counts the distinct values of
  a bounded range without blanks. 74 functions lift: math, text, the IS* family,
  DATE/YEAR/MONTH/DAY/DATEDIF, PMT/PV/FV/NPER/RATE/RRI, VLOOKUP/HLOOKUP/MATCH/INDEX, the criteria of
  SUMIF(S)/COUNTIF(S)/AVERAGEIF(S)/MAXIFS/MINIFS, the k of LARGE/SMALL/PERCENTILE/QUARTILE, RANK's
  number and NPV's rate; IFERROR, IFNA and XLOOKUP lift their first argument only
  (`IFERROR(5,A1:A3)` is 5); the Analysis ToolPak functions (EDATE, EOMONTH, WORKDAY, NETWORKDAYS,
  YEARFRAC, MROUND) lift an array but answer `#VALUE!` for a multi-cell range reference, as Excel
  does (pass `+A2:A10` in an array context; in a plain cell it is intersected like any operand).
  Arrays broadcast like the operators, each lifted argument is evaluated
  once, an element that fails to coerce is `#VALUE!`, a function's own domain error is its Excel
  error value per element (see Fixed) and a host failure (a missing sheet) fails the call. The lift is
  implemented once, at the call node (`FunctionFlags.lift` = `ArrayLift.Off | On(slots,
  rangeRefs)`, `ArgSpec.scalarSlots`/`replaceScalarSlots`, `ExprCoercer.liftSlot`, `LiftSlot`);
  the policy is not yet published in `xl functions --json`. `docs/LIMITATIONS.md` lists the
  exclusions and the known divergences (whole-column trimming, numeric-text and cached-formula
  single references); the selected IF branch is array-aware inside an array context
  (`=SUMPRODUCT(IF(TRUE,ABS(r),0))` sums every element).
- **`sheet.evaluateForRangePerCell(range[, clock, workbook])`** returns
  `RangeEvalResult(values, failures, blocked)` — the total, per-cell range evaluation behind
  `view --eval`, with `isClean` and a one-line `summary`; exported through `com.tjclp.xl.{*, given}`
  and the scripting prelude. `evaluateForRange` / `evaluateWithDependencyCheck` keep their
  fail-fast semantics (and no longer throw).
- **`QualifiedGraph` declared view**: `declaredPrecedentsOf(q)`, `declaredPrecedents(q, depth)`,
  `occupiedIn(range)`, `QualifiedGraph.DeclaredRange` (Excel display spelling via `a1`; build with
  the anchor-stripping `DeclaredRange.of`), `enum QualifiedGraph.Node { Cell, Range }` and
  `nodeOrder`, also through the scripting prelude.
- **`xl diff --cells-only`** restores the cell-only comparison and exit code.
- **`xl deps --expand`** lists a range precedent's occupied cells one by one (the pre-0.24 shape).

### Changed

- **Breaking: `--raster-output` must name a regular file** (#690): raster exports now write beside
  the output path and publish only a verified image, so a device or pipe (`--raster-output
  /dev/stdout`, a FIFO), a directory or `/` is `IO_WRITE` before any backend runs, where a backend
  used to write to it directly. Write to a file, then `cat` it. A new output in a missing or
  read-only directory is `IO_WRITE` too, instead of every backend failing into "No SVG rasterizer
  available".
- **Breaking: plain lookups follow Excel's error codes** (#670): an error-typed lookup_value is the
  answer (`VLOOKUP(1/0,…)` is `#DIV/0!`); `MATCH` reads match_type by sign (5 is 1, −3 is −1);
  approximate matching compares text case-insensitively over sorted text and never crosses text
  with numbers, across VLOOKUP/HLOOKUP/MATCH/XLOOKUP/XMATCH/LOOKUP; a blank lookup_value matches
  nothing; `INDEX` past its array caches `#REF!` (negative `#VALUE!`), and an XLOOKUP shape mismatch
  `#VALUE!`. Public API: the lookup_value slot of `MatchArgs`/`XLookupArgs` is `TExpr[CellValue]`
  and `xl functions --json` names it `cell`; `IndexArgs` is a 4-tuple over a
  `ReferenceOperators.Operand` (area_num).
- **Breaking: financial failures cache `#NUM!`** (#670 h): XIRR/IRR without a sign change, XIRR/XNPV
  count mismatches and empty ranges, zero slopes, numeric overflow and `NPV` at rate −1 were host
  failures that left the cell uncached (so `xl audit` called inherited books uncached); they are now
  the cached `#NUM!` Excel shows (unreadable flows or dates `#VALUE!`).
- **Breaking: General cell display follows Excel's 11-character rule** (#672): `view`, the renders
  and `--json` `formatted` show 0.000012345 as `1.2345E-05` and 123456789012 as `1.23457E+11`
  (values keep every digit); the E form no longer depends on the locale, and a hostile stored
  exponent renders bounded. `TEXT` keeps the text-conversion rule, also through a `General` keyword
  inside a custom code.
- **Scaling commas before the decimal point** (#672): `#,##0,.0` and `0,,.0` scale by 1,000 per
  comma, as Excel documents. Any other stray comma is literal text (`0"x",` no longer drops it) and
  `General0` renders the value once.
- **A cached formula in a text position reads its value** (#671): `=""&B1` over a cached
  `B1` is its cached value, not its formula text, without a prior recalc; boolean and date literals
  in text positions re-print as written (`=TRUE&""`).
- **Operator chains cost no nesting** (#680): the 128-level budget counts nesting (calls,
  parentheses, prefix/postfix operators), so flat `+`, `&`, comparison and `*` chains of any length
  parse, including terms ending in `%` or holding an intersection. A new per-formula budget of 1,024
  binary operators (`TooManyOperators`) bounds what still recurses; Excel's own limit is the
  8,192-character formula.
- **Parse errors read the same everywhere** (#681): the caret sits under the character it names,
  counting the `=`; `putf` leads with `the formula does not parse` and puts the formula block on its
  own lines; every gate caps the echoed formula at 80 characters (a window around the caret for a
  longer one, no caret for `FormulaTooLong`, never splitting a surrogate pair); a batch drag's
  re-parse fails with a typed `FORMULA_ERROR`; `xl eval` prints the diagnostic, not the parser's
  case class; `--stream putf` names the failing cell as the in-memory verb does; `xl audit` and
  `xl lint` render an unparseable formula on one line, line breaks flattened (#676).
- **A global flag a verb cannot take is named** (#676): `xl -f f.xlsx -s Data names` says
  `names does not take -s/--sheet` (every sheetless verb), and `xl new -o out.xlsx` says so with the
  hint `xl new out.xlsx`, each with the repaired command line; a second mistake on the line
  (`view --bogus -o x`) is reported first. The `--json` stderr `Error:` echo stays, documented.
- **Raster failures are typed** (#673): a forced backend that cannot write the format
  (`--rasterizer rsvg-convert --format jpeg`) and a conversion failure are `RASTERIZER_UNAVAILABLE`
  with the backend's stderr, not `INTERNAL`.

- **Breaking: a failed `--strict` gate writes nothing, with `-o` as with `-i`** (#677): a write
  whose recalculation reports formula-evaluation errors, iterative non-convergence, data-table seed
  warnings or an off-grid `#REF!` still exits 1 (`RECALC_GATE`) with its full summary, but the
  staged output is now discarded under `-o` too: the target is not created, and a file already at
  that path (the input itself when `-o` names `-f`) is left byte-identical. The summary reads
  `NOT saved (--strict failure): nothing written to <out>` (`-i`, and `-o` naming `-f`, keep
  `<file> left untouched`) and `--json` reports `saved: null`, `written: false`. Before, `-o`
  committed the gated workbook, so a pipeline that checked for the file instead of the exit code
  shipped formulas left uncached. One rule now governs every write — a staged file is published
  only by a run that exits 0 (`Outcome.publishesOutput`) — and the exit table (`xl --help`,
  `xl schema --json`, `generated/exit-codes.md`) says exit 1 writes nothing. It covers `put`,
  `putf`, `fill`, `copy`, `batch`, `recalc`, the structural verbs and `name add|rm`, including their
  backend-loading `--stream` runs; `--stream --strict` on the O(1) writes is still refused. A file
  left at the `-o` path by an earlier run is not replaced, so branch on the exit code or
  `data.written`. **Remedy**: to keep the file for inspection, drop `--strict` — the flag decides
  only the verdict, so the advisory run writes the identical bytes, exits 0 and prints the same
  summary.
- **Breaking (CLI contract): `deps` lists a range precedent as ONE node**: `deps` on
  `=SUMIFS(Data!B:B,Data!A:A,A1)` printed every occupied cell of both columns (18,714 lines on
  Weaver's book) and `cell` printed them all on one `Dependencies:` line. A range the formula
  reads is now one node, as Excel's Trace Precedents draws one box, carrying its counts in plain
  digits: `1  Data!A:A  range, 401 occupied cells`, plus `(N formulas)` when any of its cells are
  formulas. Ranges are spelled as Excel displays them, anchors dropped (`A:A`, `3:5`, `1:1048576`
  for the whole sheet, `B2:B9`); a one-cell range slot is a cell node; a SUMIF sum range shows at
  the size Excel reads it. Depth still counts formula hops: the walk continues through the formulas
  inside a range, a cell an earlier layer listed or covered is not listed again, and a range
  reached later is still listed once. A range with no occupied cell — even over a sheet the book
  lacks — is listed with `0 occupied cells` (it explains a zero). JSON: every node gains `kind`
  (`"cell"` | `"range"`) after `ref`; a range node is `{ref, kind, depth, formula: null,
  value: null, occupied, formulas}`; `data` gains `expand`. Dependents change only by
  `kind: "cell"`. `cell`'s `Dependencies` names each range once the same way (`=SUM(B1:B3)` gives
  `B1:B3`), still pasteable, `--json` still an array of strings. `QualifiedGraph` gains a fourth
  case-class field, `rangeCells` (no default): build graphs with `QualifiedGraph.of`.
- **`xl diff` compares sheet structure, not only cells** (a row-height-only change printed "Files
  are identical."). Besides cells, merges, comments and hyperlinks it reports rows and columns
  (effective `height`/`width` at 2 decimals — explicit, else the sheet default, else Excel's stock
  15pt / 9.140625 — `hidden`, `outlineLevel`, `collapsed`, default style), collapsed into Excel
  whole-row/column runs (`5:200`, `K:XFD`); sheet properties (`defaultRowHeight`,
  `defaultColumnWidth`, the `freezePanes` anchor, `visibility`); conditional formats and data
  validations as blocks keyed by sqref (token order, renumbered priorities, dxf ids and `xr:uid`s
  ignored when rule order is unchanged; precedence is compared across blocks as well as within
  them); the relative order of shared sheets; and defined names by scope and case-insensitive
  name (`_xlnm.*` skipped). View state, tab colour, print setup, tables, autofilter, drawings,
  `calcPr`, theme and document properties are not compared. **Behaviour change:** a
  structure-only difference is now `DIFFERENCES_FOUND` (exit 1); `--cells-only` restores the old
  scope. The JSON gains typed keys (every existing field unchanged); markdown is byte-identical
  when nothing structural changed.
- **Breaking: a plain formula cell evaluates as Excel's legacy formula (implicit intersection)**.
  Excel 365 opens a plain `<f>` (what `putf`, batch `putf`, openpyxl and most writers produce) as
  a legacy formula, and xl now computes it the same way. `PlainCellIntersectionSpec` pins xl to
  LibreOffice's recalculation of 327 plain cells inside and outside the referenced rows;
  `docs/LIMITATIONS.md` lists where xl, LibreOffice and Excel differ.
  - A reference in a value position — an operand (`=A1:A10*2`), `&`, a scalar argument
    (`=ROUND(A1:A10/7,1)`), a criterion (`COUNTIF(r,">"&r)`), the IF condition, the CHOOSE index,
    NOT's operand, a bare `=A1:A10` — reads the cell in the formula's own row (a column) or
    column (a row), and is `#VALUE!` where that row or column misses it.
  - Before, xl read the range's first cell (GH-302) or failed with "must be used within a
    function", and an IF condition over a range was an array whose first element decided every
    row.
  - An aggregate's argument (SUM, COUNT, AVERAGE, MIN, MAX, STDEV, … and AND/OR, ROWS, COLUMNS)
    keeps a reference whole but evaluates an expression as a value. **`=SUM(A1:A10*E1:E10)` in
    row 5 is now A5*E5, and `#VALUE!` outside rows 1–10, where it was the array sum** (so was
    `=SUM(ABS(r))`). Excel 365 gets the array sum for that typed formula only because it saves it
    as an array formula. Write `SUMPRODUCT(...)` for array math in one cell, with conditions as
    `--(r>0)` factors rather than IF.
  - References stay references until a position reads them. A single cell is one; IF, IFS,
    CHOOSE and SWITCH return the reference they select; OFFSET, INDIRECT and INDEX return theirs;
    a name bound to or computing a reference (a dynamic range, a scenario switch), including an
    alias chain to that name, is it; a LET
    name bound to one is it. Whole in an aggregate (`SUM(IF(A1:A10>2,A1:A10,0))` in a row where
    the condition holds is the whole-range sum, `SUM(IF(TRUE,C1,0))` skips a text C1), counted by
    ROWS without reading it, intersected in a value position.
  - A value passed to an aggregate follows Excel's typed-argument rule: TRUE is 1, `"5"` is 5, and
    other text is `#VALUE!` where it failed loudly.
  - LET never changes a value: a binding evaluates in the cell's mode and keeps a reference a
    reference (`LET(r,nmRef,SUM(r))` is `SUM(nmRef)`), including a reference returned by a nested
    LET: `LET(x,LET(y,A1:A3,y),SUM(x))` sums all three cells in a plain formula, as verified in Excel.
  - Array contexts keep dynamic-array semantics: array formulas (CSE and dynamic-array records,
    in an iterative cycle too), `evala`, SUMPRODUCT, FILTER's include, named formulas,
    conditional-format formulas, and a formula evaluated without a cell position (`xl eval`,
    `sheet.evaluateFormula(f)`: as the formula typed into a new Excel 365 cell, showing its
    top-left value, where a lifted range used to be a loud error). The raw `Evaluator.eval(expr,
    sheet)` without a cell treats a value-position range as the loud `@` failure; pass the cell,
    or use `sheet.evaluateFormula`.
  - Every path that evaluates a stored cell evaluates it at its own position by its kind: `recalc`
    and `view --eval`, `putf --no-recalc` / batch `putf` caching, `recalculateDependents`, `copy
    --values-only`, the data-table seeder's source formula (a CSE source seeds its array value),
    the JSON renderer's `--eval`, and the display of an uncached cell (`sheet.displayCell`, the
    `excel"…"` interpolator with `EvaluatingFormulaDisplay`) through the new
    `FormulaDisplayStrategy.formatAt(formula, numFmt, sheet, at)`, whose default delegates to
    `formatCached`.
  - Recalculated books change wherever a plain cell relied on the old reading.
- **Cached values move where formulas relied on the old array collapse**: recalculated books
  change for array contexts (SUMPRODUCT arguments, CSE records), and a defined name bound to a
  single blank cell now reads as that cell (LEN 0, YEAR 1900) instead of as 0. An uncached
  precedent evaluated on demand now runs at its own cell position and in its own mode (array mode
  for an ArrayFormula record), whether a reference or an aggregate, criteria or SUMPRODUCT range
  walk reads it.
- **`FunctionFlags` gains a field (`lift`)**: source-compatible, binary-incompatible for positional
  construction (the `volatile` precedent, #597).
- **`TExpr.cond` takes a `using ExprCoercer[A]`**: the IF call is brought to `A` through the slot's
  coercer, so a mistyped branch is a `Left` rather than a mistyped value.
- **A contained evaluator defect in `xl eval` / `evala` is `FORMULA_ERROR`** (still exit 3) with
  the defect wording, where it was `INTERNAL`.

### Fixed

- **Array contexts keep the reference IF and CHOOSE select** (#711): in an array formula,
  SUMPRODUCT, a conditional-format rule or a positionless `xl eval`, a whole column that IF, IFS,
  CHOOSE or SWITCH selected was materialized row by row and folded as an array, so
  `=AND(IF(TRUE,B:B,A:A))` was FALSE (its blanks counted as FALSE) where `=AND(B:B)` is TRUE, and
  `SUMPRODUCT(IF(TRUE,B:B,A:A))` read a million cells. A reference position (aggregates, AND/OR,
  ROWS) now folds the selected reference within the used range, as a plain cell already did;
  COUNTBLANK still counts the blank rows past it.
