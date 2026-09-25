# XL Current Limitations and Future Roadmap

**Last Updated**: 2026-06-15
**Current Phase**: Core domain + OOXML + streaming I/O complete; formula system complete (**108 functions** + cross-sheet support + dynamic arrays); structural editing (insert/delete rows & columns with formula rewriting); named-range & hyperlink authoring; tables + benchmarks complete; row/column serialization complete; **security hardening complete** (ZIP bomb detection, XXE prevention, formula injection guards in both in-memory and streaming writes); scripting prelude + whole-workbook recalculation + print setup authoring (0.11.0); typed bar/line/pie charts + embedded pictures (0.12.0); conditional formatting (0.12.1); LibreOffice-edit interop fixes (0.12.2).

> **Note (through 0.12.2):** Sections below largely predate the 0.10.0 "Trust & Author", 0.11.0 "Scripting", and 0.12.0–0.12.2 ("Visual"/"Clean Sweep"/"Interop") releases and may understate current capabilities. **Charts, drawings/pictures, and conditional formatting now ship** — see §12, §13, and §10 below. Hyperlinks (#9) and named ranges (#15) are **supported**; inline worksheet elements like data validation (#11) are **preserved through edits**, and **list-dropdown authoring shipped in 0.13.0** (#375); print setup is **partially supported** as of 0.11.0 (#16). 0.11.0 also added the scripting prelude (`com.tjclp.xl.scripting`), whole-workbook `recalculate` with per-cell errors, per-side borders/outlines, and sheet view settings — see the [CHANGELOG](../CHANGELOG.md). For the 0.10.0 rationale, see [archive/plan/v0.10.0-execution.md](archive/plan/v0.10.0-execution.md).
>
> **Known open issues**: the issues previously listed here (#262–#266, #271) were closed in the v0.11.1 "Totality" wave, and charts (#222) / drawings (#221) shipped in 0.12.0. For the current open set, see [GitHub Issues](https://github.com/TJC-LP/xl/issues).

This document provides a comprehensive overview of what XL can and cannot do today, with clear links to future implementation plans.

---

## What Works Today ✅

### Core Features
- ✅ **Type-safe addressing**: Column, Row, ARef with zero-overhead opaque types
- ✅ **Compile-time literals**: `ref"A1"` / `ref"A1:B10"` validated at compile time
- ✅ **Immutable domain model**: Cell, Sheet, Workbook with persistent data structures
- ✅ **Patch Monoid**: Declarative updates with lawful composition
- ✅ **Complete style system**: Fonts, colors, fills, borders, number formats, alignment
- ✅ **OOXML I/O**: Read and write valid XLSX files
- ✅ **SharedStrings Table (SST)**: Deduplication and memory efficiency
- ✅ **Styles.xml**: Component deduplication and indexing
- ✅ **Multi-sheet workbooks**: Read and write multiple sheets
- ✅ **All cell types**: Text, Number, Bool, Formula, Error, DateTime
- ✅ **Streaming Write**: True constant-memory writing with fs2-data-xml (100k+ rows)
- ✅ **Streaming Read**: True constant-memory reading with fs2-data-xml (100k+ rows)
- ✅ **Arbitrary sheet access**: Read/write any sheet by index or name
- ✅ **Excel Tables**: Structured data ranges with headers, AutoFilter, and styling (WI-10)
- ✅ **Elegant syntax**: Given conversions, batch put macro, formatted literals
- ✅ **Performance optimizations**: Inline hot paths, zero-overhead abstractions
- ✅ **Style Application**: Full end-to-end formatting with fonts, colors, fills, borders
- ✅ **DateTime Serialization**: Proper Excel serial number conversion
- ✅ **1904 date system** (GH-243): `<workbookPr date1904="1"/>` (legacy Mac Excel) is read into `WorkbookMetadata.date1904`, preserved on write, and `DateTime` cells are serialized with the 1904 epoch; epoch-aware conversions via `CellValue.excelSerialToDateTime(serial, date1904 = true)`. Display formatting, typed codec reads, and formula evaluation still assume the 1900 system — interpret raw serials with the metadata flag for 1904 files. Since GH-561 that includes the text form of a date in a text position (`&`, CONCATENATE, LEN): in a 1904 file `">="&A1` renders the 1900-system serial, 1,462 days off that file's numeric date cells, so a date criterion there selects the wrong rows rather than none. **Streaming-edit exception**: streaming writes (CLI `--stream put` or `--stream batch` via `StreamingTransform`) still serialize *new* date values with the 1900 epoch while the output keeps `date1904="1"`, so those cells land 1,462 days off — avoid streaming date puts into 1904-system files; in-memory writes (the default, via `XlsxWriter`) are epoch-correct.
- ✅ **Security Hardening**: ZIP bomb detection, XXE prevention, formula injection guards (WI-30)
- ✅ **Formula records** (GH-430, GH-435): `<f t="array" ref>` (legacy CSE), `<f t="dataTable" ref dt2D dtr r1 r2 del1 del2 ca/>` and the plain-formula calc flags `<f ca="1">` / `<f aca="1">` are modeled per cell (`FormulaKind` on `CellValue.Formula`, flags on `FormulaKind.Normal`) and re-emit byte-exactly through dirty sheet regeneration on every writer; data-table caches are pinned during evaluation/recalc (the `TABLE(...)` expression is derived display text, never parsed). A structural delete that removes a data-table input cell keeps the record and sets `del1`/`del2` with that input omitted, like Excel; only an edit tearing the interior itself degrades the cell to its cached constant. **Not covered**: evaluating/reseeding `TABLE()` results and data-table *authoring* sugar (GH-419); non-anchor group member cells stay cached constants and render un-braced (per-cell faithfulness — files that only mark the anchor and files that mark every interior cell both round-trip); shared-formula `si`/`bx` re-emission still expands per GH-370 and carries no record attrs; boolean record attrs normalize, so LibreOffice's explicit `aca="false"` re-emits omitted (the schema default).

### Developer Experience
- ✅ **Mill build system**: Fast, reliable builds
- ✅ **Scalafmt integration**: Consistent code formatting
- ✅ **GitHub Actions CI**: Automated testing
- ✅ **Comprehensive docs**: README, CLAUDE.md, reference + design guides under `docs/`
- ✅ **Property-based testing**: ScalaCheck generators for all core types
- ✅ **Law verification**: Monoid laws, round-trip laws, invariants

---

## Recently Completed ✅

### Style Application End-to-End
**Completed**: 2025-11-10 (P4 Days 1-3)
- Added `StyleRegistry` for coordinated style tracking
- High-level API: `sheet.withCellStyle(ref, style)`, `withRangeStyle(range, style)`
- Automatic registration and deduplication
- Unified style index with per-sheet remapping
- **Result**: Formatted cells now appear correctly in Excel!

### DateTime Serialization
**Completed**: 2025-11-10 (P4 Day 4)
- Implemented Excel serial number conversion
- Accounts for Excel epoch (1899-12-30) and fractional days for time
- **Result**: Date cells display correctly in Excel (no more "1899-12-30")

---

## Known Limitations (Categorized by Impact)

### 🔴 High Impact (Blocks Some Large-File Use Cases)

#### 1. Streaming Updates of Existing Workbooks Not Supported
**Status**: Not implemented as a first‑class API.

**What works today**:
- You can **read an existing workbook**, modify it in memory using the pure domain APIs (`Workbook.update`, `Sheet.put`, patches), and then write it back with `XlsxWriter`.
- When the workbook was read from a file, `SourceContext` + `ModificationTracker` enable *surgical modification*: unchanged parts are copied verbatim, changed sheets are regenerated, and unknown parts (charts, images, comments) are preserved.

**What is still missing**:
- A **streaming‑style “update this workbook in place” API** (e.g. “replace Sheet X with a new `Stream[RowData]` without loading all other sheets into memory”).

**Impact**:
- For very large multi‑sheet workbooks where you only want to append/replace one sheet, you currently need to either:
  - Use the in‑memory API (which loads all sheets), or
  - Implement custom ZIP‑level manipulation yourself.

**Workaround**:
- Use `XlsxReader.read(path)` → domain transforms → `XlsxWriter.writeWith(wb, path, config)` for correctness and preservation of unknown parts.

---

### 🟡 Medium Impact (Reduces Functionality)

#### 4. Merged Cells in Pure Row-Stream Writes
**Status**: Fully supported in the in‑memory OOXML path and `writeWorkbookStream`; not available in pure row-stream generation.

**Current State**:
- In‑memory:
  - `Sheet.mergedRanges: Set[CellRange]` tracks merged regions.
  - `OoxmlWorksheet.toXml` emits `<mergeCells>` / `<mergeCell>` for those ranges.
- In-memory workbook SAX/StAX write (`writeWorkbookStream`):
  - Delegates to the full OOXML writer and preserves merged cell metadata.
- Pure row-stream write (`writeStream`, `writeStreamsSeq`):
  - Writes rows from `Stream[RowData]`; there is no API for supplying merged cell metadata.

**Impact**:
- In‑memory read/write and CLI workbook writes preserve merges.
- Pure streaming‑generated workbooks will not contain merged ranges.

---

#### 5. Column/Row Properties ✅ NOW SERIALIZED
**Status**: Complete (via DirectSaxEmitter)
**Impact**: Column widths, row heights, hidden state, outline levels are preserved on write

**What Works**:
- `<cols>` element generated with width, hidden, outlineLevel, collapsed attributes
- `<row>` attributes include ht, customHeight, hidden, outlineLevel, collapsed
- Full round-trip preservation

---

#### 6. Formula System ✅ **PRODUCTION READY**
**Status**: Complete (WI-07, WI-08, WI-09a-h + TJC-351 cross-sheet formulas)
**Features**: Parser, evaluator, **108 functions** (including SUMIF, COUNTIF, SUMIFS, COUNTIFS, XLOOKUP, HLOOKUP, INDEX, MATCH, OFFSET, INDIRECT, XIRR, XNPV, and dynamic arrays SEQUENCE/SORT/UNIQUE/FILTER), dependency graph, cycle detection, cross-sheet references
**Phase**: WI-07, WI-08, WI-09a/b/c/d Complete + Financial Functions + Cross-Sheet Formulas

**What Works** (Production Ready — 1198 xl-evaluator tests):
```scala
import com.tjclp.xl.formula.{FormulaParser, FormulaPrinter, Evaluator}
import com.tjclp.xl.formula.SheetEvaluator.*

// Parse formulas to typed AST
FormulaParser.parse("=SUM(A1:B10)") // Right(TExpr.Call(FunctionSpecs.sum, TExpr.RangeLocation.Local(...)))
FormulaParser.parse("=IF(A1>0, \"Positive\", \"Negative\")") // Right(TExpr.Call(FunctionSpecs.ifFn, (...)))

// Evaluate formulas against sheets
sheet.evaluateFormula("=SUM(A1:A10)") // XLResult[CellValue]
sheet.evaluateCell(ref"B1") // Evaluates formula in B1 if present

// Safe evaluation with circular reference detection
sheet.evaluateWithDependencyCheck() match
  case Right(results) => // All formulas evaluated in correct order
  case Left(circularRef) => // Cycle detected: A1 → B1 → C1 → A1

// Cross-sheet formula evaluation (TJC-351)
main.evaluateFormula("=Sales!A1+10", workbook = Some(wb)) // XLResult[CellValue]
main.evaluateFormula("=SUM(Sales!A1:A10)", workbook = Some(wb)) // Cross-sheet SUM

// Cross-sheet cycle detection
val graph = DependencyGraph.fromWorkbook(workbook)
DependencyGraph.detectCrossSheetCycles(graph) match
  case Right(_) => // No cycles, safe to evaluate
  case Left(err) => // Cross-sheet circular reference detected
```

**Capabilities**:
- ✅ **Parsing** (WI-07): Typed GADT AST (TExpr), FormulaParser, FormulaPrinter, round-trip laws (57 tests)
- ✅ **Evaluation** (WI-08): Pure functional evaluator, total error handling, short-circuit semantics (58 tests)
- ✅ **108 Built-in Functions** (complete current registry, verified against `xl functions`):
  - **Aggregate** (12): SUM, COUNT, COUNTA, COUNTBLANK, AVERAGE, MEDIAN, MIN, MAX, STDEV, STDEVP, VAR, VARP
  - **Statistical** (5): LARGE, SMALL, RANK, PERCENTILE, QUARTILE
  - **Conditional** (9): SUMIF, COUNTIF, SUMIFS, COUNTIFS, AVERAGEIF, AVERAGEIFS, MAXIFS, MINIFS, SUMPRODUCT
  - **Logical / Selection** (13): IF, IFS, IFERROR, SWITCH, CHOOSE, AND, OR, NOT, ISNUMBER, ISTEXT, ISBLANK, ISERR, ISERROR
  - **Text** (12): CONCATENATE, LEFT, RIGHT, MID, LEN, UPPER, LOWER, TRIM, FIND, SUBSTITUTE, TEXT, VALUE
  - **Date** (12): TODAY, NOW, DATE, YEAR, MONTH, DAY, EOMONTH, EDATE, DATEDIF, NETWORKDAYS, WORKDAY, YEARFRAC
  - **Math** (17): ABS, ROUND, ROUNDUP, ROUNDDOWN, INT, MOD, MROUND, POWER, SQRT, LOG, LN, EXP, FLOOR, CEILING, TRUNC, SIGN, PI
  - **Random** (2, GH-115): RAND, RANDBETWEEN — volatile; deterministic via the `Rng.seeded` capability
  - **Financial** (9): NPV, IRR, XNPV, XIRR, PMT, FV, PV, RATE, NPER
  - **Lookup / Reference** (12): VLOOKUP, HLOOKUP, XLOOKUP, INDEX, MATCH, OFFSET, INDIRECT, ROW, COLUMN, ROWS, COLUMNS, ADDRESS
  - **Dynamic Arrays** (5): TRANSPOSE, SEQUENCE, SORT, UNIQUE, FILTER
- ✅ **Dependency Graph** (WI-09d - 52 tests):
  - Tarjan's SCC algorithm: O(V+E) cycle detection
  - Kahn's algorithm: O(V+E) topological sort
  - Precedent/dependent queries: O(1) lookups
  - Safe evaluation API: sheet.evaluateWithDependencyCheck()
  - Performance: 10k formula cells in <10ms
- ✅ **Operators**: +, -, *, /, <, <=, >, >=, =, <>, &, AND, OR, NOT
- ✅ **Scientific notation**: 1.5E10, 2.3E-7
- ✅ **Round-trip**: parse ∘ print = id (property-tested)
- ✅ **Cross-sheet formulas** (TJC-351, TJC-352 - 30 tests):
  - Single cell refs: `=Sales!A1`, `=Data!B2`
  - Range refs with aggregates: `=SUM(Sales!A1:A10)`, `=MIN(...)`, `=MAX(...)`, `=AVERAGE(...)`
  - VLOOKUP with cross-sheet tables: `=VLOOKUP(A1,Lookup!A1:B10,2,FALSE)` (TJC-352)
  - Arithmetic: `=Sales!A1 + Revenue!B1`
  - Workbook-level cycle detection: `DependencyGraph.fromWorkbook`, `detectCrossSheetCycles`

**Future Extensions** (Not Critical):
- ⏳ Extended function library (300+ Excel functions) - WI-09e+
- ⏳ Array formulas - Future work
- ⏳ Structured references (Table[@Column]) - Requires WI-10 integration
- ✅ Quoted sheet names in formulas (`='Q1 Report'!A1`) are parsed and printed; quoting of sheet names *shaped like cell refs* is tracked in #263
- ✅ Leading unary plus (`=+A1`, `=++A1`, `=2^+2`) parses and is **preserved through print** (0.13.0, #374) — round-trips byte-for-byte via a dedicated identity AST node (reverses #271's earlier normalization); evaluation is pure identity
- 🛡️ **Formula depth cap (GH-56 totality guard):** nesting is capped at 128 levels so pathological input returns `ParseError.NestingTooDeep` instead of `StackOverflowError`, and a formula holds at most 1024 binary operators (`ParseError.TooManyOperators`). Excel allows 64 nesting levels and no operator-chain limit (within 8192 chars); see "Operator chains (GH-680)" below.

**INDIRECT — dynamic references (GH-274)**: `INDIRECT(ref_text, [a1])` resolves A1-style text — `"B5"`, `"A1:B10"`, `"$A$1:B10"`, `"A:A"`, `"Sheet2!A1"`, `"'My Sheet'!A1:C3"` — at evaluation time and returns that reference, like OFFSET: whole to aggregates (`=SUM(INDIRECT("A1:A"&B1))`), counted by ROWS and COLUMNS without reading it (`ROWS(INDIRECT("B:B"))` is 1048576), materialized for array arithmetic and spill (`evala`), and intersected in a plain cell's value position. Unresolvable, out-of-grid, defined-name, external-workbook, or structured-reference text evaluates to `#REF!` (a value — total, never throws). Materializing clamps full column/row text to the sheet's used range and caps the result at 1,048,576 cells.

**Dependency semantics (differs from Excel):** the static graph sees INDIRECT's *arguments* (e.g. A1 in `=INDIRECT(A1)`), not its resolved *targets*. `recalculate()` evaluates INDIRECT-bearing formulas and their dependents after all other formulas (stable evaluate-last partition), and resolves not-yet-evaluated targets on demand with the same depth-100 recursion guard as cross-sheet references — so INDIRECT chains compute fresh values. Dynamic circular references (INDIRECT resolving into its own dependents) are not pre-detected by Tarjan; they surface as per-cell recursion-guard errors (cells left uncached) while the rest of the workbook still evaluates. Static cycle detection is unchanged (zero new false positives). xl's `recalculate()` is always full-workbook, so Excel's "volatile" marking is moot there; the targeted `recalculateDependents` treats INDIRECT-bearing cells as always dirty.

**Not supported (v1):** R1C1 mode (`a1=FALSE`) returns an evaluation error and leaves the cell uncached (Excel computes it on open) — deliberately not `#REF!`; consequently `INDIRECT(ADDRESS(..., FALSE), FALSE)` does not compose. INDIRECT is not accepted where a literal range is required at parse time (SUMIF/VLOOKUP/XLOOKUP/INDEX/MATCH/NPV/IRR array slots) — the same envelope as OFFSET. INDIRECT returns a reference, like OFFSET and INDEX: its values in an array formula, the whole range as an aggregate's argument, and in a plain cell's value position (a scalar argument, an operand) the cell implicit intersection picks — `=INDIRECT("A1:A10")*2` in row 5 reads A5 (see "Plain cells" below); `IFERROR(INDIRECT("A1"), x)` and `LEFT(INDIRECT("B1"), 2)` work for valid references. Structural row/column edits do not rewrite text inside INDIRECT strings (Excel parity — that is INDIRECT's purpose).

**Scenario-table seeding (GH-498/GH-506):** `seedDataTablesReport` and `xl recalc --tables` leave a table untouched and report `Skipped` when its source or precedent cone contains dynamic references such as INDIRECT or OFFSET. The warning names the affected cells; dynamic formulas outside the cone and input formulas replaced by axis values do not block seeding. Source/axis/cycle-member failures report the number of unseeded interiors, and unresolved-precedent diagnostics survive even when the source also fails. `--strict` fails on these warnings. Ordinary formula evaluation of INDIRECT/OFFSET remains supported.

**LET — lexical bindings (GH-193)**: `LET(name1, value1, ..., calculation)` is supported with let* semantics: each binding sees prior bindings; names are case-insensitive, must start with a letter or underscore, and must not be cell-ref-shaped; inner LETs shadow outer. Bindings used in typed argument positions (text, integer, boolean, number, date) coerce totally per the cell-decoder conventions — e.g. `LET(k, 2, LEFT("hey", k))` and `LET(d, A1, YEAR(d))` work, and uncoercible values produce a clean per-cell error, never an exception. Known divergences from Excel: (a) range-valued bindings work via parse-time substitution of literal ranges (`LET(r, A1:A10, SUMIF(r, ">5"))` works), but a binding whose value is a range-*returning* call (e.g. OFFSET) or a name cannot be used where a literal range is syntactically required (COUNTIF's range) — it binds the reference, which works in SUM, ROWS, SUMPRODUCT and value positions; (b) re-printed or structurally-shifted LET formulas show range bindings substituted into the body — `=LET(r, A1:A3, SUM(r)/COUNT(r))` reprints as `=LET(r, A1:A3, SUM(A1:A3)/COUNT(A1:A3))`, so structural row/column edits rewrite stored formula text with the literal ranges in place of the body's `r` usages (AST law and semantics preserved); (c) LET is a parser special form, not a registry function: it does not appear in `FunctionRegistry.allNames`, so CLI `functions` listings omit it; (d) dependency extraction over-approximates: binding values contribute their cell refs even when the binding is unused in the body; (e) duplicate-name rebinding is allowed (Excel rejects), dotted names (Excel-legal) are rejected.

**Omitted arguments (GH-603)**: an empty argument slot — `RATE(nper,,pv,fv,)`, `IF(cond,,x)`, `INDEX(rng,,2)` — parses to an explicit `TExpr.Missing`, Excel's blank. A typed slot coerces it like a blank cell (`LEFT("abc",)` is `""`, `SUM(1,,2)` is 3, `DATE(2024,,1)` is December 2023); a value slot reads it as 0 (`IF(TRUE,,5)` is 0, as in Excel); a range slot refuses it; an optional slot holds it as a *present* blank (GH-654): `VLOOKUP(x,rng,2,)` is an exact match, `MATCH(x,rng,)` exact, `RATE(…,,)` starts its iteration at 0, `LOG(10,)` is an error (base 0), `INDEX(rng,,2)` is the whole column. The dynamic-array and reference functions — OFFSET, SEQUENCE, SORT, UNIQUE, FILTER, XLOOKUP — read the empty slot as omitted, as Excel does (`SORT(rng,,-1)` sorts by its first column; `XLOOKUP(x,a,b,,0)` is `#N/A` on no match). An empty *range* slot (`SUMIF(rng,crit,)`) is absent — Excel reports an error there. Every slot prints back empty, trailing ones included; the human-facing form spells an empty slot as `,,` rather than `, ,`. Functions whose xl arity is stricter than Excel's (LEFT/RIGHT require two arguments) accept the empty slot but not its absence. Rules verified against LibreOffice's recalculation.

**`@` — implicit intersection (GH-604)**: Excel 365's `@x` is stored as `_xlfn.SINGLE(x)`; the model keeps the formula-bar spelling, `FormulaStorage` maps the two at every formula-text boundary (cell `<f>`, conditional-formatting and data-validation formulas, `<definedName>`) for any operand the parser accepts — a reference, name, call, literal, parenthesized expression or nested `@`, rescanned so `@INDEX(@rng,1)` maps whole — the parser accepts both spellings, and the printer emits `@x`. Semantics per Excel: a single cell or scalar is itself, a column vector yields the cell in the formula's row, a row vector the cell in its column, a 2-D range or a vector the formula's row/column does not cross is `#VALUE!`; a reference a defined name, a LET name or a function (OFFSET, INDIRECT, INDEX, IF, CHOOSE…) evaluates to intersects the same way, so `@f` in a plain cell is `f`; an array *value* (SEQUENCE, FILTER, arithmetic) collapses to its top-left element. A multi-cell `@range` evaluated without a cell position (`xl eval`) is a per-cell error rather than a guess. Not modeled: `@` before anything other than one primary (`@-A1` is a parse error).

**Plain cells are legacy formulas (implicit intersection)**: a plain formula cell — an `<f>` without `t="array"`, which is what `putf`, batch `putf`, `fx"…"`, openpyxl and most writers produce — is a legacy formula to Excel 365: Excel evaluates it with implicit intersection wherever the legacy evaluation applied it (and shows `@` there). xl evaluates it the same way, at its own cell:
- **A reference in a value position is intersected** with the formula's cell. Value positions are the operands of `+ - * / ^ % &` and the comparisons, scalar function arguments, criteria and lookup values, the IF/IFS condition, the CHOOSE index, NOT's operand, and the whole formula. A column reads the formula's row, a row reads its column, and anything the formula's row or column does not cross — a 2-D range included — is `#VALUE!`, an error value IFERROR and ISERROR see. `=A1:A10*2` in row 5 is A5*2; `=A:A` over a blank row is 0.
- **A reference-class argument takes a reference whole, but evaluates an expression as a value.** This covers the arguments of SUM, COUNT, COUNTA, COUNTBLANK, AVERAGE, MIN, MAX, MEDIAN, STDEV, STDEVP, VAR and VARP, of AND and OR, and of ROWS and COLUMNS. `=SUM(A1:A10*B1:B10)` in row 5 is A5*B5, and `#VALUE!` outside rows 1–10 (`=STDEV(A1:A10*1)` sees one value there, so it is `#DIV/0!`, as in Excel). The array sum needs `SUMPRODUCT(A1:A10*B1:B10)` or an array formula (Excel 365 saves a typed `=SUM(A1:A10*B1:B10)` as one; xl cannot write one from the CLI yet). A value passed to an aggregate follows Excel's rule for a typed argument: TRUE is 1, `"5"` is 5, other text is `#VALUE!`, COUNT skips it, COUNTA counts `""`.
- **References stay references until a position reads them.** A single cell is a reference; IF, IFS, CHOOSE and SWITCH return the reference they select; OFFSET, INDIRECT and INDEX return the reference they compute; a name bound to a range, or computing one (OFFSET, INDIRECT, INDEX, IF or CHOOSE over references — a dynamic range, a scenario switch), is that reference, and so is a LET name bound to one. A reference position reads it whole: `SUM(IF(A1:A10>2,A1:A10,0))` in row 5 is the whole-range sum when A5 > 2, `SUM(IF(TRUE,C1,0))` skips a text C1 as `SUM(C1)` does, `AND(C1)` ignores it as `AND(C1:C1)` does (`#VALUE!` when nothing is logical), `ROWS(OFFSET(A:A,0,1))` is 1048576 without reading a cell, and OFFSET moves from it (`OFFSET(dyn,0,1)`, `OFFSET(INDEX(A:B,0,2),1,0)`). A value position intersects it: `=IF(TRUE,A1:A10,0)` in row 5 is A5, and `=@IF(TRUE,A1:A10,0)` too — `@` intersects a returned reference exactly as the plain cell does.
- **An array value** (TRANSPOSE, SEQUENCE, FILTER) reads its top-left element in a value position and folds whole as an aggregate's argument (`SUM(SEQUENCE(3))` is 6).
- **An array constant is a value, never intersected** (GH-669): an operator over array constants and single cells keeps its array, and a lifted function lifts over one, as Excel's legacy evaluation does without CSE — `=SUM({1,2}*2)` is 6, `=SUM(ABS({-1,-2}))` is 3, `=SUM(COUNTIF(A1:A10,{1,2,3}))` counts three values; the cell shows an array result's top-left element (`={1,2}` is 1).
- **The operands of `,` (union) and ` ` (intersection) are reference positions** (GH-669): implicit intersection applies once, to the operator's result — `=A1:C10 B1:B10` in row 5 is B5, in row 12 `#VALUE!` — never to its operands.
- **LET never changes a value**: a binding evaluates in the cell's mode, and one that denotes a reference binds the reference, so `LET(x, e, b)` is `b` with `e` written in for `x` (`LET(r,nmRef,SUM(r))` is `SUM(nmRef)`, `LET(r,nmRef,r*2)` in row 5 is the row's cell doubled). A named formula (a name bound to a formula) is an array context, as Excel evaluates defined names.

Array contexts keep dynamic-array semantics: array formulas (CSE `{=…}` records and dynamic-array anchors, in an iterative cycle too), `evala`, SUMPRODUCT's arguments, FILTER's include, named formulas, conditional-format formulas (Excel evaluates them as arrays), and a formula evaluated without a cell position (`xl eval`, `sheet.evaluateFormula(f)`: the formula as typed into a new Excel 365 cell, showing its top-left value). `evaluateFormula(f, clock, wb, Some(cell))` is that plain cell's evaluation; the raw `Evaluator.eval` without a cell treats a value-position range as the loud `@` failure. `PlainCellIntersectionSpec` pins xl to LibreOffice's recalculation of 327 plain cells, inside a range's rows and outside them, and pins separately the shapes where xl follows Excel against LibreOffice.

Where xl differs from legacy Excel (as Excel 365 recalculates the plain `<f>` xl writes):
- SUMPRODUCT passes the array class down into IF, IFERROR and TRANSPOSE: `=SUMPRODUCT(IF(A1:A10>2,1,0))` counts, as LibreOffice and a formula typed into Excel 365 do, while legacy Excel — and Excel 365 opening this plain `<f>` — needs CSE there and computes the formula row's value. Inside SUMPRODUCT write conditions as factors: `=SUMPRODUCT(--(A1:A10>2))`, `=SUMPRODUCT((A1:A10>2)*A1:A10)`.
- TRANSPOSE's argument reads its whole range.
- `A1#` in a plain cell is an array value, read at its top-left in a value position.
- XLOOKUP's result is a value, not a reference.
- ROW and COLUMN over a multi-cell range return its first row or column number in every context: `=SUM(ROW(A1:A10))` in a plain cell is 1, as in legacy Excel, but `=SUMPRODUCT(ROW(A1:A10))` is 1 where Excel folds 55.
- An intersected error criterion is `#VALUE!` in COUNTIF, as in LibreOffice; Excel may count the matching error cells.

Where xl differs from LibreOffice, following Excel's rule (not verified against Excel itself):
- A named formula evaluates as an array formula: `=SUM(nm)` over `nm = IF(Sheet1!$A$1:$A$10>2,1,0)` counts every row (LibreOffice intersects the condition with the referencing cell).
- A reference's logical and numeric-text cells are skipped by SUM (Microsoft: "only numbers in that array or reference are counted"): `=SUM(C1)` and `=SUM(IF(TRUE,C1,0))` over a TRUE C1 are 0 (LibreOffice counts 1).
- A 2-D reference is `#VALUE!` even from another sheet (LibreOffice intersects both axes there).
- `+range` is an operand, so it is intersected (Excel's operand class; LibreOffice drops the plus).

**Array lifting — arrays in scalar function arguments**: Excel 365 applies a scalar function to each element of an array passed where it expects one value, and xl does the same for 74 functions (`FunctionFlags.lift`, implemented once at the call node): `=ABS(C2:C4+D2:D4)` is `{2;2;3}` under `evala`, `=SUMPRODUCT(--(B2:B4>0),ABS(C2:C4+D2:D4))` is 5, `=SUMPRODUCT(--ISNUMBER(B2:B4))` counts, `=SUMPRODUCT(1/COUNTIF(r,r))` counts distinct values (exactly so for a bounded `r` without blanks: a blank criterion counts zeros, so a blank in `r` adds 1/COUNTIF(r,0) or is `#DIV/0!`, as in Excel; for whole columns see the trimming note below). Lifted: ABS, SIGN, SQRT, INT, TRUNC, ROUND, ROUNDUP, ROUNDDOWN, CEILING, FLOOR, MOD, POWER, EXP, LN, LOG; LEN, LEFT, RIGHT, MID, UPPER, LOWER, TRIM, SUBSTITUTE, FIND, SEARCH, TEXT, VALUE, CONCATENATE; ISNUMBER, ISTEXT, ISBLANK, ISERROR, ISERR, ISNA, N, ERROR.TYPE, and IFERROR/IFNA on their value only (`IFERROR(5,A1:A3)` is 5); DATE, YEAR, MONTH, DAY, DATEDIF; PMT, PV, FV, NPER, RATE, RRI; VLOOKUP, HLOOKUP, MATCH, INDEX, XLOOKUP (its lookup value only); the criteria of SUMIF, COUNTIF, AVERAGEIF, SUMIFS, COUNTIFS, AVERAGEIFS, MAXIFS, MINIFS; the k of LARGE, SMALL, PERCENTILE, QUARTILE, the number of RANK, NPV's rate. The Analysis ToolPak lineage (EDATE, EOMONTH, WORKDAY, NETWORKDAYS, YEARFRAC, MROUND) lifts over an array but answers `#VALUE!` for a multi-cell range *reference* in a scalar argument, as Excel does: `EOMONTH(A2:A10,0)` is `#VALUE!`, `EOMONTH(+A2:A10,0)` spills. Not lifted, each for its own reason: IF, IFS, SWITCH, CHOOSE, NOT, AND, OR (their own array semantics); the aggregates and SUMPRODUCT (they take arrays); TRANSPOSE, SORT, UNIQUE, FILTER, SEQUENCE, ANCHORARRAY, SINGLE (array producers); ROW, COLUMN, ROWS, COLUMNS, CELL, OFFSET, INDIRECT, ADDRESS, HYPERLINK (reference-shaped); IRR, XIRR, XNPV; RAND, RANDBETWEEN; the no-argument functions — those keep the typed top-left collapse below. Arrays broadcast like the operators (an axis of extent 1 repeats; a shorter extent pads `#N/A`); a 1×1 array is its element; each lifted argument is evaluated once (`RAND()` inside draws once); a per-element result that is itself an array keeps its top-left element (`INDEX(B2:D4,SEQUENCE(2),0)` is `{B2;B3}`). Per element, an Excel error keeps its code and a value that fails to coerce is `#VALUE!`; the functions' own domain errors are Excel's error values too — FIND not finding its text, LEFT/RIGHT/MID/SUBSTITUTE out of their domain are `#VALUE!`, a CEILING/FLOOR/MROUND sign mismatch `#NUM!` — so `SUMPRODUCT(IFERROR(FIND("p",A1:A5),0))` and `SUMPRODUCT(--ISNUMBER(FIND("p",A1:A5)))` count per element, as in Excel; a host failure (a missing sheet, the recursion guard) fails the whole call loudly. The lifting law — element *i* of `f(range)` is `f(cell_i)` — is pinned as a property for every lifted slot over numbers, the texts `abc`, `a` and the empty string, blanks, booleans, `#N/A`, `#DIV/0!` and a date. Two known divergences sit outside that domain, both in the single-reference decoders that predate lifting, and in both the lifted element is Excel's answer: a numeric-text cell (`'5`) lifts to its number (`SUMPRODUCT(ABS(A1:A2))` over two `'5` cells is 10) while a single reference to it in a numeric argument is a loud type mismatch (`=ABS(A1)`, like `=A1+0`); and a single reference to a cached formula cell in a text argument reads the formula's text when evaluated through the library (`=LEN(A1)` over `=1+1` cached as 2 is 3; the CLI evaluates precedent formulas first and does not show it) while the lifted element reads the cached value (`LEN(A1:A2)` starts with 1). Where lifting happens: every array context — `evala`, SUMPRODUCT's arguments, FILTER's include and, inside those, IF conditions and both IF branches under an array condition (where they broadcast against it), named formulas, spill anchors, conditional-format formulas, a formula evaluated without a cell position — and ArrayFormula records (`{=…}` CSE and dynamic-array anchors), which evaluate as arrays and cache element (0,0), the Excel anchor. A plain (Normal-kind) formula cell does not spill: it is a legacy formula, evaluated by the rules in "Plain cells are legacy formulas" above — a multi-cell reference in a lifted argument, or under an operator inside one, reads the cell in the formula's row (`=ABS(C2:C4)` and `=ABS(C2:C4+D2:D4)` in row 3 read row 3; outside the range `#VALUE!`; the ToolPak functions `#VALUE!`), while an array value there keeps its top-left element. IF branches are array contexts only under an array condition: a scalar condition selects one branch, which evaluates in the IF's own mode; inside an array context the selected branch is array-aware (`=SUMPRODUCT(IF(TRUE,ABS(C2:C4),0))` is 8). SUMPRODUCT's used-extent trimming of whole columns (GH-192) reaches inside lifted arguments, powers and `&`, so `SUMPRODUCT(--(A:A>0),ABS(B:B))` sizes match. The trimming drops the rows past the sheet's used extent, so a lifted function's value at a blank is not counted for them — the same divergence as the operators' `SUMPRODUCT(--(A:A=""))`: `SUMPRODUCT(--ISBLANK(A:A))` counts only the blanks inside the used extent (0 when column A is filled to it) where Excel counts every blank row, and `SUMPRODUCT(1/COUNTIF(A:A,A:A))` is the distinct count when A is filled to the used extent where Excel's whole-column value adds 1/COUNTIF(A:A,0) for every blank row (or is `#DIV/0!` when A has no zeros) — bound the range when its blanks must count; elsewhere in an array context a whole-column reference in a lifted argument materializes every row, as the operators do, so an array formula lifting a lookup or criteria function over one (`{=SUM(COUNTIF(A:A,A:A))}`) costs rows × used rows — bound the range; in a plain cell the criterion is intersected, so it is one COUNTIF. The `&` operator broadcasts too: `=B2:B4&"x"` is an array in array contexts and intersected in a plain cell, like any operand. Array constants (`{1;2;3}`, GH-669) lift like any array, in a plain cell too. Not modeled: lifting over IF/CHOOSE/SWITCH selectors and over the other XLOOKUP slots, and the lift policy in `xl functions --json`.

**Reference operators — union `,` and intersection ` ` (GH-669)**: `SUM((A1,A2))`, `INDEX((A1:B2,D1:E2),1,1,2)`, `AREAS((A1,B1))`, `A1:C3 B2:D4`, `SUM(A:A 3:3)` and two names (`Jan Sales`) parse, print back (one space for an intersection; a bare `,` in both the human and the file form; LibreOffice's xlsx spelling `(A1~A2)` is read as `,` and never written) and evaluate as Excel does. Precedence is Excel's: `:` > space > `,` > negation > `%` > `^` > the rest, so `-A1:B2 B1:C2` negates the intersection. The operators are unregistered calls: `=UNION(…)` stays an unknown function and `xl functions` does not list them; `AREAS` is registered. Semantics:
- **Union**: the variadic aggregates (SUM, COUNT, COUNTA, AVERAGE, MIN, MAX, MEDIAN, STDEV, STDEVP, VAR, VARP) fold every area — overlaps count twice (`SUM((A1,A1))` is 2×A1), nested unions flatten, errors in an area are triaged by the aggregate's rule as one argument (`COUNT((#REF!,A1))` is 0, `SUM((#REF!,A1))` `#REF!`); INDEX's fourth argument, area_num, picks an area (omitted or empty 1, below 1 `#VALUE!`, past the last `#REF!`); AREAS counts them. A union anywhere else is `#VALUE!` — a value position (`=(A1,A2)`, `ABS((A1,A2))`, `@(A1,B1)`), ROWS/COLUMNS (Excel answers `#REF!`), OFFSET's base, COUNTBLANK (one range), and a union IF, CHOOSE, SWITCH, LET or a defined name passes to an aggregate (Excel sums it). A union's areas must share a sheet (`SUM((A1,Other!A1))` is `#VALUE!`, where LibreOffice sums).
- **Intersection**: the cells two references share, pairwise over a union's areas (`SUM((A1:A5,C1:C5) A3:C3)` sums A3 and C3). One area is an ordinary reference in every position; several are `#VALUE!` outside the multi-area consumers; none is `#NULL!` (`ISERROR`/`ERROR.TYPE` = 1/`IFERROR` see it). The operands must share a sheet (`#VALUE!`). Its dependency edges are the shared cells, so `C3: =SUM(A1:C3 A1:A3)` is not circular.
- **Operands** are references: cells, ranges, names, LET names, `#REF!` (a deletion writes `A1:B2 #REF!`), unions and intersections, and the reference IF, IFS, CHOOSE, SWITCH, OFFSET, INDIRECT and INDEX return; a literal, arithmetic, `@x`, `x#` or an array constant is a parse error, as in Excel. A space is the operator only between two such operands, spaces only (a tab or line feed never is), and never before xl's lenient infix `AND`/`OR`; `A1 (B1:C2)` is an intersection unless the name is a function (`LOG (x)` stays a call).
- **Not yet**: a union or intersection in a range-typed slot (SUMIF, COUNTIF, VLOOKUP, MATCH, LARGE, SMALL, RANK, PERCENTILE) and multi-area support in PRODUCT, SUBTOTAL, AGGREGATE, LARGE/SMALL/RANK, FREQUENCY, AND/OR and SUMPRODUCT — parse errors or `#VALUE!`, never a wrong number; a top-level bare union (a defined name such as `Print_Area` = `Sheet1!$A$1:$B$2,Sheet1!$D$1:$E$2`) stays unparseable (ReferenceScan still reads it textually), while a parenthesized union name parses and is `#VALUE!` in SUM; `INDEX(IF(…),…)`.

**Array constants, `TRUE()`/`FALSE()`, the name `NOT` and repeated-sheet ranges (GH-669)**: `{1,2;3,4}` (`,` between columns, `;` between rows; numbers, negative numbers, text, TRUE/FALSE and error values; rows of equal width; a reference or expression inside is a parse error, as in Excel) is an array value everywhere, including `INDEX({1,2;3,4},2,1)`; lookups over one (`MATCH(2,{1,2,3},0)`, `VLOOKUP(…,{…},…)`) remain parse errors. `TRUE()` and `FALSE()` are the zero-argument calls LibreOffice writes and print back as written. A bare `NOT` at the end of a formula or before a binary operator, a closer or a separator is a defined name (`=NOT+1`), not the prefix operator. `Sheet1!A1:Sheet1!B2` is the range `Sheet1!A1:B2` (printed that way); two different sheets there is a 3-D reference, still a parse error.

**Operator chains (GH-680)**: the 128-level nesting budget counts nesting — function arguments, parentheses, prefix and postfix operators, and each intersection space (GH-669: `A1 A1 A1…` builds a nested spine of intersection calls, so a chain of more than 128 is `NestingTooDeep`) — not the length of a flat binary-operator chain: `=B2+B3+…` and `=A1&A2&…` of any length parse, evaluate and print (the evaluator, printer, shifter and dependency walkers take a chain's left-nested spine in one loop). A postfix `%` and an intersection cost levels only within the one operand they build: `=1%%%…` and `=A1 A1 A1…` past 128 steps are `NestingTooDeep`, but their levels are never summed across a flat chain's terms, so `=1%+1%+…`, `=A1*5%+A1*5%+…` and `=A1 A1+A1 A1+…` of any length parse like `=B2+B3+…`. A formula holds at most 1024 binary operators (`TooManyOperators`, with the position of the first over the limit) — a guard on what still recurses along a spine, eight times the chains generated books carry.

**`x#` — the spill reference (GH-655)**: Excel 365's `A1#` is stored as `_xlfn.ANCHORARRAY(A1)`; the model keeps the formula-bar spelling, `FormulaStorage` maps the two at every formula-text boundary (nesting under `@`), the parser accepts both, and the printer emits `x#`. The value is the array the anchor's formula spills: the file's `<f t="array" ref>` span read through its cached cells when it has a cache — so a spill from a function xl does not evaluate still reads, and a legacy CSE array record (`{=…}`, the same `t="array"` record; the `cm` dynamic-array marker is not modeled) reads the same way where Excel would say `#REF!` — else the anchor's formula evaluated as an array at the anchor's position (a scalar result is `#REF!`: not a spill anchor). A constant, a blank, a spilled non-anchor cell, or a range before the `#` is `#REF!`. Readers are dynamic-dependency cells: the static graph sees the anchor, not its extent, so they recalculate last and targeted recalculation treats them as always dirty. xl never writes a spill: an xl-authored dynamic formula (`putf A1 "=SORT(B1:B3)"`) is stored as a plain `<f>`, which Excel 365 opens as a legacy formula — it shows `=@SORT(B1:B3)`, returns one value and does not spill (Microsoft, "Formula vs Formula2") — and xl evaluates it the same way, as a plain cell. `A1#` readers still evaluate such an anchor's formula as an array, where Excel has no spill to read (`#REF!`). Writing dynamic-array formulas (the `cm` cell attribute and its metadata part) is not supported yet.

**Operator precedence (GH-578)**: negation binds tighter than `^`, as in Excel: `=-2^2` is 4, `=-A1^2/2` carries Excel's sign, `=0-2^2` (binary subtraction) is -4, `=-(2^2)` keeps its grouping. Before 0.23.0 xl read `-2^2` as `-(2^2)`; a file written by an older xl that reprinted `(-2)^2` re-parses to the same tree and reprints flat as `-2^2`.

**Lenient sheet-reference parsing (GH-281, intentional)**: the formula parser accepts unquoted cell-ref-shaped sheet references (`=Q1!A1`) that Excel itself rejects (Excel requires `='Q1'!A1`). Printing always canonicalizes to the quoted form, so anything xl writes re-parses everywhere; the leniency only widens what xl can READ. Pinned by a named parser test.

**`Cell.comment` deprecated (GH-295, since 0.12.1)**: the `Option[String]` field on `Cell` was never serialized — `Sheet.comments` is the store the OOXML writer reads. Setting it now write-throughs into `Sheet.comments` on `put` (plain text, no author) instead of vanishing; migrate to `Sheet.comment(ref, Comment.plainText(...))`. The field will be removed in a future major.

**External-workbook references — closed-workbook pinning (GH-353)**: formulas referencing another workbook (`=[2]Book1!A1`, `=SUM([2]Consolidation.xlsx!D5:D9)`, `='[3]Sheet Name'!B2`) parse, print (round-trip exact), and participate in dependency graphs (contributing **no edges** — the target lives outside the workbook). xl does **not** load external workbooks or read the `xl/externalLinks/*` cache parts (they ride as preserved parts); instead, Excel's closed-workbook model is approximated by **pinning**: a formula cell whose expression touches an external workbook keeps its Excel-written cached value verbatim — never re-evaluated, never overwritten — and dependents compute from that cache like from any cached precedent. Known divergences and conventions:

- **Stale pins on mixed formulas (differs from Excel)**: a cached formula mixing local and external precedents (`=A1+[2]Book1!B1`) is pinned verbatim even when its LOCAL precedents change — Excel would recompute the cell using the externalLink cached values, which xl carries but never reads. Editing a precedent of such a cell leaves the pin (and everything downstream of it) serving the pre-edit number until the file is recalculated in Excel.
- **Uncached external cells are a loud per-cell evaluation error, deliberately not `#REF!`** ("External workbook reference … cannot be resolved"; the cell stays uncached, dependents receive the error, and the rest of a `recalculate()` pass completes — `view --eval` likewise keeps the cell's file value and names it in its one `EVAL_FAILED` warning, while the fail-fast `xl eval` surfaces the same message as the command's failure). This follows the R1C1-INDIRECT precedent above rather than the external-INDIRECT-*text* convention (which yields `#REF!` as a value): under #344 semantics raw-range aggregates *skip* error cells, so a `#REF!`-valued external cell would make `SUM` over it silently return a partial sum — a wrong number instead of a visible failure. Note the asymmetry is intentional: `INDIRECT("[Book1]Sheet1!A1")` *text* still evaluates to `#REF!`; a direct `[2]Book1!A1` reference errors.
- **Range-typed argument slots accept external ranges** (`SUMIF([2]Book1!A1:A9, ">0")`, SUMIFS/COUNTIF(S)/AVERAGEIF(S), VLOOKUP/HLOOKUP tables, SUM/MIN/MAX/…, SUMPRODUCT). Slots requiring a LOCAL literal range at parse time (XLOOKUP/MATCH/INDEX array slots — the same envelope as INDIRECT/OFFSET) reject external ranges with a message naming the construct; a CACHED cell bearing such a shape is still pinned (lexical `[n]` fallback), so `recalculate()` stays clean for it.
- **INDIRECT/OFFSET combined with an external ref** puts the cell in the dynamic evaluate-last bucket, whose caches are deliberately stripped for the pass — the cell then reports the external error on **every** recalculation even when Excel wrote a cache; failed recalculation now clears that stale cache and its affected dependents (GH-563).
- **Structural row/column edits** in this workbook never shift or void external refs (the coordinates belong to the other workbook); plain fill-drag `shift` moves them anchor-aware like sheet-qualified refs.
- `putf` accepts external formulas and stores them **uncached** (Excel computes on open) — so an xl-written external formula reports the uncached error on recalc until Excel has cached it once.

**RAND/RANDBETWEEN — volatile functions (GH-115)**: `RAND()`/`RANDBETWEEN(bottom, top)` are supported. Volatility: xl re-evaluates volatile formulas on every `recalculate`/`withCachedFormulas` pass, so cached values change per recalculation (matching Excel). Determinism: randomness is an explicit `Rng` capability (Clock pattern) — pass `Rng.seeded(seed)` to the rng-taking overloads (`sheet.evaluateFormula(f, clock, rng)`, `wb.recalculate(clock, rng)`, ...) for reproducible output; default paths use `Rng.system`. RANDBETWEEN tightens fractional bounds inward (bottom rounds up, top down) and errors when bottom > top.

**Iterative calculation — circular models (GH-373, 0.13.0; GH-492 global fixpoint, 0.20.0; GH-482 Gauss–Seidel sweep, GH-537 stall exit, unreleased)**: professional schedules (interest on average debt) are circular by design. `recalculate(IterativeCalc(maxIter, maxChange))` opt-in fixpoints declared cycles — each cyclic component is swept until every member's |Δ| < maxChange, or maxIter rounds, or (GH-537) a round consumes no randomness and replays the previous one exactly because a member fails; non-convergence keeps the last values with NO error (Excel's semantics). Within a component the sweep is **Gauss–Seidel by default** (GH-482): members evaluate in `DependencyGraph.withinComponentOrder` — Kahn on the component's induced subgraph, cutting at the top-left remaining member (sheet name, then row, then column: Excel's documented left-to-right, top-to-bottom sweep, never the A1 text, so A10 never precedes A2) when stuck — and each value is published before the next member reads it, Excel's iteration model (`A1 = B1+1`, `B1 = A1` reaches 100/100 after 100 iterations in xl; Jacobi's previous-round reads gave ~50/50). The sweep order is xl's graph order, not a calc chain read from the file; whether Excel's own chain lands on the same member for a given book has not been checked against Excel. `IterativeCalc(…, scheme = IterationScheme.Jacobi)` reproduces the 0.13.0–0.22.x trajectories. Default `recalculate()` still isolates cycles as errors. `IterativeCalc.fromCalcPr` bridges a file's authored `<calcPr>`; `Workbook.withCalcPr(CalcPr(iterativeCalculation = true, maxIterations = Some(100), maxChange = Some(BigDecimal("0.001"))))` authors it on scratch builds.

Since 0.20.0 (GH-492) the pass walks the **SCC condensation** in dependency-first order instead of splitting pre-order / one flat Jacobi over the whole cyclic core / post-order. Every precedent is a freshly computed value before anything reads it, so:

- **`converged` is a global certificate.** It is `cycles.forall(_.converged)`, and by induction over the condensation order a `true` verdict means one more whole-workbook pass would change nothing. Two caveats survive: per-component tolerances compose across the DAG only up to error amplification, and dynamic (INDIRECT/OFFSET) cycles are invisible to Tarjan, so they are neither iterated nor covered by the flag.
- **`maxIter`/`maxChange` are per component**, not per workbook. Worst-case total work is unchanged (`maxIter × |cyclic core|`) and usually far lower; `RecalcResult.cycles` carries a per-component `SccReport(members, converged, rounds, maxDelta, stalled)` so an exhausted cycle can be named instead of counted. `stalled` (GH-537) marks a component whose loop stopped early because a round consumed no randomness and replayed the previous one exactly while a member kept failing (a missing sheet, an unresolvable name — host failures, never error values): with the random sequence untouched, every further round would have been the same replay, so the engine stops instead of burning `maxIter` (the 126k-name field book spent ~20 s of every recalculation replaying 400 identical rounds). A stalled component is not converged, its failing member is in `errors`, and `iterationsUsed` can then be below `maxIter` with `converged = false`; `render`/`summary` say "stalled after N round(s)" rather than "exhausted".
- **Re-recalculating a cached circular book is safe again.** Before 0.20.0, an acyclic cell sitting BETWEEN two SCCs was evaluated only in the post-order phase, so the downstream SCC iterated against the previous generation's cache: a cached multi-SCC book walked one SCC of wavefront per pass and, on interleaved topologies (year cycles + tax-bank SCCs), parked in a stationary WRONG state reporting `converged = false` forever with `errors` empty (GH-491). The field doctrine "never re-recalculate a cached circular multi-SCC book" is **retracted** — `recalculate(IterativeCalc)` is now idempotent on a converged book: every component recognises the input as its fixpoint in round 1, and the pass is bit-exact under cold seeding (`seedFromCaches = false`) and moves by strictly less than `maxChange` — toward the true fixpoint — under the default warm seeding.
- **One iterative recalculation is one volatile generation.** `TODAY()`/`NOW()` agree inside the fixpoints and in the acyclic cells between them (previously the fixpoint pinned the clock but the pre/post phases did not). The non-iterative path is deliberately unchanged here.
- **Iterative-path values may move by less than `maxChange`** versus 0.19.x: solving per component is Gauss–Seidel *between* components, reaching the same fixpoint on a different trajectory and stopping at a different point inside the tolerance ball. Seeded-`Rng` draw sequences shift on the iterative path for the same reason. Byte-identity is claimed for the non-iterative path only.
- **The within-component sweep order is observable (GH-482).** Under the default Gauss–Seidel scheme the order members evaluate in changes the trajectory, the round count and — for a non-convergent cycle — the last values kept (the `1-B1`/`A1` oscillator's round 10 lands on (0,0), where Jacobi's landed on (1,1); a `1-B2`/`1-B1` pair now *converges* to (1,0), as a sequential sweep in that order does — whether Excel's own chain order lands on (1,0) or (0,1) has not been checked against an engine here). The order is `DependencyGraph.withinComponentOrder`, a pure function of the graph, so results stay deterministic and independent of sheet/cell insertion order, and a seeded `Rng` is repeatable. Convergent cycles reach the same fixpoint within `maxChange`, typically in fewer rounds (the average-balance idiom: one sweep propagates the whole schedule). A round-count or last-value pin from 0.22.x needs `scheme = IterationScheme.Jacobi` to reproduce.

**Iterative warm start (GH-469, 0.20.0)**: cycle members now seed from their **loaded cached number** when they have one (0 for every other shape, and as the fallback) — Excel's semantics, iterative calculation starts from the current cell values. Before 0.20.0 every member seeded to 0, which on a mutually `IF(ISERROR(…))`-guarded pair sitting at a valid numeric fixpoint made round 1 see `0/0` = `#DIV/0!`, flipped both members to the guard's text branch, and — because that text is itself a fixpoint — reported convergence with `errors` and `excelErrors` both empty while two valid caches were destroyed. Consequences:

- **`recalculate(wb) != recalculate(stripCaches(wb))` for a circular book** is now a real, documented property. For a contraction the two agree; for a genuinely multi-fixpoint nonlinear cycle they need not. That is Excel's exposure and Excel's answer.
- **Only `Number` caches seed.** A member carrying a stale error or text cache seeds 0 and heals exactly as it did before 0.20.0. This is a deliberate departure from literal Excel parity: arithmetic propagates both shapes unchanged (`#DIV/0! * 0.5 + 10` is `#DIV/0!`, `"junk" * 0.5` is `#VALUE!`), so a non-numeric seed is its own fixpoint — a perfectly healthy cycle would wedge at the poison and report `converged = true` with `errors` empty, which is precisely the silent-failure shape GH-469 was filed about.
- A book whose cyclic members carry stale but NUMERIC caches (for example one poisoned by the 0.19.1 GH-491 bug) warm-starts **from those numbers**. `IterativeCalc(maxIter, maxChange, seedFromCaches = false)` is the supported cold-start escape hatch.
- Members whose caches the dynamic (INDIRECT/OFFSET) deferral bucket strips seed 0 regardless — those caches are declared stale for the pass.
- Fresh (uncached) books are bit-identical to 0.19.x on this axis: there is nothing to warm-start from.
- `DataTableSeeder`'s circular what-if fixpoints stay **cold-seeded** on purpose: under what-if substitution the loaded caches are stale by construction.

**Number → text conversion (GH-665)**: `&`, CONCATENATE, text-typed arguments, numeric literals in text positions and `TEXT(x,"General")` render Excel's width-independent General text — 15 significant digits, trailing zeros stripped, plain while the unsigned form fits in 20 characters (`1E19` → `10000000000000000000`, `0.000123456789012346`), E notation beyond (`1E+20`, `1.23456789012346E-05`), two-digit minimum exponent. The stored value keeps its full BigDecimal precision (DECIMAL128 `1/3` stays 34 digits in the cell, in `Raw:` output and in `--json`); only the text rendering rounds. LibreOffice's `&` conversion differs above roughly 1E15 (it switches to E notation earlier and pads exponents to three digits), so it is an oracle only for the plain range; xl follows Excel's 20-character rule. The cell-display General (`xl view`, `displayCell`) is a separate rule, below.

**Cell-display General (GH-672)**: what a General cell shows follows Excel's 11-character rule, the sign uncounted — plain while the exponent is in [-4, 10], rounded to the decimals that fit (`12345678.901` → `12345678.9`, `1/3` → `0.333333333`), E notation otherwise with the significant digits that fit (`123456789012` → `1.23457E+11`, `0.000012345` → `1.2345E-05`, `1.23456789E+100` → `1.2346E+100`); the switch reads the exponent after rounding, as `%G` does (`0.0000999999999999` → `0.0001`). The shrinking Excel applies to columns narrower than 11 characters is not modelled. LibreOffice's General differs (it shows `0.000012345` and `123456789012` in full), so it is not the oracle here; the five-digit mantissa for three-digit exponents follows from the 11-character rule and awaits an Excel spot-check. The `General` keyword inside a format code follows the rule of where it is rendered: in a cell's format the display rule above, inside a `TEXT` format code the text rule (so `TEXT(123456789012,"General;-General")` is `123456789012`, as `TEXT(123456789012,"General")` is).

**Commas in number formats (GH-666, #672)**: a comma after a digit placeholder groups when an integer-part placeholder follows it and otherwise scales by 1000, so `#,##0,.0` and `0,,.0` scale as `#,##0,` does (Excel's documented rule: "a comma that follows a digit placeholder scales the number by 1,000"). LibreOffice 25.8 refuses those before-the-point codes (it renders them General) — a divergence where xl follows Excel. A comma after anything else — a literal, a space, `%`, the decimal point, `General`, or nothing at all (`,0` → `,1234`) — is literal text (`0"x",` → `12345678x,`, LibreOffice-verified; Excel unverified for the leading comma). A literal comma after `@` is verified only for the fourth (text) section of a four-section code: cell display of a text cell does not apply its format at all, so under a single-section `@,` code the cell shows `abc` where LibreOffice shows `abc,` (pre-existing); a scaling comma in a scientific or fraction section is inert. A section mixing `General` with digit placeholders (`General0`) renders the General value once; Excel leaves such codes undefined.

**No workarounds needed** - formula system is complete and production-ready!

---

#### 7. Theme Colors ✅ **RESOLVED**
**Status**: Implemented — theme color resolution works via `ThemePalette`
**Phase**: P5 (Complete)

**Current State**:
```scala
val theme = ThemePalette.default // or a parsed palette
Color.Theme(ThemeSlot.Accent1, tint = 0.5).toResolvedArgb(theme) // resolves slot → base RGB → tint → ARGB
Color.Theme(ThemeSlot.Accent1, tint = 0.5).toResolvedHex(theme)  // "#RRGGBB"
```

`ThemePalette.resolve(palette, slot, tint)` performs the slot lookup and applies the tint transformation per RGB component. Note: bare `Color.toArgb` / `toHex` (no palette) still throw on a `Theme` color by design — callers must resolve theme colors through `toResolvedArgb(theme)` / `toResolvedHex(theme)`.

---

#### 8. Shared Strings Table (SST) in Streaming Write ✅ NOW SUPPORTED (GH-223)
**Status**: Implemented — streaming writes deduplicate strings through an SST by default
**Impact**: Plain-text cells are emitted as `t="s"` index references; `xl/sharedStrings.xml` is
written after the worksheets with correct `count` (references) / `uniqueCount` (distinct) values.

**How it works** (two-pass design from `design/smart-streaming.md`):
- Pass 1: while rows stream, an `SstAccumulator` assigns SST indices in first-occurrence order —
  an index is final the moment it is assigned, so the worksheet body needs no second pass
- Pass 2: after the worksheet entries, the accumulated table is emitted as `sharedStrings.xml`

Applies to `writeStream`, `writeStreamWithAutoDetect`, `writeStreamsSeq` (one workbook-global SST
shared across sheets), `writeStreamsSeqWithAutoDetect`, `writeStreamStyled` and
`writeStreamStyledWithAutoDetect`.

**Styles** (GH-223 phase 2, GH-675): `ExcelIO.writeStreamStyled(path, sheet, styles)` and its
two-pass sibling `ExcelIO.writeStreamStyledWithAutoDetect(path, sheet, styles)` (which also computes
the `<dimension>`) take a `Vector[CellStyle]` table; `StyledRowData.cellStyles` values index into
it. The table is deduplicated into `xl/styles.xml` up front (cellXf 0 = default, then first
occurrence) and cell `s=` attributes are remapped to the emitted indices; an index that is
negative, past the table, or names a default-equal style emits no `s=` (the cell takes the
default). Every declared style is emitted whether or not a row uses it (valid OOXML — a streamed
CSV with no dates carries one unused Date xf). The component tables come from the in-memory
writer's builder, so custom formats are declared at 164+ and a stale source `numFmtId` is
re-pointed at its declaration (GH-471 parity). Formatted 100k+ row files no longer require the
in-memory path. The unstyled writers ignore `RowData.cellStyles` (source-workbook xf indices, as
readers produce them). CSV import's O(1) branch (`ImportCommands`, taken only for a workbook with
no sheets) writes through `writeStreamStyledWithAutoDetect`, so a detected ISO date column displays
as dates. The `xl` verb does not reach that branch today: `import` requires `-f` and a loaded book
always has a sheet, so `import --stream` takes the in-memory path (`STREAM_BACKEND_ONLY`), whose
date formatting #667 already fixed.

**Remaining envelope / limitations**:
- Memory is O(distinct strings) for the accumulator — the accepted envelope per the design doc
  (100k rows with 100 distinct strings keeps exactly 100 entries; 1M unique strings ≈ tens of MB)
- `SstPolicy.Never` (`WriterConfig`) keeps the previous inline-string dialect; `Auto`/`Always`
  both deduplicate (streaming cannot pre-scan, so `Auto` opts into SST — note this changes the
  default streaming output dialect from inline strings to SST as of GH-223)
- RichText cells stay `inlineStr` even in SST mode (mixed dialects are valid OOXML)
- Merged-cell emission from streaming writers (phase 3 of GH-223) is still future work
- **A preserved table is never pruned on write (#567)**: the unified writer appends new strings
  to the source's `xl/sharedStrings.xml` and re-points the edited cells — entries are never removed
  or reordered, so untouched sheets keep their `t="s"` indices and their byte-preservation. A text
  replacement, a sheet removal, a `--stream put` (which writes `inlineStr` and copies the table
  verbatim), or an edit of one sheet in a multi-sheet inline-string book (the fresh table is built
  from every sheet while the untouched sheets stay `inlineStr`) therefore leaves `<si>` entries no
  cell references: a counterparty name scrubbed from every cell still rides in the package.
  `xl lint` reports the class as `shared-string-orphan` (count + first five indices, never the
  text) at severity `hygiene` — listed, exit 0, a `LINT_HYGIENE` warning; `lint --strict` makes it
  exit 1; to drop the entries, write the workbook fresh — `XlsxWriter.write(Workbook(wb.sheets),
  out)` has no source and rebuilds the table from the cells (every sheet is regenerated, so
  untouched sheets lose byte-preservation). An opt-in on-write compaction
  (`WriterConfig.compactSharedStrings`) is a follow-up; there is no CLI compaction yet

---

### 🟢 Low Impact (Nice to Have)

#### 9. Hyperlinks ✅ NOW SUPPORTED (0.10.0)
**Status**: Implemented
**Impact**: `Cell.hyperlink` is serialized to `<hyperlinks>` + relationships and populated on read; a `hyperlink` batch op is available. Previously the model existed but write was a silent no-op.

---

#### 10. Conditional Formatting ✅ TYPED RULES SUPPORTED (#136) — with scoped limitations
**Status**: Six rule families are fully modeled: `cellIs` (all eight operators incl. between/notBetween), `expression`, 2/3-point `colorScale`, `dataBar`, `top10`, and the four text rules (contains/notContains/beginsWith/endsWith). Read into `Sheet.conditionalFormats` (typed `ConditionalFormat.Rules` envelopes with `CfRule` rules), authored via `Sheet.conditionalFormat(range, CfRule.cellIs(...), ...)` with differential formats (`Dxf`: font deltas, solid fills, borders, numFmts), and round-tripped on both writer backends. Priorities are assigned at append (`max(existing)+1, +2, ...` — above Preserved rules' priorities too); explicit priorities pass through unvalidated. The read path is typed-parse-or-Preserved at TWO granularities: an unmodeled **rule** (iconSet, timePeriod, aboveAverage, duplicate/uniqueValues, containsBlanks/Errors, any rule with a child `extLst` or out-of-whitelist attr/dxf) rides `CfRule.Preserved` verbatim while its **envelope stays typed** (sqref still shifts under structural edits); an unparseable **envelope** rides `ConditionalFormat.Preserved` whole. The write path is a reparse dirty-gate: an untouched cf model re-emits the source elements verbatim (untouched sheets stay byte-identical); a changed model regenerates with canonical emission and **append-only** dxf-table merging (existing dxf indices never move, so dxfIds baked into Preserved rules and tables stay valid). Structural edits shift typed envelopes with the merged-ranges clamp/split/drop algebra and rewrite typed rule formulas (`#REF!` on full deletion, rule kept); text-rule formulas are derived at emission so they auto-correct.

**Scope fence (v1 OUT — all ride `Preserved` byte-faithfully)**:
- No typed iconSet, timePeriod, aboveAverage, duplicate/uniqueValues, containsBlanks/Errors, autoMin/autoMax data bars, or any x14 extension content (gradient/negative/axis data bars, custom icon sets); worksheet-level `extLst` untouched.
- No dxf alignment/protection/gradient fills/font name/size; double-underline degrades to Preserved. `NumFmt.Currency` in a dxf emits but reads back as `Custom` (no distinct format-code retraction).
- **Rendering only, no recalculation role** (#497): the HTML/SVG renders and every raster built from SVG (`view --format html|svg|png|jpeg|webp|pdf`, `sheet.toSvg(range, sheet.conditionalFormatOverlay(range))`) paint cellIs, expression, text, top10, colour-scale and 2007-style data-bar rules, plus the blanks/errors/time-period rules through a render-only lift; icon sets, above/below average, duplicate/unique values and x14 data bars are not painted and are named in `CF_NOT_RENDERED`. Rules do not participate in `SheetEvaluator`/`DependencyGraph`, and the table outputs (`view --format markdown|csv|json`, `cell`) show base styles only.

**Behavioral limitations (by design)**:
- **Preserved staleness under structural edits**: `CfRule.Preserved` payload formulas and `ConditionalFormat.Preserved` sqref do NOT shift (the `Drawing.Preserved` precedent); typed envelopes around Preserved rules DO shift.
- No priority renumbering or overlap validation of explicit priorities; authoring never merges into existing blocks (a new block is appended per call, as Excel itself does).
- No streaming-path cf (`readStream` surfaces none; `writeStream` emits none); fresh-sheet SaxStax writes with cf route through the DOM-equivalent worksheet builder (DirectSaxEmitter cf deferred).
- No CLI/batch op in this wave (follow-up issue covers `xl cond-format` + batch op + skill doc).
- **Source-compat**: `Sheet` gained a 15th field (`conditionalFormats`) — source-compatible, not binary-compatible with 0.12.x.

---

#### 11. Data Validation Authoring ✅ LIST DROPDOWNS SUPPORTED (0.13.0, #375)
**Status**: List dropdowns author via `Sheet.withDataValidation(range, DataValidation.list("\"Yes,No\""))` (or `DataValidation.listOf("Yes", "No")`, or a range-ref formula), with `allowBlank`/`showDropdown` (OOXML's inverted `showDropDown` handled), multi-range sqref, structural-edit range shifting, and read-side parse into the model (read→edit→write keeps validations on rewritten sheets). Existing `dataValidations` are still **preserved through edits** since 0.10.0; unmodeled kinds (whole/decimal/date/custom, operators, prompt/error messages) survive byte-faithfully as `DataValidation.Preserved`.
**Impact**: List dropdowns are authorable programmatically; other validation kinds are preserved but not yet authorable.
**Remaining**: typed whole/decimal/date/textLength/custom authoring (list-only in v1); no CLI/batch op yet.

---

#### 12. Charts ✅ TYPED BAR/LINE/PIE SUPPORTED (#222) — with scoped limitations
**Status**: Bar (clustered/stacked/percent-stacked, column or horizontal), line, and pie charts are fully modeled: read into `Drawing.ChartFrame(anchor, Chart, name)` (`Sheet.charts`), authored (`Chart.validated`/`Chart.bar`/`line`/`pie` + `Sheet.addChart`, CLI `chart add`), and round-tripped. The read path is a typed-parse-or-Preserved hybrid: a chart parses to the typed model only when its entire `chartN.xml` fits the strict whitelist; ANY out-of-fence construct keeps the whole anchor as `Drawing.Preserved` and the part rides byte-preservation — never half-typed. Structural edits (`insertRows`/`deleteColumns`/... and the evaluator's `StructuralEditor`) shift typed chart data references, including cross-sheet; `Workbook.rename` remaps typed chart references.

**Scope fence (v1 OUT — all ride `Preserved` byte-faithfully)**:
- Scatter/area/combo/doughnut/radar/bubble/stock/3D charts; secondary axes; pivot charts; sparklines; chartex; chartsheets.
- A user-facing Axis model (titles, bounds, numFmt, date axes): axes are write-time constants derived from the chart type.
- Data labels, per-point `dPt`, trendlines, error bars, series/chart colors and `spPr` styling, real line markers, smooth lines, multi-level categories, external-workbook refs.
- **Excel-authored charts stay Preserved** (they carry `c:style`, colors/style part rels, etc.) — byte-faithful round-trip exactly as before; to change one, delete it and re-author.

**Behavioral limitations (by design)**:
- **Stale caches on clean charts**: an untouched chart's part is byte-preserved even when the cells it references changed — its embedded value cache goes stale (same as Excel-without-recalc; Excel recalculates on open). Any edit to the sheet's drawings regenerates the part with fresh caches from STORED cell values (no evaluator).
- **Full deletion drops series, not `#REF!`**: deleting all rows/columns under a series' values removes the series (categories fall back to 1..N, a deleted name cell clears the name). Excel would keep a broken `#REF!` stub; the typed model has no broken-ref state.
- **Orphaned chart parts on degenerate reorder+edit**: moving a chart to a different anchor index AND editing it in one write allocates a fresh `chartN.xml`; the old part stays on disk (never deleted — the media GC policy; orphaned parts are legal OOXML).
- **Rename and references** (#559, 0.20.0): `rename-sheet` (verb and batch op) and `SheetRenamer.rename` rewrite the sheet qualifier in cell formulas on every sheet, defined names, and typed conditional-format and data-validation formulas, preserving cached values; `Workbook.rename` alone stays tab-only (it remaps typed `ChartFrame` references and nothing else). Still not rewritten: `Preserved` chart, CF and DV payloads (opaque XML re-emitted verbatim) and hyperlink locations. A dependent text that mentions the sheet but cannot be parsed refuses the rename before anything is written — today that is 3-D ranges (`SUM(Sheet1:Sheet3!A1)`, refused for either end) and unknown functions (a name the registry does not know, with or without Excel's `_xlfn.` prefix). Known functions stored with the prefix (`_xlfn.XLOOKUP(Sheet1!A1,…)`, the form every modern file uses for post-2007 functions) rename cleanly since #576: the reader strips `_xlfn.`/`_xlws.` and the parser drops them before the registry lookup. Safe, but a file with a 3-D range or an unknown function needs that dependent rewritten by hand before the rename.
- **`_xlfn.` storage prefixes** (#556, #577, 0.20.x): the model holds the formula-bar spelling (`IFS(…)`) and the file holds Excel's storage form (`_xlfn.IFS(…)`, `_xlfn._xlws.FILTER(…)`, `_xlpm.` on LET/LAMBDA parameters) at every formula-text boundary xl writes — cell `<f>`, conditional-formatting `<formula>` and `<cfvo type="formula">`, data-validation `<formula1>`/`<formula2>`, and `<definedName>` — mapped by `FormulaStorage` on read and write, so an Excel-authored book re-serializes byte-identically. Two caveats. A file from a bare-writing producer (openpyxl) reads correctly, and any in-memory write that regenerates the worksheet heals its cells, CF blocks and DV container while every non-clean write heals `<definedName>` bodies (`workbook.xml` is always regenerated): the writer's clean-compare gates treat bare text as dirty through the same `bareFutureCalls` rule (#593), so a `put` on the sheet is enough. Not healed: text inside an unmodeled rule or validation (a `CfRule.Preserved` / `DataValidation.Preserved` payload is re-emitted verbatim as captured), x14 `<xm:f>` bodies, an untouched worksheet (copied verbatim), and every slot a `--stream` write does not patch. `xl lint` reports the bare text as `xlfn-missing` either way — the finding is short by design; this paragraph and the [`xl lint` reference](reference/cli.md#xl-lint) carry the slot-by-slot semantics. Its rule is the writer's own (`FormulaStorage.bareFutureCalls` runs the storage scanner with recording callbacks), so it also catches a half-prefixed `_xlfn.LET(x,1,x+1)` whose parameters lack `_xlpm.` — the case Excel reports as unreadable content rather than `#NAME?`. And the lint inspects those four sites plus x14's `<xm:f>` only — table-part `<calculatedColumnFormula>`/`<totalsRowFormula>` and chart references are not scanned.
- Reading a stacked bar chart requires `overlap="100"` (Excel/openpyxl write it; without it stacked bars visibly misrender, so such charts stay Preserved). Reading an empty-series pie chart (possible after structural deletion + rewrite) falls back to Preserved on the next read.
- CLI `--series-names` are literals only (cell-sourced names are library-API-only); batch JSON chart ops and streaming-write charts are out of v1. HTML/SVG/PNG rendering does not draw charts.

---

#### 13. Drawings & Images ✅ PICTURES SUPPORTED (#221) — with scoped limitations
**Status**: Embedded pictures are fully modeled: read (`Sheet.drawings`/`Sheet.pictures`), authored (`Sheet.addImage` — png/jpeg/gif/bmp natural-size sniffing; tiff/emf/wmf with an explicit `Extent`), and round-tripped with all three anchor forms (`DrawingAnchor.OneCell`/`TwoCell`/`Absolute`, `editAs`). Non-picture drawing anchors (charts, shapes, group shapes, connectors, pictures with crops/rotation/effects/hyperlinks) are captured as `Drawing.Preserved` and re-emitted verbatim in z-order; untouched drawing parts are byte-preserved whole. Media parts are content-addressed (sha-256): identical bytes never produce two media parts in one write.

**Remaining limitations (by design, GH-221 scope fence)**:
- **Streaming envelope**: drawings are an **in-memory** feature. `ExcelIO.readStream` surfaces no drawings (rows only); the streaming row writer (`writeStream`) emits none. Use `read`/`write`.
- **Fresh emission of rel-referencing fragments**: writing a workbook **without** a `SourceContext` (programmatic build, or `Drawing.Preserved` values copied onto a new sheet) drops Preserved fragments that contain relationship references (`r:embed`/`r:id`/`r:link`) — their targets only exist in the source package. Rel-free fragments (plain shapes) are emitted.
- **Orphan media is never garbage-collected**: removing pictures rewrites the drawing part but keeps all source media parts (orphans are legal OOXML).
- **Preserved anchors are not shifted** by `insertRows`/`deleteColumns`/... (typed Picture anchors are; a deleted anchor index clamps instead of dropping — Excel keeps pictures). EditAs-aware size recomputation is not attempted.
- **Accepted-and-dropped on dirty regeneration**: `a:blip/@cstate`, plain `a:xfrm`, `cNvPicPr` content (e.g. `a:picLocks`). The loss only materializes when a sheet's drawings vector actually changes; untouched parts are byte-identical.
- ~~Sheet delete/reorder + drawing edits in one write~~ **lifted (#315)**: source mappings (drawings, comments, worksheet parts, SST accounting) are keyed by sheet NAME as read — stable under delete/reorder, re-keyed by `Workbook.rename` — so structural edits combine freely with drawing/comment edits in a single write. Fresh comment parts allocate numbers above everything the source claims (never colliding with a surviving sheet's identity-mapped part). API note: `SourceContext` changed shape for this — per-sheet mappings are `Map[SheetName, _]`, a new `sheetPathMapping` field, and `markSheetDeleted(Int, SheetName)` — source- and binary-incompatible with 0.12.x for code constructing or pattern-matching `SourceContext` directly.
- **Source-/binary-compat**: `Sheet` gained a 14th field (`drawings`) — source-compatible, not binary-compatible with 0.11.x.
- Deferred, tracked: typed Chart/Shape cases (#222), SVG (svgBlip), crop/rotation/effects, picture hyperlinks, pHYs DPI sizing, media GC, `Patch` case, CLI `add-image`, DirectSaxEmitter drawing emission, string-ref `addImage` overload.

---

#### 14. Pivot Tables Not Supported
**Status**: Not implemented
**Impact**: Cannot create Pivot Tables
**Plan**: see [plan/roadmap.md](plan/roadmap.md)

**Note**: Excel Tables (structured data ranges with headers, AutoFilter, styling) are ✅ **fully supported** as of WI-10.

**Effort**: 10-15 days (for Pivot Tables only)
**LOC**: ~700

---

#### 15. Named Ranges ✅ NOW SUPPORTED (0.10.0)
**Status**: Implemented
**Impact**: `WorkbookMetadata.definedNames` is serialized to `<definedNames>` (previously read-only), with a CLI `name add` / `name rm` verb. Since 0.13.0 (#384), defined names also **resolve in formula evaluation** — `=IF(case=2,…)`, `=entry_mult*ltm_ebitda`, `=SUM(rev_range)` evaluate against workbook- and sheet-scoped names (sheet-scoped shadows global) via a dedicated `TExpr.NameRef` node, contribute dependency edges so `recalculate()` orders name-gated families correctly, and round-trip byte-faithfully; unresolvable names are clean per-cell errors. Structured references (`Table[@Column]`) inside formulas remain future work.

---

#### 16. Print Settings and Page Setup ⚠️ PARTIALLY SUPPORTED (0.11.0)
**Status**: `PageSetup` authoring shipped in 0.11.0 (#259): odd header/footer (with `&P`/`&N` codes), page margins, print area, and repeat rows — emitted in schema order and round-tripped; print area/titles ride the defined-names pipeline as sheet-scoped `_xlnm` names. Reading now populates `Sheet.pageSetup` for any sheet with `<pageMargins>`.
**Remaining**: even/first-page headers and the `fitToPage` flag — tracked in #266.

---

#### 17. Document Properties Not Written
**Status**: Not implemented
**Impact**: Missing metadata (author, title, creation date)
**Plan**: see [plan/roadmap.md](plan/roadmap.md)

**Current**:
- No `docProps/core.xml` (core properties)
- No `docProps/app.xml` (application properties)
- Excel auto-generates defaults on open

**Effort**: 1 day
**LOC**: ~80

---

### 🟣 Security & Safety (WI-30 - Production Ready)

#### 18. ZIP Bomb Protection ✅
**Status**: Implemented (WI-30)
**Impact**: Malicious XLSX files are detected and rejected

**Configuration** (`XlsxReader.ReaderConfig`):
```scala
val config = XlsxReader.ReaderConfig(
  maxCompressionRatio = 100,        // Max 100:1 ratio (default)
  maxUncompressedSize = 100_000_000L, // 100 MB max (default)
  maxEntryCount = 10_000,           // 10k files max (default)
  maxCellCount = 10_000_000L,       // 10M cells max (default)
  maxStringLength = 32_768          // 32 KB per string (default)
)

// Use permissive config for trusted files
XlsxReader.read(path, XlsxReader.ReaderConfig.permissive)
```

**Protection**:
- Compression ratio validation (detects highly compressed ZIP bombs)
- Uncompressed size tracking (prevents memory exhaustion)
- Entry count limits (prevents archive bomb variants)
- Fails early with `XLError.SecurityError` when limits exceeded

**Tests**: 10+ security tests in `ZipBombSpec.scala`

---

#### 19. XXE (XML External Entity) Protection ✅
**Status**: Implemented
**Impact**: External entity resolution is completely disabled

**Protection** (in `XmlSecurity.scala`):
```scala
// All XML parsing uses secure parser factory
factory.setFeature(XMLConstants.FEATURE_SECURE_PROCESSING, true)
factory.setFeature("http://xml.org/sax/features/external-general-entities", false)
factory.setFeature("http://xml.org/sax/features/external-parameter-entities", false)
factory.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true)
```

**Tests**: XXE protection tested in `SecuritySpec.scala`

---

#### 20. Formula Injection Guards ✅
**Status**: Implemented (WI-30, TJC-339)
**Impact**: Untrusted data can be safely written to Excel (both in-memory and streaming writes)

**API** (`CellValue.escape()` and `WriterConfig.secure`):
```scala
// Manual escaping for individual strings
CellValue.escape("=SUM(A1)")  // Returns: "'=SUM(A1)"
CellValue.escape("+1234")     // Returns: "'+1234"
CellValue.escape("-danger")   // Returns: "'-danger"
CellValue.escape("@import")   // Returns: "'@import"
CellValue.escape("Normal")    // Returns: "Normal" (unchanged)

// Automatic escaping via WriterConfig.secure (in-memory writes)
XlsxWriter.writeWith(workbook, path, WriterConfig.secure)
// All text cells starting with =, +, -, @ are automatically escaped

// Streaming writes also support formula injection escaping (TJC-339)
excel.writeStream(path, "Sheet1", config = WriterConfig.secure)(rows)
excel.writeStreamsSeq(path, sheets, config = WriterConfig.secure)
```

**Escaping Rules**:
- Text starting with `=` → `'=...` (prevents formula execution)
- Text starting with `+` → `'+...` (prevents formula prefix)
- Text starting with `-` → `'-...` (prevents formula prefix)
- Text starting with `@` → `'@...` (prevents DDE commands)
- Already-escaped text unchanged (idempotent)
- Formula cells (`CellValue.Formula`) are NOT escaped (they're real formulas)

**Unescape for Reading**:
```scala
CellValue.unescape("'=SUM(A1)")  // Returns: "=SUM(A1)"
```

**Tests**: 22 security tests in `FormulaInjectionSpec.scala`

---

#### 21. File Size Limits (Partial)
**Status**: Configurable limits in `ReaderConfig`
**Impact**: Resource exhaustion prevented via configurable limits

**Implemented**:
- `maxCellCount`: Maximum cells to process (default: 10M)
- `maxStringLength`: Maximum string length (default: 32KB)
- `maxUncompressedSize`: Maximum total uncompressed data (default: 100MB)

**Usage**:
```scala
val strictConfig = XlsxReader.ReaderConfig(
  maxCellCount = 1_000_000L,    // 1M cells max
  maxStringLength = 10_000,     // 10KB per string
  maxUncompressedSize = 50_000_000L  // 50MB max
)
XlsxReader.read(path, strictConfig)
```

---

### 🔵 Advanced Features (P6-P10)

#### 22. No Type-Class Derivation for Codecs
**Status**: Not implemented
**Impact**: Manual row/cell conversion required
**Plan**: see [plan/roadmap.md](plan/roadmap.md)

**What's Missing**:
```scala
// WANT: Automatic codec derivation
case class Person(name: String, age: Int, salary: BigDecimal)

given Codec[Person] = Codec.derived

// Read: Stream[F, RowData] → Stream[F, Person]
excel.readStream(path).through(Codec[Person].decode)

// Write: Stream[F, Person] → XLSX
people.through(Codec[Person].encode)
  .through(excel.writeStream(path, "People"))
```

**Effort**: 7-10 days
**LOC**: ~500 (type-class derivation, header binding, testing)

---

#### 23. No Path Macro for Named Cell References
**Status**: Not implemented (named ranges themselves are supported since 0.10.0 — see #15)
**Impact**: Cannot use compile-time symbolic names for cells
**Plan**: see [plan/roadmap.md](plan/roadmap.md)

**Want**:
```scala
object Paths:
  val totalRevenue = path"Summary!B10"
  val salesTable = path"Sales!A1:D100"

sheet.put(Paths.totalRevenue, formula"=SUM(${Paths.salesTable})")
```

**Effort**: 3-4 days
**LOC**: ~200

---

#### 24. No Style Literal for Inline Styling
**Status**: Not implemented (the builder DSL — `.bold`, plus 0.11.0's `.borderTop(...)`, `.indent(n)`, `range.outlined(...)` — covers most cases)
**Impact**: No single-string style literal
**Plan**: see [plan/roadmap.md](plan/roadmap.md)

**Want**:
```scala
val headerStyle = style"font-weight: bold; background: #CCCCCC; border: all thin"
```

**Effort**: 2-3 days
**LOC**: ~150

---

#### 25. XLSM Macros Not Preserved
**Status**: Not implemented
**Impact**: Opening XLSM files strips macros
**Plan**: see [plan/roadmap.md](plan/roadmap.md)

**Current**: Reading `.xlsm` treats it as `.xlsx` (ignores `vbaProject.bin`)

**Mitigation**: Should NEVER execute macros (security risk), but should preserve for round-tripping

**Effort**: 2-3 days
**LOC**: ~100

---

#### 26a. `Option[T]` Record Fields See Empty-String Cells as Present (#617)
**Status**: By design; documented
**Impact**: `derives RowCodec` decodes an `Option[String]` field as `None` only for an absent or `CellValue.Empty` cell. A cell holding the empty string — SheetJS and some exporters write `<v></v>` text cells for "blank" — is `Some("")` (and a `TypeMismatch` for `Option[Int]`), so a "sparse" column written that way is never `None`. Excel itself distinguishes `""` from blank (`ISBLANK` is FALSE, `COUNTA` counts it), so the codec does too.
**Workaround**: normalise with `.filter(_.nonEmpty)` after the read, or clear such cells (`clear --all` on the range) before reading.

#### 26b. Childless `inlineStr` Cells Read as Blank (#460)

**Impact**: openpyxl serializes `value=""` as a childless `<c t="inlineStr"/>` — out of schema
(CT_Cell wants an `<is>` with content or a `<v>`). Excel, openpyxl and xl's streaming readers treat
the cell as blank; the in-memory reader used to fail the whole workbook read with `inlineStr cell
missing <is> element and <v> element` while `xl lint` passed the same file. Since #460 the
in-memory reader reads such a cell (and a childless `<c t="str"/>`) as `CellValue.Empty`, keeping
its style index, and `xl lint` reports the shape as `empty-inline-str`. An `<is/>` that is present
but has no `<t>`/`<r>` is different: it is an inline string that is empty — `CellValue.Text("")`,
like `<is><t></t></is>` and exactly what the streaming reader returns (Excel distinguishes `""`
from blank, see #617 above) — so it is neither a read failure nor a finding. A write that
regenerates the sheet emits the blank without a `t` attribute (the writer derives `t` from the
value), healing the shape; an untouched sheet is copied verbatim.

#### 26. Named Cell Styles: Preserved, Not Yet Modeled (#610)
**Status**: Preservation fixed in 0.22.0 ([#610](https://github.com/TJC-LP/xl/issues/610)); no typed named-style model yet
**Impact**: A workbook's Cell Styles gallery, recent-colours palette and styles `extLst` survive every write; named styles cannot be authored from xl

**Fixed in 0.22.0**: a source workbook's `cellStyleXfs`, `cellStyles`, `tableStyles`, `colors` (`mruColors`/`indexedColors`) and styles-level `extLst` ride through every write of the in-memory writer verbatim — byte-identical on the default DOM backend, attribute-sorted on the StAX backend — the same opaque-passthrough contract as `dxfs`. Every `cellXf` keeps the `xfId` of the named style it derives from: the reader registers a source's cellXfs positionally in each sheet's `StyleRegistry`, so two xfs that differ only in `xfId` ("Comma 2" applied vs the same formatting typed by hand) stay two xfs through a regenerating write. A style xl authors is appended with `xfId="0"` (`Normal` + direct formatting) unless an equal direct xf already exists, which it shares; when only a named-style twin exists, the cell shares that twin. Source `cellXfs` records are emitted verbatim (`applyX` flags, `quotePrefix`, `<protection>` and attribute order intact); only the xfs xl adds are regenerated. An empty `<cellStyleXfs count="0"/>` or an out-of-range `xfId` in the source is repaired to `Normal`. An xl-authored workbook, or one written after its `sourceContext` is dropped, still emits exactly one `Normal`. `--stream put`/`putf`/`style` patch the source `styles.xml` in place and never dropped these sections.

**Remaining**:
- No typed `NamedStyle` model: a named style cannot be created, renamed, modified or applied from the API or CLI, and a cell xl restyles becomes direct formatting on `Normal` even when its source xf derived from a named style (the API cannot tell — `CellStyle` carries no `xfId`).
- cellXf-level flags are not modeled: the xfs xl itself adds carry only `applyAlignment` (no `applyFont`/`applyFill`/`applyBorder`/`applyNumberFormat`/`applyProtection`, `quotePrefix` or `<protection>`); source xfs keep theirs verbatim.
- xl-added `cellXfs` derive `fontId`/`fillId`/`borderId` by value (first equal table entry); `x14ac:knownFonts` on `<fonts>` is dropped; fonts lose attributes the `Font` model does not carry (`family`, `scheme`, `charset`).
- The streaming writer of an in-memory workbook (`ExcelIO.writeStream`) is a fresh write and emits one `Normal`.

**Mitigation**: author and modify named styles in Excel; xl preserves them.

---

## Architecture Limitations

### 27. No Streaming for Multi-Sheet Write (Requires Sequential)
**Status**: By design
**Impact**: Must write sheets in order, cannot write in parallel
**Plan**: N/A (fundamental to ZIP format)

**Why**: ZIP files are sequential - cannot write `sheet2.xml` while `sheet1.xml` is being streamed

**Current API**:
```scala
excel.writeStreamsSeq(
  path,
  Seq("Sheet1" -> rows1, "Sheet2" -> rows2)  // Sequential consumption
)
```

**Alternative Considered**: Materialize all but last sheet (breaks streaming guarantee)

**Verdict**: ACCEPTED LIMITATION (ZIP format constraint)

---

### 28. SST Materialized During Read
**Status**: By design (acceptable tradeoff)
**Impact**: SST size contributes to memory overhead
**Plan**: N/A (optimal for 95% of use cases)

**Why**:
- SST is typically small (<10MB even for 1M row files)
- Contains deduplicated strings (not all cell values)
- Needed for O(1) cell value resolution

**Memory Profile**:
- SST: ~1-10MB (worst case: 1M unique strings = ~50MB)
- Parser buffer: ~40MB
- Total: ~50-100MB constant (still excellent for large files)

**Alternative Considered**: LRU cache with on-demand loading (complex, marginal benefit)

**Verdict**: ACCEPTED TRADEOFF (pragmatic design)

---

### 29. No Concurrent Sheet Processing
**Status**: By design
**Impact**: Cannot read multiple sheets in parallel
**Plan**: N/A (would require materializing metadata)

**Current**:
```scala
// Sequential access only
excel.readStreamByIndex(path, 1)  // Sheet 1
excel.readStreamByIndex(path, 2)  // Sheet 2 (separate call)
```

**Why**: Maintaining streaming guarantee requires sequential ZIP access

**Workaround**: For parallel processing, use `excel.read()` to materialize full workbook

**Verdict**: ACCEPTED LIMITATION (streaming guarantee more valuable)

---

### 30. Unbounded `--stream view` Streams (#635, resolved in 0.22.0)
**Status**: Resolved — `--stream view --limit 0` writes its rows as the reader produces them, for csv, json (bare and inside the `--json` envelope) and markdown, in constant memory; the bytes are those of the gathered table (a property law holds every generated window in both output modes). On the 1,000,000 × 41 dogfood book the csv dump used to run 57 minutes and 13.8 GB and print nothing.

**What remains, by design**:
- **Two reads for markdown, and for csv under `--skip-empty`**: the column widths and the empty columns are facts about every row, so the window is folded once (O(columns) memory) and streamed a second time. Under `--stream` that is two passes over the worksheet — the shared-string table and the styles are parsed once and shared. Prefer csv or json for a whole-sheet dump.
- **`--json` streams too**: a JSON document is spliced into `data` element by element, and a csv or markdown table's `data.text` is escaped line by line as `ujson` escapes one string — nothing is gathered, whatever the window.
- **Nothing is written before the first row has been read**, so a source that fails before it has one (the shared-string table over `--max-size`, a worksheet the parser rejects at once) fails as every other run does — the diagnostics on stderr, the `ok: false` envelope under `--json`. **A failure after the first byte** (a worksheet rejected mid-sheet, past the 1,024-row chunk the reader delivers first) stops the output where it is and is reported on stderr with exit 3; under `--json` the envelope on stdout is left unterminated — the exit code and stderr carry the failure, and a consumer that parses stdout sees the truncation as a parse failure, never as a result.
- **A reader that closes stdout** (`xl … | head`) ends the run quietly with exit 0, as a tool killed by SIGPIPE would; a stdout that fails for another reason (no space left, an I/O error on the file it is redirected to) is `IO_WRITE`, exit 3 — the streamed bytes go through a stream of xl's own on the descriptor, so the two are told apart (a `PrintStream` would have left one sticky flag).
- `--stream cell` shares the table with `view` (#640), and every verb's `--stream` answer is published as `stream` in `xl schema --json` (#638).

---

## Roadmap to 100% Feature Parity

### Phase 4 Continuation ✅ (Complete)
**Focus**: Complete OOXML coverage for common use cases

- [x] Style application end-to-end ✅
- [x] DateTime serialization ✅
- [x] Merged cells XML ✅
- [x] Column/row properties XML ✅

**Result**: Fully functional spreadsheet library with formatting

---

### Phase 6: Codecs & Named Tuples (Priority 2)
**Effort**: 2-3 weeks
**Focus**: Ergonomic data binding

See: [plan/roadmap.md](plan/roadmap.md)

- [ ] Type-class derivation (7-10 days)
- [ ] Header row binding (3-4 days)
- [ ] Tuple/case class codecs (5-6 days)

**Result**: `Stream[F, Person]` ↔ XLSX with zero boilerplate

---

### Phase 7: Advanced Macros (Priority 3)
**Effort**: 1-2 weeks
**Focus**: Enhanced compile-time DSL

See: [plan/roadmap.md](plan/roadmap.md)

- [ ] Path macro for named references (3-4 days)
- [ ] Style literal (2-3 days)
- [ ] Formula macro with type checking (5-7 days)

**Result**: Best-in-class developer experience

---

### Phase 8: Drawings ✅ SHIPPED (0.12.0, #221)
**Focus**: Images and shapes

Embedded pictures are modeled, authored, and round-tripped (`Sheet.addImage`/`pictures`/`removeDrawing`, three anchor forms, 7-format classification, sha-deduped media); non-picture drawings (shapes, etc.) ride byte-preservation. See §13 below and [plan/roadmap.md](plan/roadmap.md) (#221). Shape *authoring* remains future.

---

### Phase 9: Charts ✅ TYPED BAR/LINE/PIE SHIPPED (0.12.0, #222)
**Focus**: Chart generation

Typed bar/line/pie charts are modeled, authored (`Chart.bar`/`line`/`pie` + `Sheet.addChart`, CLI `chart add`), and round-tripped; Excel-authored and out-of-fence charts stay byte-preserved. See §12 below and [plan/roadmap.md](plan/roadmap.md) (#222). Scatter/area/combo/3D and chart styling remain future.

---

### Phase 10: Tables & Advanced Features (Priority 6)
**Focus**: Tables ✅ (shipped), conditional formatting ✅ (0.12.1, #136 — see §10), pivot tables ⏳ (still future)

See: [plan/roadmap.md](plan/roadmap.md)

---

### Phase 11: Security & Safety ✅ (WI-30 Complete)
**Status**: Production ready (2025-12-07)
**Focus**: Security hardening

See: [plan/roadmap.md](plan/roadmap.md)

**Completed** (WI-30):
- ✅ ZIP bomb detection (`ReaderConfig` with compression ratio, size, entry count limits)
- ✅ XXE prevention (secure XML parser factory with disabled external entities)
- ✅ Formula injection guards (`CellValue.escape()`, `WriterConfig.secure`)
- ✅ Resource limits (`maxCellCount`, `maxStringLength`, `maxUncompressedSize`)

**Future work**:
- Path traversal in ZIP (currently validated but could add explicit guards)
- Fuzzing and security audit (recommended before 1.0)

---

## Feature Comparison: XL vs Apache POI

| Feature | XL Today | POI | Notes |
|---------|----------|-----|-------|
| **Core I/O** | ✅ | ✅ | XL: Pure, POI: Imperative |
| **Streaming Write** | ✅ | ✅ | XL: 88k rows/s, POI: ~30k rows/s |
| **Streaming Read** | ✅ | ✅ | XL: 55k rows/s, POI: ~40k rows/s |
| **Multi-sheet** | ✅ | ✅ | XL: Arbitrary, POI: Sequential |
| **Styles** | ✅ | ✅ | XL: Full in-memory; streaming uses minimal default styles |
| **Formulas (eval)** | ✅ | ✅ | XL: 108 functions, dependency graph, cycle detection |
| **Tables** | ✅ | ✅ | XL: Full table support with AutoFilter, structured refs |
| **Charts** | ⚠️ | ✅ | XL: typed bar/line/pie (scoped, §12); POI: full |
| **Drawings** | ⚠️ | ✅ | XL: embedded pictures (scoped, §13); POI: images/shapes |
| **Memory (100k rows)** | ✅ 50MB | ❌ 800MB | XL: 16x better |
| **Type Safety** | ✅ | ❌ | XL: Compile-time, POI: Runtime |
| **Purity** | ✅ | ❌ | XL: Pure, POI: Mutable |
| **Determinism** | ✅ | ❌ | XL: Stable diffs, POI: Non-deterministic |
| **Security** | ✅ | ⚠️ | XL: ZIP bomb, XXE, formula injection protection |

**Verdict**: XL is production-ready for data-heavy use cases (streaming, ETL) and now authors typed charts + embedded pictures (scoped — see §12/§13). POI still leads for advanced chart types, shapes, and pivot tables.

---

## Performance Benchmarks: XL vs Apache POI

*See [design/performance-investigation.md](design/performance-investigation.md) for full analysis.*

> ⚠️ **Benchmarks below were captured on an earlier release and have not been re-validated for 0.12.x — treat as indicative.** (0.12.1 additionally made codec put paths ~2.4x faster, #297.)

### Summary (Apple Silicon, JDK 21)

| Operation | 1k rows | 10k rows | 100k rows |
|-----------|---------|----------|-----------|
| **In-Memory Read** | ✅ XL +21% | ✅ XL +3% | ❌ POI +11% |
| **Streaming Read** | ✅ XL +37% | ✅ XL +2% | ❌ POI +22% |
| **Write (SaxStax)** | ✅ ~tied | ✅ XL +36% | ✅ XL +39% |

**Key Findings**:
- XL wins 5 out of 9 benchmarks, including ALL writes
- XL dominates typical workloads (<10k rows) across all operations
- 100k read gap (11-22%) is the cost of functional abstractions
- Write performance is excellent due to DirectSaxEmitter + SaxStax backend

### Root Cause: 100k Read Gap

**File**: `xl-cats-effect/src/com/tjclp/xl/io/SaxStreamingReader.scala:44`

SAX parsing is inherently synchronous - the `parser.parse()` call blocks until the entire document is parsed, then materializes all rows to a Vector before streaming begins. This adds ~15-20% overhead at scale.

**Proposed Solutions** (deferred to future work):
1. **Threaded SAX Reader**: Run parser on separate fiber with queue-based row transfer
2. **Lazy SharedStrings**: Parse SST entries on-demand during worksheet streaming

**Recommendation**: Accept current performance for v1.0. The typical workload (<10k rows) is faster across the board.

---

## Frequently Asked Questions

### Q: Can I use XL in production today?

**Yes, if your use case is**:
- Large dataset export (100k+ rows)
- ETL pipelines (streaming read/write, constant memory)
- Data generation (reports, analytics)
- Multi-sheet workbooks
- Core cell types and rich text
- Styling in in-memory workflows (full styles supported)
- Formula evaluation (108 functions, dependency graph, cycle detection)
- Excel Tables (structured data with AutoFilter, headers, styling)
- Performance-critical workloads (benchmarked vs POI)

**No/Not yet, if you need**:
- Advanced chart types (scatter/area/combo/3D), shapes, or chart styling — typed bar/line/pie charts + embedded pictures **do** ship (scoped, §12/§13)
- Pivot tables (planned), or data-validation authoring beyond list dropdowns (list dropdowns shipped in 0.13.0, #375)
- Excel macros (not planned to execute; preservation planned)

---

### Q: What's the maximum file size XL can handle?

**Streaming API**: Unlimited
- Tested: 100k rows (completes in ~3s)
- Projected: 1M rows (~30s)
- Memory: O(1) constant (~50-100MB regardless of size)
- CLI: every `--stream` read is O(1), `view --limit 0` included since 0.22.0 — its rows are written as the reader produces them (§30, #635)

**In-Memory API**: ~500k rows before OOM (8GB heap)

---

### Q: How does XL compare to other Scala Excel libraries?

**poi-scala**: Wrapper around POI (inherits memory issues)
**excel4s**: Based on POI (same limitations)
**xlsx-parser**: Read-only, not feature-complete

**XL Advantages**:
- Pure functional (no mutation)
- True streaming (constant memory)
- Type-safe (compile-time validation)
- Law-governed (property-tested)
- Deterministic (stable output)

---

### Q: How should I choose between streaming and in-memory?

- **Need full styles/metadata or random access**: use in-memory read/write (`ExcelIO`).
- **Need constant memory and sequential processing**: use streaming (`readStream`, `writeStream`); styles are minimal and strings are inline.

---

### Q: How complete is formula evaluation?

**Status**: ✅ Production Ready (as of WI-07/08/09)

**What's implemented**:
- 108 built-in functions (SUM, SUMIF, COUNTIF, XLOOKUP, HLOOKUP, INDEX, MATCH, OFFSET, INDIRECT, XIRR, XNPV, dynamic arrays, etc.)
- Full dependency graph with cycle detection
- Safe evaluation with `evaluateWithDependencyCheck()`
- Type coercion and error handling

**What's planned for future**:
- Extended function library (300+ Excel functions)
- Array formulas
- Structured references (Table[@Column])

**For most use cases**, the current implementation is sufficient. Let Excel recalculate for unsupported functions.

---

## Migration Path from Limitations

### If you hit a limitation today:

1. **Need advanced charts/shapes**: XL authors typed bar/line/pie charts + embedded pictures (#222, #221; scoped — see §12/§13) and preserves everything else through edits; reach for POI only for chart types XL doesn't model yet (scatter/area/combo/3D, shapes)
2. **Unsupported formula function**: Store the formula as a string and let Excel recalculate on open (XL evaluates the 108 built-ins)
3. **Streaming update of one sheet in a huge workbook**: Use the in-memory read → modify → write path (surgical modification preserves untouched parts)

---

## Contributing

Want to help implement these features? See:
- [docs/plan/roadmap.md](plan/roadmap.md) - Full implementation plan
- [CLAUDE.md](../CLAUDE.md) - Contribution guidelines
- [GitHub Issues](https://github.com/TJC-LP/xl/issues) - Per-feature tracking

---

## Summary

**XL Today**: Best-in-class for streaming large datasets with type safety and purity

**XL Tomorrow** (P4-P11): Feature parity with POI while maintaining performance advantages

**XL Vision**: The definitive pure functional Excel library for Scala 3

---

*Last updated 2026-06-15 (0.12.2 doc-truth pass: refreshed fixed/open status against the 0.12.0 "Visual" (charts/drawings), 0.12.1 "Clean Sweep" (conditional formatting), and 0.12.2 "Interop" releases; the earlier 0.11.0 pass is tracked under GH-272).*
