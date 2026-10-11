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

**Not supported (v1):** R1C1 mode (`a1=FALSE`) returns an evaluation error and leaves the cell uncached (Excel computes it on open) — deliberately not `#REF!`; consequently `INDIRECT(ADDRESS(..., FALSE), FALSE)` does not compose. An INDIRECT call written directly where a range is required at parse time (SUMIF/VLOOKUP/XLOOKUP/INDEX/MATCH/NPV/IRR array slots) is a parse error — the same envelope as OFFSET; a defined name or a LET name bound to the call is accepted there and resolves to the reference it returns (GH-710: `COUNTIF(dyn,">2")` with `dyn = INDIRECT("Sheet1!A1:A10")`). INDIRECT returns a reference, like OFFSET and INDEX: its values in an array formula, the whole range as an aggregate's argument, and in a plain cell's value position (a scalar argument, an operand) the cell implicit intersection picks — `=INDIRECT("A1:A10")*2` in row 5 reads A5 (see "Plain cells" below); `IFERROR(INDIRECT("A1"), x)` and `LEFT(INDIRECT("B1"), 2)` work for valid references. Structural row/column edits do not rewrite text inside INDIRECT strings (Excel parity — that is INDIRECT's purpose).

**Scenario-table seeding (GH-498/GH-506):** `seedDataTablesReport` and `xl recalc --tables` leave a table untouched and report `Skipped` when its source or precedent cone contains dynamic references such as INDIRECT or OFFSET. The warning names the affected cells; dynamic formulas outside the cone and input formulas replaced by axis values do not block seeding. Source/axis/cycle-member failures report the number of unseeded interiors, and unresolved-precedent diagnostics survive even when the source also fails. `--strict` fails on these warnings. Ordinary formula evaluation of INDIRECT/OFFSET remains supported.

**LET — lexical bindings (GH-193)**: `LET(name1, value1, ..., calculation)` is supported with let* semantics: each binding sees prior bindings; names are case-insensitive, must start with a letter or underscore, and must not be cell-ref-shaped; inner LETs shadow outer. Bindings used in typed argument positions (text, integer, boolean, number, date) coerce totally per the cell-decoder conventions — e.g. `LET(k, 2, LEFT("hey", k))` and `LET(d, A1, YEAR(d))` work, and uncoercible values produce a clean per-cell error, never an exception. Known divergences from Excel: (a) range-valued bindings work via parse-time substitution of literal ranges (`LET(r, A1:A10, SUMIF(r, ">5"))` works); a binding whose value is a reference — a cell, a name, OFFSET, INDIRECT, INDEX, IF/CHOOSE over references — binds that reference, so it works in a range slot too (GH-710: `LET(r, dyn, COUNTIF(r, ">2"))`, `LET(r, OFFSET(A1,0,0,10,1), MATCH(7, r, 0))`), and INDEX indexes a binding that holds an array (`LET(x, SEQUENCE(3), INDEX(x, 2))` is 2); a binding holding any other value is `#VALUE!` in a range slot; (b) re-printed or structurally-shifted LET formulas show range bindings substituted into the body — `=LET(r, A1:A3, SUM(r)/COUNT(r))` reprints as `=LET(r, A1:A3, SUM(A1:A3)/COUNT(A1:A3))`, so structural row/column edits rewrite stored formula text with the literal ranges in place of the body's `r` usages (AST law and semantics preserved); (c) LET is a parser special form, not a registry function: it does not appear in `FunctionRegistry.allNames`, so CLI `functions` listings omit it; (d) dependency extraction over-approximates: binding values contribute their cell refs even when the binding is unused in the body; (e) duplicate-name rebinding is allowed (Excel rejects), dotted names (Excel-legal) are rejected.

**Omitted arguments (GH-603)**: an empty argument slot — `RATE(nper,,pv,fv,)`, `IF(cond,,x)`, `INDEX(rng,,2)` — parses to an explicit `TExpr.Missing`, Excel's blank. A typed slot coerces it like a blank cell (`LEFT("abc",)` is `""`, `SUM(1,,2)` is 3, `DATE(2024,,1)` is December 2023); a value slot reads it as 0 (`IF(TRUE,,5)` is 0, as in Excel); a range slot refuses it; an optional slot holds it as a *present* blank (GH-654): `VLOOKUP(x,rng,2,)` is an exact match, `MATCH(x,rng,)` exact, `RATE(…,,)` starts its iteration at 0, `LOG(10,)` is an error (base 0), `INDEX(rng,,2)` is the whole column. The dynamic-array and reference functions — OFFSET, SEQUENCE, SORT, UNIQUE, FILTER, XLOOKUP — read the empty slot as omitted, as Excel does (`SORT(rng,,-1)` sorts by its first column; `XLOOKUP(x,a,b,,0)` is `#N/A` on no match). An empty *range* slot (`SUMIF(rng,crit,)`) is absent — Excel reports an error there. Every slot prints back empty, trailing ones included; the human-facing form spells an empty slot as `,,` rather than `, ,`. Functions whose xl arity is stricter than Excel's (LEFT/RIGHT require two arguments) accept the empty slot but not its absence. Rules verified against LibreOffice's recalculation.

**`@` — implicit intersection (GH-604)**: Excel 365's `@x` is stored as `_xlfn.SINGLE(x)`; the model keeps the formula-bar spelling, `FormulaStorage` maps the two at every formula-text boundary (cell `<f>`, conditional-formatting and data-validation formulas, `<definedName>`) for any operand the parser accepts — a reference, name, call, literal, parenthesized expression or nested `@`, rescanned so `@INDEX(@rng,1)` maps whole — the parser accepts both spellings, and the printer emits `@x`. Semantics per Excel: a single cell or scalar is itself, a column vector yields the cell in the formula's row, a row vector the cell in its column, a 2-D range or a vector the formula's row/column does not cross is `#VALUE!`; a reference a defined name, a LET name or a function (OFFSET, INDIRECT, INDEX, IF, CHOOSE…) evaluates to intersects the same way, so `@f` in a plain cell is `f`; an array *value* (SEQUENCE, FILTER, arithmetic) collapses to its top-left element. A multi-cell `@range` evaluated without a cell position (`xl eval`) is a per-cell error rather than a guess. Not modeled: `@` before anything other than one primary (`@-A1` is a parse error).

**Plain cells are legacy formulas (implicit intersection)**: a plain formula cell — an `<f>` without `t="array"`, which is what `putf`, batch `putf`, `fx"…"`, openpyxl and most writers produce — is a legacy formula to Excel 365: Excel evaluates it with implicit intersection wherever the legacy evaluation applied it (and shows `@` there). xl evaluates it the same way, at its own cell:
- **A reference in a value position is intersected** with the formula's cell. Value positions are the operands of `+ - * / ^ % &` and the comparisons, scalar function arguments, criteria and lookup values, the IF/IFS condition, the CHOOSE index, NOT's operand, and the whole formula. A column reads the formula's row, a row reads its column, and anything the formula's row or column does not cross — a 2-D range included — is `#VALUE!`, an error value IFERROR and ISERROR see. `=A1:A10*2` in row 5 is A5*2; `=A:A` over a blank row is 0.
- **A reference-class argument takes a reference whole, but evaluates an expression as a value.** This covers the arguments of SUM, COUNT, COUNTA, COUNTBLANK, AVERAGE, MIN, MAX, MEDIAN, STDEV, STDEVP, VAR and VARP, of AND and OR, and of ROWS and COLUMNS. `=SUM(A1:A10*B1:B10)` in row 5 is A5*B5, and `#VALUE!` outside rows 1–10 (`=STDEV(A1:A10*1)` sees one value there, so it is `#DIV/0!`, as in Excel). The array sum needs `SUMPRODUCT(A1:A10*B1:B10)` or an array formula (Excel 365 saves a typed `=SUM(A1:A10*B1:B10)` as one; xl cannot write one from the CLI yet). A value passed to an aggregate follows Excel's rule for a typed argument: TRUE is 1, `"5"` is 5, other text is `#VALUE!`, COUNT skips it, COUNTA counts `""`.
