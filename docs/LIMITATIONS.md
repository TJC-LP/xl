# XL Current Limitations and Future Roadmap

**Last Updated**: 2026-06-10
**Current Phase**: Core domain + OOXML + streaming I/O complete; formula system complete (**107 functions** + cross-sheet support + dynamic arrays); structural editing (insert/delete rows & columns with formula rewriting); named-range & hyperlink authoring; tables + benchmarks complete; row/column serialization complete; **security hardening complete** (ZIP bomb detection, XXE prevention, formula injection guards in both in-memory and streaming writes); scripting prelude + whole-workbook recalculation + print setup authoring (0.11.0).

> **Note (0.11.0):** Sections below largely predate the 0.10.0 "Trust & Author" and 0.11.0 "Scripting" releases and may understate current capabilities. Hyperlinks (#9) and named ranges (#15) are **supported**; inline worksheet elements like data validation (#11) are **preserved through edits** (authoring still pending); print setup is **partially supported** as of 0.11.0 (#16). 0.11.0 also added the scripting prelude (`com.tjclp.xl.scripting`), whole-workbook `recalculate` with per-cell errors, per-side borders/outlines, and sheet view settings — see the [CHANGELOG](../CHANGELOG.md). For the 0.10.0 rationale, see [archive/plan/v0.10.0-execution.md](archive/plan/v0.10.0-execution.md).
>
> **Known open issues** (GitHub): leading `=+` formula prefix rejected (#271), trailing empty number-format sections (#262), quoting of cell-ref-shaped sheet names (#263), streaming StylePatcher indent (#264), SaxStax DirectSaxEmitter metadata gaps (#265), even/first page headers + fitToPage (#266), charts (#222), drawings (#221).

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
- ✅ **1904 date system** (GH-243): `<workbookPr date1904="1"/>` (legacy Mac Excel) is read into `WorkbookMetadata.date1904`, preserved on write, and `DateTime` cells are serialized with the 1904 epoch; epoch-aware conversions via `CellValue.excelSerialToDateTime(serial, date1904 = true)`. Display formatting, typed codec reads, and formula evaluation still assume the 1900 system — interpret raw serials with the metadata flag for 1904 files. **Streaming-edit exception**: streaming writes (CLI `--stream batch` via `StreamingTransform`) still serialize *new* date values with the 1900 epoch while the output keeps `date1904="1"`, so those cells land 1,462 days off — avoid streaming date puts into 1904-system files; in-memory writes (the default, via `XlsxWriter`) are epoch-correct.
- ✅ **Security Hardening**: ZIP bomb detection, XXE prevention, formula injection guards (WI-30)

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
**Features**: Parser, evaluator, **107 functions** (including SUMIF, COUNTIF, SUMIFS, COUNTIFS, XLOOKUP, HLOOKUP, INDEX, MATCH, OFFSET, INDIRECT, XIRR, XNPV, and dynamic arrays SEQUENCE/SORT/UNIQUE/FILTER), dependency graph, cycle detection, cross-sheet references
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
- ✅ **107 Built-in Functions** (complete current registry, verified against `xl functions`):
  - **Aggregate** (12): SUM, COUNT, COUNTA, COUNTBLANK, AVERAGE, MEDIAN, MIN, MAX, STDEV, STDEVP, VAR, VARP
  - **Statistical** (5): LARGE, SMALL, RANK, PERCENTILE, QUARTILE
  - **Conditional** (9): SUMIF, COUNTIF, SUMIFS, COUNTIFS, AVERAGEIF, AVERAGEIFS, MAXIFS, MINIFS, SUMPRODUCT
  - **Logical / Selection** (13): IF, IFS, IFERROR, SWITCH, CHOOSE, AND, OR, NOT, ISNUMBER, ISTEXT, ISBLANK, ISERR, ISERROR
  - **Text** (12): CONCATENATE, LEFT, RIGHT, MID, LEN, UPPER, LOWER, TRIM, FIND, SUBSTITUTE, TEXT, VALUE
  - **Date** (12): TODAY, NOW, DATE, YEAR, MONTH, DAY, EOMONTH, EDATE, DATEDIF, NETWORKDAYS, WORKDAY, YEARFRAC
  - **Math** (16): ABS, ROUND, ROUNDUP, ROUNDDOWN, INT, MOD, POWER, SQRT, LOG, LN, EXP, FLOOR, CEILING, TRUNC, SIGN, PI
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
- ⏳ Leading `=+` formula prefix (legacy Lotus style) is rejected by the parser - #271
- 🛡️ **Formula depth cap (GH-56 totality guard):** nesting + operator-chain depth is capped at 256 levels so pathological input returns `ParseError.NestingTooDeep` instead of `StackOverflowError`. Excel allows 64 nesting levels and no operator-chain limit (within 8192 chars); xl additionally counts each operator in a *flat, paren-less* chain, so a chain of 256+ consecutive operators (e.g. `=A1+A2+…` ×256) is rejected — use `SUM`/ranges, or parenthesize to reset the budget. No real-world formula approaches this.

**INDIRECT — dynamic references (GH-274)**: `INDIRECT(ref_text, [a1])` resolves A1-style text — `"B5"`, `"A1:B10"`, `"$A$1:B10"`, `"A:A"`, `"Sheet2!A1"`, `"'My Sheet'!A1:C3"` — at evaluation time and composes with aggregates (`=SUM(INDIRECT("A1:A"&B1))`), array arithmetic, and spill (`evala`) via the same ArrayResult mechanism as OFFSET. Unresolvable, out-of-grid, defined-name, external-workbook, or structured-reference text evaluates to `#REF!` (a value — total, never throws). Full column/row text is clamped to the sheet's used range before materializing; resolved ranges are capped at 1,048,576 cells.

**Dependency semantics (differs from Excel):** the static graph sees INDIRECT's *arguments* (e.g. A1 in `=INDIRECT(A1)`), not its resolved *targets*. `recalculate()` evaluates INDIRECT-bearing formulas and their dependents after all other formulas (stable evaluate-last partition), and resolves not-yet-evaluated targets on demand with the same depth-100 recursion guard as cross-sheet references — so INDIRECT chains compute fresh values. Dynamic circular references (INDIRECT resolving into its own dependents) are not pre-detected by Tarjan; they surface as per-cell recursion-guard errors (cells left uncached) while the rest of the workbook still evaluates. Static cycle detection is unchanged (zero new false positives). xl's `recalculate()` is always full-workbook, so Excel's "volatile" marking is moot there; the targeted `recalculateDependents` treats INDIRECT-bearing cells as always dirty.

**Not supported (v1):** R1C1 mode (`a1=FALSE`) returns an evaluation error and leaves the cell uncached (Excel computes it on open) — deliberately not `#REF!`; consequently `INDIRECT(ADDRESS(..., FALSE), FALSE)` does not compose. INDIRECT is not accepted where a literal range is required at parse time (SUMIF/VLOOKUP/XLOOKUP/INDEX/MATCH/NPV/IRR array slots) — the same envelope as OFFSET. In scalar argument positions (e.g. `IFERROR(INDIRECT(...), x)`, `LEFT(INDIRECT(...), n)`) the array result is rejected by the scalar evaluator, so IFERROR returns its fallback even for valid references — shared with OFFSET, tracked as a family follow-up. Structural row/column edits do not rewrite text inside INDIRECT strings (Excel parity — that is INDIRECT's purpose).

**LET — lexical bindings (GH-193)**: `LET(name1, value1, ..., calculation)` is supported with let* semantics: each binding sees prior bindings; names are case-insensitive, must start with a letter or underscore, and must not be cell-ref-shaped; inner LETs shadow outer. Bindings used in typed argument positions (text, integer, boolean, number, date) coerce totally per the cell-decoder conventions — e.g. `LET(k, 2, LEFT("hey", k))` and `LET(d, A1, YEAR(d))` work, and uncoercible values produce a clean per-cell error, never an exception. Known divergences from Excel: (a) range-valued bindings work via parse-time substitution of literal ranges (`LET(r, A1:A10, SUMIF(r, ">5"))` works), but a binding whose value is a range-*returning* call (e.g. OFFSET) cannot be used where a literal range is syntactically required — it still works in array-tolerant positions like SUM; (b) re-printed or structurally-shifted LET formulas show range bindings substituted into the body — `=LET(r, A1:A3, SUM(r)/COUNT(r))` reprints as `=LET(r, A1:A3, SUM(A1:A3)/COUNT(A1:A3))`, so structural row/column edits rewrite stored formula text with the literal ranges in place of the body's `r` usages (AST law and semantics preserved); (c) LET is a parser special form, not a registry function: it does not appear in `FunctionRegistry.allNames`, so CLI `functions` listings omit it; (d) dependency extraction over-approximates: binding values contribute their cell refs even when the binding is unused in the body; (e) duplicate-name rebinding is allowed (Excel rejects), dotted names (Excel-legal) are rejected.

**RAND/RANDBETWEEN — volatile functions (GH-115)**: `RAND()`/`RANDBETWEEN(bottom, top)` are supported. Volatility: xl re-evaluates volatile formulas on every `recalculate`/`withCachedFormulas` pass, so cached values change per recalculation (matching Excel). Determinism: randomness is an explicit `Rng` capability (Clock pattern) — pass `Rng.seeded(seed)` to the rng-taking overloads (`sheet.evaluateFormula(f, clock, rng)`, `wb.recalculate(clock, rng)`, ...) for reproducible output; default paths use `Rng.system`. RANDBETWEEN tightens fractional bounds inward (bottom rounds up, top down) and errors when bottom > top.

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
shared across sheets), `writeStreamsSeqWithAutoDetect`, and the new `writeStreamStyled`.

**Styles** (GH-223 phase 2): `ExcelIO.writeStreamStyled(path, sheet, styles)` takes a
`Vector[CellStyle]` table; `StyledRowData.cellStyles` values index into it. The table is
deduplicated into `xl/styles.xml` (cellXf 0 = default) and cell `s=` attributes are remapped to
the emitted indices — formatted 100k+ row files no longer require the in-memory path.

**Remaining envelope / limitations**:
- Memory is O(distinct strings) for the accumulator — the accepted envelope per the design doc
  (100k rows with 100 distinct strings keeps exactly 100 entries; 1M unique strings ≈ tens of MB)
- `SstPolicy.Never` (`WriterConfig`) keeps the previous inline-string dialect; `Auto`/`Always`
  both deduplicate (streaming cannot pre-scan, so `Auto` opts into SST — note this changes the
  default streaming output dialect from inline strings to SST as of GH-223)
- RichText cells stay `inlineStr` even in SST mode (mixed dialects are valid OOXML)
- Merged-cell emission from streaming writers (phase 3 of GH-223) is still future work

---

### 🟢 Low Impact (Nice to Have)

#### 9. Hyperlinks ✅ NOW SUPPORTED (0.10.0)
**Status**: Implemented
**Impact**: `Cell.hyperlink` is serialized to `<hyperlinks>` + relationships and populated on read; a `hyperlink` batch op is available. Previously the model existed but write was a silent no-op.

---

#### 10. Conditional Formatting Not Supported
**Status**: Not implemented (preserved through edits since 0.10.0, like other unknown/inline parts)
**Impact**: Cannot author color scales, data bars, icon sets
**Plan**: see [plan/roadmap.md](plan/roadmap.md)

**Effort**: 5-7 days
**LOC**: ~300 (Rules model, XML serialization, testing)

---

#### 11. Data Validation Authoring Not Supported
**Status**: Existing `dataValidations` are **preserved through edits** since 0.10.0 (C1); no authoring API yet
**Impact**: Cannot add new dropdown lists or input validation
**Plan**: see [plan/roadmap.md](plan/roadmap.md)

**Effort**: 3-4 days
**LOC**: ~200

---

#### 12. Charts Not Supported
**Status**: Not implemented (existing charts are preserved through edits) — tracked in #222
**Impact**: Cannot generate charts
**Plan**: see [plan/roadmap.md](plan/roadmap.md)

**Effort**: 20-30 days (massive scope)
**LOC**: ~1500+

---

#### 13. Drawings (Images, Shapes) Not Supported
**Status**: Not implemented (existing drawings are preserved through edits) — tracked in #221
**Impact**: Cannot embed images or shapes
**Plan**: see [plan/roadmap.md](plan/roadmap.md)

**Effort**: 10-15 days
**LOC**: ~800

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
**Impact**: `WorkbookMetadata.definedNames` is serialized to `<definedNames>` (previously read-only), with a CLI `name add` / `name rm` verb. Structured references inside formulas remain future work.

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

### Phase 8: Drawings (Priority 4)
**Effort**: 2-3 weeks
**Focus**: Images and shapes

See: [plan/roadmap.md](plan/roadmap.md) (drawings tracked in #221)

- [ ] Image embedding (5-7 days)
- [ ] Shapes (3-4 days)
- [ ] Drawing relationships (2-3 days)

---

### Phase 9: Charts (Priority 5)
**Effort**: 4-6 weeks
**Focus**: Chart generation

See: [plan/roadmap.md](plan/roadmap.md) (charts tracked in #222)

**Very complex**: 15+ chart types, each with custom XML structure

---

### Phase 10: Tables & Advanced Features (Priority 6)
**Effort**: 3-4 weeks
**Focus**: Tables, pivots, conditional formatting

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
| **Formulas (eval)** | ✅ | ✅ | XL: 107 functions, dependency graph, cycle detection |
| **Tables** | ✅ | ✅ | XL: Full table support with AutoFilter, structured refs |
| **Charts** | ❌ | ✅ | POI: Full support |
| **Drawings** | ❌ | ✅ | POI: Images/shapes |
| **Memory (100k rows)** | ✅ 50MB | ❌ 800MB | XL: 16x better |
| **Type Safety** | ✅ | ❌ | XL: Compile-time, POI: Runtime |
| **Purity** | ✅ | ❌ | XL: Pure, POI: Mutable |
| **Determinism** | ✅ | ❌ | XL: Stable diffs, POI: Non-deterministic |
| **Security** | ✅ | ⚠️ | XL: ZIP bomb, XXE, formula injection protection |

**Verdict**: XL is production-ready for data-heavy use cases (streaming, ETL). POI better for rich formatting/charts (for now).

---

## Performance Benchmarks: XL vs Apache POI

*See [design/performance-investigation.md](design/performance-investigation.md) for full analysis.*

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
- Formula evaluation (107 functions, dependency graph, cycle detection)
- Excel Tables (structured data with AutoFilter, headers, styling)
- Performance-critical workloads (benchmarked vs POI)

**No/Not yet, if you need**:
- Charts or drawings (planned)
- Pivot tables/conditional formatting/data validation (planned)
- Excel macros (not planned to execute; preservation planned)

---

### Q: What's the maximum file size XL can handle?

**Streaming API**: Unlimited
- Tested: 100k rows (completes in ~3s)
- Projected: 1M rows (~30s)
- Memory: O(1) constant (~50-100MB regardless of size)

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
- 107 built-in functions (SUM, SUMIF, COUNTIF, XLOOKUP, HLOOKUP, INDEX, MATCH, OFFSET, INDIRECT, XIRR, XNPV, dynamic arrays, etc.)
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

1. **Need charts or drawings**: Use POI for chart/drawing *authoring* (#222, #221); XL preserves existing charts/drawings through edits
2. **Unsupported formula function**: Store the formula as a string and let Excel recalculate on open (XL evaluates the 104 built-ins)
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

*Last updated 2026-06-10 (0.11.0 doc-truth pass, GH-272: refreshed fixed/open status against the 0.10.0 and 0.11.0 releases, corrected the function registry breakdown, repointed retired plan links).*
