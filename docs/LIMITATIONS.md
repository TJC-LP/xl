# XL Current Limitations and Future Roadmap

**Last Updated**: 2025-12-18 (Cross-Sheet Formula Support - TJC-351)
**Current Phase**: Core domain + OOXML + streaming I/O complete; formula system complete (24 functions + cross-sheet support); tables + benchmarks complete; **security hardening complete** (ZIP bomb detection, XXE prevention, formula injection guards).

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
- ✅ **Security Hardening**: ZIP bomb detection, XXE prevention, formula injection guards (WI-30)

### Developer Experience
- ✅ **Mill build system**: Fast, reliable builds
- ✅ **Scalafmt integration**: Consistent code formatting
- ✅ **GitHub Actions CI**: Automated testing
- ✅ **Comprehensive docs**: README, CLAUDE.md, 29 plan files
- ✅ **Property-based testing**: ScalaCheck generators for all core types
- ✅ **Law verification**: Monoid laws, round-trip laws, invariants

---

## Recently Completed ✅ (This Session)

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

#### 4. Merged Cells in Streaming Writes
**Status**: Fully supported in the in‑memory OOXML path; not emitted by streaming writers.

**Current State**:
- In‑memory:
  - `Sheet.mergedRanges: Set[CellRange]` tracks merged regions.
  - `OoxmlWorksheet.toXml` emits `<mergeCells>` / `<mergeCell>` for those ranges.
- Streaming write (`writeStreamTrue`, `writeStreamsSeqTrue`):
  - Only writes `sheetData` with plain rows and cells; no merged cell metadata is currently generated.

**Impact**:
- In‑memory read/write round‑trips preserve merges.
- Pure streaming‑generated workbooks will not contain merged ranges.

---

#### 5. Column/Row Properties Not Serialized
**Status**: Tracked but not written
**Impact**: Column widths, row heights, hidden state lost on write
**Plan**: [P4 Continuation - Dimension Metadata](plan/11-ooxml-mapping.md#column-row-properties)
**Phase**: P4

**Current State**:
- `ColumnProperties(width, hidden)` and `RowProperties(height, hidden)` defined
- `Sheet.columnProps` and `Sheet.rowProps` track state
- Patches work: `Patch.SetColumnProperties`, `Patch.SetRowProperties`
- BUT: Not serialized to `<cols>` or `<row ht="">` in XML

**What's Missing**:
```xml
<!-- Should emit: -->
<cols>
  <col min="1" max="1" width="15.5" customWidth="1"/>
</cols>

<row r="1" ht="25.5" customHeight="1">
  ...
</row>
```

**Effort**: 3-4 hours
**LOC**: ~60 changes in Worksheet.scala

---

#### 6. Formula System ✅ **PRODUCTION READY**
**Status**: Complete (WI-07, WI-08, WI-09a/b/c/d + financial functions + TJC-351 cross-sheet formulas)
**Features**: Parser, evaluator, 24 functions (including NPV, IRR, VLOOKUP), dependency graph, cycle detection, cross-sheet references
**Plan**: [Formula System](plan/formula-system.md)
**Phase**: WI-07, WI-08, WI-09a/b/c/d Complete + Financial Functions + Cross-Sheet Formulas

**What Works** (Production Ready - 280+ tests):
```scala
import com.tjclp.xl.formula.{FormulaParser, FormulaPrinter, Evaluator}
import com.tjclp.xl.formula.SheetEvaluator.*

// Parse formulas to typed AST
FormulaParser.parse("=SUM(A1:B10)") // Right(TExpr.FoldRange(...))
FormulaParser.parse("=IF(A1>0, \"Positive\", \"Negative\")") // Right(TExpr.If(...))

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
- ✅ **24 Built-in Functions** (WI-09a/b/c + financial functions - 78+ tests):
  - **Aggregate** (5): SUM, COUNT, AVERAGE, MIN, MAX
  - **Logical** (4): IF, AND, OR, NOT
  - **Text** (6): CONCATENATE, LEFT, RIGHT, LEN, UPPER, LOWER
  - **Date** (6): TODAY, NOW, DATE, YEAR, MONTH, DAY
  - **Financial** (3): NPV, IRR, VLOOKUP
- ✅ **Dependency Graph** (WI-09d - 52 tests):
  - Tarjan's SCC algorithm: O(V+E) cycle detection
  - Kahn's algorithm: O(V+E) topological sort
  - Precedent/dependent queries: O(1) lookups
  - Safe evaluation API: sheet.evaluateWithDependencyCheck()
  - Performance: 10k formula cells in <10ms
- ✅ **Operators**: +, -, *, /, <, <=, >, >=, =, <>, &, AND, OR, NOT
- ✅ **Scientific notation**: 1.5E10, 2.3E-7
- ✅ **Round-trip**: parse ∘ print = id (property-tested)
- ✅ **Cross-sheet formulas** (TJC-351 - 26 tests):
  - Single cell refs: `=Sales!A1`, `=Data!B2`
  - Range refs with SUM: `=SUM(Sales!A1:A10)`
  - Arithmetic: `=Sales!A1 + Revenue!B1`
  - Workbook-level cycle detection: `DependencyGraph.fromWorkbook`, `detectCrossSheetCycles`

**Future Extensions** (Not Critical):
- ⏳ Extended function library (300+ Excel functions) - WI-09e+
- ⏳ Array formulas - Future work
- ⏳ Structured references (Table[@Column]) - Requires WI-10 integration
- ⏳ Quoted sheet names in formulas (`='Q1 Report'!A1`) - Parser support pending
- ⏳ Cross-sheet COUNT/AVERAGE/MIN/MAX - Only SUM implemented for cross-sheet ranges

**No workarounds needed** - formula system is complete and production-ready!

---

#### 7. Theme Colors Not Resolved
**Status**: Color.Theme exists but always returns 0
**Impact**: Theme-based colors don't work
**Plan**: [P5 - Styles](plan/05-styles-and-themes.md#theme-resolution)
**Phase**: P5 (Deferred)

**Current State**:
```scala
Color.Theme(ThemeSlot.Accent1, tint = 0.5)
// toArgb returns: 0x00000000 (black, wrong!)
// Should resolve: Theme palette → apply tint → ARGB
```

**What's Missing**:
- Parse `xl/theme/theme1.xml`
- Extract theme color palette
- Resolve `ThemeSlot` → base RGB
- Apply tint transformation
- Cache resolved colors

**Effort**: 1-2 days
**LOC**: ~100 in Theme parsing + resolution

**Workaround**: Use `Color.Rgb()` directly with explicit ARGB values

---

#### 8. No Shared Strings Table (SST) in Streaming Write
**Status**: Streaming write always uses inline strings
**Impact**: Larger file sizes for repeated strings
**Plan**: [P6 - SST Streaming Integration](plan/13-streaming-and-performance.md#sst-streaming)
**Phase**: P6 (Future)

**Current State**:
- `writeStreamTrue` uses `type="inlineStr"` for all text cells
- Each cell carries full string (no deduplication)
- File size: ~2x larger for files with repeated strings

**Example**:
```scala
// 100k rows with same company name
RowData(i, Map(0 -> CellValue.Text("ACME Corporation")))
// TODAY: "ACME Corporation" written 100k times (~1.7MB)
// WITH SST: Written once in SST, cells reference index (~200KB)
```

**Why It's Hard**:
- Streaming write consumes rows once (cannot scan twice)
- Two-pass approach: 1) collect strings, 2) write with indices (breaks streaming)
- Online SST building: Track duplicates as stream flows (complex state)

**Effort**: 4-5 days
**LOC**: ~200 in SST streaming builder

**Workaround**: Acceptable for most use cases (disk space is cheap)

---

### 🟢 Low Impact (Nice to Have)

#### 9. Hyperlinks Not Supported
**Status**: Not implemented
**Impact**: Cannot create clickable links in cells
**Plan**: [P8 - Hyperlinks](plan/11-ooxml-mapping.md#hyperlinks)
**Phase**: P8 (Future)

**Effort**: 2-3 days
**LOC**: ~120 (Hyperlink model, relationships, worksheet XML)

---

#### 10. Conditional Formatting Not Supported
**Status**: Not implemented
**Impact**: Cannot add color scales, data bars, icon sets
**Plan**: [P10 - Conditional Formatting](plan/08-tables-and-pivots.md#conditional-formatting)
**Phase**: P10 (Future)

**Effort**: 5-7 days
**LOC**: ~300 (Rules model, XML serialization, testing)

---

#### 11. Data Validation Not Supported
**Status**: Not implemented
**Impact**: Cannot add dropdown lists, input validation
**Plan**: [P10 - Data Validation](plan/08-tables-and-pivots.md#data-validation)
**Phase**: P10 (Future)

**Effort**: 3-4 days
**LOC**: ~200

---

#### 12. Charts Not Supported
**Status**: Not implemented
**Impact**: Cannot generate charts
**Plan**: [P9 - Charts](plan/07-charts.md)
**Phase**: P9 (Future - Very Complex)

**Effort**: 20-30 days (massive scope)
**LOC**: ~1500+

---

#### 13. Drawings (Images, Shapes) Not Supported
**Status**: Not implemented
**Impact**: Cannot embed images or shapes
**Plan**: [P8 - Drawings](plan/06-drawings.md)
**Phase**: P8 (Future)

**Effort**: 10-15 days
**LOC**: ~800

---

#### 14. Pivot Tables Not Supported
**Status**: Not implemented
**Impact**: Cannot create Pivot Tables
**Plan**: [P10 - Advanced Features](plan/advanced-features.md)
**Phase**: P10 (Future)

**Note**: Excel Tables (structured data ranges with headers, AutoFilter, styling) are ✅ **fully supported** as of WI-10.

**Effort**: 10-15 days (for Pivot Tables only)
**LOC**: ~700

---

#### 15. Named Ranges Not Supported
**Status**: Not implemented
**Impact**: Cannot define or use named ranges
**Plan**: [P7 - Named Ranges](plan/11-ooxml-mapping.md#named-ranges)
**Phase**: P7 (Future)

**Effort**: 1-2 days
**LOC**: ~100

---

#### 16. Print Settings and Page Setup Not Supported
**Status**: Not implemented
**Impact**: Default print settings used
**Plan**: [P10 - Page Setup](plan/11-ooxml-mapping.md#page-setup)
**Phase**: P10 (Future)

**Effort**: 2-3 days
**LOC**: ~150

---

#### 17. Document Properties Not Written
**Status**: Not implemented
**Impact**: Missing metadata (author, title, creation date)
**Plan**: [P10 - Document Properties](plan/11-ooxml-mapping.md#document-properties)
**Phase**: P10 (Future)

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
**Status**: Implemented (WI-30)
**Impact**: Untrusted data can be safely written to Excel

**API** (`CellValue.escape()` and `WriterConfig.secure`):
```scala
// Manual escaping for individual strings
CellValue.escape("=SUM(A1)")  // Returns: "'=SUM(A1)"
CellValue.escape("+1234")     // Returns: "'+1234"
CellValue.escape("-danger")   // Returns: "'-danger"
CellValue.escape("@import")   // Returns: "'@import"
CellValue.escape("Normal")    // Returns: "Normal" (unchanged)

// Automatic escaping via WriterConfig.secure
XlsxWriter.writeWith(workbook, path, WriterConfig.secure)
// All text cells starting with =, +, -, @ are automatically escaped
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
**Plan**: [P6 - Codecs](plan/09-codecs-and-named-tuples.md)
**Phase**: P6 (High Priority)

**What's Missing**:
```scala
// WANT: Automatic codec derivation
case class Person(name: String, age: Int, salary: BigDecimal)

given Codec[Person] = Codec.derived

// Read: Stream[F, RowData] → Stream[F, Person]
excel.readStream(path).through(Codec[Person].decode)

// Write: Stream[F, Person] → XLSX
people.through(Codec[Person].encode)
  .through(excel.writeStreamTrue(path, "People"))
```

**Effort**: 7-10 days
**LOC**: ~500 (type-class derivation, header binding, testing)

---

#### 23. No Path Macro for Named Cell References
**Status**: Not implemented
**Impact**: Cannot use symbolic names for cells
**Plan**: [P7 - Advanced Macros](plan/17-macros-and-syntax.md#path-macro)
**Phase**: P7 (Future)

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
**Status**: Not implemented
**Impact**: Verbose style definitions
**Plan**: [P7 - Advanced Macros](plan/17-macros-and-syntax.md#style-literal)
**Phase**: P7 (Future)

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
**Plan**: [P11 - Security](plan/23-security.md#macro-preservation)
**Phase**: P11 (Future)

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
excel.writeStreamsSeqTrue(
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

### Phase 4 Continuation (Priority 1 - Next Sprint)
**Effort**: 1-2 weeks
**Focus**: Complete OOXML coverage for common use cases

- [ ] Style application end-to-end (2-3 days) - **HIGH IMPACT**
- [ ] DateTime serialization (1 day) - **HIGH IMPACT**
- [ ] Merged cells XML (3 hours) - MEDIUM IMPACT
- [ ] Column/row properties XML (4 hours) - MEDIUM IMPACT

**Result**: Fully functional spreadsheet library with formatting

---

### Phase 6: Codecs & Named Tuples (Priority 2)
**Effort**: 2-3 weeks
**Focus**: Ergonomic data binding

See: [plan/09-codecs-and-named-tuples.md](plan/09-codecs-and-named-tuples.md)

- [ ] Type-class derivation (7-10 days)
- [ ] Header row binding (3-4 days)
- [ ] Tuple/case class codecs (5-6 days)

**Result**: `Stream[F, Person]` ↔ XLSX with zero boilerplate

---

### Phase 7: Advanced Macros (Priority 3)
**Effort**: 1-2 weeks
**Focus**: Enhanced compile-time DSL

See: [plan/17-macros-and-syntax.md](plan/17-macros-and-syntax.md)

- [ ] Path macro for named references (3-4 days)
- [ ] Style literal (2-3 days)
- [ ] Formula macro with type checking (5-7 days)

**Result**: Best-in-class developer experience

---

### Phase 8: Drawings (Priority 4)
**Effort**: 2-3 weeks
**Focus**: Images and shapes

See: [plan/06-drawings.md](plan/06-drawings.md)

- [ ] Image embedding (5-7 days)
- [ ] Shapes (3-4 days)
- [ ] Drawing relationships (2-3 days)

---

### Phase 9: Charts (Priority 5)
**Effort**: 4-6 weeks
**Focus**: Chart generation

See: [plan/07-charts.md](plan/07-charts.md)

**Very complex**: 15+ chart types, each with custom XML structure

---

### Phase 10: Tables & Advanced Features (Priority 6)
**Effort**: 3-4 weeks
**Focus**: Tables, pivots, conditional formatting

See: [plan/08-tables-and-pivots.md](plan/08-tables-and-pivots.md)

---

### Phase 11: Security & Safety ✅ (WI-30 Complete)
**Status**: Production ready (2025-12-07)
**Focus**: Security hardening

See: [plan/23-security.md](plan/23-security.md)

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
| **Formulas (eval)** | ✅ | ✅ | XL: 24 functions, dependency graph, cycle detection |
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

## Frequently Asked Questions

### Q: Can I use XL in production today?

**Yes, if your use case is**:
- Large dataset export (100k+ rows)
- ETL pipelines (streaming read/write, constant memory)
- Data generation (reports, analytics)
- Multi-sheet workbooks
- Core cell types and rich text
- Styling in in-memory workflows (full styles supported)
- Formula evaluation (24 functions, dependency graph, cycle detection)
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
- **Need constant memory and sequential processing**: use streaming (`readStream`, `writeStreamTrue`); styles are minimal and strings are inline.

---

### Q: How complete is formula evaluation?

**Status**: ✅ Production Ready (as of WI-07/08/09)

**What's implemented**:
- 24 built-in functions (SUM, IF, VLOOKUP, NPV, IRR, etc.)
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

1. **Styles not working**: Wait for P4 completion (2-3 days) OR use external formatting tools
2. **Need charts**: Use POI for now, plan to migrate when P9 complete
3. **Formula evaluation**: Let Excel handle it (store formulas as strings)
4. **Selective updates**: Use `read()` + `write()` for small files, wait for P6 for large files

---

## Contributing

Want to help implement these features? See:
- [docs/plan/18-roadmap.md](plan/18-roadmap.md) - Full implementation plan
- [CLAUDE.md](../CLAUDE.md) - Contribution guidelines
- Each feature links to detailed design doc in `docs/plan/`

---

## Summary

**XL Today**: Best-in-class for streaming large datasets with type safety and purity

**XL Tomorrow** (P4-P11): Feature parity with POI while maintaining performance advantages

**XL Vision**: The definitive pure functional Excel library for Scala 3

---

*Last updated by Claude Code session on 2025-12-18 (TJC-351 Cross-Sheet Formula Support)*
