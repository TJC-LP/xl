# XL Project Status

**Last Updated**: 2026-06-10

## Current State

> **For detailed phase completion status and roadmap, see [plan/roadmap.md](plan/roadmap.md)**

### What Works (Production-Ready)

**New in 0.11.0 "Scripting"** (2026-06-10):
- ✅ **Scripting prelude** `com.tjclp.xl.scripting.{*, given}` — ONE import for scripts: core API + DSL + compile-time literals + formula evaluation + sync `Excel` + streaming `ExcelIO` + smart detection (`String.toFormatted`) + `.unsafe` boundary
- ✅ **Range fill**: `range := value` puts the value in every cell (Excel Ctrl+Enter semantics; previously a silent no-op)
- ✅ **Total ARef navigation**: `down`/`up`/`right`/`left` (default 1) for Either-free loops
- ✅ **`Workbook.upsert(name, f)`** — total update-or-create counterpart of `update`
- ✅ **`Workbook.recalculate(clock)`** — total whole-workbook recalculation returning `RecalcResult` (cached workbook + per-sheet values + per-cell `CellEvalError`s; cycles isolated, acyclic remainder still evaluates)
- ✅ **`FormattedParsers.detect`** — total smart value detection (currency/accounting/percent/ISO date/number/boolean/text), promoted from CLI-internal code
- ✅ **Per-side borders + outlines**: `CellStyle.borderTop/borderBottom/borderLeft/borderRight` and `range.outlined(style[, color])` via the new `Patch.MergeBorder`
- ✅ **Alignment indent**: `CellStyle.indent(n)`
- ✅ **Sheet view settings**: `SheetView(showGridLines, zoomScale)` on `Sheet.viewSettings` (SVG renderer respects gridline suppression)
- ✅ **Print setup extensions**: `PageSetup` gains `headerFooter`, `margins`, `printArea`, `repeatRows` (sheet-scoped `_xlnm` defined names; even/first headers + fitToPage tracked in #266)
- ✅ **xl-scripting skill** (`plugin/skills/xl-scripting/`) — SKILL.md + API reference + runnable recipes, compile-verified on CI
- ✅ **Anti-rot CI**: examples job (`scripts/test-examples.sh`) + skill-verify workflow (`scripts/verify-skill-snippets.sh`)

**Core Features**:
- ✅ Type-safe addressing (Column, Row, ARef with 64-bit packing)
- ✅ Compile-time validated literals: `ref"A1"` and `ref"A1:B10"`
- ✅ Immutable domain model (Cell, Sheet, Workbook)
- ✅ Patch Monoid for declarative updates
- ✅ Complete style system (Font, Fill, Border, Color, NumFmt, Align)
- ✅ StylePatch Monoid for style composition
- ✅ StyleRegistry for per-sheet style management
- ✅ **End-to-end XLSX read/write** (creates real Excel files)
- ✅ **Surgical modification** (read → modify → write preserves unknown parts: charts, images, drawings, comments, and inline worksheet elements — dataValidations, sheetProtection, autoFilter — preserved through edits as of 0.10.0 / C1)
- ✅ **Structural editing** (0.10.0): insert/delete rows & columns shift cells, merges, row/col properties, freeze panes, and rewrite all affected formulas (cross-sheet) with `#REF!` generation
- ✅ **Named ranges & hyperlinks authoring** (0.10.0): `DefinedName` and `Cell.hyperlink` are now serialized (previously read-only)
- ✅ Hybrid write optimization (11x speedup for unmodified workbooks, 2-5x for partial modifications)
- ✅ Shared Strings Table (SST) deduplication
- ✅ Styles.xml with component deduplication
- ✅ Multi-sheet workbooks
- ✅ All cell types: Text, Number, Bool, Formula, Error, DateTime
- ✅ RichText support (multiple formats within one cell)
- ✅ DateTime serialization (Excel serial number conversion)
- ✅ **Excel Tables** (structured data ranges with headers, AutoFilter, and styling)
- ✅ **True streaming I/O** (constant memory, 100k+ rows)

**Ergonomics & Type Safety**:
- ✅ Given conversions: `sheet.put(ref"A1", "Hello")` (no wrapper needed)
- ✅ Batch put via varargs `Sheet.put(ref -> value, ...)`
- ✅ Formatted literals: `money"$1,234.56"`, `percent"45.5%"`, `date"2025-11-10"`
- ✅ **String interpolation**: `ref"$sheet!$cell"`, `money"$$${amount}"` with runtime validation
- ✅ Compile-time optimization for literal interpolations (zero runtime overhead)
- ✅ **CellCodec[A]** for 9 primitive types (String, Int, Long, Double, BigDecimal, Boolean, LocalDate, LocalDateTime, RichText)
- ✅ Batch `Sheet.put` with auto-inferred formatting (former `putMixed` API)
- ✅ `readTyped[A]` for type-safe cell reading
- ✅ **Optics** module (Lens, Optional, focus DSL)
- ✅ RichText DSL: `"Bold".bold.red + " normal " + "Italic".italic.blue`
- ✅ HTML export: `sheet.toHtml(range"A1:B10")`
- ✅ **Formula Parsing** (WI-07 complete): TExpr GADT, FormulaParser, FormulaPrinter with round-trip verification and scientific notation
- ✅ **Formula Evaluation** (WI-08 complete): Pure functional evaluator with total error handling, short-circuit semantics, and Excel-compatible behavior
- ✅ **Function Library**: **105 built-in functions** (aggregate, conditional, logical, text, date, financial, lookup, math, statistical, dynamic arrays), extensible type class parser, evaluation API. 0.10.0 added IFS, SWITCH, CHOOSE, LARGE, SMALL, RANK, PERCENTILE, QUARTILE, HLOOKUP, MAXIFS, MINIFS, OFFSET, and the spill functions SEQUENCE/SORT/UNIQUE/FILTER (#76, #120, #122); GH-274 added INDIRECT (dynamic text-to-reference resolution with deferred-bucket recalculation).
- ✅ **Dependency Graph** (WI-09d complete): Circular reference detection (Tarjan's SCC), topological sort (Kahn's algorithm), safe evaluation with cycle detection
- ✅ **Cross-Sheet Formula References** (TJC-351): Single cell refs (`=Sales!A1`), range refs (`=SUM(Sales!A1:A10)`), arithmetic with cross-sheet refs, workbook-level cycle detection (`DependencyGraph.fromWorkbook`)

**Performance** (JMH Benchmarked - WI-15):
- ✅ **Streaming reads: 35% faster than POI for small files** (0.887ms vs 1.357ms @ 1k rows)
- ✅ **Streaming reads: Competitive with POI for large files** (8.408ms vs 7.773ms @ 10k rows - within 8%)
- ✅ **In-memory reads: 26% faster than POI for small files** (1.225ms vs 1.650ms @ 1k rows)
- ✅ Inline hot paths (SAX parser: 3.8x speedup vs fs2-data-xml)
- ✅ Zero-overhead opaque types
- ✅ Macros compile away (no runtime parsing)
- ⚠️ Writes: POI 49% faster (future optimization work - Phase 3)

**Streaming API**:
- ✅ Excel[F[_]] algebra trait
- ✅ ExcelIO[IO] interpreter
- ✅ `readStream` / `readSheetStream` / `readStreamByIndex` – constant‑memory streaming read (fs2.io.readInputStream + fs2‑data‑xml)
- ✅ `writeStream` / `writeStreamsSeq` – constant‑memory streaming write (fs2‑data‑xml)
- ✅ `writeWorkbookStream` – lower-allocation SAX/StAX write for in-memory workbooks; preserves merges, comments, tables, row/column properties, and freeze panes
- ✅ **`writeFast`** – SAX/StAX streaming write (opt-in via `ExcelIO.writeFast()` or `WriterConfig(backend = XmlBackend.SaxStax)`)
- ✅ Benchmark: 100k rows in ~1.8s read (~10MB constant memory) / ~1.1s write (~10MB constant memory)

**Output Configuration** (P6.7 Complete):
- ✅ WriterConfig with compression and prettyPrint options
- ✅ Compression.Deflated default (5-10x smaller files)
- ✅ WriterConfig.debug for debugging (STORED + prettyPrint)
- ✅ Backward compatible API (writeWith for custom config)
- ✅ 4 compression tests verify behavior

**Infrastructure**:
- ✅ Mill build system
- ✅ Scalafmt 3.10.1 integration
- ✅ GitHub Actions CI pipeline
- ✅ Comprehensive documentation (README.md, CLAUDE.md)

### Test Coverage

**3005+ tests** (verified via `./mill __.test`, 2026-06-10):

| Module | Tests | Covers |
|--------|-------|--------|
| xl-evaluator | 1279 | parser, evaluator, 105-function library, dependency graph, cross-sheet formulas, recalculation, structural editing |
| xl-core | 971 | addressing laws, Patch/StylePatch monoids, codecs, optics, RichText, interpolation, render (HTML/SVG), styles DSL |
| xl-ooxml | 381 | round-trips (cells, styles, tables, comments, hyperlinks), compression, security (XXE, ZIP bomb), preservation |
| xl-cli | 308 | command parsing, batch ops, view/eval/export, streaming mode |
| xl-cats-effect | 76 | streaming I/O, O(1) memory verification, SAX/StAX write |
| xl-agent | 54 | benchmark engine, skill abstraction |
| xl (prelude) | 17 | external-consumer probes (`xl/test/src/xlprelude/`) |
| xl-testkit | 0 | placeholder (no sources yet) |

See [reference/testing-guide.md](reference/testing-guide.md) for suite structure and testing patterns.

---

## Current Limitations & Known Issues

### XML Serialization

**Formula System** (WI-07, WI-08, WI-09a/b/c/d - Production Ready):
- ✅ **Parsing** (WI-07): Typed AST (TExpr GADT), FormulaParser, FormulaPrinter, round-trip verification, 57 tests
- ✅ **Evaluation** (WI-08): Pure functional evaluator, total error handling, short-circuit semantics, 58 tests
- ✅ **Function Library** (WI-09a-h + TJC-1055 complete): **105 built-in functions**, extensible type class parser, evaluation API
  - **Aggregate** (12): SUM, COUNT, COUNTA, COUNTBLANK, AVERAGE, MEDIAN, MIN, MAX, STDEV, STDEVP, VAR, VARP
  - **Statistical** (5): LARGE, SMALL, RANK, PERCENTILE, QUARTILE
  - **Conditional** (9): SUMIF, COUNTIF, SUMIFS, COUNTIFS, AVERAGEIF, AVERAGEIFS, MAXIFS, MINIFS, SUMPRODUCT
  - **Logical / Selection** (13): IF, IFS, IFERROR, SWITCH, CHOOSE, AND, OR, NOT, ISNUMBER, ISTEXT, ISBLANK, ISERR, ISERROR
  - **Text** (12): CONCATENATE, LEFT, RIGHT, MID, LEN, UPPER, LOWER, TRIM, FIND, SUBSTITUTE, TEXT, VALUE
  - **Date** (12): TODAY, NOW, DATE, YEAR, MONTH, DAY, EOMONTH, EDATE, DATEDIF, NETWORKDAYS, WORKDAY, YEARFRAC
  - **Math** (16): ABS, ROUND, ROUNDUP, ROUNDDOWN, INT, MOD, POWER, SQRT, LOG, LN, EXP, FLOOR, CEILING, TRUNC, SIGN, PI
  - **Financial** (9): NPV, IRR, XNPV, XIRR, PMT, FV, PV, RATE, NPER
  - **Lookup / Reference** (12): VLOOKUP, HLOOKUP, XLOOKUP, INDEX, MATCH, OFFSET, INDIRECT, ROW, COLUMN, ROWS, COLUMNS, ADDRESS
  - **Dynamic Arrays** (5): TRANSPOSE, SEQUENCE, SORT, UNIQUE, FILTER
  - FunctionSpec registry: macro-collected specs with extensible registry
  - APIs: sheet.evaluateFormula(), sheet.evaluateCell(), sheet.evaluateAllFormulas()
  - Clock trait for pure date/time functions (deterministic testing)
- ✅ **Dependency Graph** (WI-09d): Circular reference detection + topological sort, 52 tests
  - Tarjan's SCC algorithm: O(V+E) cycle detection with early exit
  - Kahn's algorithm: O(V+E) topological sort for correct evaluation order
  - Precedent/dependent queries: O(1) lookups via adjacency lists
  - Safe evaluation: sheet.evaluateWithDependencyCheck() (production-ready)
  - Performance: Handles 10k formula cells in <10ms
- ⚠️ Merged cells are supported by the in-memory OOXML path and `writeWorkbookStream`. Pure row-stream generation (`writeStream` / `writeStreamsSeq`) has no merge API.
- ✅ Hyperlinks serialized as of 0.10.0 (`Cell.hyperlink` → `<hyperlinks>` + worksheet relationships; populated on read).
- ✅ Column/row properties (width, height, hidden, outlineLevel, collapsed) are fully serialized via DirectSaxEmitter.

### Style System

**Minor Limitations**:
- ✅ Theme colors resolved via `Color.toResolvedArgb(theme)` / `toResolvedHex(theme)` (slot lookup + tint application through `ThemePalette.resolve`)
- ⚠️  StyleRegistry requires explicit initialization per sheet (design choice for purity)

### OOXML Coverage

**Missing Parts** (not critical for MVP):
- ❌ docProps/core.xml, docProps/app.xml (metadata)
- ⚠️ xl/theme/theme1.xml (theme palette) — preserved from source on round-trip, not generated for new workbooks
- ❌ xl/calcChain.xml (formula calculation order)
- ✅ Worksheet relationships (`_rels/sheetN.xml.rels`) — written when a sheet has comments, tables, or hyperlinks
- ⚠️ Print settings, page setup — partial as of 0.11.0: header/footer (odd), margins, print area, repeat rows (#259); even/first headers + fitToPage tracked in #266
- ❌ Conditional formatting
- ❌ Data validation (preserved through edits, but no authoring API yet)
- ✅ Named ranges (authoring shipped in 0.10.0: `DefinedName` serialization + CLI `name add/rm`)

### Streaming I/O Limitations

**Row-stream write path** (✅ Working):
- ✅ True constant-memory row streaming with `writeStream` / `writeStreamsSeq`
- ✅ O(1) memory regardless of file size
- ⚠️  No SST support (inline strings only - larger files)
- ⚠️  Minimal styles (default only - no rich formatting)
- ⚠️  No row-stream API for workbook metadata such as merged ranges, comments, tables, and freeze panes

**In-memory workbook SAX/StAX write path** (✅ Working):
- ✅ `writeWorkbookStream` writes an already-materialized `Workbook` through the SAX/StAX backend
- ✅ Preserves full workbook metadata handled by the OOXML writer, including merges, comments, tables, row/column properties, and freeze panes
- ⚠️  Not a row-input streaming API; the `Workbook` is already in memory

**Read Path** (✅ P6.6 Complete):
- ✅ **True constant-memory streaming** - uses `fs2.io.readInputStream`
- ✅ O(1) memory for worksheet data (unlimited rows supported)
- ✅ Streams worksheet XML incrementally (4KB chunks)
- ⚠️  SharedStrings Table (SST) materialized in memory (~10MB typical, scales with unique strings)
- ✅ Large files (500k+ rows) process without OOM
- ✅ Memory tests verify O(1) behavior

**Result**:
- Both streaming **read and write** achieve constant memory for worksheet data ✅
- 500k rows: ~10-20MB memory (worksheet streaming + SST materialized)
- 1M+ rows supported without memory issues (unless >100k unique strings)
- **Design tradeoff**: SST materialization acceptable for most use cases (text typically <10MB)

### Security & Safety

**Implemented**:
- ✅ ZIP bomb detection
- ✅ XXE (XML External Entity) prevention
- ✅ Formula injection guards in in-memory and streaming writes

**Remaining**:
- ❌ XLSM macro preservation policy and tests (macros are never executed)

**Implemented (continued)**:
- ✅ Configurable file size limits via CLI `--max-size <MB>` (default 100MB; `--max-size 0` = unlimited)

### Advanced Features

**Completed** (P6, P7, P8, P31, WI-07/08/09, WI-10, WI-15, WI-17):
- ✅ P6: CellCodec primitives (9 types with auto-formatting)
- ✅ P7: String interpolation Phase 1 (runtime validation for all macros)
- ✅ P8: String interpolation Phase 2 (compile-time optimization)
- ✅ P31: Optics, RichText, HTML export, enhanced ergonomics
- ✅ **Formula System** (WI-07/08/09): Parser, evaluator, 105 functions, dependency graph, cycle detection
- ✅ **Excel Tables** (WI-10): Structured data with headers, AutoFilter, styling
- ✅ **Benchmarks** (WI-15): JMH performance suite (XL vs POI)
- ✅ **SAX Write** (WI-17): Fast SAX/StAX streaming write path
- ✅ **Security Hardening** (WI-30): ZIP bomb detection, XXE prevention, formula injection guards

**Not Started** (Future):
- ❌ P6b: Full case class codec derivation (Magnolia/Shapeless)
- ❌ P9: Advanced macros (path macro, style literal)
- ❌ P10: Drawings (images, shapes)
- ❌ P11: Charts
- ❌ Pivot Tables (remaining part of P12)

---

## Next Steps

> **For detailed roadmap and future plans, see [plan/roadmap.md](plan/roadmap.md)**

### Priority 1: P6.5 - Performance & Quality Polish

**Focus**: Address PR review feedback
- Optimize style indexOf from O(n²) to O(1)
- Extract whitespace check utilities
- Add error path tests
- Full round-trip integration tests

### Priority 2: P6b - Full Codec Derivation

**Focus**: Automatic case class mapping
- Derive RowCodec[A] for case classes
- Header-based column binding
- Type-safe bulk operations

### Priority 3: P9 - Advanced Macros

**Focus**: Additional compile-time validation
- `path` macro for file path validation
- `style` literal for CellStyle DSLs
- Enhanced diagnostics

### Priority 4: P10-P13 - Advanced Features

**Focus**: Drawings, Charts, Tables, Security
- See [plan/roadmap.md](plan/roadmap.md) for detailed breakdown

---

## File Structure

### Completed Modules
```
xl/src/com/tjclp/xl/
└── scripting.scala        ✅ One-import scripting prelude (com.tjclp.xl.scripting)

xl-core/src/com/tjclp/xl/
├── addressing/            ✅ Opaque types (Column, Row, ARef packing), CellRange, SheetName
├── cells/                 ✅ Cell, CellValue, CellError, Comment
├── sheets/                ✅ Sheet, SheetView, PageSetup
├── workbooks/             ✅ Workbook, WorkbookMetadata, DefinedName
├── patch/                 ✅ Patch Monoid (incl. MergeBorder)
├── styles/                ✅ CellStyle, Font, Fill, Border, Color, NumFmt, StylePatch, style DSL
├── codec/                 ✅ CellCodec (9 primitive types), readTyped/readTypedOr/readTypedOpt
├── macros/                ✅ ref"", fx"", money"" … compile-time literals (lives in xl-core; no separate xl-macros module)
├── optics/                ✅ Lens, Optional, focus DSL
├── richtext/              ✅ TextRun, RichText, DSL extensions
├── formatted/             ✅ Formatted literals + FormattedParsers.detect
├── render/                ✅ HTML/SVG renderers (sheet.toHtml / toSvg)
└── dsl/                   ✅ Ergonomic patch operators (:=, ++, outlined)

xl-ooxml/src/com/tjclp/xl/ooxml/
├── ContentTypes.scala     ✅ [Content_Types].xml
├── Relationships.scala    ✅ .rels files
├── Workbook.scala         ✅ xl/workbook.xml
├── worksheet/             ✅ xl/worksheets/sheet#.xml (RichText, merges, hyperlinks, page setup)
├── SharedStrings.scala    ✅ xl/sharedStrings.xml (SST with RichText)
├── Styles.scala           ✅ xl/styles.xml
├── XlsxWriter.scala       ✅ ZIP assembly + surgical modification
└── XlsxReader.scala       ✅ ZIP parsing + security limits

xl-cats-effect/src/com/tjclp/xl/io/
├── Excel.scala            ✅ Algebra trait
├── ExcelIO.scala          ✅ Interpreter with true streaming
└── Sax/StAX + fs2 writers ✅ Event-based streaming read/write
```

### Completed Modules (Additional)
- `xl-evaluator/` ✅ **Complete** (WI-07/08/09 - formula parsing, evaluation, 105 functions, dependency graph, structural editing, recalculation)
- `xl-cli/` ✅ **Complete** (stateless `xl` CLI: 40 subcommands, 21 batch ops, rendering, streaming mode)
- `xl-agent/` ✅ **Complete** (AI agent benchmark runner)
- `xl-benchmarks/` ✅ **Complete** (WI-15 - JMH performance benchmarks)

### Not Started (Future Phases)
- `xl-testkit/` (law helpers, golden test framework)
- `xl-drawings/` (P8 - images, shapes)
- `xl-charts/` (P9 - chart generation)

---

## Technical Debt

### Completed ✅
1. ~~StreamingXmlWriter compilation~~ - ✅ fs2-data-xml integration complete
2. ~~DateTime serialization~~ - ✅ Excel serial number conversion implemented
3. ~~Cell → CellStyle linkage~~ - ✅ StyleRegistry provides sheet-level style management
4. ~~Comments~~ - ✅ Full OOXML round-trip (xl/commentsN.xml + VML drawings), rich text support, 12+ tests
5. ~~Merged cells~~ - ✅ `mergedRanges` serialized via `<mergeCells>` (in-memory + `writeWorkbookStream` paths)
6. ~~Column/row properties~~ - ✅ width/height/hidden/outline serialized via DirectSaxEmitter
7. ~~Hyperlinks~~ - ✅ serialized with worksheet relationships (0.10.0)

### Remaining
1. **Theme resolution** - Improve Theme color ARGB approximations (currently functional but not perfect)

---

## Performance Results (Actual)

### JMH Benchmark Results (WI-15) - XL vs Apache POI

**XL vs Apache POI** (Apple Silicon M-series, JDK 25):

#### Streaming Reads (SAX Parser - Production Recommendation)
| Rows | POI | XL | Result |
|------|-----|----|--------|
| **1,000** | 1.357 ± 0.076 ms | **0.887 ± 0.060 ms** | ✨ **XL 35% faster** |
| **10,000** | 7.773 ± 0.590 ms | 8.408 ± 0.153 ms | Competitive (XL within 8%) |

#### In-Memory Reads (For Modification Workflows)
| Rows | POI | XL | Result |
|------|-----|----|--------|
| **1,000** | 1.650 ± 0.055 ms | **1.225 ± 0.086 ms** | ✨ **XL 26% faster** |
| **10,000** | 13.784 ± 0.377 ms | 14.115 ± 1.250 ms | Competitive (XL within 2%) |

#### Writes
| Rows | POI | XL | Result |
|------|-----|----|--------|
| **1,000** | 1.280 ± 0.041 ms | 1.906 ± 0.245 ms | POI 49% faster |
| **10,000** | 10.228 ± 0.417 ms | 15.248 ± 1.315 ms | POI 49% faster |

**Key Findings**:
- ✨ **XL is fastest for small-medium files** (< 5k rows): 35% faster streaming, 26% faster in-memory
- ✅ **XL competitive on large files**: Within 8% of POI on 10k row streaming reads
- 🔧 **Write optimization**: Future work (Phase 3) - POI currently 49% faster
- 💾 **Constant memory**: Streaming uses O(1) memory regardless of file size
- ⚡ **SAX parser**: 3.8x speedup vs previous fs2-data-xml implementation

**Recommendation**: Use `ExcelIO.readStream()` for production workloads (fastest for <5k rows, constant memory).

### Streaming Implementation (P5 + P6.6 Complete) ✅

**Memory characteristics**:
- **Write**: **~10MB constant memory** (O(1)) ✅
- **Read**: **~10MB constant memory** (O(1)) ✅ (P6.6 fixed with fs2.io.readInputStream)
- **Scalability**: Can handle 1M+ rows without OOM ✅

**Performance characteristics** (validated with JMH):
- **Streaming vs in-memory**: Streaming 1.7x faster for large files (8.4ms streaming vs 14.1ms in-memory @ 10k rows)
- **XL vs POI**: XL is fastest for small files (35% faster @ 1k rows), competitive for large files (within 8% @ 10k rows)

### Comparison to Apache POI (True Streaming Read + Write)

| Operation | XL | Apache POI SXSSF | Improvement |
|-----------|-----|------------------|-------------|
| **Write 100k** | 1.1s @ 10MB | ~5s @ 800MB | **4.5x faster, 80x less memory** ✅ |
| **Write 1M** | ~11s @ 10MB | ~50s @ 800MB | **4.5x faster, constant memory** ✅ |
| **Read 100k** | 1.8s @ 10MB | ~8s @ 1GB | **4.4x faster, 100x less memory** ✅ |
| **Read 500k** | ~9s @ 10MB | OOM @ 1GB+ | **Constant memory vs OOM** ✅ |

**Note**: All performance claims now verified for both reads **and** writes. True O(1) streaming achieved.

**Result for Writes**: Exceeded goal of 3-5x throughput, achieved 80x memory improvement with constant memory.

---

## Commands for Next Session

```bash
# Quick start
./mill __.compile
./mill __.test

# Work on streaming
./mill xl-cats-effect.compile
./mill xl-cats-effect.test

# Format code
./mill __.reformat

# Create sample file (once fixed)
./mill xl-ooxml.test.runMain com.tjclp.xl.ooxml.Demo
```

---

## Critical Success Factors

1. **Purity maintained** - Core is 100% pure, zero side effects
2. **Laws verified** - All Monoids tested with property-based tests
3. **Deterministic output** - Same input = same bytes (stable diffs)
4. **Zero overhead** - Opaque types, inline, compile-time macros
5. **Real files** - Creates valid XLSX that Excel/LibreOffice opens
6. **Type safety** - Opaque types prevent mixing units; codecs enforce type correctness
7. **Performance** - ~35% faster than Apache POI on streaming reads (JMH validated), writes currently ~49% slower, 80x less memory
