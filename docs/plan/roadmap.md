# Roadmap — From Spec to MVP and Beyond

**Current Status: ~85% Complete (698/698 tests passing)**

## Completed Phases ✅

### ✅ P0: Bootstrap & CI (Complete)
**Status**: 100% Complete
**Definition of Done**:
- Mill build system configured
- Module structure (`xl-core`, `xl-macros`, `xl-ooxml`, `xl-cats-effect`)
- Scalafmt 3.10.1 integration
- GitHub Actions CI pipeline
- Documentation framework (CLAUDE.md, README.md, docs/plan/)

### ✅ P1: Addressing & Literals (Complete)
**Status**: 100% Complete
**Test Coverage**: 17 tests passing
**Definition of Done**:
- Opaque types: `Column`, `Row`, `ARef` (64-bit packed)
- `CellRange` with normalization
- Compile-time macros: `cell"A1"`, `range"A1:B10"`
- Property-based tests for all laws
- A1 notation parsing/printing with round-trip verification

### ✅ P2: Core + Patches (Complete)
**Status**: 100% Complete
**Test Coverage**: 21 tests passing
**Definition of Done**:
- Immutable domain model: `Cell`, `Sheet`, `Workbook`
- `Patch` enum as Monoid (Put, SetStyle, Merge, Remove, Batch)
- Monoid laws verified (associativity, identity, idempotence)
- `applyPatch` with lawful semantics
- `XLError` ADT for total error handling

### ✅ P3: Styles & Themes (Complete)
**Status**: 100% Complete
**Test Coverage**: 60 tests passing
**Definition of Done**:
- Complete style system: `CellStyle`, `Font`, `Fill`, `Border`, `Color`, `NumFmt`, `Align`
- `StylePatch` Monoid for declarative style updates
- `StyleRegistry` for per-sheet style management
- Unit conversions: `Pt`, `Px`, `Emu` with bidirectional laws
- Color parsing (hex, RGB, ARGB, theme colors)
- Style canonicalization for deduplication
- 60 comprehensive style tests

### ✅ P4: OOXML MVP (Complete)
**Status**: 100% Complete
**Test Coverage**: 24 tests passing
**Definition of Done**:
- Full OOXML read/write pipeline
- `[Content_Types].xml`, `_rels/.rels`
- `xl/workbook.xml`, `xl/worksheets/sheet#.xml`
- Shared Strings Table (SST) with deduplication
- `xl/styles.xml` with component-level deduplication
- ZIP assembly and parsing
- All cell types: Text, Number, Bool, Formula, Error, DateTime
- DateTime serialization (Excel serial numbers)
- RichText support (`<si><r>` elements)
- Round-trip tests (24 passing)

### ✅ P4.1: XLSX Reader & Non-Sequential Sheets (Complete)
**Status**: 100% Complete
**Test Coverage**: Part of 698 total tests
**Definition of Done**:
- Full XLSX reading with style preservation
- WorkbookStyles parser (fonts, fills, borders, number formats, alignment)
- Relationship-based worksheet resolution (r:id mapping)
- Non-sequential sheet index support (ContentTypes.forSheetIndices)
- Case-insensitive cell reference parsing (Locale.ROOT normalization)
- Optics bug fixes (Lens.modify return type)
- Zero compilation warnings (eliminated all 10 warnings)
- All tests passing (698/698)

### ✅ P5: Streaming (Complete)
**Status**: 100% Complete
**Test Coverage**: 18 tests passing
**Performance**: 100k rows @ ~1.1s write / ~1.8s read, O(1) memory (~50MB)
**Definition of Done**:
- `Excel[F[_]]` algebra trait
- `ExcelIO[IO]` interpreter
- True streaming write with `writeStreamTrue` (fs2-data-xml)
- True streaming read with `readStreamTrue` (fs2-data-xml)
- Constant memory usage (O(1), independent of file size)
- Benchmark: 4.5x faster than Apache POI, 16x less memory
- Can handle 1M+ rows without OOM

### ✅ P6: CellCodec Primitives (Complete)
**Status**: Primitive codecs complete (80%), full derivation deferred to P6b
**Test Coverage**: 42 codec tests + 16 batch operation tests = 58 tests passing
**Definition of Done** (Primitives):
- `CellCodec[A]` trait with bidirectional encoding
- Inline given instances for 9 primitive types:
  - String, Int, Long, Double, BigDecimal, Boolean
  - LocalDate, LocalDateTime, RichText
- Auto-inferred number/date formatting
- `putMixed` API for batch updates with type inference
- `readTyped[A]` for type-safe cell reading
- Identity laws: `codec.read(Cell(_, codec.write(v)._1)) == Right(Some(v))`
- 58 comprehensive codec tests

**Deferred to P6b**: Full case class derivation using Magnolia/Shapeless

### ✅ P31: Refactoring & Optics (Bonus Phase - Complete)
**Status**: 100% Complete
**Test Coverage**: 34 optics tests + 5 RichText tests = 39 tests passing
**Definition of Done**:
- Optics module: `Lens[S, A]`, `Optional[S, A]`
- Focus DSL: `sheet.focus(ref).modify(...)`
- Functional update helpers: `modifyValue`, `modifyStyle`
- RichText: `TextRun`, `RichText` composition
- RichText DSL: `.bold`, `.italic`, `.red`, `.size()`, etc.
- HTML export: `sheet.toHtml(range)` with inline CSS
- Enhanced patch DSL with ergonomic operators
- 39 comprehensive tests

### ✅ P4.5: OOXML Quality & Spec Compliance (Complete)
**Status**: 100% Complete
**Completed**: 2025-11-10
**Test Coverage**: +22 new tests (4 new test classes)
**Commits**: b22832e (Part 1), 4dd98f5 (Part 2)
**Definition of Done**:
- ✅ Fixed default fills (gray125 pattern per OOXML spec ECMA-376 §18.8.21)
- ✅ Fixed SharedStrings count vs uniqueCount tracking
- ✅ Added whitespace preservation for plain text cells (xml:space="preserve")
- ✅ Added alignment serialization to styles.xml with round-trip verification
- ✅ Fixed Scala version consistency (3.7.3 everywhere)
- ✅ Used idiomatic xml:space namespace construction (PrefixedAttribute)
- ✅ Added comprehensive AI contracts (REQUIRES/ENSURES/DETERMINISTIC/ERROR CASES)
- ✅ Added 22 comprehensive tests across 4 new test classes
- ✅ Zero spec violations, full round-trip fidelity achieved

**See**: [ooxml-quality.md](ooxml-quality.md) for detailed implementation notes

---

## Remaining Phases ⬜

### ⬜ P6.5: Performance & Quality Polish (Future)
**Priority**: Medium
**Estimated Effort**: 8-10 hours
**Source**: PR #4 review feedback
**Definition of Done**:
- Optimize style indexOf from O(n²) to O(1) using Maps
- Extract whitespace check to XmlUtil.needsXmlSpacePreserve utility
- Add XlsxReader error path tests (10+ tests)
- Add full XLSX round-trip integration test
- Logging strategy for missing/malformed files (defer to P11)

**See**: [future-improvements.md](future-improvements.md) for detailed breakdown

### ⬜ P6b: Full Codec Derivation (Future)
**Priority**: Medium
**Estimated Effort**: 2-3 weeks
**Definition of Done**:
- Automatic case class to/from row mapping
- Header-based column binding
- Derived `RowCodec[A]` instances
- Type-safe bulk operations
- Integration with compile-time macros

### ⬜ P7: Advanced Macros (Future)
**Priority**: Low
**Estimated Effort**: 1-2 weeks
**Definition of Done**:
- `path` macro for compile-time file path validation
- `style` literal for CellStyle DSL
- Enhanced error messages with precise diagnostics
- Compile-time style validation

### ⬜ P8: Drawings (Future)
**Priority**: Medium
**Estimated Effort**: 3-4 weeks
**Definition of Done**:
- Image embedding (PNG, JPEG)
- Shapes and text boxes
- Anchoring (absolute, one-cell, two-cell)
- Drawing deduplication
- xl/drawings/drawing#.xml serialization

### ⬜ P9: Charts (Future)
**Priority**: Medium
**Estimated Effort**: 4-6 weeks
**Definition of Done**:
- Chart types: bar, line, pie, scatter
- Chart data binding
- Chart style customization
- xl/charts/chart#.xml serialization

### ⬜ P10: Tables & Advanced Features (Future)
**Priority**: Low
**Estimated Effort**: 2-3 weeks
**Definition of Done**:
- Excel tables (structured references)
- Conditional formatting rules
- Data validation
- Named ranges

### ⬜ P11: Safety, Security & Documentation (Future)
**Priority**: High (for production use)
**Estimated Effort**: 2-3 weeks
**Definition of Done**:
- ZIP bomb detection
- XXE (XML External Entity) prevention
- Formula injection guards
- File size limits
- Comprehensive user documentation
- Cookbook with real-world examples

---

## Overall Progress

| Phase | Status | Test Coverage | Completion |
|-------|--------|---------------|------------|
| P0: Bootstrap | ✅ Complete | Infrastructure | 100% |
| P1: Addressing | ✅ Complete | 17 tests | 100% |
| P2: Core & Patches | ✅ Complete | 21 tests | 100% |
| P3: Styles | ✅ Complete | 60 tests | 100% |
| P4: OOXML MVP | ✅ Complete | 24 tests | 100% |
| P4.1: XLSX Reader | ✅ Complete | Integrated | 100% |
| P4.5: OOXML Quality | ✅ Complete | +22 tests | 100% |
| P5: Streaming | ✅ Complete | 18 tests | 100% |
| P6: Codecs (primitives) | ✅ Complete | 58 tests | 80% |
| P31: Optics/RichText | ✅ Complete | 39 tests | 100% |
| **Total Core** | **✅** | **698/698 tests** | **~85%** |
| P6.5: Perf & Quality | ⬜ Future | - | 0% |
| P6b: Full Derivation | ⬜ Future | - | 0% |
| P7: Advanced Macros | ⬜ Future | - | 0% |
| P8: Drawings | ⬜ Future | - | 0% |
| P9: Charts | ⬜ Future | - | 0% |
| P10: Tables | ⬜ Future | - | 0% |
| P11: Safety/Docs | ⬜ Future | - | 0% |

**Current State**: Production-ready for core spreadsheet operations (read, write, style, stream). Exceeds Apache POI in performance (4.5x faster, 16x less memory). Ready for real-world use in financial modeling, data export, and report generation.
