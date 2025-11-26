# Changelog

All notable changes to the XL project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added - High-Fidelity Rendering

**Theme Color Resolution** (`xl-ooxml/src/com/tjclp/xl/ooxml/ThemeParser.scala`)
- Parse OOXML theme XML for accurate color rendering in exports
- Resolve theme color indices (Dark1, Light1, Accent1-6) to actual hex values
- Fix ECMA-376 spec compliance: theme="0" → Light1, theme="1" → Dark1

**SVG Renderer** (`xl-core/src/com/tjclp/xl/render/SvgRenderer.scala`)
- Two-pass rendering: backgrounds first, then text layer (prevents overlap)
- All 14 Excel border styles with per-side rendering
- Optional gridlines (default: off)
- Row/column sizing from sheet metadata
- Text overflow into adjacent empty cells
- Rich text with proper font size conversion (pt to px)
- Hidden row/column filtering

**HTML Renderer** (`xl-core/src/com/tjclp/xl/render/HtmlRenderer.scala`)
- `HAlign.General` for content-aware alignment (numbers right, text left)
- Text overflow into adjacent empty cells
- Print scaling support for PDF-matching output
- Row/column sizing from sheet metadata

**Shared Utilities** (`xl-core/src/com/tjclp/xl/render/RenderUtils.scala`)
- Extracted common rendering logic (color resolution, sizing, overflow)
- `showLabels` parameter for row/column headers

### Added - Row/Column Properties

**Core Properties** (`xl-core/src/com/tjclp/xl/sheets/`)
- `RowProperties(height, hidden, outlineLevel, collapsed)`
- `ColumnProperties(width, hidden, outlineLevel, collapsed)`
- `col"A"` compile-time column macro

**Builder DSL** (`xl-core/src/com/tjclp/xl/dsl/RowColumnDsl.scala`)
- Fluent builders: `row(1).height(20).hidden`, `column("A").width(15)`
- Integration with Sheet API

**OOXML Serialization** (`xl-ooxml/src/com/tjclp/xl/ooxml/`)
- `<cols>` element serialization for column properties
- `<row>` attributes for height, hidden, outline, collapsed

### Added - Compile-Time API Validation

**Macros** (`xl-core/src/com/tjclp/xl/macros/`)
- `PutLiteral.scala`: Compile-time validation for `sheet.put("A1", value)`
- `SheetLiteral.scala`: Compile-time validation for Sheet/Workbook construction
- String-ref variants: `put`, `style`, `merge` with compile-time ref parsing

### Changed - API & Rendering

- `Sheet.put(ref, value)` and similar methods now support string refs with compile-time validation
- Removed unnecessary `XLResult` wrappers from infallible operations (cleaner API)
- Renderers moved from `html/` and `svg/` to unified `render/` package

### Fixed

- `AVERAGE` formula returning `Text("(a, b)")` tuple instead of `Number(mean)`
- Theme color index mapping (ECMA-376 spec compliance)
- Font size conversion in SVG (Excel pt to CSS px)
- Rich text rendering with proper text wrapping

### Breaking Changes - API Simplification

**Compile-time safe migration**:
- `sheet.put(ref"A1", value)` now returns `Sheet` directly (was `XLResult[Sheet]`)
- Most `.unsafe` calls no longer needed for infallible operations
- Migration: Remove `.unsafe` calls that now compile without them

### Internal Changes

- `HtmlRenderer` and `SvgRenderer` moved to `xl-core/src/com/tjclp/xl/render/`
- New `RenderUtils` for shared rendering logic
- `ThemeParser` added to xl-ooxml for theme color extraction

### Testing

- 698 tests (was ~660)
- New regression tests for font rendering
- SVG renderer tests: 754+ lines of coverage
- HTML renderer tests: 386+ lines of coverage
- Row/column operations tests: 386 lines

---

### Added - P6: Cell-level Codecs

**Type-Safe Cell Operations** (`xl-core/src/com/tjclp/xl/codec/`)
- Cell-level bidirectional codecs for 9 primitive types: String, Int, Long, Double, BigDecimal, Boolean, LocalDate, LocalDateTime, RichText
- `putMixed` batch update API with auto-inferred formatting
- `readTyped[A]` for type-safe cell reading with Either[CodecError, Option[A]]
- Auto-format inference: LocalDate → NumFmt.Date, LocalDateTime → NumFmt.DateTime, BigDecimal → NumFmt.Decimal
- Identity laws: `codec.read(Cell(ref, codec.write(value)._1)) == Right(Some(value))`

**RichText DSL** (`xl-core/src/com/tjclp/xl/richtext.scala`)
- Multi-run formatted text within single cells (OOXML `<r>` elements)
- Composable with `+` operator: `"Bold".bold.red + " normal " + "Italic".italic.blue`
- Formatting methods: `.bold`, `.italic`, `.underline`, `.size(pt)`, `.fontFamily(name)`
- Color shortcuts: `.red`, `.green`, `.blue`, `.black`, `.white`, `.withColor(color)`
- Given conversions: `String → TextRun`, `TextRun → RichText`
- Full OOXML round-trip support with `xml:space="preserve"`

**HTML Export** (`xl-core/src/com/tjclp/xl/html/`)
- `sheet.toHtml(range)` - Export cell ranges to HTML tables
- Inline CSS preservation of cell styles and rich text formatting
- Optional style inclusion: `toHtml(range, includeStyles = false)`
- Use cases: Dashboards, reporting, email generation

### Added - P31: Ergonomics & Purity Enhancements

**Type Safety** (Phase 1)
- `StyleId` opaque type replacing raw Int for style indices
- Zero runtime overhead with compile-time safety
- Prevents accidental mixing of indices across different domains

**Patch DSL** (`xl-core/src/com/tjclp/xl/dsl.scala` - Phase 2)
- `:=` operator for cell assignment with automatic primitive conversions
- `++` operator for patch composition (alternative to Cats `|+|`)
- CellRange operators: `.merge`, `.unmerge`, `.remove` (returns Patch)
- `PatchBatch` object for varargs patch construction
- ARef styling methods: `.styled(style)`, `.styleId(id)`, `.clearStyle`
- Usage: `import com.tjclp.xl.dsl.*`

**Range Combinators** (`xl-core/src/com/tjclp/xl/SheetExtensions.scala` - Phase 2)
- `fillBy(range)(f: (Column, Row) => CellValue)` - Fill using Excel coordinates
- `tabulate(range)(f: (Int, Int) => CellValue)` - Fill using 0-based indices
- `putRow(row, startCol, values)` - Horizontal value placement with strict bounds (XLResult)
- `putCol(col, startRow, values)` - Vertical value placement with strict bounds (XLResult)
- `putRange(range, values)` - Row-major placement with value-count validation
- Row-major deterministic iteration for all combinators

**Deterministic Iteration** (`xl-core/src/com/tjclp/xl/SheetExtensions.scala` - Phase 2)
- `cellsSorted: Vector[Cell]` - Row-major sorted cells
- `rowsSorted: Vector[(Row, Vector[Cell])]` - Cells grouped by row
- `columnsSorted: Vector[(Column, Vector[Cell])]` - Cells grouped by column

**Formula Literal Macro** (`xl-macros/src/com/tjclp/xl/CellRangeLiterals.scala` - Phase 2)
- `fx"..."` compile-time validated formula literals
- Minimal validation (balanced parentheses, no interpolation)
- Emits `CellValue.Formula(...)` directly

**Pure Error Channels** (`xl-cats-effect/src/com/tjclp/xl/io/` - Phase 3)
- `ExcelR[F]` trait for explicit error handling without exceptions
- All methods return `F[XLResult[A]]` or `Stream[F, Either[XLError, A]]`
- Methods: `readR`, `writeR`, `readStreamR`, `writeStreamR`, etc.
- Coexists with `Excel[F]` for gradual migration
- No breaking changes to existing API

**Configurable Writer** (`xl-ooxml/src/com/tjclp/xl/ooxml/XlsxWriter.scala` - Phase 3)
- `SstPolicy` enum: Auto (default), Always, Never
- `WriterConfig(sstPolicy, prettyPrint)` for power users
- `writeWith(workbook, path, config)` method
- `XmlUtil.compact()` for minimal XML output (~30% smaller files)
- Backward compatible: `write()` delegates to `writeWith()` with defaults

**Optics Library** (`xl-core/src/com/tjclp/xl/optics.scala` - Phase 4)
- `Lens[S, A]` - Total optics with get, set, modify, update, andThen
- `Optional[S, A]` - Partial optics with getOption, modify, getOrElse
- Predefined optics: `sheetCells`, `cellAt`, `cellValue`, `cellStyleId`, `sheetName`, `mergedRanges`
- Focus DSL: `sheet.focus(ref)`, `sheet.modifyCell(ref)(f)`, `sheet.modifyValue(ref)(f)`, `sheet.modifyStyleId(ref)(f)`
- Law-governed: get-put, put-get, put-put
- Zero external dependencies (no Monocle)

### Changed
- `Sheet.putRange(range, values)` now validates the supplied value count and returns `XLResult[Sheet]`, emitting `XLError.ValueCountMismatch` when the count does not match the range size.
- `Sheet.putRow` and `Sheet.putCol` now return `XLResult[Sheet]`, enforce Excel's column/row limits, and buffer the supplied values to provide deterministic size semantics.

### Breaking Changes
- `Sheet.putRange` return type changed from `Sheet` to `XLResult[Sheet]`; callers must handle the result explicitly.
- `Sheet.putRow` and `Sheet.putCol` now return `XLResult[Sheet]` instead of `Sheet`.

### Internal Changes
- StyleId opaque type used internally (transparent to users at API boundary)
- StyleRegistry now uses StyleId instead of raw Int
- OOXML layer performs StyleId ↔ Int conversions at boundaries
- `xl-macros` macros split into focused sources (`CellRangeLiterals.scala`, `BatchPutMacro.scala`, `FormattedLiterals.scala`) for maintainability.
- `Sheet` extensions and auxiliary data classes moved to `SheetProperties.scala` and `SheetExtensions.scala` to keep core data structures small.

### Testing
- 263 tests (229 original + 34 optics tests)
- All passing with property-based tests for laws
- Optics: 3 lens laws + 5 optional behaviors + 26 usage tests

### Added - P6.8: Surgical Modification & Unified Write Path

**Complete**: 2025-11-18 (PR #16)

**Surgical Modification Infrastructure** (`xl-core/src/com/tjclp/xl/`, `xl-ooxml/src/com/tjclp/xl/ooxml/`)
- **SourceContext**: Tracks source file metadata (path, SHA-256 fingerprint, PartManifest, ModificationTracker)
- **ModificationTracker**: Monoid-based dirty tracking for sheets (marks modified, deleted, reordered)
- **PartManifest**: Distinguishes parsed vs unparsed ZIP entries for byte-perfect preservation
- **PreservedPartStore**: Lazy streaming access to preserved parts with Resource-safe cleanup
- **RelationshipGraph**: Tracks OOXML part relationships to prevent broken rId references

**Hybrid Write Strategy** (`xl-ooxml/src/com/tjclp/xl/ooxml/XlsxWriter.scala`)
- **Verbatim copy optimization**: 11x speedup for unmodified workbooks (byte-for-byte with fingerprint validation)
- **Surgical write**: Only regenerate modified sheets, copy rest byte-perfect (2-5x speedup)
- **Intelligent SST handling**: Preserve original SST unless new strings introduced (prevents inline string corruption)
- **Full regeneration fallback**: No SourceContext → standard write path (backwards compatible)

**Unified Style Indexing** (`xl-ooxml/src/com/tjclp/xl/ooxml/Styles.scala`)
- **Single public API**: `StyleIndex.fromWorkbook(wb)` auto-detects source context
- **Automatic optimization**: Preserves styles for surgical writes, full deduplication otherwise
- **Transparent to users**: Existing code works unchanged, optimization automatic

**Security & Correctness Fixes**
- **XXE hardening**: Migrated RichText XML parsing to `XmlSecurity.parseSafe` (prevents XXE attacks)
  - Fixed: `SharedStrings.scala` (SST richtext rPr parsing)
  - Fixed: `Worksheet.scala` (worksheet richtext rPr parsing)
- **SST normalization fix**: Normalize strings consistently in both modified/preserved sets (prevents false "new string" detection)
- **Whitespace preservation**: Added `getTextPreservingWhitespace()` helper for exact whitespace round-trips
  - Partial fix: Reading preserves whitespace correctly
  - Known issue: Writing still has edge cases (tracked in Issue #17)

**Test Coverage**: 44 surgical modification tests (all passing)
- Core surgical write behavior (6 tests)
- Reader passthrough (5 tests)
- ModificationTracker (19 tests: 6 basic + 13 property-based Monoid laws)
- Workbook modification tracking (6 tests)
- PartManifest & RelationshipGraph (8 tests)
- Corruption regressions (4 tests)
- Real-world file preservation (1 test)

**Performance**
- Unmodified workbooks: 11x faster (verbatim copy)
- Partially modified: 2-5x faster (surgical write)
- Memory: 30-50MB savings (lazy part loading)
