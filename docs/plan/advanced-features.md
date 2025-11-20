# Advanced Features Plan (P10-P12 + Benchmarks)

**Status**: ⬜ Not Started (Future Work)
**Last Updated**: 2025-11-20
**Purpose**: Consolidated plan for advanced Excel features beyond core I/O

> **Note**: This document consolidates planning for drawings, charts, tables/pivots, and performance benchmarks. For implementation scaffolds, see [docs/reference/implementation-scaffolds.md](../reference/implementation-scaffolds.md).

---

## Overview

This plan covers advanced Excel features that build upon the core OOXML foundation (P0-P8 complete). These features are prioritized for Phases 10-12 and include:

1. **Drawings & Images** (P10) - Pictures, shapes, anchors
2. **Charts** (P11) - Data visualization with typed grammar
3. **Tables & Pivot Tables** (P12) - Structured data and aggregations
4. **Performance Benchmarks** (Infrastructure) - JMH testing and POI comparisons

**Dependencies**: Most features depend on Phase 4 (Formula Engine & Dependency Graph) for proper data binding and relationship management.

---

## 1. Drawings & Graphics (P10)

**Status**: ⬜ Not Started
**Estimated Effort**: 10-15 days
**Priority**: Medium
**LOC**: ~800

### Scope

**Anchors**:
- `TwoCell(fromRow, fromCol, fromDx, fromDy) → (toRow, toCol, toDx, toDy)`
- `OneCell(atRow, atCol, dx, dy, widthEmu, heightEmu)`
- **Normalization**: Ensure `(from ≤ to)` lexicographically; clamp negative EMUs to 0

**Pictures**:
- `ImageId = sha256(bytes)` → Deduplicate single media part
- Alt text, rotation degrees, lock aspect ratio modeled as pure fields
- Support: PNG, JPEG formats

**Shapes**:
- `Rect/Ellipse/Line/Polygon` + `Stroke`, `SolidFill | NoFill`
- Pure affine transforms with composition law; identity is no-op
- Text boxes with RichText content

### OOXML Mapping

- Drawings reside in `xl/drawings/drawing#.xml` with anchors:
  - `xdr:twoCellAnchor` for cell-relative positioning
  - `xdr:oneCellAnchor` for fixed positioning
- Media stored under `/xl/media/*.png|*.jpeg`
- Relations via `drawing#.xml.rels` and `workbook.xml.rels`

### Implementation Strategy

1. Define pure AST for drawings (Anchor, Picture, Shape types)
2. Implement XmlCodec[OoxmlDrawing] for serialization
3. Add media deduplication via SHA-256 content hashing
4. Wire into existing relationship graph
5. Add round-trip tests with embedded images

**See**: `docs/reference/implementation-scaffolds.md` Section B.6 for code examples

---

## 2. Charts (P11)

**Status**: ⬜ Not Started
**Estimated Effort**: 20-30 days
**Priority**: Medium
**LOC**: ~1500+

### Scope

**Chart Types**:
- **Marks**: `Column|Bar` (clustered/stacked), `Line` (smooth/markers), `Area` (stacked), `Scatter`, `Pie`
- **Axes**: Category/Value/Time
- **Scales**: Linear/Log10/Time

**Series Specification**:
- `(mark, encoding{x, y, color?, size?}, name?)`
- Multiple series per chart
- Data binding to cell ranges or tables

### Semantics

**Normalization**:
- Series order → by `name` (stable sort)
- Color palette resolved via theme
- Deterministic legend ordering

**Stacking**:
- Grouped by X axis
- Y-values summed
- Gaps handled as zeros or missing per policy

**Legend & Titles**:
- Pure values
- Printed deterministically

### OOXML Mapping

- Charts in `xl/charts/chart#.xml`:
  - `c:barChart|lineChart|pieChart|scatterChart`
  - `c:plotArea` for chart content
  - `c:legend` for legend positioning
  - Axes configuration
- Data ranges emit:
  - `c:numRef` for numeric data
  - `c:strRef` for string data
  - With `f` formulas referencing sheet ranges

### Implementation Strategy

1. Define pure Chart AST (ChartType, Series, Axis, Legend types)
2. Implement typed grammar for chart specifications
3. Add XmlCodec[OoxmlChartPart] for ChartML serialization
4. Wire chart data bindings to formula dependency graph
5. Add comprehensive round-trip tests
6. Support common chart types first (bar, line, pie), then extend

**See**: `docs/reference/implementation-scaffolds.md` Section B.6 for code examples

---

## 3. Tables & Pivot Tables (P12)

**Status**: ⬜ Not Started
**Estimated Effort**: 15-20 days
**Priority**: Low
**LOC**: ~1000

### Scope

**Excel Tables**:
- `TableSpec(range, style, showFilterButtons, totalsRow)`
- Banding computed by parity (rows, columns, or both)
- Independent of content (structural only)
- Structured references for formulas

**Pivot Tables**:
- `PivotSpec(sourceRange or cache, rows: List[Field], cols: List[Field], values: List[Aggregation])`
- **Aggregations**: `Sum/Avg/Count/Min/Max` over numeric fields
- Pivot cache for data materialization
- Refresh on workbook open

**Slicers**:
- `SlicerSpec(field, selectionSet)` → pure filter state
- Serialization to `slicer*.xml` (total)
- Interactions adjust pivot cache query

### OOXML Mapping

**Tables**:
- `xl/tables/table#.xml` for table definitions
- Structured references via `definedNames`
- AutoFilter integration

**Pivots**:
- `xl/pivotTables/pivotTable#.xml` for pivot structure
- `xl/pivotCache/pivotCacheDefinition#.xml` for data cache
- `xl/pivotCache/pivotCacheRecords#.xml` for materialized data

### Implementation Strategy

1. Define pure AST for tables (TableSpec, ColumnDef, TotalsRow)
2. Define pure AST for pivots (PivotSpec, Field, Aggregation)
3. Implement structured reference resolution in formula parser
4. Add pivot cache materialization (can be lazy)
5. Wire into dependency graph for refresh semantics
6. Add comprehensive round-trip tests

**See**: `docs/reference/implementation-scaffolds.md` Section B.5 for code examples

---

## 4. Performance Benchmarks (Infrastructure)

**Status**: ⬜ Not Started
**Estimated Effort**: 1-2 weeks
**Priority**: High (for performance claims validation)

### Scope

**JMH Benchmarks**:
- Parse 100k-row sheet (no styles) → measure allocs/op and throughput
- Write 100k-row sheet with SST hot cache
- Patch application microbench:
  - Single cell update
  - Full row update
  - Full column update
- Style deduplication and canonicalization

**Comparisons**:
- **Apache POI** streaming (SXSSF) baseline for read/write
- Comparable features only (no chart/drawing comparisons until implemented)
- Memory footprint comparisons

**Metrics to Track**:
- **Throughput**: rows/second for read/write operations
- **Latency**: p50/p95/p99 latencies for operations
- **Memory**: Max RSS, heap allocation rates
- **GC**: GC pause counts and durations
- **Overhead**: Pure functional overhead vs imperative (should be zero)

### Infrastructure

**Build Integration**:
- `xl-benchmarks` module with sbt-jmh plugin
- CI integration for regression detection
- Threshold alerts (fail build if >10% regression)

**Baseline Establishment**:
1. Run benchmarks on current codebase (P0-P8 complete)
2. Establish baseline performance targets
3. Track improvements/regressions per phase

**Comparison Matrix**:
```
| Operation          | XL      | Apache POI | Improvement |
|--------------------|---------|------------|-------------|
| Write 100k rows    | ~1.1s   | ~5s        | 4.5x faster |
| Read 100k rows     | ~1.8s   | ~8s        | 4.4x faster |
| Memory (100k rows) | ~50MB   | ~800MB     | 16x less    |
```

### Implementation Strategy

1. Set up `xl-benchmarks` module with JMH dependencies
2. Implement core operation benchmarks (read, write, patch)
3. Add POI comparison benchmarks
4. Create CI job for regression testing
5. Document benchmark results in STATUS.md
6. Add performance regression alerts

**See**: `docs/reference/implementation-scaffolds.md` Section A.3 for JMH code examples

---

## Implementation Dependencies

### Phase Dependencies

**Drawings** (P10):
- ✅ P4 OOXML foundation complete
- ✅ P6.8 Surgical modification (unknown part preservation)
- ⬜ Relationship graph (Phase 2 of strategic plan)

**Charts** (P11):
- ✅ P4 OOXML foundation complete
- ⬜ P10 Drawings (charts often contain embedded shapes)
- ⬜ Phase 4 Formula/Graph (for data binding)
- ⬜ Tables (P12) optional but beneficial

**Tables** (P12):
- ✅ P4 OOXML foundation complete
- ⬜ Phase 2 Named ranges + relationships
- ⬜ Phase 4 Formula engine (for structured references)

**Benchmarks**:
- ✅ Can start immediately
- ✅ Current implementation provides baseline
- Expand as features added

### Recommended Implementation Order

1. **Benchmarks** (parallel with any other work) - establishes performance baseline
2. **Drawings** - foundational for charts, relatively isolated
3. **Tables** - needed for chart data binding, depends on named ranges
4. **Charts** - final integration, depends on drawings + tables + formulas

---

## Success Criteria

**Drawings Complete** when:
- ✅ Can embed PNG/JPEG images with two-cell and one-cell anchors
- ✅ Can create basic shapes (rectangle, ellipse, line)
- ✅ Media deduplication via SHA-256 works correctly
- ✅ Round-trip tests preserve all drawing properties
- ✅ Surgical modification preserves existing drawings

**Charts Complete** when:
- ✅ Can create bar, line, pie, scatter charts
- ✅ Data binding to cell ranges and tables works
- ✅ Chart styling (colors, fonts, legend) is fully configurable
- ✅ Charts participate in dependency graph (refresh on data change)
- ✅ Round-trip tests preserve all chart properties

**Tables Complete** when:
- ✅ Can create Excel tables with structured references
- ✅ Formulas can reference tables via structured notation
- ✅ AutoFilter and totals row work correctly
- ✅ Table styles and banding serialize properly
- ✅ Pivot tables can be created with row/column/value fields

**Benchmarks Complete** when:
- ✅ JMH benchmarks run in CI
- ✅ Performance regression alerts configured
- ✅ All claimed performance improvements verified
- ✅ POI comparison benchmarks establish competitive advantage

---

## Related Documentation

- **Strategic Plan**: [strategic-implementation-plan.md](strategic-implementation-plan.md) - Phase 5 execution strategy
- **Code Scaffolds**: [docs/reference/implementation-scaffolds.md](../reference/implementation-scaffolds.md) - Implementation patterns
- **Roadmap**: [roadmap.md](roadmap.md) - Granular phase tracking (P10-P12)
- **OOXML Research**: [docs/reference/ooxml-research.md](../reference/ooxml-research.md) - Spec references

---

**Last Updated**: 2025-11-20 (consolidated from 4 separate plan files)
