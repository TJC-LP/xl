# Advanced Features (Drawings, Charts, Tables, Benchmarks)

**Status**: ⬜ Not Started
**Priority**: Medium-High
**Estimated Effort**: 10-15 weeks total
**Last Updated**: 2025-11-20

---

## Metadata

| Field | Value |
|-------|-------|
| **Owner Modules** | `xl-ooxml`, `xl-testkit` (benchmarks), `xl-evaluator` (chart data binding) |
| **Touches Files** | New files in `xl/drawings/`, `xl/charts/`, `xl/tables/`, benchmark module |
| **Dependencies** | P0-P8 complete; Charts/Tables benefit from WI-07/WI-08 (formula engine) |
| **Enables** | Full Excel feature parity (visual elements, data structures) |
| **Parallelizable With** | WI-07/WI-08 (formula) — independent modules; WI-15 can start immediately |
| **Merge Risk** | Low (new OOXML parts, minimal overlap with existing code) |

---

## Work Items

| ID | Description | Type | Files | Status | PR |
|----|-------------|------|-------|--------|----|
| `WI-10` | Table Support | Feature | `xl/ooxml/OoxmlTable.scala` | ⏳ Not Started | - |
| `WI-11` | Chart Model | Feature | `xl/ooxml/OoxmlChart.scala` | ⏳ Not Started | - |
| `WI-12` | Drawing Layer | Feature | `xl/ooxml/OoxmlDrawing.scala` | ⏳ Not Started | - |
| `WI-13` | Pivot Tables | Feature | `xl/ooxml/OoxmlPivot.scala` | ⏳ Not Started | - |
| `WI-15` | Benchmark Suite | Infra | `xl-benchmarks/` (new module) | ⏳ Not Started | - |

---

## Dependencies

### Prerequisites (Complete)
- ✅ P0-P8: Foundation complete (OOXML I/O, styles, streaming operational)
- ✅ P6.8: Surgical modification (preserves unknown parts during development)

### Recommended Order
1. **WI-15** (Benchmarks) — Can start immediately, establishes performance baseline
2. **WI-12** (Drawings) — Foundational for charts, relatively isolated
3. **WI-10** (Tables) — Needed for chart data binding, depends on named ranges (defer or implement inline)
4. **WI-11** (Charts) — Final integration, depends on drawings + tables recommended
5. **WI-13** (Pivot Tables) — Complex, requires tables + formula engine

### File Conflicts
- **Low risk**: All work items create new OOXML parts (new files)
- **Medium risk**: WI-11 (Charts) may modify Worksheet.scala for chart relationships
- **High risk**: WI-13 (Pivots) touches tables + formula evaluation (coordinate with WI-08)

### Safe Parallelization
- ✅ WI-15 (Benchmarks) + any other work — completely independent
- ✅ WI-10 (Tables) + WI-12 (Drawings) — different OOXML parts
- ✅ WI-07/WI-08 (Formula) + WI-15 (Benchmarks) — different modules

---

## Worktree Strategy

**Branch naming**: `<feature>` (e.g., `drawings`, `charts`, `tables`, `benchmarks`)

**Merge order**:
1. WI-15 (Benchmarks) — anytime (independent)
2. WI-12 (Drawings) — before WI-11 (charts depend on drawings)
3. WI-10 (Tables) — before WI-13 (pivots depend on tables)
4. WI-11 (Charts) — after WI-12, optionally after WI-10
5. WI-13 (Pivots) — last (depends on tables + formulas)

**Conflict resolution**:
- If multiple agents working on different WI-XX items, coordinate merge order
- Use PR drafts to signal work in progress

---

## Execution Algorithm

### WI-15: Benchmark Suite (Start Immediately)
```
1. Create worktree: `gtr create WI-15-benchmarks`
2. Set up xl-benchmarks module in build.mill
3. Add JMH dependencies
4. Create benchmark suites:
   - ReadWriteBenchmark (streaming vs in-memory)
   - PatchBenchmark (single cell vs batch)
   - StyleBenchmark (deduplication, indexing)
5. Add POI comparison benchmarks
6. Configure CI job for regression alerts
7. Run benchmarks: `./mill xl-benchmarks.runJmh`
8. Document results in STATUS.md
9. Create PR: "feat(benchmarks): add JMH performance testing suite"
10. Update roadmap: WI-15 → ✅ Complete, WI-16 → 🔵 Available
```

### WI-12: Drawing Layer (After WI-15)
```
1. Create worktree: `gtr create WI-12-drawings`
2. Define Drawing AST in xl-ooxml:
   - Anchor (TwoCell, OneCell)
   - Picture (image bytes, anchor, alt text)
   - Shape (Rect, Ellipse, Line, Polygon)
3. Implement OoxmlDrawing serialization
4. Add media deduplication (SHA-256)
5. Wire into relationship graph
6. Add round-trip tests with embedded images
7. Run tests: `./mill xl-ooxml.test`
8. Create PR: "feat(ooxml): add drawing layer support"
9. Update roadmap: WI-12 → ✅ Complete
```

### WI-10: Table Support (Parallel with WI-12)
```
1. Create worktree: `gtr create WI-10-tables`
2. Define TableSpec AST:
   - Range, style, headers, totals row
   - AutoFilter configuration
3. Implement OoxmlTable serialization
4. Add structured reference support (basic or defer to formula engine)
5. Add round-trip tests
6. Run tests: `./mill xl-ooxml.test`
7. Create PR: "feat(ooxml): add Excel table support"
8. Update roadmap: WI-10 → ✅ Complete, WI-13 → recalculate dependencies
```

### WI-11: Chart Model (After WI-12, optionally WI-10)
```
1. Create worktree: `gtr create WI-11-charts`
2. Define Chart AST:
   - ChartType (Bar, Line, Pie, Scatter)
   - Series (data binding, styling)
   - Axes (Category, Value, Time)
   - Legend positioning
3. Implement OoxmlChart serialization
4. Wire chart data bindings (use cell ranges initially)
5. Add round-trip tests for each chart type
6. Run tests: `./mill xl-ooxml.test`
7. Create PR: "feat(ooxml): add chart model support"
8. Update roadmap: WI-11 → ✅ Complete
```

### WI-13: Pivot Tables (Last — Requires Tables + Formulas)
```
1. Create worktree: `gtr create WI-13-pivots`
2. Define PivotSpec AST:
   - Source range/cache
   - Row/column/value fields
   - Aggregations (Sum, Avg, Count, Min, Max)
3. Implement pivot cache materialization
4. Implement OoxmlPivot serialization
5. Wire into dependency graph for refresh
6. Add comprehensive round-trip tests
7. Run tests: `./mill xl-ooxml.test`
8. Create PR: "feat(ooxml): add pivot table support"
9. Update roadmap: WI-13 → ✅ Complete
```

---

## Detailed Design

### 1. Drawings & Graphics (P10)

**Status**: ⬜ Not Started
**Estimated Effort**: 10-15 days
**Priority**: Medium
**LOC**: ~800

#### Scope

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

#### OOXML Mapping

- Drawings reside in `xl/drawings/drawing#.xml` with anchors:
  - `xdr:twoCellAnchor` for cell-relative positioning
  - `xdr:oneCellAnchor` for fixed positioning
- Media stored under `/xl/media/*.png|*.jpeg`
- Relations via `drawing#.xml.rels` and `workbook.xml.rels`

#### Implementation Strategy

1. Define pure AST for drawings (Anchor, Picture, Shape types)
2. Implement XmlCodec[OoxmlDrawing] for serialization
3. Add media deduplication via SHA-256 content hashing
4. Wire into existing relationship graph
5. Add round-trip tests with embedded images

**See**: `docs/reference/implementation-scaffolds.md` Section B.6 for code examples

---

### 2. Charts (P11)

**Status**: ⬜ Not Started
**Estimated Effort**: 20-30 days
**Priority**: Medium
**LOC**: ~1500+

#### Scope

**Chart Types**:
- **Marks**: `Column|Bar` (clustered/stacked), `Line` (smooth/markers), `Area` (stacked), `Scatter`, `Pie`
- **Axes**: Category/Value/Time
- **Scales**: Linear/Log10/Time

**Series Specification**:
- `(mark, encoding{x, y, color?, size?}, name?)`
- Multiple series per chart
- Data binding to cell ranges or tables

#### Semantics

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

#### OOXML Mapping

- Charts in `xl/charts/chart#.xml`:
  - `c:barChart|lineChart|pieChart|scatterChart`
  - `c:plotArea` for chart content
  - `c:legend` for legend positioning
  - Axes configuration
- Data ranges emit:
  - `c:numRef` for numeric data
  - `c:strRef` for string data
  - With `f` formulas referencing sheet ranges

#### Implementation Strategy

1. Define pure Chart AST (ChartType, Series, Axis, Legend types)
2. Implement typed grammar for chart specifications
3. Add XmlCodec[OoxmlChartPart] for ChartML serialization
4. Wire chart data bindings to formula dependency graph
5. Add comprehensive round-trip tests
6. Support common chart types first (bar, line, pie), then extend

**See**: `docs/reference/implementation-scaffolds.md` Section B.6 for code examples

---

### 3. Tables & Pivot Tables (P12)

**Status**: ⬜ Not Started
**Estimated Effort**: 15-20 days
**Priority**: Low
**LOC**: ~1000

#### Scope

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

#### OOXML Mapping

**Tables**:
- `xl/tables/table#.xml` for table definitions
- Structured references via `definedNames`
- AutoFilter integration

**Pivots**:
- `xl/pivotTables/pivotTable#.xml` for pivot structure
- `xl/pivotCache/pivotCacheDefinition#.xml` for data cache
- `xl/pivotCache/pivotCacheRecords#.xml` for materialized data

#### Implementation Strategy

1. Define pure AST for tables (TableSpec, ColumnDef, TotalsRow)
2. Define pure AST for pivots (PivotSpec, Field, Aggregation)
3. Implement structured reference resolution in formula parser
4. Add pivot cache materialization (can be lazy)
5. Wire into dependency graph for refresh semantics
6. Add comprehensive round-trip tests

**See**: `docs/reference/implementation-scaffolds.md` Section B.5 for code examples

---

### 4. Performance Benchmarks (Infrastructure)

**Status**: ⬜ Not Started
**Estimated Effort**: 1-2 weeks
**Priority**: High (for performance claims validation)

#### Scope

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

#### Infrastructure

**Build Integration**:
- `xl-benchmarks` module with sbt-jmh plugin
- CI integration for regression detection
- Threshold alerts (fail build if >10% regression)

**Baseline Establishment**:
1. Run benchmarks on current codebase (P0-P8 complete)
2. Establish baseline performance targets
3. Track improvements/regressions per phase

**Comparison Matrix** (Current estimates):
```
| Operation          | XL      | Apache POI | Improvement |
|--------------------|---------|------------|-------------|
| Write 100k rows    | ~1.1s   | ~5s        | 4.5x faster |
| Read 100k rows     | ~1.8s   | ~8s        | 4.4x faster |
| Memory (100k rows) | ~50MB   | ~800MB     | 16x less    |
```

#### Implementation Strategy

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

**Drawings** (WI-12):
- ✅ P4 OOXML foundation complete
- ✅ P6.8 Surgical modification (unknown part preservation)
- Recommended: Relationship graph extensions (minimal)

**Charts** (WI-11):
- ✅ P4 OOXML foundation complete
- Recommended: WI-12 (Drawings) — charts often contain embedded shapes
- Optional: WI-07/WI-08 (Formula) — for dynamic data binding
- Optional: WI-10 (Tables) — for table-based chart data

**Tables** (WI-10):
- ✅ P4 OOXML foundation complete
- Optional: Named ranges support (can implement inline)
- Optional: WI-07/WI-08 (Formula) — for structured references

**Pivots** (WI-13):
- ✅ P4 OOXML foundation complete
- Required: WI-10 (Tables) — pivot source data
- Required: WI-08 (Formula Evaluator) — for aggregations
- High dependency complexity

**Benchmarks** (WI-15):
- ✅ Can start immediately
- ✅ Current implementation provides baseline
- Expand as features added

### Recommended Implementation Order

1. **WI-15** (Benchmarks) — parallel with any other work, establishes performance baseline
2. **WI-12** (Drawings) — foundational for charts, relatively isolated
3. **WI-10** (Tables) — needed for chart data binding
4. **WI-11** (Charts) — final integration, benefits from WI-12 + WI-10
5. **WI-13** (Pivots) — most complex, requires WI-10 + WI-08

---

## Definition of Done

### Per Feature

**Drawings Complete** (WI-12) when:
- [x] Can embed PNG/JPEG images with two-cell and one-cell anchors
- [x] Can create basic shapes (rectangle, ellipse, line)
- [x] Media deduplication via SHA-256 works correctly
- [x] Round-trip tests preserve all drawing properties
- [x] Surgical modification preserves existing drawings

**Charts Complete** (WI-11) when:
- [x] Can create bar, line, pie, scatter charts
- [x] Data binding to cell ranges works
- [x] Chart styling (colors, fonts, legend) is fully configurable
- [x] Round-trip tests preserve all chart properties
- [x] Surgical modification preserves existing charts

**Tables Complete** (WI-10) when:
- [x] Can create Excel tables with headers and totals
- [x] AutoFilter works correctly
- [x] Table styles and banding serialize properly
- [x] Round-trip tests preserve table structure

**Pivots Complete** (WI-13) when:
- [x] Can create pivot tables with row/column/value fields
- [x] Aggregations work (Sum, Avg, Count, Min, Max)
- [x] Pivot cache materializes correctly
- [x] Round-trip tests preserve pivot structure

**Benchmarks Complete** (WI-15) when:
- [x] JMH benchmarks run in CI
- [x] Performance regression alerts configured
- [x] All claimed performance improvements verified (4.5x faster, 16x less memory)
- [x] POI comparison benchmarks establish competitive advantage

### Overall Phase

- [ ] All WI-10 through WI-15 marked ✅ Complete
- [ ] 200+ tests added (50 per feature + benchmarks)
- [ ] STATUS.md updated with new capabilities
- [ ] LIMITATIONS.md updated (remove what's now implemented)
- [ ] Examples added to docs/reference/examples.md
- [ ] Zero WartRemover errors
- [ ] Code formatted

---

## Module Ownership

**Per Work Item**:
- **WI-10** (Tables): `xl-ooxml` (OoxmlTable.scala, TableSpec.scala)
- **WI-11** (Charts): `xl-ooxml` (OoxmlChart.scala, ChartSpec.scala, ChartML serialization)
- **WI-12** (Drawings): `xl-ooxml` (OoxmlDrawing.scala, DrawingSpec.scala, media/)
- **WI-13** (Pivots): `xl-ooxml` + `xl-evaluator` (PivotSpec.scala, pivot cache, aggregations)
- **WI-15** (Benchmarks): `xl-benchmarks` (new module, JMH suites)

**Shared Files** (potential conflicts):
- `xl-ooxml/src/com/tjclp/xl/ooxml/Worksheet.scala` — **Medium risk** (WI-11, WI-12 add relationships)
- `xl-ooxml/src/com/tjclp/xl/ooxml/Relationships.scala` — **Medium risk** (all features add relationship types)

---

## Merge Risk Assessment

**WI-10** (Tables): **Low**
- New files only
- Minimal Worksheet.scala changes

**WI-11** (Charts): **Medium**
- Touches Worksheet.scala (chart relationships)
- Complex OOXML (ChartML is verbose)

**WI-12** (Drawings): **Low**
- New files, isolated OOXML part
- Media deduplication is self-contained

**WI-13** (Pivots): **High**
- Depends on tables + formula evaluator
- Complex data flow (source → cache → pivot)
- Cross-module coordination required

**WI-15** (Benchmarks): **None**
- New module, zero overlap with features

---

## Success Criteria

### Functional Requirements
- ✅ Can embed images in worksheets
- ✅ Can create charts with data binding
- ✅ Can create Excel tables with structured references
- ✅ Can create pivot tables with aggregations
- ✅ Benchmark suite validates performance claims

### Non-Functional Requirements
- ✅ Deterministic output (all XML sorted, stable)
- ✅ Round-trip fidelity (read → write → read = identity)
- ✅ Surgical modification preserves all advanced features
- ✅ Performance maintained (no regressions from adding features)

---

## Related Documentation

- **Strategic Plan**: [strategic-implementation-plan.md](strategic-implementation-plan.md) - Phase 5 (Advanced Features)
- **Code Scaffolds**: [docs/reference/implementation-scaffolds.md](../reference/implementation-scaffolds.md) - Implementation patterns
- **Roadmap**: [roadmap.md](roadmap.md) - Granular phase tracking (WI-10 through WI-15)
- **OOXML Research**: [docs/reference/ooxml-research.md](../reference/ooxml-research.md) - Spec references
- **Performance**: [docs/reference/performance-guide.md](../reference/performance-guide.md) - Benchmarking guidance

---

## Notes

- **Surgical modification** already preserves all these features when reading existing XLSX files — this plan is about **first-class AST support** for creation/modification
- **Formula dependency** is soft — basic features (static charts, simple tables) can work without evaluator
- **Benchmarks** should start ASAP to establish baseline before optimization work begins
- **Estimated timeline**: 10-15 weeks total if serialized; 6-8 weeks if parallelized across 2-3 agents

---

**Last Updated**: 2025-11-20 (restructured with work items table)
