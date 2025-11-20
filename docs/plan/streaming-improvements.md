# Streaming I/O Improvements

**Status**: 🟡 Partially Complete (P6.6-P6.7 ✅, P7.5 ⬜ Future)
**Priority**: Medium (P7.5 future enhancements)
**Estimated Effort**: 3-4 weeks (P7.5 only)
**Last Updated**: 2025-11-20

---

## Metadata

| Field | Value |
|-------|-------|
| **Owner Modules** | `xl-cats-effect` (streaming I/O), `xl-ooxml` (two-phase writer) |
| **Touches Files** | `ExcelIO.scala`, `StreamingXmlWriter.scala`, `TwoPhaseWriter.scala` (new) |
| **Dependencies** | P0-P8 complete, P6.6-P6.7 complete |
| **Enables** | Full-feature streaming (SST + styles at O(1) memory) |
| **Parallelizable With** | WI-07 (formula), WI-10 (tables), WI-11 (charts) — different modules |
| **Merge Risk** | Low (new streaming implementation, existing in-memory path unchanged) |

---

## Work Items

| ID | Description | Type | Files | Status | PR |
|----|-------------|------|-------|--------|----|
| `P6.6` | Fix streaming reader memory leak | Fix | `ExcelIO.scala`, `StreamingXmlReader.scala` | ✅ Done | (2025-11-13) |
| `P6.7` | Compression defaults (DEFLATED) | Config | `WriterConfig.scala`, `XlsxWriter.scala` | ✅ Done | (2025-11-14) |
| `WI-16` | Two-phase streaming writer | Feature | `TwoPhaseWriter.scala` (new) | ⏳ Not Started | - |
| `WI-17` | SST heuristic improvement | Optimize | `XlsxWriter.scala` | ⏳ Not Started | - |
| `WI-18` | Merged cells serialization | Feature | `OoxmlWorksheet.scala` | ⏳ Not Started | - |
| `WI-19` | Column/row properties serialization | Feature | `OoxmlWorksheet.scala` | ⏳ Not Started | - |
| `WI-20` | Query API (RowQuery filtering) | Feature | `QueryApi.scala` (new) | ⏳ Not Started | - |

---

## Dependencies

### Prerequisites (Complete)
- ✅ P6.6: Streaming reader fixed (fs2.io.readInputStream)
- ✅ P6.7: Compression defaults (DEFLATED)

### Enables
- Full-feature streaming for large datasets (100k+ rows with rich formatting)
- Query API for streaming filters/transforms

### File Conflicts
- **Low risk**: WI-16 (Two-phase writer) is new implementation
- **Low risk**: WI-18/WI-19 modify OoxmlWorksheet (additive)
- **Medium risk**: WI-20 (Query API) touches xl-core (Sheet extensions)

### Safe Parallelization
- ✅ WI-07/WI-08 (Formula) — different module
- ✅ WI-10 (Tables) — different OOXML parts
- ✅ WI-15 (Benchmarks) — infrastructure

---

## Worktree Strategy

**Branch naming**: `streaming` or `WI-16-two-phase-writer`

**Merge order**:
1. WI-16 (Two-phase writer) — foundation for rich streaming
2. WI-17 (SST heuristic) — optimization, can parallel
3. WI-18/WI-19 (Merged cells, col/row props) — can parallel
4. WI-20 (Query API) — after two-phase writer

**Conflict resolution**: Minimal — mostly new files

---

## Execution Algorithm

### WI-16: Two-Phase Streaming Writer (3-4 weeks)
```
1. Create worktree: `gtr create WI-16-two-phase-writer`
2. Design phase 1 (registry building):
   - RegistryBuilder for styles
   - RegistryBuilder for SST
   - Stream once to collect metadata
3. Design phase 2 (write with indices):
   - Reorder ZIP writing (ContentTypes after registries)
   - Use stable indices in worksheet XML
4. Implement disk-backed SST option (MapDB)
5. Add memory tests (< 100MB for 1M rows)
6. Add round-trip tests (SST + styles)
7. Run tests: `./mill xl-cats-effect.test`
8. Create PR: "feat(streaming): add two-phase writer with SST/styles"
9. Update roadmap: WI-16 → ✅ Complete
```

### WI-17: SST Heuristic Improvement (1 hour)
```
1. Create worktree: `gtr create WI-17-sst-heuristic`
2. Update shouldUseSST logic:
   - Calculate deduplication ratio
   - Use 20% savings threshold
3. Add tests verifying heuristic
4. Run tests: `./mill xl-ooxml.test`
5. Create PR: "optimize(ooxml): improve SST heuristic"
6. Update roadmap: WI-17 → ✅ Complete
```

### WI-18: Merged Cells Serialization (2-3 hours)
```
1. Create worktree: `gtr create WI-18-merged-cells`
2. Add mergeCells XML emission in OoxmlWorksheet
3. Add round-trip tests
4. Run tests: `./mill xl-ooxml.test`
5. Create PR: "feat(ooxml): serialize merged cells to XML"
6. Update roadmap: WI-18 → ✅ Complete
```

### WI-19: Column/Row Properties (3-4 hours)
```
1. Create worktree: `gtr create WI-19-col-row-props`
2. Add <cols> and row @ht serialization in OoxmlWorksheet
3. Add round-trip tests
4. Run tests: `./mill xl-ooxml.test`
5. Create PR: "feat(ooxml): serialize column/row properties"
6. Update roadmap: WI-19 → ✅ Complete
```

### WI-20: Query API (4-5 days)
```
1. Create worktree: `gtr create WI-20-query-api`
2. Design RowQuery DSL (filter, map, groupBy)
3. Implement streaming query operators
4. Add tests for query composition
5. Run tests: `./mill xl-core.test`
6. Create PR: "feat(core): add query API for streaming transforms"
7. Update roadmap: WI-20 → ✅ Complete
```

---

## Design

### Current State (P6.6-P6.7 Complete)

**Write Path** (✅ Working):
- True constant-memory streaming with `writeStreamTrue`
- O(1) memory (~10MB) regardless of file size
- ✅ DEFLATED compression (5-10x smaller files)
- **Limitations**:
  - No SST support (inline strings → larger files)
  - Minimal styles (default only → no rich formatting)

**Read Path** (✅ Fixed in P6.6):
- ✅ True constant-memory with `fs2.io.readInputStream`
- O(1) memory (~50MB) regardless of file size
- ✅ 100k+ row files handled without OOM
- SST materialized in memory (acceptable tradeoff)

### P7.5: Two-Phase Streaming Writer (Not Started)

**Problem**: Current streaming writer can't support SST or styles because:
1. `[Content_Types].xml` written first (before knowing if SST needed)
2. `Styles.xml` requires style registry (built during worksheet write)
3. SST requires string deduplication (need to see all strings first)

**Solution**: Two-phase approach with deferred static parts

#### Phase 1: Scan Data (Build Registries)
```scala
// First pass: collect style IDs and strings without writing
case class RegistryBuilder(
  styles: mutable.LinkedHashMap[CellStyle, StyleId],
  strings: mutable.LinkedHashMap[String, Int],
  totalStrings: Int
):
  def registerStyle(style: CellStyle): (RegistryBuilder, StyleId) = ???
  def registerString(s: String): (RegistryBuilder, Int) = ???

// Stream through data once to build registries
val (styleReg, sstReg) = dataStream
  .fold((RegistryBuilder.empty, RegistryBuilder.empty)) { case ((styles, strings), row) =>
    // Register styles and strings, don't emit yet
    ???
  }
  .compile.lastOrError
```

#### Phase 2: Write with Stable Indices
```
1. [Content_Types].xml (with SST and styles overrides)
2. Relationships (xl/_rels/workbook.xml.rels)
3. Workbook (xl/workbook.xml)
4. Worksheets (xl/worksheets/*.xml) - use stable style/string indices
5. Styles.xml (xl/styles.xml) - from styleReg
6. SharedStrings.xml (xl/sharedStrings.xml) - from sstReg
```

**Memory Considerations**:
- Style registry: O(unique styles) - typically < 1000 styles = ~100KB
- SST registry: O(unique strings) - could be large!
  - For 1M strings: ~50-100MB in memory
  - **Alternative**: Use disk-backed map (MapDB, Chronicle Map)

**Disk-Backed SST** (for very large datasets):
```scala
import org.mapdb.*

val sstMap = DBMaker
  .fileDB("temp-sst.db")
  .fileMmapEnable()
  .make()
  .hashMap("strings", Serializer.STRING, Serializer.INTEGER)
  .create()

// Use like normal map, backed by disk
sstMap.put("some string", 0)
```

**Tradeoffs**:
- **Pro**: Full features (SST, styles) with near-constant memory
- **Pro**: Can handle 1M+ rows with rich formatting
- **Con**: Two passes over data (slower than single-pass)
- **Con**: More complex than current streaming writer

**When to Use**:
- Large datasets (100k+ rows) with full styling needs
- Use disk-backed SST if > 100k unique strings

### Minor Improvements

#### WI-17: Improve SST Heuristic
**Current**: `shouldUseSST = totalCount > 100 && uniqueCount < totalCount`
**Problem**: Too simplistic, adds SST overhead for mostly-unique datasets

**Better**:
```scala
def shouldUseSST(totalCount: Int, uniqueCount: Int): Boolean =
  val deduplicationRatio = (totalCount - uniqueCount).toDouble / totalCount
  totalCount > 100 && deduplicationRatio > 0.2  // 20% savings threshold
```

**Impact**: Avoids SST overhead when strings are mostly unique

#### WI-18: Merged Cells Serialization
**Current**: `mergedRanges` tracked but not written to XML
**Fix**: Emit `<mergeCells>` in worksheet.xml

```scala
// Add to OoxmlWorksheet.toXml
if mergedRanges.nonEmpty then
  val mergeCellsElem = elem("mergeCells", "count" -> mergedRanges.size.toString)(
    mergedRanges.map(r => elem("mergeCell", "ref" -> r.toA1)())
  )
  // Add to worksheet children
```

**Impact**: Merged cells preserved in round-trip

#### WI-19: Column/Row Properties Serialization
**Current**: Widths/heights tracked but not written
**Fix**: Emit `<cols>` and `row ht=""` attributes

```scala
// Add to OoxmlWorksheet.toXml
val colsElem = if columnProperties.nonEmpty then
  elem("cols")(
    columnProperties.map { case (col, props) =>
      elem("col",
        "min" -> col.index1.toString,
        "max" -> col.index1.toString,
        "width" -> props.width.toString,
        "customWidth" -> "1"
      )()
    }.toSeq*
  )
else null
```

**Impact**: Column widths/row heights preserved

#### WI-20: Query API
**Goal**: Streaming filters and transformations

**Design**:
```scala
trait RowQuery:
  def filter(predicate: RowData => Boolean): RowQuery
  def map(f: RowData => RowData): RowQuery
  def groupBy(key: RowData => String): Map[String, Stream[F, RowData]]

// Usage
ExcelIO.readStream(path)
  .query
  .filter(_.cells.get(0).exists(_.asInt > 1000))
  .map(row => row.copy(cells = row.cells + (10 -> calculated)))
  .through(ExcelIO.writeStreamTrue(outputPath, "Filtered"))
```

**Impact**: Ergonomic streaming transformations

---

## Definition of Done

### Completed (P6.6-P6.7) ✅
- [x] Streaming reader uses fs2.io.readInputStream (O(1) memory)
- [x] Compression defaults to DEFLATED
- [x] WriterConfig for compression/prettyPrint
- [x] Memory tests pass (< 50MB for 100k rows)
- [x] Compression tests verify 5-10x reduction

### Future Work (P7.5)
- [ ] Two-phase writer implemented
- [ ] SST + styles in streaming write
- [ ] Disk-backed SST option
- [ ] Query API for streaming transforms
- [ ] Minor improvements (merged cells, col/row props, SST heuristic)
- [ ] 30+ tests added
- [ ] Performance maintained (< 100MB for 1M rows)

---

## Module Ownership

**Primary**: `xl-cats-effect` (ExcelIO.scala, TwoPhaseWriter.scala)

**Secondary**: `xl-ooxml` (OoxmlWorksheet.scala for merged cells/props)

**Test Files**: `xl-cats-effect/test` (StreamingSpec, MemorySpec, CompressionSpec)

---

## Merge Risk Assessment

**Risk Level**: Low

**Rationale**:
- Two-phase writer is new implementation (doesn't break existing)
- Minor improvements are additive to OoxmlWorksheet
- Query API is new API surface
- No breaking changes to existing streaming

---

## Related Documentation

- **Roadmap**: `docs/plan/roadmap.md` (WI-16 through WI-20)
- **Design**: `docs/design/io-modes.md` (streaming vs in-memory architecture)
- **Performance**: `docs/reference/performance-guide.md` (mode selection guidance)
- **Strategic**: `docs/plan/strategic-implementation-plan.md` (Phase 1: Hardening & Streaming)

---

## Notes

- **P6.6-P6.7 complete** — constant-memory read, compression defaults operational
- **P7.5 deferred** — two-phase writer is enhancement, not critical (in-memory path works for rich formatting)
- **Estimated timeline**: 3-4 weeks for P7.5 if needed
- **Priority**: Medium (in-memory path adequate for most use cases)
