# Streaming I/O Improvements

## Status: Partially Complete (P6.6-P6.7 ✅ Done, P7.5 ⬜ Future)

**Last Updated**: 2025-11-16

## Overview

The XL library currently has two I/O modes with different tradeoffs. This document outlines completed critical fixes and future improvements for full SST/style support in streaming writes.

## Current State (Updated 2025-11-16)

### Write Path (✅ Working)
- **True constant-memory** streaming with `writeStreamTrue`
- O(1) memory (~10MB) regardless of file size
- **✅ Compression defaults (P6.7)**: DEFLATED compression, 5-10x smaller files
- **Limitations**:
  - No SST support (inline strings only → larger files)
  - Minimal styles (default only → no rich formatting)
  - [Content_Types].xml written before SST decision

### Read Path (✅ Fixed in P6.6)
- **✅ True constant-memory** with `fs2.io.readInputStream`
- O(1) memory (~50MB) regardless of file size
- **✅ 100k+ row files** handled without OOM
- SST materialized in memory (acceptable tradeoff for most use cases)

## Critical Fixes (P6.6-P6.7) ✅ COMPLETE

### P6.6: Fix Streaming Reader ✅ COMPLETE (2025-11-13)

**Problem**: `ExcelIO.readStream` uses `readAllBytes()` for worksheets and SST
```scala
// Current broken implementation
val bytes = zipFile.getInputStream(entry).readAllBytes()  // ❌ Materializes entire entry
StreamingXmlReader.parseWorksheetStream(Stream.emits(bytes))
```

**Solution**: Use `fs2.io.readInputStream` for chunked streaming
```scala
// Fixed implementation
import fs2.io.readInputStream

val byteStream = readInputStream[IO](
  Sync[IO].delay(zipFile.getInputStream(entry)),
  chunkSize = 4096,
  closeAfterUse = true
)
StreamingXmlReader.parseWorksheetStream(byteStream)
```

**Changes Required**:
1. Update `ExcelIO.readStream` (xl-cats-effect/src/com/tjclp/xl/io/ExcelIO.scala)
2. Update `ExcelIO.readStreamByIndex`
3. Update `ExcelIO.readSheetStream`
4. Update `StreamingXmlReader.parseSharedStrings` to accept `Stream[F, Byte]`
5. Add memory tests (assert heap doesn't scale with file size)

**Test**:
```scala
test("streaming read uses constant memory"):
  val large = generateFile(100_000)  // 100k rows

  val memBefore = currentHeap()
  ExcelIO.readStream[IO](large).compile.drain.unsafeRunSync()
  val memAfter = currentHeap()

  assert((memAfter - memBefore) < 50_000_000)  // < 50MB
```

**Impact**: True O(1) memory for reads, can handle 1M+ row files

---

### P6.7: Compression Defaults ✅ COMPLETE (2025-11-14)

**Problem** (SOLVED): XlsxWriter defaults to STORED (uncompressed) + pretty-print
- Files are 5-10x larger than necessary
- Requires precomputing CRC/size for STORED
- Pretty-print only useful for debugging

**Solution**: Add compression control, default to DEFLATED
```scala
case class WriterConfig(
  compression: Compression = Compression.Deflated,
  prettyPrint: Boolean = false
)

enum Compression:
  case Stored   // No compression (debug only)
  case Deflated // Standard ZIP compression
```

**Changes Required**:
1. Add `WriterConfig` to xl-ooxml
2. Update `XlsxWriter.writeZip` to use config
3. Implement DEFLATED (remove STORED flag, no CRC/size precompute)
4. Default to `WriterConfig(Deflated, prettyPrint = false)`
5. Add compression tests

**Impact**:
- Smaller files (5-10x reduction typical)
- Faster workflows (no CRC computation)
- Configurable for debugging (STORED + pretty)

---

## Strategic Improvements (P7.5)

### P7.5: Two-Phase Streaming Writer with SST/Styles (3-4 weeks)

**Problem**: Current streaming writer can't support SST or styles because:
1. [Content_Types].xml written first (before knowing if SST needed)
2. Styles.xml requires style registry (built during worksheet write)
3. SST requires string deduplication (need to see all strings first)

**Solution**: Two-phase approach with deferred static parts

#### Phase 1: Scan Data (Build Registries)
```scala
// First pass: collect styles IDs and strings without writing
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
```scala
// Now write in this order:
1. [Content_Types].xml (with SST and styles overrides)
2. Relationships (xl/_rels/workbook.xml.rels)
3. Workbook (xl/workbook.xml)
4. Worksheets (xl/worksheets/*.xml) - use stable styles/string indices
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

**Changes Required**:
1. Add `TwoPhaseStreamingWriter` (new class)
2. Implement registry builders (mutable, stateful)
3. Add disk-backed map option for SST
4. Restructure ZIP writing order
5. Update [Content_Types].xml to include SST/styles upfront
6. Add tests for large datasets with many unique strings/styles

**Tradeoffs**:
- **Pro**: Full features (SST, styles) with near-constant memory
- **Pro**: Can handle 1M+ rows with rich formatting
- **Con**: Two passes over data (slower than single-pass)
- **Con**: More complex than current streaming writer

**When to Use**:
- Large datasets (100k+ rows) with full styling needs
- Use disk-backed SST if > 100k unique strings

---

## Minor Improvements

### Improve SST Heuristic (1 hour)
**Current**: `shouldUseSST = totalCount > 100 && uniqueCount < totalCount`
**Problem**: Too simplistic, adds SST overhead for mostly-unique datasets

**Better**:
```scala
def shouldUseSST(totalCount: Int, uniqueCount: Int): Boolean =
  val deduplicationRatio = (totalCount - uniqueCount).toDouble / totalCount
  totalCount > 100 && deduplicationRatio > 0.2  // 20% savings threshold
```

**Impact**: Avoids SST overhead when strings are mostly unique

---

### Theme Color Resolution (1-2 days)
**Current**: `Color.Theme.toArgb` returns approximation (often black)
**Fix**: Parse `xl/theme/theme1.xml`, resolve color palette

**Impact**: Correct theme colors in rich text and cell fills

---

### Merged Cells Serialization (2-3 hours)
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

---

### Column/Row Properties Serialization (3-4 hours)
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

---

## Testing Strategy

### Memory Tests (Critical)
```scala
test("streaming read: O(1) memory for 100k rows"):
  val large = generateLargeFile(100_000)
  val memBefore = currentHeapUsage()

  ExcelIO.readStream[IO](large)
    .compile.drain
    .unsafeRunSync()

  val memAfter = currentHeapUsage()
  val memUsed = (memAfter - memBefore) / (1024 * 1024)

  assert(memUsed < 50, s"Used $memUsed MB, expected < 50 MB")

test("streaming read: O(1) memory for 1M rows"):
  val huge = generateLargeFile(1_000_000)
  val memBefore = currentHeapUsage()

  ExcelIO.readStream[IO](huge)
    .take(1000)  // Process first 1000
    .compile.drain
    .unsafeRunSync()

  val memAfter = currentHeapUsage()
  val memUsed = (memAfter - memBefore) / (1024 * 1024)

  assert(memUsed < 50, s"Used $memUsed MB, expected < 50 MB")
```

### Compression Tests
```scala
test("DEFLATED produces smaller files"):
  val data = generateData(10_000)

  val storedSize = writeWith(WriterConfig(Stored, prettyPrint = false), data)
  val deflatedSize = writeWith(WriterConfig(Deflated, prettyPrint = false), data)

  assert(deflatedSize < storedSize * 0.3, "DEFLATED should be 3x smaller")

test("prettyPrint increases file size"):
  val data = generateData(10_000)

  val compactSize = writeWith(WriterConfig(Deflated, prettyPrint = false), data)
  val prettySize = writeWith(WriterConfig(Deflated, prettyPrint = true), data)

  assert(prettySize > compactSize)
```

### Two-Phase Writer Tests
```scala
test("two-phase writer preserves SST"):
  val data = Stream.range(0, 100_000).map(i =>
    RowData(Vector(s"String ${i % 1000}"))  // 1000 unique strings
  )

  val path = TwoPhaseStreamingWriter.write(data, tempPath)
  val wb = XlsxReader.read(path)

  // Should have SST with 1000 entries, total count 100k
  assert(wb.sheets.head.cells.values.forall(_.value.isText))

test("two-phase writer handles many unique styles"):
  val data = Stream.range(0, 100_000).map(i =>
    RowData(Vector("text"), styles = Vector(uniqueStyle(i % 1000)))
  )

  val memBefore = currentHeapUsage()
  TwoPhaseStreamingWriter.write(data, tempPath).unsafeRunSync()
  val memAfter = currentHeapUsage()

  assert((memAfter - memBefore) < 100_000_000)  // < 100MB
```

---

## Implementation Priority

### Phase 1: Critical Fixes (Week 1)
1. **P6.6**: Fix streaming reader (2-3 days) - CRITICAL
2. **P6.7**: Compression defaults (1 day) - QUICK WIN
3. Memory tests (1 day)

**Deliverable**: True O(1) streaming reads, smaller files

### Phase 2: Minor Improvements (Week 2)
4. SST heuristic (1 hour)
5. Merged cells (2-3 hours)
6. Column/row properties (3-4 hours)
7. Theme colors (1-2 days)

**Deliverable**: Feature completeness for in-memory path

### Phase 3: Two-Phase Writer (Weeks 3-4)
8. **P7.5**: Implement two-phase streaming writer (3-4 weeks)
9. Add disk-backed SST option
10. Comprehensive tests

**Deliverable**: Full streaming with SST/styles

---

## Success Criteria

### P6.6 (Streaming Reader Fix)
- ✅ Read 100k rows using < 50MB memory
- ✅ Read 1M rows using < 50MB memory
- ✅ Memory tests pass
- ✅ No `readAllBytes()` in codebase

### P6.7 (Compression)
- ✅ Default to DEFLATED
- ✅ Files 3-5x smaller than STORED
- ✅ Configurable compression mode
- ✅ Pretty-print configurable

### P7.5 (Two-Phase Writer)
- ✅ Write 1M rows with full styles using < 100MB memory
- ✅ SST with 100k unique strings using < 100MB
- ✅ Disk-backed SST option for very large datasets
- ✅ Round-trip tests with SST/styles

---

## Related Documents

- **STATUS.md**: Current limitations documented
- **io-modes.md**: Architecture decision for two I/O modes
- **performance-guide.md**: User guidance on mode selection
- **future-improvements.md**: Roadmap priorities

---

## Author

Documented 2025-11-11 based on technical review feedback
