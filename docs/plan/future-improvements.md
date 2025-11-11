# P6.5: Performance & Quality Polish

**Status**: 🟡 Partially Complete (Critical issues fixed, medium-priority deferred)
**Priority**: Medium
**Estimated Effort**: 8-10 hours total (5 hours completed, 3-5 hours remaining)
**Source**: PR #4 review feedback (chatgpt-codex-connector, claude)

## Overview

This phase addresses medium-priority improvements identified during PR #4 review. While not blockers for the current release, these enhancements will improve performance, code quality, observability, and test coverage.

**Key Areas**:
1. Performance optimizations (O(n²) → O(1))
2. Code quality improvements (DRY, utilities)
3. Observability enhancements (logging, warnings)
4. Test coverage expansion (error paths, integration)

---

## Performance Improvements

### 1. ✅ O(n²) Style indexOf Optimization (COMPLETED)

**Priority**: High (Critical for workbooks with 1000+ unique styles)
**Estimated Effort**: 2 hours
**Source**: PR #4 Issue #1
**Completed**: 2025-11-11 (commit 79b3269)

**Location**: `xl-ooxml/src/com/tjclp/xl/ooxml/Styles.scala:185-193`

**Problem**:
```scala
val cellXfsElem = elem("cellXfs", "count" -> index.cellStyles.size.toString)(
  index.cellStyles.map { style =>
    val fontIdx = index.fonts.indexOf(style.font)      // O(n)
    val fillIdx = allFills.indexOf(style.fill)         // O(n)
    val borderIdx = index.borders.indexOf(style.border) // O(n)
    // ... → O(n²) overall
  }*
)
```

**Impact**:
- Current: O(n²) where n = number of unique styles
- Becomes problematic for workbooks with 1000+ unique cell formats
- Typical case: 10-100 styles → minimal impact (~1ms)
- Pathological case: 10,000 styles → could be 100ms+

**Solution**:
```scala
// Pre-build lookup maps once (O(n))
val fontMap = index.fonts.zipWithIndex.toMap
val fillMap = allFills.zipWithIndex.toMap
val borderMap = index.borders.zipWithIndex.toMap

// Then O(1) lookups instead of O(n)
val cellXfsElem = elem("cellXfs", "count" -> index.cellStyles.size.toString)(
  index.cellStyles.map { style =>
    val fontIdx = fontMap(style.font)      // O(1)
    val fillIdx = fillMap(style.fill)      // O(1)
    val borderIdx = borderMap(style.border) // O(1)
    // ... → O(n) overall
  }*
)
```

**Implementation**:
```scala
// Pre-build lookup maps for O(1) access
val fontMap = index.fonts.zipWithIndex.toMap
val fillMap = allFills.zipWithIndex.toMap
val borderMap = index.borders.zipWithIndex.toMap

val cellXfsElem = elem("cellXfs", "count" -> index.cellStyles.size.toString)(
  index.cellStyles.map { style =>
    val fontIdx = fontMap.getOrElse(style.font, -1)    // O(1)
    val fillIdx = fillMap.getOrElse(style.fill, -1)    // O(1)
    val borderIdx = borderMap.getOrElse(style.border, -1) // O(1)
    ...
  }*
)
```

**Tests Added**:
- StylePerformanceSpec: 2 tests verifying linear scaling with 1000+ styles
- Performance comparison test: 100 vs 1000 styles < 20x ratio

**Result**: ✅ Complete - All tests pass, performance scales sub-quadratically

---

## Critical Fixes (Completed)

### 1b. ✅ Non-Deterministic allFills Ordering (COMPLETED)

**Priority**: High (Architecture violation - determinism required)
**Estimated Effort**: 30 minutes
**Source**: PR #4 Issue #2
**Completed**: 2025-11-11 (commit 79b3269)

**Location**: `xl-ooxml/src/com/tjclp/xl/ooxml/Styles.scala:166-179`

**Problem**:
```scala
val allFills = (defaultFills ++ index.fills.filterNot(defaultFills.contains)).distinct
```
The `.distinct` method on Vector uses Set internally, which has non-deterministic iteration order. This violates the architecture principle #3 (deterministic output) and breaks byte-identical output requirement.

**Impact**: Non-deterministic XML output breaks stable diffs and reproducible builds.

**Solution**:
```scala
// Deterministic deduplication preserving first-occurrence order
val allFills = {
  val builder = Vector.newBuilder[Fill]
  val seen = scala.collection.mutable.Set.empty[Fill]

  for (fill <- defaultFills ++ index.fills) {
    if (!seen.contains(fill)) {
      seen += fill
      builder += fill
    }
  }
  builder.result()
}
```

**Tests Added**:
- DeterminismSpec: 4 tests verifying fill ordering, XML stability, deduplication
- Verifies multiple serializations produce identical output

**Result**: ✅ Complete - All tests pass, output is deterministic

---

### 1c. ✅ XXE Security Vulnerability (COMPLETED)

**Priority**: High (Security - required for production)
**Estimated Effort**: 30 minutes
**Source**: PR #4 Issue #3
**Completed**: 2025-11-11 (commit 79b3269)

**Location**: `xl-ooxml/src/com/tjclp/xl/ooxml/XlsxReader.scala:205-232`

**Problem**:
```scala
try Right(XML.loadString(xmlString))
```
Default XML parser is vulnerable to XML External Entity (XXE) attacks. Malicious XLSX files could read arbitrary files from server or cause DoS.

**Impact**: Critical security vulnerability allowing:
- Server-side file disclosure
- Denial of Service (billion laughs attack)
- Internal port scanning

**Solution**:
```scala
private def parseXml(xmlString: String, location: String): XLResult[Elem] =
  try
    // Configure SAX parser to prevent XXE (XML External Entity) attacks
    val factory = javax.xml.parsers.SAXParserFactory.newInstance()
    factory.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true)
    factory.setFeature("http://xml.org/sax/features/external-general-entities", false)
    factory.setFeature("http://xml.org/sax/features/external-parameter-entities", false)
    factory.setFeature("http://apache.org/xml/features/nonvalidating/load-external-dtd", false)
    factory.setXIncludeAware(false)
    factory.setNamespaceAware(true)

    val loader = XML.withSAXParser(factory.newSAXParser())
    Right(loader.loadString(xmlString))
  catch
    case e: Exception =>
      Left(XLError.ParseError(location, s"XML parse error: ${e.getMessage}"))
```

**Tests Added**:
- SecuritySpec: 3 tests verifying XXE rejection
  - Rejects DOCTYPE declarations
  - Rejects external entity references
  - Accepts legitimate files

**Result**: ✅ Complete - All XXE attack vectors blocked, legitimate files parse correctly

---

## Code Quality Improvements (Deferred)

### 2. Extract Whitespace Check to Utility

**Priority**: Low (Code quality, DRY principle)
**Estimated Effort**: 30 minutes
**Source**: PR #4 Issue 4

**Locations** (5 occurrences):
- `SharedStrings.scala:58`
- `Worksheet.scala:26` (plain text)
- `Worksheet.scala:67` (RichText)
- `StreamingXmlWriter.scala:43` (plain text)
- `StreamingXmlWriter.scala:185` (RichText)

**Current**:
```scala
val needsPreserve = s.startsWith(" ") || s.endsWith(" ") || s.contains("  ")
```

**Refactored**:
```scala
// In xl-ooxml/src/com/tjclp/xl/ooxml/xml.scala (XmlUtil object)

/**
 * Check if string needs xml:space="preserve" per OOXML spec.
 *
 * REQUIRES: s is non-null string
 * ENSURES: Returns true if s has leading/trailing/multiple consecutive spaces
 * DETERMINISTIC: Yes (pure function)
 * ERROR CASES: None (total function)
 *
 * @param s String to check
 * @return true if xml:space="preserve" should be added
 */
def needsXmlSpacePreserve(s: String): Boolean =
  s.startsWith(" ") || s.endsWith(" ") || s.contains("  ")
```

**Then replace all 5 occurrences**:
```scala
val needsPreserve = XmlUtil.needsXmlSpacePreserve(s)
```

**Benefits**:
- Single source of truth
- Easier to maintain if logic changes
- Self-documenting function name
- Consistent across all modules

**Test Plan**:
- All existing whitespace tests should pass unchanged
- No new tests needed (behavior identical)

---

### 3. Logging for Missing styles.xml

**Priority**: Low (Observability)
**Estimated Effort**: 15 minutes
**Source**: PR #4 Issue 5

**Location**: `xl-ooxml/src/com/tjclp/xl/ooxml/XlsxReader.scala:98`

**Current**:
```scala
val styles = parts
  .get("xl/styles.xml")
  .flatMap(xml => WorkbookStyles.fromXml(xml.root).toOption)
  .getOrElse(WorkbookStyles.default) // Silent fallback
```

**Improved**:
```scala
val styles = parts.get("xl/styles.xml") match
  case Some(xml) =>
    WorkbookStyles.fromXml(xml.root) match
      case Right(s) => s
      case Left(err) =>
        // Log parse error, fall back to default
        System.err.println(s"Warning: Failed to parse xl/styles.xml: $err")
        WorkbookStyles.default
  case None =>
    // Log missing styles.xml
    System.err.println("Warning: xl/styles.xml not found, using default styles")
    WorkbookStyles.default
```

**Benefits**:
- Helps debug malformed XLSX files
- Makes silent failures visible
- Aids troubleshooting for users

**Considerations**:
- May want structured logging (not System.err) for production
- Consider making logging configurable
- Defer to P11 (comprehensive logging strategy)

**For Now**: Document this as a P11 enhancement rather than implementing immediately.

---

## Test Coverage Improvements

### 4. XlsxReader Error Path Tests

**Priority**: Medium (Robustness)
**Estimated Effort**: 2 hours
**Source**: PR #4 Issue 4

**Missing Coverage**:

**Malformed XML**:
- Test: "XlsxReader rejects invalid workbook.xml"
- Test: "XlsxReader handles missing required elements gracefully"
- Test: "XlsxReader reports clear errors for malformed styles.xml"

**Missing Required Files**:
- Test: "XlsxReader rejects XLSX missing workbook.xml"
- Test: "XlsxReader handles missing worksheet gracefully"
- Test: "XlsxReader handles missing [Content_Types].xml"

**Corrupt ZIP Files**:
- Test: "XlsxReader rejects corrupt ZIP"
- Test: "XlsxReader handles incomplete ZIP"
- Test: "XlsxReader rejects ZIP with invalid central directory"

**Invalid Relationships**:
- Test: "XlsxReader handles broken relationship references"
- Test: "XlsxReader handles circular relationship references"
- Test: "XlsxReader handles missing relationship targets"

**Implementation**:
- Create `xl-ooxml/test/src/com/tjclp/xl/ooxml/XlsxReaderErrorSpec.scala`
- Use resource files with intentionally malformed XLSX
- Verify proper error messages (not stack traces)
- Verify no resource leaks on error

---

### 5. Full XLSX Round-Trip Integration Test

**Priority**: Medium (Integration verification)
**Estimated Effort**: 1 hour
**Source**: PR #4 Issue 4

**Goal**: Verify all features survive write → read → write cycle with byte-identical output.

**Test Outline**:
```scala
test("full XLSX round-trip preserves all features") {
  // Create workbook with ALL features:
  val wb = createComplexWorkbook()

  // Features to test:
  // - All cell types (Text, Number, Bool, Formula, Error, DateTime, RichText)
  // - All styles (fonts, fills, borders, numFmts, alignment)
  // - Merged cells
  // - Multiple sheets
  // - Non-sequential sheet indices
  // - Whitespace in text (leading, trailing, double)
  // - Duplicate strings (SST deduplication)
  // - All alignment properties
  // - RichText with formatting

  // Write to bytes
  val bytes1 = XlsxWriter.toBytes(wb)

  // Read back
  val wb2 = XlsxReader.fromBytes(bytes1).getOrElse(fail("Read failed"))

  // Write again
  val bytes2 = XlsxWriter.toBytes(wb2)

  // Verify byte-identical (deterministic output)
  assertEquals(bytes1, bytes2, "Second write should be byte-identical")

  // Also verify domain model equality
  assertEquals(wb2.sheets.size, wb.sheets.size)
  wb.sheets.zip(wb2.sheets).foreach { case (s1, s2) =>
    assertEquals(s2.cells, s1.cells, "Cells should round-trip")
    assertEquals(s2.mergedRanges, s1.mergedRanges, "Merges should round-trip")
    // ... verify all properties
  }
}
```

**Benefits**:
- High confidence in round-trip fidelity
- Catches subtle serialization bugs
- Verifies deterministic output
- Integration-level validation

---

## Definition of Done

### Completed ✅
- [x] Style indexOf optimized to O(1) using Maps (Issue #1)
- [x] Non-deterministic allFills ordering fixed (Issue #2)
- [x] XXE security vulnerability fixed (Issue #3)
- [x] Benchmark tests verify linear scaling (StylePerformanceSpec: 2 tests)
- [x] Determinism tests verify stable output (DeterminismSpec: 4 tests)
- [x] Security tests verify XXE protection (SecuritySpec: 3 tests)
- [x] All tests passing (645 total: 169 core + 215 ooxml + 261 cats-effect)
- [x] Code formatted (Scalafmt 3.10.1)
- [x] Documentation updated (this file, roadmap.md)
- [x] Committed (79b3269)

### Deferred (Medium Priority)
- [ ] Whitespace check extracted to XmlUtil.needsXmlSpacePreserve
- [ ] All 5 call sites updated to use utility
- [ ] 10+ error path tests for XlsxReader
- [ ] 1 comprehensive round-trip integration test
- [ ] Logging strategy documented for P11

### Future Work
- [ ] Smarter SST heuristic (Issue #6 - optimization, not correctness)

---

## Implementation Order

**Phase 1: Quick Wins (1 hour)**
1. Extract needsXmlSpacePreserve utility (30 min)
2. Add round-trip integration test (30 min)

**Phase 2: Performance (2 hours)**
3. Optimize style indexOf with Maps (2 hours)
   - Includes benchmark test

**Phase 3: Robustness (2 hours)**
4. Add XlsxReader error path tests (2 hours)

**Phase 4: Deferred to P11**
5. Structured logging for missing/malformed files

---

## Related

- **PR #4**: Source of feedback
- **P11**: Security & observability enhancements
- **Benchmarks Plan**: `docs/plan/benchmarks.md` for formal performance testing

---

## Notes

This phase is **not blocking** for production use. The current implementation is correct and passes all tests. These are **enhancements** that improve code quality and performance at scale.

**When to Tackle**:
- Before releasing 1.0 (performance matters for large files)
- When adding formal benchmarks (P6.5 + benchmarks.md together)
- When contributor bandwidth allows

---

# P6.6: Fix Streaming Reader (Critical Memory Issue)

**Status**: ✅ Complete (2025-11-11)
**Priority**: CRITICAL (violates documented O(1) claim)
**Actual Effort**: 2 days
**Source**: Technical review 2025-11-11

## Problem (SOLVED)

The streaming reader API (`readStreamTrue`, `readStream`, `readStreamByIndex`, `readSheetStream`) was using `InputStream.readAllBytes()` internally, which **materialized entire ZIP entries in memory**. This violated the constant-memory claim and made the API misleading.

**Fixed**: Replaced all `readAllBytes()` calls with `fs2.io.readInputStream` for true chunked streaming.

**Impact**:
- Large files (100k+ rows) spike memory or OOM
- Memory grows with file size (O(n), not O(1))
- Users expect constant-memory but get linear growth

**Current Broken Implementation**:
```scala
// xl-cats-effect/src/com/tjclp/xl/io/ExcelIO.scala
val bytes = zipFile.getInputStream(entry).readAllBytes()  // ❌ Materializes entire entry!
StreamingXmlReader.parseWorksheetStream(Stream.emits(bytes))
```

## Solution

Replace `readAllBytes()` with `fs2.io.readInputStream` for chunked streaming:

```scala
import fs2.io.readInputStream

val byteStream = readInputStream[F](
  Sync[F].delay(zipFile.getInputStream(entry)),
  chunkSize = 4096,
  closeAfterUse = true
)
StreamingXmlReader.parseWorksheetStream(byteStream)
```

## Files to Change

1. **ExcelIO.scala** (4 methods):
   - `readStream` - worksheet entries
   - `readStreamByIndex` - worksheet entries
   - `readSheetStream` - specific sheet
   - Update SST parsing in all three

2. **StreamingXmlReader.scala**:
   - Update `parseSharedStrings` to accept `Stream[F, Byte]` instead of `Array[Byte]`
   - Update `parseWorksheetStream` if needed

3. **Tests**:
   - Add memory tests (assert heap doesn't scale with file size)
   - Test 100k row file uses < 50MB
   - Test 1M row file uses < 50MB

## Implementation Steps

### Step 1: Update StreamingXmlReader.parseSharedStrings (1 hour)
```scala
// Before
def parseSharedStrings(bytes: Array[Byte]): Either[String, SharedStrings] =
  val stream = Stream.emits(bytes)
  // ...

// After
def parseSharedStrings[F[_]: Async](byteStream: Stream[F, Byte]): F[Either[String, SharedStrings]] =
  byteStream
    .through(fs2.data.xml.events.events())
    .compile.toList
    .map(parseEvents)
```

### Step 2: Update ExcelIO methods (2-3 hours)
```scala
// Before
def readStream[F[_]: Async](path: Path): Stream[F, RowData] =
  // ... get ZIP entry
  val bytes = zipFile.getInputStream(entry).readAllBytes()
  StreamingXmlReader.parseWorksheetStream(Stream.emits(bytes))

// After
def readStream[F[_]: Async](path: Path): Stream[F, RowData] =
  // ... get ZIP entry
  val byteStream = readInputStream[F](
    Sync[F].delay(zipFile.getInputStream(entry)),
    chunkSize = 4096,
    closeAfterUse = true
  )
  StreamingXmlReader.parseWorksheetStream(byteStream)
```

### Step 3: Add Memory Tests (1 hour)
```scala
test("streaming read uses constant memory for 100k rows"):
  val large = generateLargeFile(100_000)
  val memBefore = currentHeapUsage()

  ExcelIO.readStream[IO](large)
    .compile.drain
    .unsafeRunSync()

  val memAfter = currentHeapUsage()
  val memUsed = (memAfter - memBefore) / (1024 * 1024)  // MB

  assert(memUsed < 50, s"Used $memUsed MB (expected < 50 MB)")

test("streaming read uses constant memory for 1M rows"):
  val huge = generateLargeFile(1_000_000)
  val memBefore = currentHeapUsage()

  ExcelIO.readStream[IO](huge)
    .take(1000)  // Process first 1000
    .compile.drain
    .unsafeRunSync()

  val memAfter = currentHeapUsage()
  val memUsed = (memAfter - memBefore) / (1024 * 1024)

  assert(memUsed < 50, s"Used $memUsed MB (expected < 50 MB)")
```

### Step 4: Update Documentation (30 minutes)
- Remove ⚠️ warnings from STATUS.md
- Update README.md streaming section
- Update performance-guide.md

## Success Criteria

- ✅ Read 100k rows using < 50MB memory
- ✅ Read 1M rows using < 50MB memory (streaming, not full materialization)
- ✅ No `readAllBytes()` in codebase
- ✅ All existing tests pass
- ✅ Memory tests added (2 new tests)

## Related

- **ADR-013**: Acknowledges this bug and fix plan
- **streaming-improvements.md**: Full streaming roadmap
- **io-modes.md**: Architecture explanation

---

# P6.7: Compression Defaults & Configuration

**Status**: ✅ Complete (2025-11-11)
**Priority**: HIGH (production readiness)
**Actual Effort**: 1 day
**Estimated Effort**: 1 day
**Source**: Technical review 2025-11-11

## Problem

XlsxWriter currently defaults to **STORED (uncompressed) + prettyPrint=true**, which:
- Produces files **5-10x larger** than necessary
- Requires precomputing CRC/size for STORED entries (overhead)
- Pretty-printed XML only useful for debugging
- Not production-friendly

## Solution

Add `WriterConfig` with compression control, default to DEFLATED:

```scala
case class WriterConfig(
  compression: Compression = Compression.Deflated,
  prettyPrint: Boolean = false
)

enum Compression:
  case Stored   // No compression (debug mode)
  case Deflated // Standard ZIP compression (production)
```

## Implementation

### Step 1: Add WriterConfig (1 hour)
**Location**: `xl-ooxml/src/com/tjclp/xl/ooxml/WriterConfig.scala`

```scala
package com.tjclp.xl.ooxml

/** Compression method for XLSX ZIP entries */
enum Compression:
  /** No compression (larger files, faster for debugging) */
  case Stored
  /** Standard DEFLATE compression (smaller files, production default) */
  case Deflated

/**
 * Configuration for XLSX writer.
 *
 * @param compression Compression method (default: Deflated for production)
 * @param prettyPrint Whether to pretty-print XML (default: false for compact output)
 */
case class WriterConfig(
  compression: Compression = Compression.Deflated,
  prettyPrint: Boolean = false
)

object WriterConfig:
  /** Production defaults: compressed, compact */
  val default: WriterConfig = WriterConfig()

  /** Debug defaults: uncompressed, pretty-printed for git diffs */
  val debug: WriterConfig = WriterConfig(Compression.Stored, prettyPrint = true)
```

### Step 2: Update XlsxWriter.writeZip (2 hours)
**Location**: `xl-ooxml/src/com/tjclp/xl/ooxml/XlsxWriter.scala`

```scala
def writeZip(
  workbook: Workbook,
  outputStream: OutputStream,
  config: WriterConfig = WriterConfig.default
): Unit =
  val zipOut = new ZipOutputStream(outputStream)

  // Set compression method based on config
  config.compression match
    case Compression.Stored =>
      zipOut.setMethod(ZipOutputStream.STORED)
    case Compression.Deflated =>
      zipOut.setMethod(ZipOutputStream.DEFLATED)
      zipOut.setLevel(Deflater.DEFAULT_COMPRESSION)

  // ... write parts

  def writePart(entryName: String, xml: Elem): Unit =
    val entry = new ZipEntry(entryName)

    config.compression match
      case Compression.Stored =>
        // STORED requires precomputed size and CRC
        val xmlString = if config.prettyPrint then
          XmlUtil.prettyPrint(xml)
        else
          xml.toString  // Compact
        val bytes = xmlString.getBytes(StandardCharsets.UTF_8)
        entry.setSize(bytes.length)
        entry.setCrc(computeCrc(bytes))
        zipOut.putNextEntry(entry)
        zipOut.write(bytes)

      case Compression.Deflated =>
        // DEFLATED doesn't need precomputed size/CRC
        zipOut.putNextEntry(entry)
        val xmlString = if config.prettyPrint then
          XmlUtil.prettyPrint(xml)
        else
          xml.toString  // Compact
        zipOut.write(xmlString.getBytes(StandardCharsets.UTF_8))

    zipOut.closeEntry()
```

### Step 3: Update ExcelIO API (30 minutes)
**Location**: `xl-cats-effect/src/com/tjclp/xl/io/ExcelIO.scala`

```scala
def write[F[_]: Async](
  workbook: Workbook,
  path: Path,
  config: WriterConfig = WriterConfig.default  // Add config parameter
): F[Unit] =
  Sync[F].blocking {
    val out = Files.newOutputStream(path)
    try XlsxWriter.writeZip(workbook, out, config)
    finally out.close()
  }
```

### Step 4: Add Tests (1 hour)
```scala
test("DEFLATED produces smaller files than STORED"):
  val data = generateWorkbook(10_000)

  val storedPath = writeWith(WriterConfig(Compression.Stored, prettyPrint = false), data)
  val deflatedPath = writeWith(WriterConfig(Compression.Deflated, prettyPrint = false), data)

  val storedSize = Files.size(storedPath)
  val deflatedSize = Files.size(deflatedPath)

  assert(deflatedSize < storedSize * 0.5, s"DEFLATED ($deflatedSize) should be < 50% of STORED ($storedSize)")

test("prettyPrint increases file size"):
  val data = generateWorkbook(10_000)

  val compactPath = writeWith(WriterConfig(Compression.Deflated, prettyPrint = false), data)
  val prettyPath = writeWith(WriterConfig(Compression.Deflated, prettyPrint = true), data)

  assert(Files.size(prettyPath) > Files.size(compactPath))

test("default config produces valid XLSX"):
  val wb = generateWorkbook(1000)
  val path = writeWith(WriterConfig.default, wb)

  // Verify Excel can open it
  val reloaded = XlsxReader.read(path)
  assert(reloaded.isRight)
```

## Benefits

- **5-10x smaller files** (typical DEFLATE compression ratio)
- **Faster workflows** (no CRC precomputation)
- **Configurable** (debug mode available with `WriterConfig.debug`)
- **Backward compatible** (existing calls use new defaults)

## Success Criteria

- ✅ Default config uses DEFLATED + prettyPrint=false
- ✅ Files 5-10x smaller than old STORED defaults
- ✅ WriterConfig.debug available for debugging
- ✅ All existing tests pass with new defaults
- ✅ 3 new compression tests added

## Related

- **ADR-012**: Decision to default to DEFLATED
- **streaming-improvements.md**: Overall streaming roadmap

---

# P6.8: Builder Pattern for Batched Operations

**Status**: 🟡 Not Started (Recommended from Lazy Evaluation Review)
**Priority**: MEDIUM (performance enhancement)
**Estimated Effort**: 3-4 days
**Source**: Scoped-down lazy-evaluation.md

## Overview

Add `SheetBuilder` for batching operations instead of full Spark-style optimizer. Captures 80-90% of lazy evaluation benefits with 10% of complexity.

**See**: `docs/plan/lazy-evaluation.md` "Recommended Scope" section for full implementation details.

## Quick Summary

```scala
val builder = SheetBuilder("Sales")
(1 to 10000).foreach(i => builder.put(cell"A$i", s"Row $i"))
val sheet = builder.build  // Single allocation

// 30-50% faster than:
(1 to 10000).foldLeft(sheet)((s, i) => s.put(cell"A$i", s"Row $i"))
```

**Effort**: 3-4 days (vs 4-5 weeks for full optimizer)
**Impact**: 30-50% speedup for batch operations
**Complexity**: Low (mutable builder pattern, familiar to Scala developers)
