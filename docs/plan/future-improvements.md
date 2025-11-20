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

✅ **Status**: Complete (see `xl-ooxml/test/src/com/tjclp/xl/ooxml/XlsxReaderErrorSpec.scala`)

- Added 10 targeted regression tests covering malformed XML, missing parts, corrupt ZIPs, invalid relationships, and bad shared-string indices.
- Each fixture is built on the fly (no external files) and asserts precise `XLError` messages or `CellError` fallback behavior.
- Scenarios now include:
  - Missing/malformed `workbook.xml`
  - Missing `sheets` node
  - Malformed `styles.xml`
  - Missing worksheet parts and relationship targets
  - Missing `[Content_Types].xml`
  - Non-ZIP / truncated ZIP payloads
  - Invalid shared strings references producing `#REF!`

---

### 5. Full XLSX Round-Trip Integration Test

✅ **Status**: Complete (`xl-ooxml/test/src/com/tjclp/xl/ooxml/OoxmlRoundTripSpec.scala`)

- Builds a 3-sheet workbook exercising:
  - All cell value variants (Text, Number, Bool, DateTime, Error)
  - Shared strings with deduplication
  - Rich styling (fonts, fills, borders, alignments)
  - Merged ranges and multi-sheet relationships
- Asserts:
  - Domain equality between original and round-tripped workbooks (with special handling for Excel date serials).
  - Byte-identical output across write → read → write by normalizing ZIP metadata (fixed timestamps) and preserving style registry ordering.
- Added diagnostics to highlight offending ZIP entries if the deterministic guarantee regresses.

---

## Definition of Done

### Completed ✅
- [x] Style indexOf optimized to O(1) using Maps (Issue #1)
- [x] Non-deterministic allFills ordering fixed (Issue #2)
- [x] XXE security vulnerability fixed (Issue #3)
- [x] Benchmark tests verify linear scaling (StylePerformanceSpec: 2 tests)
- [x] Determinism tests verify stable output (DeterminismSpec: 4 tests)
- [x] Security tests verify XXE protection (SecuritySpec: 3 tests)
- [x] Whitespace preservation check centralized in `XmlUtil.needsXmlSpacePreserve`
- [x] Error-path regression suite for XlsxReader (10 tests in `XlsxReaderErrorSpec`)
- [x] Full-feature round-trip integration test with deterministic ZIP output
- [x] Reader warnings surfaced via `XlsxReader.readWithWarnings` with effectful logging hook in `ExcelIO`
- [x] All tests passing (636 total as of 2025-11-16)
- [x] Code formatted (Scalafmt 3.10.1)
- [x] Documentation updated (this file, roadmap.md)
- [x] Committed (79b3269)

### Deferred (Medium Priority)
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
# Quick Wins and Low-Effort Improvements

This document captures “small but high‑leverage” improvements that are good starter tasks and keep the codebase and documentation sharp.

The items here are intentionally scoped to be safe and incremental; most should not affect core semantics.

## Documentation Wins

- **Keep top‑level status/docs authoritative**
  - Treat `docs/STATUS.md`, `docs/LIMITATIONS.md`, `docs/design/architecture.md`, and the root `README.md` as the main external story.
  - When behavior changes, update these first; older deep‑dive docs can be left as historical context.

- **Deprecate `cell"..."/putMixed` terminology in docs**
  - Many reference docs still use the old `cell"A1"` / `range"A1:B10"` / `putMixed` names.
  - Bring them in line with the current API:
    - Use `ref"A1"` / `ref"A1:B10"` from `com.tjclp.xl.syntax.*`.
    - Refer to “batch `Sheet.put(ref -> value, ...)`” instead of `putMixed`.

- **Tighten Quick Start and migration guides**
  - `docs/QUICK-START.md` and `docs/reference/migration-from-poi.md` contain pre‑ref API examples and some non‑compiling snippets.
  - Rewrite them around:
    - `import com.tjclp.xl.*` (single import ergonomics).
    - `Sheet("Name")` / `Workbook.empty` + `Workbook.put` / `Workbook.update`.
    - `ref"...", fx"...", money"...", percent"..."` macros.

- **Clarify streaming trade‑offs in one place**
  - Use `docs/design/io-modes.md` + `docs/reference/performance-guide.md` as the canonical description of:
    - In‑memory vs streaming I/O.
    - What is preserved (styles, merges, SST) in each path.
    - When to choose which mode.

## Small Code-Facing Wins (Behavior Already Implemented)

These are **documentation alignment** tasks first; implementation is already in place:

- **Merged cells**
  - In‑memory path (`XlsxReader`/`XlsxWriter` + `OoxmlWorksheet`) already round‑trips `Sheet.mergedRanges` via `<mergeCells>`.
  - Make sure all docs say:
    - “Merged cells are supported in in‑memory I/O.”
    - “Streaming writers do not currently emit merges.”

- **Streaming read behavior**
  - `ExcelIO.readStream` / `readSheetStream` / `readStreamByIndex` now use `fs2.io.readInputStream` and fs2‑data‑xml.
  - Clean up any lingering references to the old “broken streaming read” implementation; treat P6.6 as done and focus on feature coverage trade‑offs instead. ✅ (done)

## Small Code-Facing Wins (Potential Future Changes)

These are implementation candidates that are safe and localized, but are *not* implemented yet.

- **`Formatted.putFormatted` should reuse the style‑aware `Sheet.put`**
  - Today `Formatted.putFormatted` writes only the value; `Sheet.put(ref -> formatted)` already creates and registers a `CellStyle` with the right `NumFmt`.
  - Implementation win: delegate to the existing `Sheet.put` path so `putFormatted` preserves number formats consistently.

- **Public stream‑based writer helper**
  - `XlsxWriter` already supports writing to an `OutputStreamTarget`; `writeToBytes` currently uses a temp file.
  - Implementation win: add a `writeToStream` helper and re‑implement `writeToBytes` without touching disk (use `ByteArrayOutputStream` instead).

- **Workbook helpers for “update with Either”**
  - Common pattern: `XlsxReader.read(path)` → `Workbook` → sheet‑level operations that return `XLResult[Sheet]` → `Workbook.put`.
  - Implementation win: add a helper like `Workbook.updateEither(name)(Sheet => XLResult[Sheet])` to reduce boilerplate in callers.

## How to Use This List

- Treat this file as a backlog of “good first issues” and cleanups.
- When you land one of these improvements:
  - Update this document and `docs/STATUS.md` if the item materially changes behavior.
  - Prefer small, isolated PRs that tackle one bullet at a time.
