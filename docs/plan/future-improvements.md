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
