# Performance & Quality Polish

**Status**: 🟡 Partially Complete
**Priority**: Medium
**Estimated Effort**: 3-5 hours remaining (5 hours completed)
**Last Updated**: 2025-11-20

---

## Metadata

| Field | Value |
|-------|-------|
| **Owner Modules** | `xl-ooxml`, `xl-core` (XmlUtil), `xl-testkit` |
| **Touches Files** | `Styles.scala`, `XmlUtil.scala`, test files |
| **Dependencies** | P0-P8 complete |
| **Enables** | WI-16 (streaming optimizations need benchmarks first) |
| **Parallelizable With** | WI-07 (formula), WI-10 (tables), WI-11 (charts) — different modules |
| **Merge Risk** | Low (utilities and tests, minimal core changes) |

---

## Work Items

| ID | Description | Type | Files | Status | PR |
|----|-------------|------|-------|--------|----|
| `P6.5-01` | O(n²) style indexOf optimization | Perf | `Styles.scala` | ✅ Done | (79b3269) |
| `P6.5-02` | Fix non-deterministic allFills ordering | Fix | `Styles.scala` | ✅ Done | (79b3269) |
| `P6.5-03` | XXE security vulnerability fix | Security | `XlsxReader.scala` | ✅ Done | (79b3269) |
| `P6.5-04` | Extract needsXmlSpacePreserve utility | Quality | `XmlUtil.scala` | ✅ Done | (commit ref) |
| `P6.5-05` | XlsxReader error path tests | Test | `XlsxReaderErrorSpec.scala` | ✅ Done | (commit ref) |
| `P6.5-06` | Full round-trip integration test | Test | `OoxmlRoundTripSpec.scala` | ✅ Done | (commit ref) |
| `P6.5-07` | Logging strategy documentation | Docs | `error-model-and-safety.md` | ⏳ Deferred to P13 | - |

---

## Dependencies

### Prerequisites (Complete)
- ✅ P0-P8: Foundation (all core features operational)

### Enables
- WI-15: Benchmark Suite (infrastructure for measuring improvements)
- WI-16: Streaming Optimizations (needs perf baselines first)

### File Conflicts
- None — changes are localized to utilities and tests

### Safe Parallelization
- ✅ WI-07 (Formula Parser) — different module (xl-evaluator)
- ✅ WI-10 (Tables) — different feature area
- ✅ WI-11 (Charts) — different feature area

---

## Worktree Strategy

**Branch naming**: `perf-polish` or `P6.5-<item-id>`

**Merge order**: Sequential for safety (small PRs, low risk)

**Conflict resolution**: Minimal risk — utilities don't conflict with feature work

---

## Execution Algorithm

### Completed Work Items (Historical)
The following work items were completed in previous sessions:

**P6.5-01 through P6.5-06** (✅ Complete):
- Style indexOf optimization (O(n²) → O(1))
- Non-deterministic fills ordering fix
- XXE security hardening
- XmlUtil.needsXmlSpacePreserve utility
- Error path regression tests (10 tests)
- Full round-trip integration test

See git history for implementation details.

### Remaining Work (Deferred to P13)

**P6.5-07: Logging Strategy** (⏳ Deferred to P13):
```
1. Defer to P13 (comprehensive observability strategy)
2. Document logging requirements in error-model-and-safety.md
3. Wait for structured logging decision (not System.err)
```

---

## Definition of Done

### Completed ✅
- [x] Style indexOf optimized to O(1) using Maps
- [x] Non-deterministic allFills ordering fixed
- [x] XXE security vulnerability fixed
- [x] Whitespace check utility (needsXmlSpacePreserve)
- [x] Error path tests added (10 tests in XlsxReaderErrorSpec)
- [x] Full round-trip integration test
- [x] All tests passing (680+ total)
- [x] Code formatted
- [x] Documentation updated

### Deferred to P13
- [ ] Structured logging strategy

---

## Module Ownership

**Primary**: `xl-ooxml` (Styles.scala, XlsxReader.scala, XmlUtil.scala)

**Secondary**: `xl-core` (XmlUtil extensions)

**Test Files**: `xl-ooxml/test` (StylePerformanceSpec, DeterminismSpec, SecuritySpec, XlsxReaderErrorSpec, OoxmlRoundTripSpec)

---

## Merge Risk Assessment

**Risk Level**: Low

**Rationale**:
- Most changes are utilities and tests
- No central domain model changes
- Additive changes only (no refactoring of existing code)

---

## Implementation Details (Historical)

### P6.5-01: O(n²) Style indexOf Optimization

**Goal**: Optimize style component lookups from O(n²) to O(1)

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

**Solution**:
```scala
// Pre-build lookup maps once (O(n))
val fontMap = index.fonts.zipWithIndex.toMap
val fillMap = allFills.zipWithIndex.toMap
val borderMap = index.borders.zipWithIndex.toMap

// Then O(1) lookups
val cellXfsElem = elem("cellXfs", "count" -> index.cellStyles.size.toString)(
  index.cellStyles.map { style =>
    val fontIdx = fontMap(style.font)      // O(1)
    val fillIdx = fillMap(style.fill)      // O(1)
    val borderIdx = borderMap(style.border) // O(1)
  }*
)
```

**Tests**:
- StylePerformanceSpec: 2 tests verifying linear scaling with 1000+ styles

**Result**: ✅ Complete (commit 79b3269)

---

### P6.5-02: Non-Deterministic allFills Ordering

**Goal**: Fix non-deterministic `.distinct` causing unstable XML output

**Problem**:
```scala
val allFills = (defaultFills ++ index.fills.filterNot(defaultFills.contains)).distinct
```

The `.distinct` method uses Set internally with non-deterministic iteration order.

**Solution**:
```scala
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

**Tests**:
- DeterminismSpec: 4 tests verifying fill ordering, XML stability

**Result**: ✅ Complete (commit 79b3269)

---

### P6.5-03: XXE Security Vulnerability

**Goal**: Prevent XML External Entity attacks

**Problem**: Default XML parser allows DOCTYPE declarations and external entity references

**Solution**:
```scala
private def parseXml(xmlString: String, location: String): XLResult[Elem] =
  try
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

**Tests**:
- SecuritySpec: 3 tests verifying XXE rejection

**Result**: ✅ Complete (commit 79b3269)

---

### P6.5-04: Extract needsXmlSpacePreserve Utility

**Goal**: DRY — centralize whitespace check logic

**Files**: 5 occurrences → 1 utility function

**Solution**:
```scala
// In XmlUtil object
def needsXmlSpacePreserve(s: String): Boolean =
  s.startsWith(" ") || s.endsWith(" ") || s.contains("  ")
```

**Result**: ✅ Complete

---

### P6.5-05: XlsxReader Error Path Tests

**Goal**: Comprehensive error handling test coverage

**Tests Added** (10 new tests in `XlsxReaderErrorSpec.scala`):
- Malformed XML
- Missing parts
- Corrupt ZIPs
- Invalid relationships
- Bad shared-string indices

**Result**: ✅ Complete

---

### P6.5-06: Full Round-Trip Integration Test

**Goal**: End-to-end round-trip verification with deterministic output

**Test**: `OoxmlRoundTripSpec.scala`
- Builds 3-sheet workbook with all cell types, styles, merges
- Asserts domain equality and byte-identical output

**Result**: ✅ Complete

---

### P6.5-07: Logging Strategy (Deferred to P13)

**Goal**: Structured logging for missing/malformed files

**Deferred Rationale**:
- Requires comprehensive logging strategy (not System.err)
- Part of broader observability work in P13
- Not blocking for current use cases

**Action**: Document in `error-model-and-safety.md` as P13 work item

---

## Quick Wins (Backlog)

These are low-effort improvements for future sessions:

### Documentation Wins
- **Deprecate cell"..."/putMixed in docs** — update examples to use `ref"A1"` and batch `put`
- **Tighten Quick Start** — rewrite with current API
- **Clarify streaming trade-offs** — use io-modes.md as canonical reference

### Small Code Wins
- **Formatted.putFormatted** — delegate to style-aware Sheet.put
- **writeToStream helper** — avoid temp file in writeToBytes
- **Workbook.updateEither** — reduce boilerplate for sheet operations

---

## Related Documentation

- **Roadmap**: `docs/plan/roadmap.md` (Phase P6.5)
- **PR #4**: Source of feedback
- **Design**: `docs/design/wartremover-policy.md`, `io-modes.md`
- **Tests**: `xl-ooxml/test/src/com/tjclp/xl/ooxml/*Spec.scala`

---

## Notes

This phase is **not blocking** for production use. All critical fixes (performance, security, determinism) are complete. Remaining work is enhancements and documentation polish.

**Status Summary**:
- ✅ P6.5-01 through P6.5-06: Complete
- ⏸ P6.5-07: Deferred to P13
- 📋 Quick Wins: Backlog for future sessions
