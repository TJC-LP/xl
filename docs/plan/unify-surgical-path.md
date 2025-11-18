# Unify Surgical and Normal Write Paths

**Status**: Ready for Implementation
**Priority**: P1 (High-value refactoring, medium risk)
**Branch**: feat/surgical-mod (or experimental branch)
**Estimated Effort**: 5-7 hours

---

## Current State (Starting Point)

### What We've Accomplished

As of commit 839b27e, we've fixed the critical surgical write bugs:

1. ✅ **NumFmt ID preservation**: Style 283 has numFmtId=39 (was 0)
2. ✅ **Style count stability**: 647 styles preserved (was 648)
3. ✅ **Style index accuracy**: 0 cell shifts (was 238 corrupted)
4. ✅ **AST overhaul**: WorkbookStyles preserves fonts/fills/borders

**Validation**: `data/validate-surgical.sc` shows 0 critical regressions

### The Current Architecture

**Two separate code paths:**

1. **Normal write** (no source):
   - `StyleIndex.fromWorkbook(wb)` - full deduplication
   - File: `xl-ooxml/src/com/tjclp/xl/ooxml/Styles.scala:58-130`
   - 73 LOC

2. **Surgical write** (with source):
   - `StyleIndex.fromWorkbookSurgical(wb, modified, path)` - preserve original
   - File: `xl-ooxml/src/com/tjclp/xl/ooxml/Styles.scala:162-265`
   - 104 LOC

**Dispatch logic:**
```scala
// xl-ooxml/src/com/tjclp/xl/ooxml/XlsxWriter.scala:708-714
val (styleIndex, sheetRemappings) = sourceContext match
  case Some(ctx) =>
    StyleIndex.fromWorkbookSurgical(workbook, tracker.modifiedSheets, ctx.sourcePath)
  case None =>
    StyleIndex.fromWorkbook(workbook)
```

**Problem**:
- Users must choose between methods
- 177 LOC with ~23% duplication
- Dispatch logic split across files
- Preservation is API concern, not implementation detail

---

## The Goal: Unified Architecture

### Vision

**Single public method with automatic optimization:**

```scala
// Single call site (XlsxWriter.scala)
val (styleIndex, sheetRemappings) = StyleIndex.fromWorkbook(workbook)
```

**Internal dispatch based on context:**

```scala
// StyleIndex.scala (new design)
object StyleIndex:
  def fromWorkbook(wb: Workbook): (StyleIndex, Map[Int, Map[Int, Int]]) =
    wb.sourceContext match
      case Some(ctx) =>
        fromWorkbookWithSource(wb, ctx)
      case None =>
        fromWorkbookWithoutSource(wb)

  private def fromWorkbookWithSource(...) = // Current fromWorkbookSurgical logic
  private def fromWorkbookWithoutSource(...) = // Current fromWorkbook logic
```

### Benefits

1. **Simpler API**: 1 public method (not 2)
2. **No user choice**: Automatic optimization
3. **Less duplication**: 177 LOC → ~185 LOC (shared helpers)
4. **Clearer semantics**: Preservation is transparent
5. **Single responsibility**: StyleIndex owns strategy, not XlsxWriter

---

## Implementation Plan

### Phase 1: Unify StyleIndex Methods (Core)

**Estimated time**: 3 hours
**Risk**: Medium

#### Files to Modify

**Primary file**: `xl-ooxml/src/com/tjclp/xl/ooxml/Styles.scala`

**Changes:**

1. **Rename existing methods** (make private):
   ```scala
   // Line 58: Rename fromWorkbook → fromWorkbookWithoutSource
   private def fromWorkbookWithoutSource(wb: Workbook): (StyleIndex, Map[Int, Map[Int, Int]]) =
     // Existing fromWorkbook logic (lines 58-130)

   // Line 162: Rename fromWorkbookSurgical → fromWorkbookWithSource
   private def fromWorkbookWithSource(
     wb: Workbook,
     ctx: SourceContext
   ): (StyleIndex, Map[Int, Map[Int, Int]]) =
     val tracker = ctx.modificationTracker
     // Existing fromWorkbookSurgical logic (lines 162-265)
     // But take sourcePath from ctx, not as parameter
   ```

2. **Add new public dispatcher**:
   ```scala
   /**
    * Build unified style index from workbook with automatic optimization.
    *
    * Strategy (automatic based on workbook.sourceContext):
    *   - With source: Preserve original styles (byte-perfect surgical modification)
    *   - Without source: Full deduplication (optimal compression)
    *
    * Users don't choose - method transparently optimizes based on available context.
    */
   def fromWorkbook(wb: Workbook): (StyleIndex, Map[Int, Map[Int, Int]]) =
     wb.sourceContext match
       case Some(ctx) =>
         fromWorkbookWithSource(wb, ctx)
       case None =>
         fromWorkbookWithoutSource(wb)
   ```

3. **Update fromWorkbookWithSource signature**:
   - Remove `modifiedSheetIndices: Set[Int]` parameter (get from ctx.modificationTracker)
   - Remove `sourcePath: Path` parameter (get from ctx.sourcePath)
   - Add `ctx: SourceContext` parameter

**Secondary file**: `xl-ooxml/src/com/tjclp/xl/ooxml/XlsxWriter.scala`

**Changes:**

1. **Remove dispatch logic** (lines 708-714):
   ```scala
   // BEFORE (7 lines):
   val (styleIndex, sheetRemappings) = sourceContext match
     case Some(ctx) =>
       StyleIndex.fromWorkbookSurgical(workbook, tracker.modifiedSheets, ctx.sourcePath)
     case None =>
       StyleIndex.fromWorkbook(workbook)

   // AFTER (1 line):
   val (styleIndex, sheetRemappings) = StyleIndex.fromWorkbook(workbook)
   ```

---

### Phase 2: Extract Shared Logic (Optional Cleanup)

**Estimated time**: 1 hour
**Risk**: Low

**Potential shared helpers:**

```scala
// Both methods have similar deduplication loops
private def buildRemappingTable(
  sheets: Vector[Sheet],
  sheetIndices: Set[Int],
  unifiedIndex: Map[String, ?],
  unifiedStyles: Vector[CellStyle]
): Map[Int, Map[Int, Int]] = ...

// Both methods build component vectors
private def deduplicateComponents(
  styles: Vector[CellStyle]
): (Vector[Font], Vector[Fill], Vector[Border], Vector[(Int, NumFmt)]) = ...
```

**Defer to later**: Can extract helpers in follow-up PR if duplication bothers us.

---

### Phase 3: Update Documentation

**Estimated time**: 1-2 hours
**Risk**: Low

#### Files to Update

1. **`CLAUDE.md`** (lines 186-201 - "Reviewing PR Feedback" section)
   - Remove references to "surgical modification" as user-facing concept
   - Update to: "XL automatically preserves structure when reading from existing files"
   - Example:
     ```scala
     // Old docs:
     // "Use fromWorkbookSurgical for byte-perfect preservation"

     // New docs:
     // "XL automatically preserves structure from source files (transparent optimization)"
     ```

2. **`docs/design/domain-model.md`**
   - Update StyleIndex section to describe unified behavior
   - Remove split path explanation

3. **Method documentation** (already good):
   - The new `fromWorkbook` doc explains automatic optimization
   - Private method docs can stay detailed

**Note**: Tests don't need updates (they call through XlsxWriter.write, not StyleIndex directly)

---

### Phase 4: Regression Testing

**Estimated time**: 30 minutes
**Risk**: Low (automated)

#### Test Scenarios

**Scenario 1: New workbook (no source)**
```scala
val wb = Workbook.empty("Sheet1")
  .flatMap(_.put(ref"A1" -> "Hello"))
XlsxWriter.write(wb, output)
```
**Expected**: Full deduplication, optimal XLSX

**Scenario 2: Modified workbook (with source)**
```scala
val wb = XlsxReader.read(source).flatMap { original =>
  original("Sheet1").flatMap(_.put(ref"A1" -> "Modified"))
}
XlsxWriter.write(wb, output)
```
**Expected**: Preserve original styles, regenerate only modified sheet

**Scenario 3: Syndigo real-world**
```bash
cd data && scala-cli run surgical-demo.sc
cd data && scala-cli run validate-surgical.sc
```
**Expected**:
- 647 styles (not 648)
- 0 style index shifts
- 0 style property mismatches

#### Automated Tests

Run full suite:
```bash
./mill __.compile
./mill __.test
```

**Expected**: 636/636 tests passing (no regressions)

**Critical test suites:**
- `StyleSpec` - canonicalKey, NumFmt.fromId
- `OoxmlRoundTripSpec` - round-trip preservation
- `XlsxWriterCorruptionRegressionSpec` - surgical write correctness
- `XlsxWriterSurgicalSpec` - surgical behavior (should still work)
- `XlsxWriterStyleMetadataSpec` - style preservation

---

## Implementation Checklist

### Phase 1: Core Refactor (Must Do)
- [ ] Rename `fromWorkbook` → `fromWorkbookWithoutSource` (make private)
- [ ] Rename `fromWorkbookSurgical` → `fromWorkbookWithSource` (make private)
- [ ] Update `fromWorkbookWithSource` signature (take ctx: SourceContext)
- [ ] Add public `fromWorkbook(wb: Workbook)` dispatcher
- [ ] Update XlsxWriter.scala to use single `fromWorkbook()` call
- [ ] Compile: `./mill __.compile`
- [ ] Test: `./mill __.test` (all pass)

### Phase 2: Documentation (Should Do)
- [ ] Update CLAUDE.md (remove surgical user-facing terminology)
- [ ] Update method documentation
- [ ] Update docs/design/domain-model.md if needed

### Phase 3: Validation (Must Do)
- [ ] Run Syndigo demo: `cd data && scala-cli run surgical-demo.sc`
- [ ] Run validation: `cd data && scala-cli run validate-surgical.sc`
- [ ] Verify 0 regressions
- [ ] Test new workbook creation (no source)
- [ ] Test modified workbook (with source)

### Phase 4: Commit
- [ ] Format code: `./mill __.reformat`
- [ ] Commit: "Refactor: Unify surgical and normal write paths into single API"
- [ ] Push to feat/surgical-mod

---

## Files Reference

### Primary Files (Core Changes)

1. **`xl-ooxml/src/com/tjclp/xl/ooxml/Styles.scala`**
   - Lines 58-130: `fromWorkbook` (rename to `fromWorkbookWithoutSource`)
   - Lines 162-265: `fromWorkbookSurgical` (rename to `fromWorkbookWithSource`)
   - Add new public `fromWorkbook` dispatcher after line 57

2. **`xl-ooxml/src/com/tjclp/xl/ooxml/XlsxWriter.scala`**
   - Lines 708-714: Remove dispatch logic
   - Replace with: `val (styleIndex, sheetRemappings) = StyleIndex.fromWorkbook(workbook)`

### Documentation Files (Optional)

3. **`CLAUDE.md`**
   - Section "Reviewing PR Feedback" (lines ~186-201)
   - Remove "surgical modification" user-facing terminology

4. **`docs/design/domain-model.md`**
   - StyleIndex section
   - Update to describe unified behavior

### Test Files (Should Not Need Changes)

- Tests call `XlsxWriter.write()`, not `StyleIndex` directly
- Tests should pass transparently (behavior unchanged)
- May update test names/comments to remove "surgical" terminology (cosmetic)

---

## Success Criteria

### Must Pass
1. ✅ All 636 tests pass
2. ✅ Syndigo validation shows 0 regressions
3. ✅ Single public method (`StyleIndex.fromWorkbook`)
4. ✅ No user-facing API changes (transparent refactor)

### Should Verify
5. ✅ Code formatted (`./mill __.checkFormat`)
6. ✅ New workbook creation still works
7. ✅ Modified workbook preservation still works
8. ✅ Documentation updated

---

## Potential Slash Command

**File**: `.claude/commands/unify-surgical.md`

```markdown
Read the plan at docs/plan/unify-surgical-path.md and implement the StyleIndex unification.

Steps:
1. Read xl-ooxml/src/com/tjclp/xl/ooxml/Styles.scala (current implementation)
2. Rename existing methods to private helpers
3. Add public dispatcher that checks wb.sourceContext
4. Update XlsxWriter.scala to use single method
5. Compile and test
6. Validate with Syndigo
7. Commit if all tests pass

Expected outcome:
- 1 public method (not 2)
- All 636 tests pass
- 0 Syndigo regressions
- Cleaner API
```

---

## Context for Fresh Session

### What's Been Done (Don't Redo)

1. ✅ NumFmt.builtInById map completed (IDs 23-48 added)
2. ✅ CellStyle.numFmtId field added (for preservation)
3. ✅ WorkbookStyles extended to store fonts/fills/borders
4. ✅ fromWorkbookSurgical preserves original components
5. ✅ buildStyleRegistry starts empty (not with default)
6. ✅ Tests added for canonicalKey and NumFmt.fromId
7. ✅ Validation script created (validate-surgical.sc)

### What Needs Doing (This Plan)

1. ❌ Unify `fromWorkbook` and `fromWorkbookSurgical` into single API
2. ❌ Remove dispatch logic from XlsxWriter
3. ❌ Update documentation to remove "surgical" user-facing terminology

### Key Insights to Remember

**Critical Design Principle:**
- `canonicalKey` = visual equivalence (for deduplication)
- `numFmtId` = byte-perfect preservation (separate concern)
- **Never include numFmtId in canonicalKey!**

**Why Surgical Preservation Works Now:**
- WorkbookStyles stores original fonts/fills/borders
- fromWorkbookSurgical uses those original vectors (no deduplication)
- Exact fontId/fillId/borderId indices preserved

**Why We Had 648 Styles:**
- StyleRegistry.default added extra style when source style 0 had different properties
- Fixed by starting buildStyleRegistry with empty registry

---

## Step-by-Step Implementation Guide

### Step 1: Rename Existing Methods

**File**: `xl-ooxml/src/com/tjclp/xl/ooxml/Styles.scala`

**Line 58**: Change method signature
```scala
// BEFORE:
def fromWorkbook(wb: Workbook): (StyleIndex, Map[Int, Map[Int, Int]]) =

// AFTER:
private def fromWorkbookWithoutSource(wb: Workbook): (StyleIndex, Map[Int, Map[Int, Int]]) =
```

**Line 162**: Change method signature and parameters
```scala
// BEFORE:
@SuppressWarnings(...)
def fromWorkbookSurgical(
  wb: Workbook,
  modifiedSheetIndices: Set[Int],
  sourcePath: java.nio.file.Path
): (StyleIndex, Map[Int, Map[Int, Int]]) =

// AFTER:
private def fromWorkbookWithSource(
  wb: Workbook,
  ctx: SourceContext
): (StyleIndex, Map[Int, Map[Int, Int]]) =
  val tracker = ctx.modificationTracker
  val sourcePath = ctx.sourcePath
  val modifiedSheetIndices = tracker.modifiedSheets
  // ... rest of method unchanged
```

### Step 2: Add Public Dispatcher

**File**: `xl-ooxml/src/com/tjclp/xl/ooxml/Styles.scala`

**Location**: After line 57 (before the old fromWorkbook)

```scala
  /**
   * Build unified style index from workbook with automatic optimization.
   *
   * Strategy (automatic based on workbook.sourceContext):
   *   - **With source**: Preserve original styles for byte-perfect surgical modification
   *   - **Without source**: Full deduplication for optimal compression
   *
   * Users don't choose the strategy - the method transparently optimizes based on available
   * context. This enables read-modify-write workflows to preserve structure automatically while
   * allowing programmatic creation to produce optimal output.
   *
   * @param wb
   *   The workbook to index
   * @return
   *   (StyleIndex for writing, Map[sheetIndex -> styleId remapping])
   */
  def fromWorkbook(wb: Workbook): (StyleIndex, Map[Int, Map[Int, Int]]) =
    wb.sourceContext match
      case Some(ctx) =>
        // Has source: surgical mode (preserve original structure)
        fromWorkbookWithSource(wb, ctx)
      case None =>
        // No source: full deduplication (optimal compression)
        fromWorkbookWithoutSource(wb)
```

### Step 3: Update XlsxWriter Dispatch

**File**: `xl-ooxml/src/com/tjclp/xl/ooxml/XlsxWriter.scala`

**Lines 708-714**: Replace dispatch with single call

```scala
// BEFORE (7 lines):
val (styleIndex, sheetRemappings) = sourceContext match
  case Some(ctx) =>
    // Surgical: preserve original styles, deduplicate only modified sheets
    StyleIndex.fromWorkbookSurgical(workbook, tracker.modifiedSheets, ctx.sourcePath)
  case None =>
    // No source: full deduplication across all sheets
    StyleIndex.fromWorkbook(workbook)

// AFTER (1 line):
val (styleIndex, sheetRemappings) = StyleIndex.fromWorkbook(workbook)
```

**Note**: Remove the comment about surgical vs normal - it's now an implementation detail.

### Step 4: Compile and Test

```bash
cd /Users/rcaputo3/git/xl

# Compile (should succeed)
./mill __.compile

# Run all tests (should all pass)
./mill __.test

# Expected: 636/636 tests passing
```

If compilation fails:
- Check that `SourceContext` is imported in Styles.scala
- Verify method signatures match exactly

If tests fail:
- Check which tests fail
- Likely need to update test that directly calls `fromWorkbookSurgical`
- Search for: `grep -r "fromWorkbookSurgical" xl-*/test/`

### Step 5: Validate with Syndigo

```bash
# Publish to ivy2Local
./mill __.publishLocal

# Run Syndigo demo
cd data
scala-cli run surgical-demo.sc

# Run validation
scala-cli run validate-surgical.sc
```

**Expected output:**
```
Style count: 647 (not 648)
Style index shifts: 0 (not 238)
Style property mismatches: 0
Value differences: 5 (whitespace only, acceptable)
```

### Step 6: Commit

```bash
git add -A
git commit -m "Refactor: Unify surgical and normal write paths into single API

Merge StyleIndex.fromWorkbook and fromWorkbookSurgical into single method with
automatic optimization based on workbook.sourceContext.

Benefits:
- Single public method (was 2 separate methods)
- No user choice required (automatic optimization)
- Cleaner API surface (preservation is transparent)
- Removed dispatch logic from XlsxWriter

Implementation:
- Renamed fromWorkbook → fromWorkbookWithoutSource (private)
- Renamed fromWorkbookSurgical → fromWorkbookWithSource (private)
- Added public fromWorkbook dispatcher (checks sourceContext internally)
- Updated XlsxWriter to use unified method

Behavior unchanged:
- With source: Preserves original styles (byte-perfect)
- Without source: Full deduplication (optimal)

Tests: 636/636 passing
Validated: Syndigo output (0 regressions)

🤖 Generated with [Claude Code](https://claude.com/claude-code)

Co-Authored-By: Claude <noreply@anthropic.com>"

git push
```

---

## Troubleshooting

### If Tests Fail

**Check for direct calls to renamed methods:**
```bash
grep -r "fromWorkbookSurgical" xl-*/test/
```

**Update any direct test calls:**
```scala
// If test directly calls fromWorkbookSurgical:
// BEFORE:
StyleIndex.fromWorkbookSurgical(wb, modifiedSheets, path)

// AFTER:
StyleIndex.fromWorkbook(wb)  // Assumes wb has sourceContext
```

### If Validation Fails

**Check Syndigo output:**
```bash
cd data

# Check style count
unzip -p syndigo-surgical-output.xlsx xl/styles.xml | \
  xmllint --format - | grep '<cellXfs count='

# Expected: <cellXfs count="647">

# Run full validation
scala-cli run validate-surgical.sc
```

**If regressions appear:**
- Verify `fromWorkbookWithSource` logic unchanged
- Check that sourcePath/modifiedSheets extracted correctly from ctx
- Add debug output to trace execution

---

## Validation Metrics

**Before Unification:**
- Public methods: 2
- LOC duplication: 40 lines (23%)
- User choice: Required (which method?)

**After Unification:**
- Public methods: 1
- LOC duplication: 0 lines
- User choice: None (automatic)

**Quality gates:**
- ✅ Tests: 636/636 passing
- ✅ Syndigo validation: 0 regressions
- ✅ Style count: 647
- ✅ Style shifts: 0
- ✅ Style properties: All match

---

## Future Work (Out of Scope)

**Potential further unification:**
1. Extract shared deduplication logic into helpers (Phase 2)
2. Unify `fromDomain` methods in OoxmlWorksheet (different purpose, not priority)
3. Create `OoxmlPreservationLayer` trait for all components (over-engineering)

**Recommendation**: Stop after Phase 1. The 23% duplication is acceptable given clear separation of concerns.

---

## References

- **Current implementation**: Commits 3e1362b, a892b7f, 839b27e
- **Validation tool**: `data/validate-surgical.sc`
- **Test suite**: 636 tests covering all scenarios
- **Related docs**:
  - `docs/plan/numfmt-id-preservation.md` (completed fixes)
  - `docs/design/domain-model.md` (StyleIndex architecture)

---

## Quick Start (For Fresh Session)

```bash
# 1. Read this plan
cat docs/plan/unify-surgical-path.md

# 2. Read current implementation
code xl-ooxml/src/com/tjclp/xl/ooxml/Styles.scala

# 3. Make changes per Step 1-3

# 4. Test
./mill __.compile && ./mill __.test

# 5. Validate
cd data && scala-cli run validate-surgical.sc

# 6. Commit and push
git add -A && git commit -m "..." && git push
```

**Estimated time**: 5-7 hours (3 hours core + 2 hours docs/validation)
