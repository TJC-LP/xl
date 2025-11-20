# Type Class Consolidation for Easy Mode `put()` API

> **Status**: ✅ Completed – This plan has been fully implemented in PR #20.
> This document is retained as a historical design/implementation record.
> **Archived**: 2025-11-20

**Status**: ✅ Implemented (PR #20)
**Author**: Claude Code Agent
**Date**: 2025-11-19 (Design) / 2025-11-20 (Implementation)
**Context**: PR #19 review identified 120 lines of duplicated `put()` overloads
**Goal**: Reduce 36 overloads to ~4-6 generic methods using type classes
**Result**: 120 LOC → 50 LOC (58% reduction), 583 → 586 tests (+3 for NumFmt bug fix)

---

## Executive Summary

The Easy Mode API in `xl-core/src/com/tjclp/xl/extensions.scala` contains significant duplication:
- **18 overloads on `Sheet`** (9 unstyled + 9 styled)
- **18 overloads on `XLResult[Sheet]`** (9 unstyled + 9 styled, chainable variants)
- **Total**: 120 lines of nearly identical code

This proposal introduces a lightweight type class to consolidate these overloads while maintaining:
- ✅ **Type inference** (same user experience)
- ✅ **Error quality** (same error messages)
- ✅ **Zero overhead** (inline given instances)
- ✅ **Auto-format inference** (LocalDate → Date format, BigDecimal → Decimal format)
- ✅ **Extensibility** (users can add custom types)

**Expected reduction**: 120 LOC → 30-40 LOC (67-75% reduction)

---

## 1. Type Class Design

### 1.1 Name Selection

**Chosen**: Reuse existing `CellWriter[A]`

**Rationale**:
- **Already exists**: `CellWriter[A]` trait in `xl-core/src/com/tjclp/xl/codec/CellWriter.scala`
- **DRY principle**: Don't duplicate identical interfaces
- **Infrastructure**: All 9 inline given instances already exist in `CellCodec`
- **Separation of concerns**: Codec for batch operations, extensions leverage writer interface
- **Consistency**: Aligns with existing `CellReader[A]` for symmetry

**Alternatives considered**:
- ❌ `CellValueEncoder[A]` - Would duplicate existing `CellWriter` interface
- ❌ `ToCellValue[A]` - Too Java-ish, not idiomatic Scala 3
- ❌ New trait - Violates DRY principle

### 1.2 Interface (Already Exists)

```scala
package com.tjclp.xl.codec

import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.styles.CellStyle

/**
 * Write a typed value to CellValue with optional style hint.
 *
 * Part of the cell codec system. Also used by Easy Mode extensions
 * for generic put() operations with automatic NumFmt inference.
 */
trait CellWriter[A]:
  /**
   * Encode value to CellValue with optional style hint.
   *
   * The optional CellStyle is used for NumFmt auto-inference:
   * - LocalDate/LocalDateTime → Date/DateTime format
   * - BigDecimal → Decimal format
   * - Numeric types → General format
   * - Text/Boolean/RichText → None (no special formatting)
   */
  def write(a: A): (CellValue, Option[CellStyle])

object CellWriter:
  /** Summon the writer instance for type A */
  def apply[A](using w: CellWriter[A]): CellWriter[A] = w
```

### 1.3 Relationship to Existing CellCodec

**Key insight**: `CellCodec[A]` extends `CellWriter[A]`, so all codec instances are already writer instances.

**Benefit**: Zero new code needed for type class instances!

```scala
// In CellCodec.scala (already exists)
trait CellCodec[A] extends CellReader[A] with CellWriter[A]

// All 9 inline given instances are already CellWriter instances
inline given CellCodec[String]        // ✓ Also CellWriter[String]
inline given CellCodec[Int]           // ✓ Also CellWriter[Int]
// ... etc for all 9 types
```

---

## 2. Given Instances

### 2.1 Instance Location

**Location**: Re-export from `CellCodec` object in `xl-core/src/com/tjclp/xl/codec/CellCodec.scala`

**Why**:
- All 9 instances already exist as `inline given CellCodec[T]`
- `CellCodec[A]` extends `CellWriter[A]`, so all instances are already `CellWriter` instances
- Zero code duplication

### 2.2 Instance List (Already Implemented)

All 9 types already have inline given instances:

```scala
// Already exists in CellCodec.scala (lines 26-168)
inline given CellWriter[String]        // → CellValue.Text(s), None
inline given CellWriter[Int]           // → CellValue.Number(n), Some(NumFmt.General)
inline given CellWriter[Long]          // → CellValue.Number(n), Some(NumFmt.General)
inline given CellWriter[Double]        // → CellValue.Number(n), Some(NumFmt.General)
inline given CellWriter[BigDecimal]    // → CellValue.Number(n), Some(NumFmt.Decimal)
inline given CellWriter[Boolean]       // → CellValue.Bool(b), None
inline given CellWriter[LocalDate]     // → CellValue.DateTime(date.atStartOfDay), Some(NumFmt.Date)
inline given CellWriter[LocalDateTime] // → CellValue.DateTime(dt), Some(NumFmt.DateTime)
inline given CellWriter[RichText]      // → CellValue.RichText(rt), None
```

**Special case: RichText**
- RichText already wraps `CellValue.RichText`, so encoding is identity-like
- No cell-level style (formatting is in TextRun properties)

### 2.3 Zero-Overhead Guarantee

**Mechanism**: `inline given` instances ensure zero runtime overhead
- Compiler inlines all instance calls
- No boxing, no vtable lookups
- Equivalent to hand-written overloads

**Verification**: Compare bytecode of type class version vs explicit overloads (future work, not blocking)

---

## 3. API Design

### 3.1 Method Signatures

**File**: `xl-core/src/com/tjclp/xl/extensions.scala`

```scala
package com.tjclp.xl

import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.codec.CellWriter
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.richtext.RichText
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.workbooks.Workbook

import java.time.{LocalDate, LocalDateTime}

@SuppressWarnings(Array("org.wartremover.warts.Throw"))
object extensions:

  // Helper (unchanged)
  private def toXLResult[A](
    either: Either[String, A],
    ref: String,
    errorType: String
  ): XLResult[A] =
    either.left.map(msg => XLError.InvalidCellRef(ref, s"$errorType: $msg"))

  // ========== Sheet Extensions: Generic Put (Data Only) ==========

  extension (sheet: Sheet)
    /**
     * Put typed value at cell reference.
     *
     * Supports: String, Int, Long, Double, BigDecimal, Boolean,
     * LocalDate, LocalDateTime, RichText. Auto-infers NumFmt based on type.
     *
     * @tparam A The value type (must have CellWriter instance)
     * @param cellRef Cell reference like "A1"
     * @param value Typed value to write
     * @return XLResult[Sheet] for chaining
     */
    def put[A: CellWriter](cellRef: String, value: A): XLResult[Sheet] =
      toXLResult(ARef.parse(cellRef), cellRef, "Invalid cell reference").map { ref =>
        val (cellValue, styleOpt) = summon[CellWriter[A]].write(value)
        val updated = sheet.put(ref, cellValue)
        styleOpt.fold(updated)(style => updated.withCellStyle(ref, style))
      }

  // ========== Sheet Extensions: Generic Put (Styled) ==========

  extension (sheet: Sheet)
    /**
     * Put typed value with inline styling.
     *
     * Merges auto-inferred style (from type) with explicit CellStyle.
     * Explicit style takes precedence for conflicting properties.
     *
     * @tparam A The value type (must have CellWriter instance)
     * @param cellRef Cell reference like "A1"
     * @param value Typed value to write
     * @param cellStyle Explicit cell style to apply
     * @return XLResult[Sheet] for chaining
     */
    def put[A: CellWriter](cellRef: String, value: A, cellStyle: CellStyle): XLResult[Sheet] =
      toXLResult(ARef.parse(cellRef), cellRef, "Invalid cell reference").map { ref =>
        val (cellValue, autoStyleOpt) = summon[CellWriter[A]].write(value)
        val updated = sheet.put(ref, cellValue)

        // Merge auto-inferred style with explicit style (explicit wins)
        val finalStyle = autoStyleOpt match
          case Some(autoStyle) =>
            // Preserve NumFmt from auto-inference if explicit style doesn't override
            if cellStyle.numFmt.isEmpty then cellStyle.copy(numFmt = autoStyle.numFmt)
            else cellStyle
          case None => cellStyle

        updated.withCellStyle(ref, finalStyle)
      }

  // ========== XLResult[Sheet] Extensions: Chainable Operations ==========

  extension (result: XLResult[Sheet])
    /**
     * Put typed value (chainable).
     *
     * Chains after previous XLResult[Sheet] operations. All types supported.
     */
    @annotation.targetName("putGenericChainable")
    def put[A: CellWriter](cellRef: String, value: A): XLResult[Sheet] =
      result.flatMap(_.put(cellRef, value))

    /**
     * Put typed value with style (chainable).
     *
     * Chains after previous XLResult[Sheet] operations with inline styling.
     */
    @annotation.targetName("putGenericStyledChainable")
    def put[A: CellWriter](cellRef: String, value: A, cellStyle: CellStyle): XLResult[Sheet] =
      result.flatMap(_.put(cellRef, value, cellStyle))

  // ... rest of extensions unchanged (style, cell, range, merge, etc.)
```

### 3.2 Error Message Quality

**Concern**: Generic methods might produce worse error messages than explicit overloads.

**Testing strategy**:
```scala
// Test error messages for type mismatches
val result1 = sheet.put("A1", unsupportedType) // Should fail at compile time
val result2 = sheet.put("InvalidRef", "Value") // Should produce clear runtime error

// Expected error for InvalidRef:
// XLError.InvalidCellRef("InvalidRef", "Invalid cell reference: Expected column letter...")
```

**Verification**: Add tests in `EasyModeSyntaxSpec.scala` to verify error messages are unchanged.

### 3.3 Type Inference Behavior

**Success cases** (should infer correctly):
```scala
sheet.put("A1", "Hello")                      // A = String
sheet.put("A1", 42)                           // A = Int
sheet.put("A1", LocalDate.now())              // A = LocalDate
sheet.put("A1", "Title", headerStyle)         // A = String
```

**Ambiguity cases** (should require explicit type):
```scala
sheet.put("A1", null)  // ❌ Compile error: ambiguous type (good!)
```

**Conversion cases**:
```scala
val x: Long = 42
sheet.put("A1", x)  // A = Long (correct, no Int conversion)
```

### 3.4 NumFmt Auto-Inference Strategy

**Current behavior** (from `CellCodec.scala`):
- `LocalDate` → `Some(CellStyle.default.withNumFmt(NumFmt.Date))`
- `LocalDateTime` → `Some(CellStyle.default.withNumFmt(NumFmt.DateTime))`
- `BigDecimal` → `Some(CellStyle.default.withNumFmt(NumFmt.Decimal))`
- `Int/Long/Double` → `Some(CellStyle.default.withNumFmt(NumFmt.General))`
- `String/Boolean/RichText` → `None` (no special formatting)

**IMPORTANT BUG FIX**: Current explicit overloads in `extensions.scala` DON'T apply auto-inferred NumFmt! They just put the value without styling. The new type class implementation will fix this.

**Preservation in new API**:
```scala
// Unstyled: Apply auto-inferred style if present
val (cellValue, styleOpt) = writer.write(value)
styleOpt.fold(updated)(style => updated.withCellStyle(ref, style))

// Styled: Merge auto-inferred NumFmt with explicit style
val finalStyle = autoStyleOpt match
  case Some(autoStyle) =>
    if cellStyle.numFmt.isEmpty then cellStyle.copy(numFmt = autoStyle.numFmt)
    else cellStyle
  case None => cellStyle
```

This ensures dates/decimals get proper formatting even when styled.

---

## 4. Implementation Strategy

### 4.1 Step-by-Step Plan

**Phase 1: Add CellWriter.apply** (2 minutes)
1. Add `CellWriter.apply[A]` summon method to `CellWriter` companion object
2. No need to re-export instances (already available via `CellCodec` instances)

**Phase 2: Replace Sheet extensions** (10 minutes)
1. Remove 9 unstyled overloads (lines 52-94)
2. Replace with single generic `put[A: CellWriter]` method
3. Remove 9 styled overloads (lines 100-169)
4. Replace with single generic `put[A: CellWriter](_, _, CellStyle)` method

**Phase 3: Replace XLResult extensions** (10 minutes)
1. Remove 18 chainable overloads (lines 259-347)
2. Replace with 2 generic chainable methods

**Phase 4: Testing** (20 minutes)
1. Compile and verify zero errors
2. Run existing test suite (should all pass)
3. Add new tests for type inference edge cases
4. Add tests for NumFmt auto-application (bug fix verification)

**Total effort**: ~42 minutes

### 4.2 Backward Compatibility

**Breaking changes**: NONE

**Reason**:
- Generic methods have identical signatures to explicit overloads
- Type inference selects correct instance automatically
- Error messages remain the same (same validation logic)
- Runtime behavior improved (fixes missing NumFmt application bug)

**Verification**:
```bash
./mill __.compile  # Should compile with zero errors
./mill __.test     # All 263 tests should pass
```

### 4.3 NumFmt Auto-Inference Enhancement

**Current logic** (in explicit overloads):
```scala
// LocalDate overload (line 82-84)
def put(cellRef: String, value: LocalDate): XLResult[Sheet] =
  toXLResult(ARef.parse(cellRef), cellRef, "Invalid cell reference")
    .map(ref => sheet.put(ref, CellValue.DateTime(value.atStartOfDay)))
```

**Problem**: Current explicit overloads DON'T apply NumFmt! They just put the value.

**New logic** (with type class):
```scala
def put[A: CellWriter](cellRef: String, value: A): XLResult[Sheet] =
  toXLResult(ARef.parse(cellRef), cellRef, "Invalid cell reference").map { ref =>
    val (cellValue, styleOpt) = summon[CellWriter[A]].write(value)
    val updated = sheet.put(ref, cellValue)
    styleOpt.fold(updated)(style => updated.withCellStyle(ref, style))
  }
```

**Improvement**: New version DOES apply auto-inferred NumFmt (fixing a bug in current implementation!)

**Verification**: Add test to ensure LocalDate cells have Date format after put().

---

## 5. Testing Strategy

### 5.1 Compile-Time Tests

**File**: `xl-core/test/src/com/tjclp/xl/EasyModeTypeInferenceSpec.scala` (NEW)

```scala
package com.tjclp.xl

import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.CellStyle
import java.time.LocalDate
import munit.FunSuite

class EasyModeTypeInferenceSpec extends FunSuite:

  test("put() should infer String type") {
    val sheet = Sheet(SheetName("Test")).unsafe
    val result: Either[_, Sheet] = sheet.put("A1", "Hello")
    assert(result.isRight)
  }

  test("put() should infer Int type") {
    val sheet = Sheet(SheetName("Test")).unsafe
    val result: Either[_, Sheet] = sheet.put("A1", 42)
    assert(result.isRight)
  }

  test("put() should infer LocalDate type") {
    val sheet = Sheet(SheetName("Test")).unsafe
    val result: Either[_, Sheet] = sheet.put("A1", LocalDate.of(2025, 11, 19))
    assert(result.isRight)
  }

  test("put() with style should infer String type") {
    val sheet = Sheet(SheetName("Test")).unsafe
    val style = CellStyle.default
    val result: Either[_, Sheet] = sheet.put("A1", "Title", style)
    assert(result.isRight)
  }

  // Compile-time check: This should NOT compile (no CellWriter[List[String]])
  // test("put() should reject unsupported types") {
  //   val sheet = Sheet(SheetName("Test")).unsafe
  //   val result = sheet.put("A1", List("a", "b")) // ❌ Compile error
  // }
```

### 5.2 Runtime Tests

**File**: `xl-core/test/src/com/tjclp/xl/EasyModeSyntaxSpec.scala` (UPDATE)

Add tests for:
1. **NumFmt preservation**: Verify LocalDate cells have Date format (NEW - bug fix verification)
2. **Style merging**: Verify explicit style + auto NumFmt merges correctly
3. **Error messages**: Verify parse errors have same quality
4. **Chainable operations**: Verify XLResult[Sheet] chaining works

```scala
test("put() should auto-infer Date format for LocalDate (BUG FIX)") {
  val sheet = Sheet(SheetName("Test")).unsafe
  val result = sheet.put("A1", LocalDate.of(2025, 11, 19))

  result match
    case Right(updated) =>
      val ref = ARef.parse("A1").toOption.get
      val cell = updated.cells(ref)
      // Verify cell has Date format applied (this is the bug fix!)
      assert(cell.styleId.isDefined, "Cell should have style")
      val style = updated.styleRegistry.get(cell.styleId.get)
      assertEquals(style.numFmt, NumFmt.Date)
    case Left(err) => fail(s"Unexpected error: $err")
}

test("put() with style should merge auto NumFmt with explicit style") {
  val sheet = Sheet(SheetName("Test")).unsafe
  val boldStyle = CellStyle.default.bold
  val result = sheet.put("A1", LocalDate.of(2025, 11, 19), boldStyle)

  result match
    case Right(updated) =>
      val ref = ARef.parse("A1").toOption.get
      val cell = updated.cells(ref)
      val style = updated.styleRegistry.get(cell.styleId.get)
      // Verify both bold and Date format applied
      assertEquals(style.font.bold, Some(true))
      assertEquals(style.numFmt, NumFmt.Date)
    case Left(err) => fail(s"Unexpected error: $err")
}

test("put() should produce clear error for invalid ref") {
  val sheet = Sheet(SheetName("Test")).unsafe
  val result = sheet.put("InvalidRef", "Value")

  result match
    case Left(XLError.InvalidCellRef(ref, msg)) =>
      assertEquals(ref, "InvalidRef")
      assert(msg.contains("Invalid cell reference"))
    case _ => fail("Expected InvalidCellRef error")
}
```

### 5.3 Performance Tests

**Goal**: Verify zero overhead from type class abstraction

**Method**: Compare bytecode or microbenchmarks (future work, not blocking)

**Hypothesis**: `inline given` instances eliminate all overhead

**Verification**:
```bash
# Compile with -Xprint:typer to inspect generated code
./mill xl-core.compile -Xprint:typer | grep "put.*String"
```

Expected: Type class version generates identical bytecode to explicit overload.

---

## 6. Migration Path

### 6.1 Zero Breaking Changes

**Reason**: Generic signatures are compatible with explicit overloads

**User code** (unchanged):
```scala
import com.tjclp.xl.*

val sheet = Sheet("Test")
  .put("A1", "Hello")        // Still works
  .put("B1", 42)             // Still works
  .put("C1", LocalDate.now()) // Still works (now with Date format!)
```

### 6.2 Documentation Updates

**Files to update**:
1. `docs/design/easy-mode-api.md` - Mention type class consolidation (if exists)
2. `CLAUDE.md` - Update "Common Patterns" section with type class pattern
3. Scaladoc in `extensions.scala` - Add type class explanation

**Example Scaladoc**:
```scala
/**
 * Put typed value at cell reference.
 *
 * Uses CellWriter type class for automatic type support. Supports:
 * String, Int, Long, Double, BigDecimal, Boolean, LocalDate, LocalDateTime, RichText.
 *
 * Auto-infers NumFmt based on type (LocalDate → Date, BigDecimal → Decimal, etc.).
 *
 * @tparam A The value type (must have CellWriter instance)
 * @param cellRef Cell reference like "A1"
 * @param value Typed value to write
 * @return XLResult[Sheet] for chaining
 */
def put[A: CellWriter](cellRef: String, value: A): XLResult[Sheet]
```

---

## 7. Tradeoffs Analysis

### 7.1 Pros

1. **67-75% code reduction**: 120 LOC → 30-40 LOC
2. **Extensibility**: Users can add custom types:
   ```scala
   given CellWriter[CustomType] with
     def write(c: CustomType) = (CellValue.Text(c.toString), None)
   ```
3. **Consistency**: Single implementation, no drift between overloads
4. **Bug fix**: Current explicit overloads don't apply NumFmt (new version does!)
5. **Maintainability**: Add new types in one place (CellCodec companion)
6. **Infrastructure reuse**: Leverages existing `CellWriter` from codec system

### 7.2 Cons

1. **IDE navigation**: "Go to definition" shows generic method, not type-specific impl
   - **Mitigation**: Scaladoc lists all supported types explicitly
2. **Type class complexity**: Requires understanding implicits/givens
   - **Mitigation**: Users never see type class, just use `put()`
3. **Error message complexity**: Missing instance errors mention "CellWriter"
   - **Mitigation**: Comprehensive Scaladoc + clear error in CellWriter companion
4. **Slightly longer signatures**: `put[A: CellWriter]` vs `put(...: String)`
   - **Mitigation**: Users never write type parameter (inferred)

### 7.3 Overall Assessment

**Recommendation**: PROCEED with type class consolidation

**Rationale**:
- Pros heavily outweigh cons
- Cons are mostly theoretical (users won't notice in practice)
- Fixes existing bug (missing NumFmt application)
- Aligns with Scala 3 best practices
- Leverages existing infrastructure (DRY)

---

## 8. Alternative Approaches

### 8.1 Macro-Based Generation

**Approach**: Use macro to generate all 18 overloads at compile time

**Pros**:
- Preserves explicit overloads (IDE navigation)
- Zero runtime overhead (same as type class)

**Cons**:
- More complex implementation (macro code)
- Less extensible (users can't add types)
- Harder to maintain (macro hygiene)

**Decision**: REJECT (type classes are simpler and more extensible)

### 8.2 Build-Time Code Generation

**Approach**: Use sbt/Mill plugin to generate overloads from template

**Pros**:
- Explicit overloads in source
- Full IDE support

**Cons**:
- Build complexity
- Generated code in source tree (pollution)
- Not extensible by users

**Decision**: REJECT (not worth build complexity)

### 8.3 Keep Explicit Overloads

**Approach**: Accept duplication as cost of clarity

**Pros**:
- Simple, explicit
- Perfect IDE support

**Cons**:
- 120 lines of duplication
- Bug risk (overloads can drift)
- Not extensible
- Current implementation has bug (missing NumFmt)

**Decision**: REJECT (duplication is not justified, especially with existing bug)

---

## 9. Success Criteria

### 9.1 Compilation

- ✅ All modules compile with zero errors
- ✅ No new warnings introduced

### 9.2 Tests

- ✅ All 263 existing tests pass
- ✅ New type inference tests pass
- ✅ NumFmt auto-application tests pass (bug fix verification)
- ✅ Style merging tests pass

### 9.3 Documentation

- ✅ Scaladoc updated with type class explanation
- ✅ CLAUDE.md updated with new pattern
- ✅ Design doc (this file) captures all decisions

### 9.4 User Experience

- ✅ Type inference works for all 9 supported types
- ✅ Error messages are clear and actionable
- ✅ Chainable operations work seamlessly
- ✅ NumFmt auto-application works (bug fixed)

---

## 10. Implementation Checklist

**Phase 1: CellWriter companion** (2 min)
- [ ] Add `CellWriter.apply[A]` method to summon instances
- [ ] Verify all 9 inline given instances are accessible

**Phase 2: Sheet extensions** (10 min)
- [ ] Remove 9 unstyled overloads with generic method
- [ ] Remove 9 styled overloads with generic method
- [ ] Add NumFmt auto-application logic
- [ ] Test compilation

**Phase 3: XLResult extensions** (10 min)
- [ ] Remove 18 chainable overloads with 2 generic methods
- [ ] Test compilation

**Phase 4: Testing** (20 min)
- [ ] Run existing test suite (all 263 tests)
- [ ] Add type inference tests
- [ ] Add NumFmt preservation tests (bug fix verification)
- [ ] Add style merging tests
- [ ] Add error message tests

**Phase 5: Documentation** (10 min)
- [ ] Update Scaladoc in extensions.scala
- [ ] Update CLAUDE.md with type class pattern
- [ ] Update any design docs

**Total effort**: ~52 minutes

---

## 11. Conclusion

This type class consolidation eliminates 67-75% of duplicated code while maintaining full backward compatibility and actually improving functionality (NumFmt auto-application bug fix). The implementation is straightforward, leveraging existing `CellWriter` infrastructure, and aligns with Scala 3 best practices.

**Key Benefits**:
- **Code reduction**: 120 LOC → 30-40 LOC
- **Bug fix**: Adds missing NumFmt auto-application
- **Extensibility**: Users can add custom types
- **Maintainability**: Single implementation, no drift
- **Infrastructure reuse**: Leverages existing codec system

**Recommendation**: APPROVE and implement in follow-up PR after PR #19 merges.

---

## 12. References

- PR #19: Easy Mode API implementation
- `xl-core/src/com/tjclp/xl/extensions.scala`: Current implementation with duplication
- `xl-core/src/com/tjclp/xl/codec/CellWriter.scala`: Existing type class
- `xl-core/src/com/tjclp/xl/codec/CellCodec.scala`: Existing inline given instances
- Automated review feedback: Identified duplication as optional improvement
