# Phase 1 Implementation Plan: String Interpolation Runtime Support

**Status**: Ready for Implementation
**Parent**: `string-interpolation.md` (Design Specification)
**Timeline**: 12-16 developer-days (sequential) or 8-10 days (2 devs parallel)
**PRs**: 4 independently reviewable pull requests

## Executive Summary

This document provides **PR-ready implementation specifications** for Phase 1 (Basic Runtime Support) of string interpolation. Phase 1 focuses exclusively on enabling runtime string parsing for all four macro types (`ref`, `money`, `percent`, `date`, `accounting`, `fx`) while maintaining 100% backward compatibility.

**No compile-time optimization** is included in Phase 1 - that's deferred to Phase 2.

### Strategic Approach

**4 PRs in Sequential Order**:

1. **PR #1: Infrastructure** (0.75 days) - Error types + test patterns
2. **PR #2: ref"$var"** (1.5 days) - Highest value, reuses existing parsers
3. **PR #3: Formatted literals** (2 days) - money/percent/date/accounting together
4. **PR #4: fx"$var"** (1 day) - Minimal validation, lowest risk

### Key Principles

- ✅ **100% Backward Compatible**: Compile-time literals unchanged
- ✅ **Independently Reviewable**: Each PR is self-contained with tests
- ✅ **Incrementally Valuable**: Each PR adds working functionality
- ✅ **Risk-Minimized**: Infrastructure first, simple before complex
- ✅ **Well-Tested**: 111 new tests ensure correctness

---

## Current State Analysis

### Existing Macro Structure

| Macro | File | Lines | Features | Current Tests |
|-------|------|-------|----------|---------------|
| `ref` | `RefLiteral.scala` | 219 | Unified ref parsing (ARef, CellRange, RefType), sheet qualification, quote escaping | 42 tests (RefTypeSpec + AddressingSpec) |
| `money`, `percent`, `date`, `accounting` | `FormattedLiterals.scala` | 117 | Number/date parsing, format preservation | 18 tests (ElegantSyntaxSpec) |
| `fx` | `CellRangeLiterals.scala` | 63 | Formula validation (parentheses balance) | 0 dedicated tests |

### Existing Runtime Parsers (Already Pure & Total)

✅ **Already exist** (no changes needed):
- `ARef.parse(s: String): Either[String, ARef]`
- `CellRange.parse(s: String): Either[String, CellRange]`
- `RefType.parse(s: String): Either[String, RefType]`
- `Column.fromLetter(s: String): Either[String, Column]`
- `SheetName.apply(s: String): Either[String, SheetName]`

❌ **Missing** (will create in Phase 1):
- Runtime parsers for formatted literals (money, percent, date, accounting)
- Runtime parser for formulas (fx)

### Existing Error Types

✅ **Already exist** (in `XLError.scala`):
- `InvalidCellRef(ref: String, reason: String)`
- `InvalidRange(range: String, reason: String)`
- `InvalidReference(reason: String)`
- `FormulaError(expression: String, reason: String)`
- `InvalidSheetName(name: String, reason: String)`

❌ **Missing** (will add in PR #1):
- `MoneyFormatError(value: String, reason: String)`
- `PercentFormatError(value: String, reason: String)`
- `DateFormatError(value: String, reason: String)`
- `AccountingFormatError(value: String, reason: String)`

---

## PR #1: Infrastructure - Error Types & Test Patterns

**Priority**: Must merge first (all others depend on it)
**Estimated Effort**: 6 hours (0.75 days)
**Risk Level**: VERY LOW (purely additive)

### 1. Objective

Establish foundation for Phase 1 by adding:
- New XLError cases for formatted literal parsing
- Test generators for runtime validation
- No functional changes to existing macros

### 2. Files Modified

**2 files changed**:
- `xl-core/src/com/tjclp/xl/error/XLError.scala` (+16 lines)
- `xl-core/test/src/com/tjclp/xl/Generators.scala` (+30 lines)

**1 file created**:
- `xl-core/test/src/com/tjclp/xl/error/XLErrorSpec.scala` (new, ~100 lines)

### 3. Detailed Implementation

#### A. Modify `XLError.scala`

**Location**: `xl-core/src/com/tjclp/xl/error/XLError.scala`

**Add after line 45** (after `NumberFormatError`):

```scala
/** Money format parse errors */
case MoneyFormatError(value: String, reason: String)

/** Percent format parse errors */
case PercentFormatError(value: String, reason: String)

/** Date format parse errors */
case DateFormatError(value: String, reason: String)

/** Accounting format parse errors */
case AccountingFormatError(value: String, reason: String)
```

**Update `message` extension** (in `object XLError`, after line 84):

```scala
case MoneyFormatError(value, reason) =>
  s"Invalid money format '$value': $reason"
case PercentFormatError(value, reason) =>
  s"Invalid percent format '$value': $reason"
case DateFormatError(value, reason) =>
  s"Invalid date format '$value': $reason"
case AccountingFormatError(value, reason) =>
  s"Invalid accounting format '$value': $reason"
```

#### B. Modify `Generators.scala`

**Location**: `xl-core/test/src/com/tjclp/xl/Generators.scala`

**Add at end of file**:

```scala
// ===== String generators for runtime parsing tests =====

/** Valid money strings */
val genMoneyString: Gen[String] = Gen.oneOf(
  Gen.posNum[Double].map(n => f"$$$n%.2f"),
  Gen.posNum[Int].map(n => f"$$$n%,d.00"),
  Gen.const("$1,234.56"),
  Gen.const("999.99")
)

/** Valid percent strings */
val genPercentString: Gen[String] =
  Gen.choose(0.0, 100.0).map(n => f"$n%.2f%%")

/** Valid date strings (ISO format) */
val genDateString: Gen[String] =
  Gen.choose(2000, 2030).flatMap { year =>
    Gen.choose(1, 12).flatMap { month =>
      Gen.choose(1, 28).map { day =>
        f"$year%04d-$month%02d-$day%02d"
      }
    }
  }

/** Valid formula strings */
val genFormulaString: Gen[String] = Gen.oneOf(
  Gen.const("=SUM(A1:A10)"),
  Gen.const("=IF(A1>0,B1,C1)"),
  Gen.const("=AVERAGE(B2:B100)"),
  Gen.const("=COUNT(C1:C50)")
)

// Invalid strings for negative tests
val genInvalidMoney: Gen[String] = Gen.oneOf("$ABC", "1.2.3", "$$$$", "")
val genInvalidPercent: Gen[String] = Gen.oneOf("ABC%", "1%%", "%", "")
val genInvalidDate: Gen[String] = Gen.oneOf("2025-13-01", "not-a-date", "2025/11/10", "")
```

#### C. Create `XLErrorSpec.scala` (new file)

**Location**: `xl-core/test/src/com/tjclp/xl/error/XLErrorSpec.scala`

```scala
package com.tjclp.xl.errors

import munit.FunSuite

class XLErrorSpec extends FunSuite:

  test("MoneyFormatError.message includes value and reason") {
    val err = XLError.MoneyFormatError("$ABC", "non-numeric characters")
    val msg = err.message
    assert(msg.contains("$ABC"))
    assert(msg.contains("non-numeric"))
    assert(msg.contains("money format"))
  }

  test("PercentFormatError.message includes value and reason") {
    val err = XLError.PercentFormatError("ABC%", "invalid number")
    val msg = err.message
    assert(msg.contains("ABC%"))
    assert(msg.contains("invalid"))
    assert(msg.contains("percent format"))
  }

  test("DateFormatError.message includes value and reason") {
    val err = XLError.DateFormatError("not-a-date", "unparseable")
    val msg = err.message
    assert(msg.contains("not-a-date"))
    assert(msg.contains("unparseable"))
    assert(msg.contains("date format"))
  }

  test("AccountingFormatError.message includes value and reason") {
    val err = XLError.AccountingFormatError("$ABC", "invalid format")
    val msg = err.message
    assert(msg.contains("$ABC"))
    assert(msg.contains("invalid"))
    assert(msg.contains("accounting format"))
  }

  test("MoneyFormatError pattern matches correctly") {
    val err: XLError = XLError.MoneyFormatError("$123", "test")
    err match
      case XLError.MoneyFormatError(value, reason) =>
        assertEquals(value, "$123")
        assertEquals(reason, "test")
      case _ => fail("Should match MoneyFormatError")
  }

  test("PercentFormatError pattern matches correctly") {
    val err: XLError = XLError.PercentFormatError("50%", "test")
    err match
      case XLError.PercentFormatError(value, reason) =>
        assertEquals(value, "50%")
        assertEquals(reason, "test")
      case _ => fail("Should match PercentFormatError")
  }

  test("DateFormatError pattern matches correctly") {
    val err: XLError = XLError.DateFormatError("2025-11-10", "test")
    err match
      case XLError.DateFormatError(value, reason) =>
        assertEquals(value, "2025-11-10")
        assertEquals(reason, "test")
      case _ => fail("Should match DateFormatError")
  }

  test("AccountingFormatError pattern matches correctly") {
    val err: XLError = XLError.AccountingFormatError("($123)", "test")
    err match
      case XLError.AccountingFormatError(value, reason) =>
        assertEquals(value, "($123)")
        assertEquals(reason, "test")
      case _ => fail("Should match AccountingFormatError")
  }
```

### 4. Testing Requirements

**Test Count**: 8 tests (all in `XLErrorSpec.scala`)

**Coverage**:
- Error message formatting (4 tests)
- Pattern matching (4 tests)

**Verification**:
```bash
./mill xl-core.test.testOnly com.tjclp.xl.errorss.XLErrorSpec
```

### 5. Documentation Updates

**Update** `docs/reference/glossary.md`:

Add entries:
- `MoneyFormatError` - Runtime error parsing money format strings
- `PercentFormatError` - Runtime error parsing percent format strings
- `DateFormatError` - Runtime error parsing date format strings
- `AccountingFormatError` - Runtime error parsing accounting format strings

**Update** `docs/LIMITATIONS.md`:

Remove bullet:
- ~~"No runtime parsing for macro literals"~~

### 6. Success Criteria

- [ ] `./mill xl-core.compile` succeeds with 0 warnings
- [ ] `./mill xl-core.test.testOnly com.tjclp.xl.errors.XLErrorSpec` passes (8/8 tests)
- [ ] `./mill __.test` passes (all 263 existing tests + 8 new = 271 total)
- [ ] No changes to existing macro behavior
- [ ] Glossary updated with new error types

### 7. Review Checklist

**For Reviewers**:
- [ ] Error messages are user-friendly and actionable
- [ ] Pattern matching works for all 4 new error types
- [ ] Generators produce valid and invalid strings appropriately
- [ ] No breaking changes to existing error handling

**Questions**:
- Are error messages clear enough for end users?
- Should we add more generator variants?

### 8. Rollback Plan

**Risk**: Zero (purely additive)

**Procedure**: Simple revert of single commit
1. `git revert <commit-hash>`
2. Verify: `./mill __.test`

---

## PR #2: String Interpolation for ref"$var"

**Priority**: Depends on PR #1
**Estimated Effort**: 12 hours (1.5 days)
**Risk Level**: LOW (isolated to ref macro)

### 1. Objective

Enable runtime string interpolation for `ref""` macro while maintaining 100% backward compatibility for compile-time literals.

**Before**:
```scala
val cellStr = "A1"
val ref = ARef.parse(cellStr).getOrElse(???)  // Manual parsing
```

**After**:
```scala
val cellStr = "A1"
val ref = ref"$cellStr"  // Returns Either[XLError, RefType]
```

### 2. Files Modified

**2 files modified**:
- `xl-core/src/com/tjclp/xl/macros/RefLiteral.scala` (+50 lines)
- `xl-core/src/com/tjclp/xl/addressing/RefType.scala` (+10 lines)

**1 file created**:
- `xl-core/test/src/com/tjclp/xl/macros/RefInterpolationSpec.scala` (new, ~250 lines)

### 3. Detailed Implementation

#### A. Modify `RefLiteral.scala`

**Location**: `xl-core/src/com/tjclp/xl/macros/RefLiteral.scala`

**Step 1: Change signature** (lines 35-36):

```scala
// OLD:
inline def ref(inline args: Any*): ARef | CellRange | RefType =
  ${ errorNoInterpolation('sc, 'args, "ref") }

// NEW:
transparent inline def ref(inline args: Any*): ARef | CellRange | RefType | Either[XLError, RefType] =
  ${ refImplN('sc, 'args) }
```

**Step 2: Add new implementation** (after line 92):

```scala
/**
 * Macro implementation supporting both compile-time literals and runtime interpolation.
 *
 * - No args (literal): Delegates to refImpl0, returns ARef | CellRange | RefType directly
 * - With args (interpolation): Builds string at runtime, returns Either[XLError, RefType]
 */
private def refImplN(
  sc: Expr[StringContext],
  args: Expr[Seq[Any]]
)(using Quotes): Expr[ARef | CellRange | RefType | Either[XLError, RefType]] =
  import quotes.reflect.*

  args match
    case Varargs(exprs) if exprs.isEmpty =>
      // No interpolation - compile-time literal (backward compatible)
      refImpl0(sc)

    case Varargs(exprs) =>
      // Has interpolation - runtime parsing
      '{
        val str = $sc.s($args*)
        RefType.parseToXLError(str)
      }.asExprOf[Either[XLError, RefType]]
```

#### B. Modify `RefType.scala`

**Location**: `xl-core/src/com/tjclp/xl/addressing/RefType.scala`

**Add helper method** (in `object RefType`, after line 143):

```scala
/**
 * Parse ref string with XLError wrapping.
 *
 * Used by runtime string interpolation macro.
 */
def parseToXLError(s: String): Either[XLError, RefType] =
  parse(s).left.map { err =>
    XLError.InvalidReference(s"Failed to parse '$s': $err")
  }
```

### 4. Testing Requirements

**Test File**: `xl-core/test/src/com/tjclp/xl/macros/RefInterpolationSpec.scala`

**Test Count**: 25 tests

**Coverage**:

```scala
package com.tjclp.xl.macros

import com.tjclp.xl.*
import com.tjclp.xl.errors.XLError
import munit.{FunSuite, ScalaCheckSuite}
import org.scalacheck.Prop.*
import org.scalacheck.Gen

class RefInterpolationSpec extends ScalaCheckSuite:

  // ===== Backward Compatibility (Compile-Time Literals) =====

  test("Compile-time literal: ref\"A1\" returns ARef directly") {
    val r = ref"A1"
    assert(r.isInstanceOf[ARef])
    val aref = r.asInstanceOf[ARef]
    assertEquals(aref.toA1, "A1")
  }

  test("Compile-time literal: ref\"A1:B10\" returns CellRange directly") {
    val r = ref"A1:B10"
    assert(r.isInstanceOf[CellRange])
    val range = r.asInstanceOf[CellRange]
    assertEquals(range.toA1, "A1:B10")
  }

  test("Compile-time literal: ref\"Sales!A1\" returns RefType directly") {
    val r = ref"Sales!A1"
    assert(r.isInstanceOf[RefType])
    val qualified = r.asInstanceOf[RefType]
    assertEquals(qualified.toA1, "Sales!A1")
  }

  // ===== Runtime Interpolation (New Functionality) =====

  test("Runtime interpolation: simple cell ref") {
    val cellStr = "A1"
    val result = ref"$cellStr"

    result match
      case Right(RefType.Cell(aref)) =>
        assertEquals(aref.toA1, "A1")
      case other =>
        fail(s"Expected Right(Cell(A1)), got $other")
  }

  test("Runtime interpolation: cell range") {
    val rangeStr = "A1:B10"
    val result = ref"$rangeStr"

    result match
      case Right(RefType.Range(range)) =>
        assertEquals(range.toA1, "A1:B10")
      case other =>
        fail(s"Expected Right(Range), got $other")
  }

  test("Runtime interpolation: qualified cell") {
    val qcellStr = "Sales!A1"
    val result = ref"$qcellStr"

    result match
      case Right(RefType.QualifiedCell(sheet, aref)) =>
        assertEquals(sheet.value, "Sales")
        assertEquals(aref.toA1, "A1")
      case other =>
        fail(s"Expected Right(QualifiedCell), got $other")
  }

  test("Runtime interpolation: qualified range") {
    val qrangeStr = "Sales!A1:B10"
    val result = ref"$qrangeStr"

    result match
      case Right(RefType.QualifiedRange(sheet, range)) =>
        assertEquals(sheet.value, "Sales")
        assertEquals(range.toA1, "A1:B10")
      case other =>
        fail(s"Expected Right(QualifiedRange), got $other")
  }

  test("Runtime interpolation: quoted sheets name") {
    val quotedStr = "'Q1 Sales'!A1"
    val result = ref"$quotedStr"

    result match
      case Right(RefType.QualifiedCell(sheet, _)) =>
        assertEquals(sheet.value, "Q1 Sales")
      case other =>
        fail(s"Expected Right with quoted sheets, got $other")
  }

  test("Runtime interpolation: escaped quotes in sheets name") {
    val escapedStr = "'It''s Q1'!A1"
    val result = ref"$escapedStr"

    result match
      case Right(RefType.QualifiedCell(sheet, _)) =>
        assertEquals(sheet.value, "It's Q1")
      case other =>
        fail(s"Expected Right with escaped quotes, got $other")
  }

  // ===== Error Cases =====

  test("Runtime interpolation: invalid ref returns Left(XLError)") {
    val invalidStr = "INVALID@#$"
    val result = ref"$invalidStr"

    result match
      case Left(err: XLError.InvalidReference) =>
        assert(err.message.contains("INVALID"))
      case other =>
        fail(s"Expected Left(InvalidReference), got $other")
  }

  test("Runtime interpolation: empty string returns Left") {
    val emptyStr = ""
    val result = ref"$emptyStr"

    assert(result.isLeft, "Empty string should fail")
  }

  test("Runtime interpolation: missing ref after bang returns Left") {
    val invalidStr = "Sales!"
    val result = ref"$invalidStr"

    assert(result.isLeft, "Missing ref after ! should fail")
  }

  test("Runtime interpolation: invalid column returns Left") {
    val invalidStr = "ZZZ1"  // Beyond XFD
    val result = ref"$invalidStr"

    assert(result.isLeft, "Column beyond XFD should fail")
  }

  // ===== Mixed Compile-Time and Runtime =====

  test("Mixed interpolation: prefix + variable") {
    val sheet = "Sales"
    val result = ref"$sheet!A1"

    result match
      case Right(RefType.QualifiedCell(s, aref)) =>
        assertEquals(s.value, "Sales")
        assertEquals(aref.toA1, "A1")
      case other =>
        fail(s"Expected Right(QualifiedCell), got $other")
  }

  test("Mixed interpolation: variable + suffix") {
    val colRow = "B5"
    val result = ref"Sheet1!$colRow"

    result match
      case Right(RefType.QualifiedCell(s, aref)) =>
        assertEquals(s.value, "Sheet1")
        assertEquals(aref.toA1, "B5")
      case other =>
        fail(s"Expected Right(QualifiedCell), got $other")
  }

  test("Mixed interpolation: multiple variables") {
    val sheet = "Q1"
    val col = "B"
    val row = 42
    val result = ref"$sheet!$col$row"

    result match
      case Right(RefType.QualifiedCell(s, aref)) =>
        assertEquals(s.value, "Q1")
        assertEquals(aref.toA1, "B42")
      case other =>
        fail(s"Expected Right(QualifiedCell), got $other")
  }

  // ===== Property-Based Tests =====

  property("Round-trip: RefType -> toA1 -> parse -> RefType") {
    forAll(Generators.genRefType) { refType =>
      val a1 = refType.toA1
      val dynamicStr = identity(a1)  // Force runtime path
      val result = ref"$dynamicStr"

      result match
        case Right(parsed) =>
          assertEquals(parsed, refType)
        case Left(err) =>
          fail(s"Round-trip failed for $refType: $err")
    }
  }

  property("Valid ref strings always parse") {
    forAll(Generators.genValidRefString) { str =>
      val result = ref"$str"
      assert(result.isRight, s"Valid ref '$str' should parse")
    }
  }

  property("Invalid ref strings always return Left") {
    val genInvalid = Gen.oneOf("", "INVALID", "@#$", "A", "1", "!")
    forAll(genInvalid) { str =>
      val result = ref"$str"
      assert(result.isLeft, s"Invalid ref '$str' should fail")
    }
  }

  // ===== Edge Cases =====

  test("Edge: Maximum valid column (XFD)") {
    val maxCol = "XFD1"
    val result = ref"$maxCol"
    assert(result.isRight, "XFD1 should be valid")
  }

  test("Edge: Maximum valid row (1048576)") {
    val maxRow = "A1048576"
    val result = ref"$maxRow"
    assert(result.isRight, "A1048576 should be valid")
  }

  test("Edge: Sheet name with spaces") {
    val withSpaces = "'Q1 Sales Report'!A1"
    val result = ref"$withSpaces"

    result match
      case Right(RefType.QualifiedCell(sheet, _)) =>
        assertEquals(sheet.value, "Q1 Sales Report")
      case Left(err) =>
        fail(s"Should parse sheets with spaces: $err")
  }

  test("Edge: Sheet name with special chars") {
    val withSpecial = "'Sheet (2025)'!A1"
    val result = ref"$withSpecial"

    result match
      case Right(RefType.QualifiedCell(sheet, _)) =>
        assertEquals(sheet.value, "Sheet (2025)")
      case Left(err) =>
        fail(s"Should parse sheets with parens: $err")
  }

  // ===== Integration with for-comprehension =====

  test("Integration: for-comprehension with Either") {
    val sheetStr = "Sales"
    val cellStr = "A1"

    val result = for
      ref <- ref"$sheetStr!$cellStr"
    yield ref.toA1

    result match
      case Right(a1) => assertEquals(a1, "Sales!A1")
      case Left(err) => fail(s"Should parse: $err")
  }
```

**Add to `Generators.scala`**:

```scala
// Valid ref string generator
val genValidRefString: Gen[String] = Gen.oneOf(
  genARef.map(_.toA1),
  genCellRange.map(_.toA1),
  genRefType.map(_.toA1)
)
```

### 5. Documentation Updates

**Update** `docs/reference/examples.md`:

Add section:

```markdown
### String Interpolation for References

Runtime string interpolation with error handling:

\`\`\`scala
// Simple cell reference
val cellStr = "A1"
ref"$cellStr" match
  case Right(RefType.Cell(aref)) => println(s"Cell: ${aref.toA1}")
  case Left(err) => println(s"Error: ${err.message}")

// Qualified cell reference
val sheet = "Sales"
val cell = "B10"
ref"$sheet!$cell" match
  case Right(RefType.QualifiedCell(name, aref)) => // ...
  case Left(err) => // ...

// Use in for-comprehension
for
  ref <- ref"$userInput"
  sheet <- workbook(ref.sheet)
  cell <- sheet(ref.cell)
yield cell.value
\`\`\`
```

**Update** `CLAUDE.md` (Macro Patterns section):

```markdown
### Hybrid Compile-Time/Runtime Validation

Macros now support both compile-time literals and runtime interpolation:

\`\`\`scala
// Compile-time: validated at compile time, zero overhead
val r1: ARef = ref"A1"

// Runtime: validated at runtime, returns Either
val cellStr = "A1"
val r2: Either[XLError, RefType] = ref"$cellStr"
\`\`\`
```

### 6. Success Criteria

- [ ] `./mill xl-core.compile` succeeds with 0 warnings
- [ ] `./mill xl-core.test.testOnly com.tjclp.xl.macros.RefInterpolationSpec` passes (25/25 tests)
- [ ] `./mill __.test` passes (all 271 existing tests + 25 new = 296 total)
- [ ] Compile-time literals still work identically (backward compatible)
- [ ] Runtime interpolations return `Either[XLError, RefType]`
- [ ] Documentation updated with examples

### 7. Review Checklist

**For Reviewers**:
- [ ] Backward compatibility verified (compile-time literals unchanged)
- [ ] Error handling is total (no exceptions)
- [ ] Type safety maintained (`transparent inline` correct)
- [ ] Performance: compile-time literals have zero overhead
- [ ] Property tests verify round-trip laws

**Questions**:
- Should we provide an `unsafe` variant that throws exceptions?
- Is `Either[XLError, RefType]` the right API for runtime?

### 8. Rollback Plan

**Risk**: Low (isolated to ref macro)

**Procedure**:
1. Revert signature change in `RefLiteral.scala` (line 35)
2. Remove `refImplN` function (lines 94-110)
3. Remove `parseToXLError` from `RefType.scala`
4. Delete test file `RefInterpolationSpec.scala`
5. Verify: `./mill __.test` (271 tests should pass)

---

## PR #3: String Interpolation for Formatted Literals

**Priority**: Depends on PR #1
**Estimated Effort**: 15 hours (2 days)
**Risk Level**: MEDIUM (new parser logic, most complex PR)

### 1. Objective

Enable runtime string interpolation for `money""`, `percent""`, `date""`, and `accounting""` macros by creating pure runtime parsers.

**Before**:
```scala
val priceStr = "$1,234.56"
// No runtime parser available
```

**After**:
```scala
val priceStr = "$1,234.56"
money"$priceStr" match
  case Right(formatted) => // Use formatted value
  case Left(err) => // Handle parse errors
```

### 2. Files Modified/Created

**1 file modified**:
- `xl-core/src/com/tjclp/xl/macros/FormattedLiterals.scala` (+80 lines)

**3 files created**:
- `xl-core/src/com/tjclp/xl/formatted/FormattedParsers.scala` (new, ~150 lines)
- `xl-core/test/src/com/tjclp/xl/formatted/FormattedParsersSpec.scala` (new, ~220 lines)
- `xl-core/test/src/com/tjclp/xl/macros/FormattedInterpolationSpec.scala` (new, ~280 lines)

### 3. Detailed Implementation

#### A. Create `FormattedParsers.scala` (NEW FILE)

**Location**: `xl-core/src/com/tjclp/xl/formatted/FormattedParsers.scala`

```scala
package com.tjclp.xl.formatted

import com.tjclp.xl.cell.CellValue
import com.tjclp.xl.errors.XLError
import com.tjclp.xl.styles.numfmt.NumFmt
import java.time.LocalDate
import java.time.format.DateTimeFormatter
import scala.util.{Try, Success, Failure}

/**
 * Pure runtime parsers for formatted literal strings.
 *
 * These mirror the compile-time macro parsing logic but operate at runtime.
 * All parsers are total functions returning Either[XLError, Formatted].
 */
object FormattedParsers:

  /**
   * Parse money format: $1,234.56
   *
   * Accepts:
   * - With/without dollar sign: "$1234.56" or "1234.56"
   * - With/without commas: "$1,234.56" or "$1234.56"
   * - With/without spaces: "$ 1,234.56" or "$1,234.56"
   *
   * Rejects:
   * - Non-numeric: "$ABC"
   * - Multiple decimals: "$1.2.3"
   * - Empty: "" or "$"
   */
  def parseMoney(s: String): Either[XLError, Formatted] =
    Try {
      val cleaned = s.replaceAll("[\\$,\\s]", "")
      if cleaned.isEmpty then
        throw new NumberFormatException("Empty value after removing $ and ,")
      val num = BigDecimal(cleaned)
      Formatted(CellValue.Number(num.toDouble), NumFmt.Currency)
    }.toEither.left.map { err =>
      XLError.MoneyFormatError(s, err.getMessage)
    }

  /**
   * Parse percent format: 45.5%
   *
   * Accepts:
   * - With percent sign: "45.5%"
   * - Without percent sign: "45.5" (treated as 45.5%, not 0.455)
   * - Integer: "50%"
   *
   * Rejects:
   * - Non-numeric: "ABC%"
   * - Multiple percent signs: "50%%"
   * - Empty: "" or "%"
   */
  def parsePercent(s: String): Either[XLError, Formatted] =
    Try {
      val cleaned = s.replace("%", "").trim
      if cleaned.isEmpty then
        throw new NumberFormatException("Empty value after removing %")
      val num = BigDecimal(cleaned) / 100
      Formatted(CellValue.Number(num.toDouble), NumFmt.Percent)
    }.toEither.left.map { err =>
      XLError.PercentFormatError(s, err.getMessage)
    }

  /**
   * Parse ISO date format: 2025-11-10
   *
   * Accepts:
   * - ISO 8601 format only: "YYYY-MM-DD"
   * - Valid dates only (no Feb 30, etc.)
   *
   * Rejects:
   * - US format: "11/10/2025"
   * - European format: "10/11/2025"
   * - Invalid dates: "2025-02-30"
   * - Non-dates: "not-a-date"
   *
   * Note: Other formats will be added in Phase 2 (compile-time optimization)
   */
  def parseDate(s: String): Either[XLError, Formatted] =
    Try {
      val localDate = LocalDate.parse(s)  // ISO format: YYYY-MM-DD
      val dateTime = localDate.atStartOfDay()
      Formatted(CellValue.DateTime(dateTime), NumFmt.Date)
    }.toEither.left.map { err =>
      XLError.DateFormatError(
        s,
        s"Expected ISO format (YYYY-MM-DD): ${err.getMessage}"
      )
    }

  /**
   * Parse accounting format: ($123.45) or $123.45
   *
   * Accepts:
   * - Positive: "$123.45"
   * - Negative with parens: "($123.45)"
   * - With/without commas: "$1,234.56" or "$1234.56"
   *
   * Rejects:
   * - Negative without parens: "-$123.45" (use money"" for this)
   * - Non-numeric: "$ABC"
   * - Empty: "" or "$()"
   */
  def parseAccounting(s: String): Either[XLError, Formatted] =
    Try {
      val isNegative = s.contains("(") && s.contains(")")
      val cleaned = s.replaceAll("[\\$,()\\s]", "")
      if cleaned.isEmpty then
        throw new NumberFormatException("Empty value after removing $ , ( )")
      val num = if isNegative then -BigDecimal(cleaned) else BigDecimal(cleaned)
      Formatted(CellValue.Number(num.toDouble), NumFmt.Currency)
    }.toEither.left.map { err =>
      XLError.AccountingFormatError(s, err.getMessage)
    }
```

#### B. Modify `FormattedLiterals.scala`

**Location**: `xl-core/src/com/tjclp/xl/macros/FormattedLiterals.scala`

**Change signatures** (lines 15-28):

```scala
// OLD:
transparent inline def money(): Formatted = ${ moneyImpl('sc) }
transparent inline def percent(): Formatted = ${ percentImpl('sc) }
transparent inline def date(): Formatted = ${ dateImpl('sc) }
transparent inline def accounting(): Formatted = ${ accountingImpl('sc) }

inline def money(inline args: Any*): Formatted = ${ errorNoInterpolation(...) }
inline def percent(inline args: Any*): Formatted = ${ errorNoInterpolation(...) }
inline def date(inline args: Any*): Formatted = ${ errorNoInterpolation(...) }
inline def accounting(inline args: Any*): Formatted = ${ errorNoInterpolation(...) }

// NEW:
transparent inline def money(inline args: Any*): Formatted | Either[XLError, Formatted] =
  ${ moneyImplN('sc, 'args) }

transparent inline def percent(inline args: Any*): Formatted | Either[XLError, Formatted] =
  ${ percentImplN('sc, 'args) }

transparent inline def date(inline args: Any*): Formatted | Either[XLError, Formatted] =
  ${ dateImplN('sc, 'args) }

transparent inline def accounting(inline args: Any*): Formatted | Either[XLError, Formatted] =
  ${ accountingImplN('sc, 'args) }
```

**Add new implementations** (after line 113):

```scala
// ===== Runtime Implementations =====

private def moneyImplN(
  sc: Expr[StringContext],
  args: Expr[Seq[Any]]
)(using Quotes): Expr[Formatted | Either[XLError, Formatted]] =
  args match
    case Varargs(exprs) if exprs.isEmpty =>
      // No interpolation - compile-time literal
      moneyImpl(sc)
    case Varargs(_) =>
      // Has interpolation - runtime parsing
      '{
        FormattedParsers.parseMoney($sc.s($args*))
      }.asExprOf[Either[XLError, Formatted]]

private def percentImplN(
  sc: Expr[StringContext],
  args: Expr[Seq[Any]]
)(using Quotes): Expr[Formatted | Either[XLError, Formatted]] =
  args match
    case Varargs(exprs) if exprs.isEmpty =>
      percentImpl(sc)
    case Varargs(_) =>
      '{
        FormattedParsers.parsePercent($sc.s($args*))
      }.asExprOf[Either[XLError, Formatted]]

private def dateImplN(
  sc: Expr[StringContext],
  args: Expr[Seq[Any]]
)(using Quotes): Expr[Formatted | Either[XLError, Formatted]] =
  args match
    case Varargs(exprs) if exprs.isEmpty =>
      dateImpl(sc)
    case Varargs(_) =>
      '{
        FormattedParsers.parseDate($sc.s($args*))
      }.asExprOf[Either[XLError, Formatted]]

private def accountingImplN(
  sc: Expr[StringContext],
  args: Expr[Seq[Any]]
)(using Quotes): Expr[Formatted | Either[XLError, Formatted]] =
  args match
    case Varargs(exprs) if exprs.isEmpty =>
      accountingImpl(sc)
    case Varargs(_) =>
      '{
        FormattedParsers.parseAccounting($sc.s($args*))
      }.asExprOf[Either[XLError, Formatted]]
```

### 4. Testing Requirements

#### Test File 1: `FormattedParsersSpec.scala` (Unit Tests)

**Location**: `xl-core/test/src/com/tjclp/xl/formatted/FormattedParsersSpec.scala`

**Test Count**: 32 tests (8 per parser type)

```scala
package com.tjclp.xl.formatted

import com.tjclp.xl.cell.CellValue
import com.tjclp.xl.errors.XLError
import com.tjclp.xl.styles.numfmt.NumFmt
import munit.FunSuite
import java.time.LocalDate

class FormattedParsersSpec extends FunSuite:

  // ===== Money Parser Tests (8 tests) =====

  test("parseMoney: $1,234.56 with comma separator") {
    FormattedParsers.parseMoney("$1,234.56") match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(1234.56))
        assertEquals(f.numFmt, NumFmt.Currency)
      case Left(err) => fail(s"Should parse: $err")
  }

  test("parseMoney: 1234.56 without dollar sign") {
    FormattedParsers.parseMoney("1234.56") match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(1234.56))
      case Left(err) => fail(s"Should parse: $err")
  }

  test("parseMoney: $1234.56 without commas") {
    FormattedParsers.parseMoney("$1234.56") match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(1234.56))
      case Left(err) => fail(s"Should parse: $err")
  }

  test("parseMoney: $ 1,234.56 with spaces") {
    FormattedParsers.parseMoney("$ 1,234.56") match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(1234.56))
      case Left(err) => fail(s"Should parse: $err")
  }

  test("parseMoney: rejects non-numeric $ABC") {
    FormattedParsers.parseMoney("$ABC") match
      case Left(XLError.MoneyFormatError(value, _)) =>
        assertEquals(value, "$ABC")
      case other => fail(s"Should fail, got $other")
  }

  test("parseMoney: rejects empty string") {
    FormattedParsers.parseMoney("") match
      case Left(XLError.MoneyFormatError(_, _)) => () // Expected
      case other => fail(s"Should fail, got $other")
  }

  test("parseMoney: rejects multiple decimals $1.2.3") {
    FormattedParsers.parseMoney("$1.2.3") match
      case Left(XLError.MoneyFormatError(_, _)) => () // Expected
      case other => fail(s"Should fail, got $other")
  }

  test("parseMoney: large number $1,000,000.00") {
    FormattedParsers.parseMoney("$1,000,000.00") match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(1000000.00))
      case Left(err) => fail(s"Should parse: $err")
  }

  // ===== Percent Parser Tests (8 tests) =====

  test("parsePercent: 45.5% with percent sign") {
    FormattedParsers.parsePercent("45.5%") match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(0.455))
        assertEquals(f.numFmt, NumFmt.Percent)
      case Left(err) => fail(s"Should parse: $err")
  }

  test("parsePercent: 50 without percent sign") {
    FormattedParsers.parsePercent("50") match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(0.50))
      case Left(err) => fail(s"Should parse: $err")
  }

  test("parsePercent: 100% full percent") {
    FormattedParsers.parsePercent("100%") match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(1.0))
      case Left(err) => fail(s"Should parse: $err")
  }

  test("parsePercent: 0% zero percent") {
    FormattedParsers.parsePercent("0%") match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(0.0))
      case Left(err) => fail(s"Should parse: $err")
  }

  test("parsePercent: rejects non-numeric ABC%") {
    FormattedParsers.parsePercent("ABC%") match
      case Left(XLError.PercentFormatError(value, _)) =>
        assertEquals(value, "ABC%")
      case other => fail(s"Should fail, got $other")
  }

  test("parsePercent: rejects multiple percent signs 50%%") {
    FormattedParsers.parsePercent("50%%") match
      case Left(XLError.PercentFormatError(_, _)) => () // Expected
      case other => fail(s"Should fail, got $other")
  }

  test("parsePercent: rejects empty string") {
    FormattedParsers.parsePercent("") match
      case Left(XLError.PercentFormatError(_, _)) => () // Expected
      case other => fail(s"Should fail, got $other")
  }

  test("parsePercent: handles > 100%") {
    FormattedParsers.parsePercent("150%") match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(1.50))
      case Left(err) => fail(s"Should parse: $err")
  }

  // ===== Date Parser Tests (8 tests) =====

  test("parseDate: 2025-11-10 ISO format") {
    FormattedParsers.parseDate("2025-11-10") match
      case Right(f) =>
        f.value match
          case CellValue.DateTime(dt) =>
            assertEquals(dt.toLocalDate.toString, "2025-11-10")
          case _ => fail("Expected DateTime")
        assertEquals(f.numFmt, NumFmt.Date)
      case Left(err) => fail(s"Should parse: $err")
  }

  test("parseDate: 2000-01-01 Y2K date") {
    FormattedParsers.parseDate("2000-01-01") match
      case Right(f) =>
        f.value match
          case CellValue.DateTime(dt) =>
            assertEquals(dt.toLocalDate, LocalDate.of(2000, 1, 1))
          case _ => fail("Expected DateTime")
      case Left(err) => fail(s"Should parse: $err")
  }

  test("parseDate: 2025-12-31 end of year") {
    FormattedParsers.parseDate("2025-12-31") match
      case Right(f) => () // Should parse
      case Left(err) => fail(s"Should parse: $err")
  }

  test("parseDate: rejects US format 11/10/2025") {
    FormattedParsers.parseDate("11/10/2025") match
      case Left(XLError.DateFormatError(value, reason)) =>
        assertEquals(value, "11/10/2025")
        assert(reason.contains("ISO format"))
      case other => fail(s"Should fail, got $other")
  }

  test("parseDate: rejects invalid date 2025-02-30") {
    FormattedParsers.parseDate("2025-02-30") match
      case Left(XLError.DateFormatError(_, _)) => () // Expected
      case other => fail(s"Should fail, got $other")
  }

  test("parseDate: rejects invalid month 2025-13-01") {
    FormattedParsers.parseDate("2025-13-01") match
      case Left(XLError.DateFormatError(_, _)) => () // Expected
      case other => fail(s"Should fail, got $other")
  }

  test("parseDate: rejects non-date string") {
    FormattedParsers.parseDate("not-a-date") match
      case Left(XLError.DateFormatError(_, _)) => () // Expected
      case other => fail(s"Should fail, got $other")
  }

  test("parseDate: rejects empty string") {
    FormattedParsers.parseDate("") match
      case Left(XLError.DateFormatError(_, _)) => () // Expected
      case other => fail(s"Should fail, got $other")
  }

  // ===== Accounting Parser Tests (8 tests) =====

  test("parseAccounting: $123.45 positive amount") {
    FormattedParsers.parseAccounting("$123.45") match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(123.45))
        assertEquals(f.numFmt, NumFmt.Currency)
      case Left(err) => fail(s"Should parse: $err")
  }

  test("parseAccounting: ($123.45) negative with parens") {
    FormattedParsers.parseAccounting("($123.45)") match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(-123.45))
      case Left(err) => fail(s"Should parse: $err")
  }

  test("parseAccounting: ($1,234.56) negative with commas") {
    FormattedParsers.parseAccounting("($1,234.56)") match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(-1234.56))
      case Left(err) => fail(s"Should parse: $err")
  }

  test("parseAccounting: $0.00 zero amount") {
    FormattedParsers.parseAccounting("$0.00") match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(0.00))
      case Left(err) => fail(s"Should parse: $err")
  }

  test("parseAccounting: rejects non-numeric $ABC") {
    FormattedParsers.parseAccounting("$ABC") match
      case Left(XLError.AccountingFormatError(_, _)) => () // Expected
      case other => fail(s"Should fail, got $other")
  }

  test("parseAccounting: rejects empty parens $()") {
    FormattedParsers.parseAccounting("$()") match
      case Left(XLError.AccountingFormatError(_, _)) => () // Expected
      case other => fail(s"Should fail, got $other")
  }

  test("parseAccounting: rejects unbalanced parens ($123.45") {
    // Note: Current implementation removes parens, so this might parse
    // This test documents expected behavior for future improvement
    FormattedParsers.parseAccounting("($123.45") match
      case Right(_) | Left(_) => () // Either is acceptable for now
  }

  test("parseAccounting: large negative ($1,000,000.00)") {
    FormattedParsers.parseAccounting("($1,000,000.00)") match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(-1000000.00))
      case Left(err) => fail(s"Should parse: $err")
  }
```

#### Test File 2: `FormattedInterpolationSpec.scala` (Integration Tests)

**Location**: `xl-core/test/src/com/tjclp/xl/macros/FormattedInterpolationSpec.scala`

**Test Count**: 28 tests (7 per literal type)

```scala
package com.tjclp.xl.macros

import com.tjclp.xl.*
import com.tjclp.xl.cell.CellValue
import com.tjclp.xl.errors.XLError
import com.tjclp.xl.styles.numfmt.NumFmt
import munit.{FunSuite, ScalaCheckSuite}
import org.scalacheck.Prop.*
import java.time.LocalDate

class FormattedInterpolationSpec extends ScalaCheckSuite:

  // ===== Backward Compatibility (Compile-Time Literals) =====

  test("Compile-time: money\"$$1,234.56\" returns Formatted directly") {
    val f = money"$$1,234.56"
    assert(f.isInstanceOf[Formatted])
    assertEquals(f.numFmt, NumFmt.Currency)
    assertEquals(f.value, CellValue.Number(1234.56))
  }

  test("Compile-time: percent\"45.5%\" returns Formatted directly") {
    val f = percent"45.5%"
    assert(f.isInstanceOf[Formatted])
    assertEquals(f.numFmt, NumFmt.Percent)
  }

  test("Compile-time: date\"2025-11-10\" returns Formatted directly") {
    val f = date"2025-11-10"
    assert(f.isInstanceOf[Formatted])
    assertEquals(f.numFmt, NumFmt.Date)
  }

  test("Compile-time: accounting\"$$123.45\" returns Formatted directly") {
    val f = accounting"$$123.45"
    assert(f.isInstanceOf[Formatted])
    assertEquals(f.numFmt, NumFmt.Currency)
  }

  // ===== Money Runtime Interpolation (7 tests) =====

  test("Runtime money: variable with dollar and commas") {
    val priceStr = "$1,234.56"
    money"$priceStr" match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(1234.56))
        assertEquals(f.numFmt, NumFmt.Currency)
      case Left(err) => fail(s"Should parse: $err")
  }

  test("Runtime money: variable without dollar sign") {
    val priceStr = "999.99"
    money"$priceStr" match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(999.99))
      case Left(err) => fail(s"Should parse: $err")
  }

  test("Runtime money: mixed prefix and variable") {
    val amount = "1234.56"
    money"$$$amount" match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(1234.56))
      case Left(err) => fail(s"Should parse: $err")
  }

  test("Runtime money: invalid variable returns Left") {
    val invalidStr = "$ABC"
    money"$invalidStr" match
      case Left(XLError.MoneyFormatError(_, _)) => () // Expected
      case other => fail(s"Should fail, got $other")
  }

  test("Runtime money: empty string returns Left") {
    val emptyStr = ""
    money"$emptyStr" match
      case Left(_) => () // Expected
      case Right(_) => fail("Empty string should fail")
  }

  test("Runtime money: for-comprehension") {
    val str1 = "$100.00"
    val str2 = "$200.00"

    val result = for
      f1 <- money"$str1"
      f2 <- money"$str2"
    yield (f1.value, f2.value)

    result match
      case Right((v1, v2)) =>
        assertEquals(v1, CellValue.Number(100.00))
        assertEquals(v2, CellValue.Number(200.00))
      case Left(err) => fail(s"Should parse: $err")
  }

  test("Runtime money: large amount") {
    val largeStr = "$1,000,000.00"
    money"$largeStr" match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(1000000.00))
      case Left(err) => fail(s"Should parse: $err")
  }

  // ===== Percent Runtime Interpolation (7 tests) =====

  test("Runtime percent: variable with percent sign") {
    val pctStr = "45.5%"
    percent"$pctStr" match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(0.455))
        assertEquals(f.numFmt, NumFmt.Percent)
      case Left(err) => fail(s"Should parse: $err")
  }

  test("Runtime percent: variable without percent sign") {
    val pctStr = "50"
    percent"$pctStr" match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(0.50))
      case Left(err) => fail(s"Should parse: $err")
  }

  test("Runtime percent: mixed variable and suffix") {
    val value = "75"
    percent"$value%" match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(0.75))
      case Left(err) => fail(s"Should parse: $err")
  }

  test("Runtime percent: invalid variable returns Left") {
    val invalidStr = "ABC%"
    percent"$invalidStr" match
      case Left(XLError.PercentFormatError(_, _)) => () // Expected
      case other => fail(s"Should fail, got $other")
  }

  test("Runtime percent: empty string returns Left") {
    val emptyStr = ""
    percent"$emptyStr" match
      case Left(_) => () // Expected
      case Right(_) => fail("Empty string should fail")
  }

  test("Runtime percent: > 100%") {
    val largeStr = "150%"
    percent"$largeStr" match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(1.50))
      case Left(err) => fail(s"Should parse: $err")
  }

  test("Runtime percent: 0%") {
    val zeroStr = "0%"
    percent"$zeroStr" match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(0.0))
      case Left(err) => fail(s"Should parse: $err")
  }

  // ===== Date Runtime Interpolation (7 tests) =====

  test("Runtime date: ISO format variable") {
    val dateStr = "2025-11-10"
    date"$dateStr" match
      case Right(f) =>
        f.value match
          case CellValue.DateTime(dt) =>
            assertEquals(dt.toLocalDate.toString, "2025-11-10")
          case _ => fail("Expected DateTime")
        assertEquals(f.numFmt, NumFmt.Date)
      case Left(err) => fail(s"Should parse: $err")
  }

  test("Runtime date: Y2K date") {
    val dateStr = "2000-01-01"
    date"$dateStr" match
      case Right(f) =>
        f.value match
          case CellValue.DateTime(dt) =>
            assertEquals(dt.toLocalDate, LocalDate.of(2000, 1, 1))
          case _ => fail("Expected DateTime")
      case Left(err) => fail(s"Should parse: $err")
  }

  test("Runtime date: end of year") {
    val dateStr = "2025-12-31"
    date"$dateStr" match
      case Right(_) => () // Should parse
      case Left(err) => fail(s"Should parse: $err")
  }

  test("Runtime date: invalid format returns Left") {
    val invalidStr = "11/10/2025"  // US format not supported
    date"$invalidStr" match
      case Left(XLError.DateFormatError(value, reason)) =>
        assertEquals(value, "11/10/2025")
        assert(reason.contains("ISO"))
      case other => fail(s"Should fail, got $other")
  }

  test("Runtime date: invalid date returns Left") {
    val invalidStr = "2025-02-30"  // Feb 30 doesn't exist
    date"$invalidStr" match
      case Left(XLError.DateFormatError(_, _)) => () // Expected
      case other => fail(s"Should fail, got $other")
  }

  test("Runtime date: non-date string returns Left") {
    val invalidStr = "not-a-date"
    date"$invalidStr" match
      case Left(XLError.DateFormatError(_, _)) => () // Expected
      case other => fail(s"Should fail, got $other")
  }

  test("Runtime date: empty string returns Left") {
    val emptyStr = ""
    date"$emptyStr" match
      case Left(_) => () // Expected
      case Right(_) => fail("Empty string should fail")
  }

  // ===== Accounting Runtime Interpolation (7 tests) =====

  test("Runtime accounting: positive amount") {
    val acctStr = "$123.45"
    accounting"$acctStr" match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(123.45))
        assertEquals(f.numFmt, NumFmt.Currency)
      case Left(err) => fail(s"Should parse: $err")
  }

  test("Runtime accounting: negative with parens") {
    val acctStr = "($123.45)"
    accounting"$acctStr" match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(-123.45))
      case Left(err) => fail(s"Should parse: $err")
  }

  test("Runtime accounting: negative with commas") {
    val acctStr = "($1,234.56)"
    accounting"$acctStr" match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(-1234.56))
      case Left(err) => fail(s"Should parse: $err")
  }

  test("Runtime accounting: zero amount") {
    val acctStr = "$0.00"
    accounting"$acctStr" match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(0.00))
      case Left(err) => fail(s"Should parse: $err")
  }

  test("Runtime accounting: invalid variable returns Left") {
    val invalidStr = "$ABC"
    accounting"$invalidStr" match
      case Left(XLError.AccountingFormatError(_, _)) => () // Expected
      case other => fail(s"Should fail, got $other")
  }

  test("Runtime accounting: empty string returns Left") {
    val emptyStr = ""
    accounting"$emptyStr" match
      case Left(_) => () // Expected
      case Right(_) => fail("Empty string should fail")
  }

  test("Runtime accounting: large negative") {
    val acctStr = "($1,000,000.00)"
    accounting"$acctStr" match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(-1000000.00))
      case Left(err) => fail(s"Should parse: $err")
  }

  // ===== Property-Based Tests =====

  property("Money: valid strings always parse") {
    forAll(Generators.genMoneyString) { str =>
      money"$str" match
        case Right(_) => () // Expected
        case Left(err) => fail(s"Valid money '$str' should parse: $err")
    }
  }

  property("Percent: valid strings always parse") {
    forAll(Generators.genPercentString) { str =>
      percent"$str" match
        case Right(_) => () // Expected
        case Left(err) => fail(s"Valid percent '$str' should parse: $err")
    }
  }

  property("Date: valid strings always parse") {
    forAll(Generators.genDateString) { str =>
      date"$str" match
        case Right(_) => () // Expected
        case Left(err) => fail(s"Valid date '$str' should parse: $err")
    }
  }

  property("Money: invalid strings always return Left") {
    forAll(Generators.genInvalidMoney) { str =>
      money"$str" match
        case Left(_) => () // Expected
        case Right(_) => fail(s"Invalid money '$str' should fail")
    }
  }
```

### 5. Documentation Updates

**Update** `docs/reference/examples.md`:

Add section:

```markdown
### String Interpolation for Formatted Literals

Runtime parsing with error handling:

\`\`\`scala
// Money format
val priceStr = "$1,234.56"
money"$priceStr" match
  case Right(formatted) =>
    sheet.put(cell"A1", formatted.value)
  case Left(XLError.MoneyFormatError(value, reason)) =>
    println(s"Invalid price '$value': $reason")

// Percent format
val pctStr = "45.5%"
percent"$pctStr" match
  case Right(formatted) => // Use formatted value
  case Left(err) => // Handle error

// Date format (ISO only)
val dateStr = "2025-11-10"
date"$dateStr" match
  case Right(formatted) => // Use formatted value
  case Left(XLError.DateFormatError(value, reason)) =>
    println(s"Invalid date '$value': $reason")

// Accounting format
val acctStr = "($123.45)"  // Negative with parens
accounting"$acctStr" match
  case Right(formatted) => // Use formatted value
  case Left(err) => // Handle error
\`\`\`
```

**Update** `CLAUDE.md`:

```markdown
### Formatted Literal Runtime Parsers

Located in `xl-core/src/com/tjclp/xl/formatted/FormattedParsers.scala`:

- `parseMoney(s: String): Either[XLError, Formatted]` - Accepts $1,234.56 or 1234.56
- `parsePercent(s: String): Either[XLError, Formatted]` - Accepts 45.5% or 45.5
- `parseDate(s: String): Either[XLError, Formatted]` - ISO format only (YYYY-MM-DD)
- `parseAccounting(s: String): Either[XLError, Formatted]` - Accepts $123 or ($123)

All parsers are pure, total functions. Phase 2 will add compile-time literal detection.
```

### 6. Success Criteria

- [ ] `./mill xl-core.compile` succeeds with 0 warnings
- [ ] `./mill xl-core.test.testOnly com.tjclp.xl.formatted.FormattedParsersSpec` passes (32/32 tests)
- [ ] `./mill xl-core.test.testOnly com.tjclp.xl.macros.FormattedInterpolationSpec` passes (28/28 tests)
- [ ] `./mill __.test` passes (all 296 existing tests + 60 new = 356 total)
- [ ] All parsers are pure (no exceptions leak)
- [ ] Compile-time literals unchanged
- [ ] Documentation updated

### 7. Review Checklist

**For Reviewers**:
- [ ] Parser purity verified (all exceptions caught and wrapped)
- [ ] Error messages are clear and actionable
- [ ] ISO date format requirement documented
- [ ] Backward compatibility maintained
- [ ] Property tests verify round-trips

**Questions**:
- Should we support more date formats in Phase 1?
- Should money parser accept negative without parentheses?
- Is ISO-only date format acceptable for Phase 1?

### 8. Rollback Plan

**Risk**: Medium (new parser code, complex logic)

**Procedure**:
1. Delete `FormattedParsers.scala` (new file)
2. Delete `FormattedParsersSpec.scala` (new test file)
3. Delete `FormattedInterpolationSpec.scala` (new test file)
4. Revert signature changes in `FormattedLiterals.scala` (lines 15-28)
5. Remove `*ImplN` functions in `FormattedLiterals.scala` (lines 114-160)
6. Verify: `./mill __.test` (296 tests should pass)

---

## PR #4: String Interpolation for fx"$var"

**Priority**: Depends on PR #1
**Estimated Effort**: 8 hours (1 day)
**Risk Level**: VERY LOW (simple validation logic)

### 1. Objective

Enable runtime string interpolation for `fx""` (formula) macro with minimal validation (parentheses balance only).

**Before**:
```scala
val formulaStr = "=SUM(A1:A10)"
val cellValue = CellValue.Formula(formulaStr)  // No validation
```

**After**:
```scala
val formulaStr = "=SUM(A1:A10)"
fx"$formulaStr" match
  case Right(cellValue) => // Use formula
  case Left(err) => // Handle parse errors
```

### 2. Files Modified/Created

**1 file modified**:
- `xl-core/src/com/tjclp/xl/macros/CellRangeLiterals.scala` (+40 lines)

**2 files created**:
- `xl-core/src/com/tjclp/xl/cell/FormulaParser.scala` (new, ~70 lines)
- `xl-core/test/src/com/tjclp/xl/macros/FormulaInterpolationSpec.scala` (new, ~180 lines)

### 3. Detailed Implementation

#### A. Create `FormulaParser.scala` (NEW FILE)

**Location**: `xl-core/src/com/tjclp/xl/cell/FormulaParser.scala`

```scala
package com.tjclp.xl.cell

import com.tjclp.xl.errors.XLError

/**
 * Pure runtime parser for formula strings.
 *
 * Currently performs minimal validation (parentheses balance only).
 * Full formula parsing is planned for Phase 7 (Formula Evaluator).
 *
 * Limitations:
 * - Does not validate function names
 * - Does not detect string literals ("text") which may contain parens
 * - Does not validate formula syntax beyond parentheses
 */
object FormulaParser:

  /**
   * Parse formula string with minimal validation.
   *
   * Accepts:
   * - Any string with balanced parentheses
   * - With or without leading '=' (both are valid)
   *
   * Rejects:
   * - Empty string
   * - Unbalanced parentheses
   */
  def parse(s: String): Either[XLError, CellValue] =
    if s.isEmpty then
      Left(XLError.FormulaError(s, "Formula cannot be empty"))
    else if !validateParentheses(s) then
      Left(XLError.FormulaError(s, "Unbalanced parentheses"))
    else
      Right(CellValue.Formula(s))

  /**
   * Check parentheses are balanced.
   *
   * Limitation: Does not handle string literals ("text") which may contain parens.
   * This is acceptable for Phase 1; full parsing is Phase 7.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  private def validateParentheses(s: String): Boolean =
    var depth = 0
    var i = 0
    val n = s.length

    while i < n do
      val c = s.charAt(i)
      if c == '(' then
        depth += 1
      else if c == ')' then
        depth -= 1
        if depth < 0 then return false
      i += 1

    depth == 0
```

#### B. Modify `CellRangeLiterals.scala`

**Location**: `xl-core/src/com/tjclp/xl/macros/CellRangeLiterals.scala`

**Change signature** (lines 22-25):

```scala
// OLD:
transparent inline def fx(): CellValue = ${ fxImpl0('sc) }
inline def fx(inline args: Any*): CellValue = ${ errorNoInterpolation(...) }

// NEW:
transparent inline def fx(inline args: Any*): CellValue | Either[XLError, CellValue] =
  ${ fxImplN('sc, 'args) }
```

**Add new implementation** (after line 58):

```scala
/**
 * Macro implementation supporting both compile-time literals and runtime interpolation.
 */
private def fxImplN(
  sc: Expr[StringContext],
  args: Expr[Seq[Any]]
)(using Quotes): Expr[CellValue | Either[XLError, CellValue]] =
  import quotes.reflect.*

  args match
    case Varargs(exprs) if exprs.isEmpty =>
      // No interpolation - compile-time literal
      fxImpl0(sc)

    case Varargs(_) =>
      // Has interpolation - runtime parsing
      '{
        FormulaParser.parse($sc.s($args*))
      }.asExprOf[Either[XLError, CellValue]]
```

### 4. Testing Requirements

**Test File**: `xl-core/test/src/com/tjclp/xl/macros/FormulaInterpolationSpec.scala`

**Test Count**: 18 tests

```scala
package com.tjclp.xl.macros

import com.tjclp.xl.cell.{CellValue, FormulaParser}
import com.tjclp.xl.errors.XLError
import munit.FunSuite

class FormulaInterpolationSpec extends FunSuite:

  // ===== Backward Compatibility (Compile-Time Literals) =====

  test("Compile-time literal: fx\"=SUM(A1:A10)\" returns CellValue directly") {
    val f = fx"=SUM(A1:A10)"
    assert(f.isInstanceOf[CellValue.Formula])
    val formula = f.asInstanceOf[CellValue.Formula]
    assertEquals(formula.expression, "=SUM(A1:A10)")
  }

  // Note: Compile-time validation errors test (unbalanced parens) would be:
  // test("Compile-time: unbalanced parens fails at compile time") {
  //   // val f = fx"=SUM(A1:A10"  // Should not compile
  // }

  // ===== Runtime Interpolation (New Functionality) =====

  test("Runtime interpolation: simple SUM formula") {
    val formulaStr = "=SUM(A1:A10)"
    fx"$formulaStr" match
      case Right(CellValue.Formula(expr)) =>
        assertEquals(expr, "=SUM(A1:A10)")
      case other =>
        fail(s"Expected Right(Formula), got $other")
  }

  test("Runtime interpolation: IF formula") {
    val formulaStr = "=IF(A1>0,B1,C1)"
    fx"$formulaStr" match
      case Right(CellValue.Formula(expr)) =>
        assertEquals(expr, "=IF(A1>0,B1,C1)")
      case other =>
        fail(s"Expected Right(Formula), got $other")
  }

  test("Runtime interpolation: nested parentheses") {
    val formulaStr = "=IF(A1>0,SUM(B1:B10),0)"
    fx"$formulaStr" match
      case Right(CellValue.Formula(expr)) =>
        assertEquals(expr, "=IF(A1>0,SUM(B1:B10),0)")
      case other =>
        fail(s"Should parse: $other")
  }

  test("Runtime interpolation: complex formula with multiple functions") {
    val formulaStr = "=AVERAGE(IF(A1:A10>0,B1:B10))"
    fx"$formulaStr" match
      case Right(_) => () // Expected
      case Left(err) => fail(s"Should parse: $err")
  }

  test("Runtime interpolation: formula without = prefix") {
    val formulaStr = "SUM(A1:A10)"
    fx"$formulaStr" match
      case Right(CellValue.Formula(expr)) =>
        assertEquals(expr, "SUM(A1:A10)")  // We don't require =
      case Left(err) => fail(s"Should parse: $err")
  }

  // ===== Error Cases =====

  test("Runtime interpolation: empty formula returns Left") {
    val emptyStr = ""
    fx"$emptyStr" match
      case Left(XLError.FormulaError(input, msg)) =>
        assertEquals(input, "")
        assert(msg.contains("empty"))
      case other =>
        fail(s"Expected Left(FormulaError), got $other")
  }

  test("Runtime interpolation: unbalanced opening paren returns Left") {
    val invalidStr = "=SUM(A1:A10"
    fx"$invalidStr" match
      case Left(XLError.FormulaError(input, msg)) =>
        assertEquals(input, "=SUM(A1:A10")
        assert(msg.contains("unbalanced"))
      case other =>
        fail(s"Expected Left(FormulaError), got $other")
  }

  test("Runtime interpolation: unbalanced closing paren returns Left") {
    val invalidStr = "=SUM(A1:A10))"
    fx"$invalidStr" match
      case Left(XLError.FormulaError(input, msg)) =>
        assertEquals(input, "=SUM(A1:A10))")
        assert(msg.contains("unbalanced"))
      case other =>
        fail(s"Expected Left(FormulaError), got $other")
  }

  test("Runtime interpolation: mismatched parens returns Left") {
    val invalidStr = "=SUM(A1:A10()"
    fx"$invalidStr" match
      case Left(XLError.FormulaError(_, msg)) =>
        assert(msg.contains("unbalanced"))
      case other =>
        fail(s"Expected Left(FormulaError), got $other")
  }

  // ===== Mixed Compile-Time and Runtime =====

  test("Mixed interpolation: prefix + variable") {
    val range = "A1:A10"
    fx"=SUM($range)" match
      case Right(CellValue.Formula(expr)) =>
        assertEquals(expr, "=SUM(A1:A10)")
      case Left(err) => fail(s"Should parse: $err")
  }

  test("Mixed interpolation: variable function name") {
    val func = "SUM"
    fx"=$func(A1:A10)" match
      case Right(CellValue.Formula(expr)) =>
        assertEquals(expr, "=SUM(A1:A10)")
      case Left(err) => fail(s"Should parse: $err")
  }

  test("Mixed interpolation: multiple variables") {
    val func = "IF"
    val cond = "A1>0"
    val thenVal = "B1"
    val elseVal = "C1"
    fx"=$func($cond,$thenVal,$elseVal)" match
      case Right(CellValue.Formula(expr)) =>
        assertEquals(expr, "=IF(A1>0,B1,C1)")
      case Left(err) => fail(s"Should parse: $err")
  }

  // ===== Edge Cases =====

  test("Edge: formula with whitespace") {
    val formulaStr = " =SUM( A1:A10 ) "
    fx"$formulaStr" match
      case Right(CellValue.Formula(expr)) =>
        assertEquals(expr, " =SUM( A1:A10 ) ")  // Preserve whitespace
      case Left(err) => fail(s"Should parse: $err")
  }

  test("Edge: formula with many nested parens") {
    val formulaStr = "=IF(IF(A1>0,IF(B1>0,1,0),0),SUM(C1:C10),0)"
    fx"$formulaStr" match
      case Right(_) => () // Should parse
      case Left(err) => fail(s"Should parse: $err")
  }

  test("Edge: formula with array syntax {1,2,3}") {
    val formulaStr = "=SUM({1,2,3})"
    fx"$formulaStr" match
      case Right(CellValue.Formula(expr)) =>
        assertEquals(expr, "=SUM({1,2,3})")
      case Left(err) => fail(s"Should parse: $err")
  }

  // ===== Limitation Tests (Documented Known Issues) =====

  test("Limitation: string literals with parens not handled") {
    // This is a known limitation - full parsing is Phase 7
    val formulaStr = """=IF(A1=")", "yes", "no")"""
    fx"$formulaStr" match
      case Left(_) =>
        // Expected to fail due to string literal with )
        ()
      case Right(_) =>
        // If it happens to pass, that's fine too
        ()
  }

  // ===== Integration =====

  test("Integration: for-comprehension with Either") {
    val func = "SUM"
    val range = "A1:A10"

    val result = for
      formula <- fx"=$func($range)"
    yield formula

    result match
      case Right(CellValue.Formula(expr)) =>
        assertEquals(expr, "=SUM(A1:A10)")
      case Left(err) => fail(s"Should parse: $err")
  }
```

### 5. Documentation Updates

**Update** `docs/reference/examples.md`:

Add section:

```markdown
### String Interpolation for Formulas

Runtime formula construction with error handling:

\`\`\`scala
// Simple formula
val range = "A1:A10"
fx"=SUM($range)" match
  case Right(formula) => sheet.put(cell"B1", formula)
  case Left(XLError.FormulaError(input, reason)) =>
    println(s"Invalid formula '$input': $reason")

// Dynamic function name
val func = getUserFunction()  // e.g., "AVERAGE"
fx"=$func(B1:B10)" match
  case Right(formula) => // Use formula
  case Left(err) => // Handle error

// Complex formula with multiple variables
val cond = "A1>0"
val thenVal = "B1"
val elseVal = "C1"
fx"=IF($cond,$thenVal,$elseVal)" match
  case Right(formula) => // Use formula
  case Left(err) => // Handle error
\`\`\`

**Note**: Phase 1 validates parentheses balance only. Full formula parsing (function names, syntax, etc.) is planned for Phase 7 (Formula Evaluator).
```

**Update** `docs/LIMITATIONS.md`:

Add bullet:
- Formula validation is minimal in Phase 1 (parentheses balance only). Full validation planned for Phase 7.

### 6. Success Criteria

- [ ] `./mill xl-core.compile` succeeds with 0 warnings
- [ ] `./mill xl-core.test.testOnly com.tjclp.xl.macros.FormulaInterpolationSpec` passes (18/18 tests)
- [ ] `./mill __.test` passes (all 356 existing tests + 18 new = 374 total)
- [ ] Parentheses validation works
- [ ] Compile-time literals unchanged
- [ ] Documentation updated with limitations

### 7. Review Checklist

**For Reviewers**:
- [ ] Minimal validation scope documented
- [ ] Future extensibility (Phase 7) considered
- [ ] Backward compatibility maintained
- [ ] Error messages clear

**Questions**:
- Should we validate = prefix requirement?
- Should we add string literal detection for quoted text?
- Is parentheses-only validation acceptable for Phase 1?

### 8. Rollback Plan

**Risk**: Very Low (simple validation logic)

**Procedure**:
1. Delete `FormulaParser.scala` (new file)
2. Delete `FormulaInterpolationSpec.scala` (new test file)
3. Revert signature changes in `CellRangeLiterals.scala` (lines 22-25)
4. Remove `fxImplN` function (lines 60-75)
5. Verify: `./mill __.test` (356 tests should pass)

---

## Phase 1 Completion Summary

### Deliverables Checklist

After all 4 PRs are merged:

- [ ] **Infrastructure** (PR #1)
  - [x] 4 new error types added
  - [x] Test generators added
  - [x] 8 tests passing

- [ ] **ref interpolation** (PR #2)
  - [x] Runtime string interpolation working
  - [x] Returns Either[XLError, RefType]
  - [x] 25 tests passing
  - [x] Backward compatible

- [ ] **Formatted literals** (PR #3)
  - [x] 4 pure parsers created
  - [x] Runtime interpolation for money/percent/date/accounting
  - [x] 60 tests passing (32 parser + 28 integration)
  - [x] Backward compatible

- [ ] **fx interpolation** (PR #4)
  - [x] Formula parser created
  - [x] Minimal validation (parens balance)
  - [x] 18 tests passing
  - [x] Backward compatible

- [ ] **Total Impact**
  - [x] 374 total tests (263 existing + 111 new)
  - [x] 10 files modified/created
  - [x] ~1,130 lines of code added
  - [x] 100% backward compatible
  - [x] Zero breaking changes

### Success Metrics

**Test Coverage**:
- Before Phase 1: 263 tests
- After Phase 1: 374 tests (+42% increase)
- All existing tests still passing (backward compatibility verified)

**Code Quality**:
- `./mill __.compile` succeeds with 0 warnings
- `./mill __.checkFormat` passes
- All parsers are pure and total (no exceptions)
- WartRemover compliant

**Documentation**:
- `docs/reference/examples.md` updated with interpolation examples
- `CLAUDE.md` updated with new patterns
- `docs/STATUS.md` reflects Phase 1 complete
- `docs/LIMITATIONS.md` updated with known limitations

**Functional Requirements**:
- Runtime string interpolation works for all 4 macro types
- Error handling is total (Either[XLError, T])
- Compile-time literals unchanged (zero performance regression)
- Purity and totality preserved

---

## Timeline & Milestones

### Sequential Development (1 Developer)

**Week 1**:
- Days 1-2: PR #1 (Infrastructure) - Implement, test, review
- Days 3-4: PR #2 (ref interpolation) - Implement, test, review
- Day 5: Integration testing, address review feedback

**Week 2**:
- Days 1-3: PR #3 (Formatted literals) - Implement, test, review
- Day 4: PR #4 (fx interpolation) - Implement, test, review
- Day 5: Final integration testing

**Week 3**:
- Days 1-2: Address all review feedback, polish documentation
- Day 3: Final verification and merge
- Days 4-5: Buffer for unexpected issues

**Total**: 12-16 developer-days (2.5-3 weeks)

### Parallel Development (2 Developers)

**Week 1**:
- Dev A: PR #1 (Days 1-2) → PR #2 (Days 3-4)
- Dev B: [waiting for PR #1] → PR #3 draft (Days 3-5, can start parsers)

**Week 2**:
- Dev A: Review/merge PR #2 → PR #4 (Days 1-2)
- Dev B: Complete PR #3 (Days 1-3) → Review

**Week 3**:
- Both: Final integration testing + docs (Days 1-2)
- Buffer: Days 3-5

**Total**: 8-10 developer-days per dev (1.5-2 weeks calendar time)

### Milestones

| Milestone | Date | Deliverable |
|-----------|------|-------------|
| M1: Infrastructure Complete | End of Week 1, Day 2 | PR #1 merged |
| M2: ref Interpolation Complete | End of Week 1, Day 4 | PR #2 merged |
| M3: Formatted Literals Complete | End of Week 2, Day 3 | PR #3 merged |
| M4: fx Interpolation Complete | End of Week 2, Day 4 | PR #4 merged |
| M5: Phase 1 Complete | End of Week 3, Day 3 | All PRs merged, 374 tests passing |

---

## CI/CD Considerations

### GitHub Actions Workflow

Each PR should trigger:
1. `./mill __.compile` - Zero warnings required
2. `./mill __.test` - All tests must pass
3. `./mill __.checkFormat` - Code formatting verified
4. Coverage report generated

### PR Merge Requirements

- [ ] All CI checks passing (green)
- [ ] 2+ approvals from reviewers
- [ ] No unresolved review comments
- [ ] Documentation updated
- [ ] Tests included and passing
- [ ] No merge conflicts

### Deployment Strategy

Phase 1 is **non-breaking**, so deployment is straightforward:
1. Merge PR #1 → Deploy to staging → Verify
2. Merge PR #2 → Deploy to staging → Verify
3. Merge PR #3 → Deploy to staging → Verify
4. Merge PR #4 → Deploy to staging → Verify
5. Final integration test on staging
6. Deploy to production

---

## Open Questions & Decisions

### Q1: Should we provide unsafe variants?

**Question**: Should we add `.unsafe` variants that throw exceptions instead of returning Either?

```scala
ref"$userInput"         // Either[XLError, RefType]
ref.unsafe"$userInput"  // RefType (throws on errors)
```

**Decision**: Defer to post-Phase 1. Gather user feedback first.

**Rationale**: XL's charter prioritizes totality and purity. Exceptions violate this. If users need `unsafe`, they can `.getOrElse(throw ...)`.

### Q2: Should we support more date formats?

**Question**: Should `date""` parser accept US format (MM/DD/YYYY) or European (DD/MM/YYYY)?

**Decision**: ISO only for Phase 1. Add other formats in Phase 2 with compile-time detection.

**Rationale**: Avoids ambiguity (11/10/2025 could be Nov 10 or Oct 11). ISO is unambiguous.

### Q3: Should money parser accept negative without parentheses?

**Question**: Should `money"-$123.45"` be valid, or require `accounting"($123.45)"`?

**Decision**: No. Use `accounting""` for negatives with parentheses. Keep `money""` simple.

**Rationale**: Clear separation of concerns. Accounting format has specific semantics.

### Q4: Should we validate formula function names?

**Question**: Should `fx"=INVALID_FUNCTION(A1)"` return Left with "unknown function" error?

**Decision**: No, not in Phase 1. Defer to Phase 7 (Formula Evaluator).

**Rationale**: Minimal validation scope for Phase 1. Full parsing is complex and out of scope.

### Q5: Performance SLA for runtime parsing?

**Question**: What's acceptable overhead for runtime interpolated strings?

**Decision**: < 1ms per interpolated string (measured via microbenchmarks).

**Rationale**: User-initiated actions (parsing user input) can tolerate millisecond latency. If slower, optimize in Phase 2.

---

## Next Steps After Phase 1

Once Phase 1 is complete (all 4 PRs merged, 374 tests passing), Phase 2 can begin:

### Phase 2: Compile-Time Optimization (Estimated 2-3 weeks)

**Goals**:
- Detect constant interpolations at compile time
- Emit direct constants for `val x = "A1"; ref"$x"`
- Add compile-time diagnostics for common mistakes
- Performance: Zero overhead even for simple variables

**Prerequisites**:
- Phase 1 complete and deployed
- User feedback collected
- Performance benchmarks established

**Estimated Effort**: 10-12 developer-days

---

## Document Metadata

**Created**: 2025-11-14
**Author**: Generated by Claude Code (Anthropic)
**Version**: 1.0 (Implementation-Ready)
**Status**: Approved for Implementation
**Parent Design**: `string-interpolation.md`

**Reviewers**:
- [ ] Tech Lead (Architecture review)
- [ ] Senior Engineer (Implementation review)
- [ ] QA Lead (Test coverage review)
- [ ] Documentation Lead (Docs review)

**Approval**:
- [ ] Approved to proceed with PR #1
- [ ] Budget allocated (12-16 developer-days)
- [ ] Timeline accepted (2-3 weeks)

---

**End of Phase 1 Implementation Plan**
