package com.tjclp.xl.macros

import com.tjclp.xl.*
import com.tjclp.xl.addressing.{ARef, CellRange, RefType, SheetName}
import com.tjclp.xl.error.XLError
import munit.{FunSuite, ScalaCheckSuite}
import org.scalacheck.Prop.*

class RefInterpolationSpec extends ScalaCheckSuite:

  // ===== Backward Compatibility (Compile-Time Literals) =====

  test("Compile-time literal: ref\"A1\" returns ARef directly") {
    val r = ref"A1"
    r match
      case aref: ARef =>
        assertEquals(aref.toA1, "A1")
      case other => fail(s"Expected ARef, got $other")
  }

  test("Compile-time literal: ref\"A1:B10\" returns CellRange directly") {
    val r = ref"A1:B10"
    r match
      case range: CellRange =>
        assertEquals(range.toA1, "A1:B10")
      case other => fail(s"Expected CellRange, got $other")
  }

  test("Compile-time literal: ref\"Sales!A1\" returns RefType directly") {
    val r = ref"Sales!A1"
    r match
      case qualified: RefType =>
        assertEquals(qualified.toA1, "Sales!A1")
      case other => fail(s"Expected RefType, got $other")
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

  test("Runtime interpolation: quoted sheet name") {
    val quotedStr = "'Q1 Sales'!A1"
    val result = ref"$quotedStr"

    result match
      case Right(RefType.QualifiedCell(sheet, _)) =>
        assertEquals(sheet.value, "Q1 Sales")
      case other =>
        fail(s"Expected Right with quoted sheet, got $other")
  }

  test("Runtime interpolation: escaped quotes in sheet name") {
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

  property("Invalid ref strings always return Left") {
    val genInvalid = org.scalacheck.Gen.oneOf("", "INVALID", "@#$", "A", "1", "!")
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
        fail(s"Should parse sheet with spaces: $err")
      case other => fail(s"Unexpected result: $other")
  }

  test("Edge: Sheet name with special chars") {
    val withSpecial = "'Sheet (2025)'!A1"
    val result = ref"$withSpecial"

    result match
      case Right(RefType.QualifiedCell(sheet, _)) =>
        assertEquals(sheet.value, "Sheet (2025)")
      case Left(err) =>
        fail(s"Should parse sheet with parens: $err")
      case other => fail(s"Unexpected result: $other")
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
