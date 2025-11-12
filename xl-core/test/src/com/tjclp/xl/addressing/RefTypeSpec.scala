package com.tjclp.xl.addressing

import com.tjclp.xl.Generators.{genARef, genCellRange, genSheetName}
import munit.ScalaCheckSuite
import org.scalacheck.Prop.*
import org.scalacheck.{Arbitrary, Gen}

/** Property tests for RefType parsing and behavior */
class RefTypeSpec extends ScalaCheckSuite:

  // ========== Generators ==========

  given Arbitrary[RefType.Cell] = Arbitrary(genARef.map(RefType.Cell.apply))
  given Arbitrary[RefType.Range] = Arbitrary(genCellRange.map(RefType.Range.apply))
  given Arbitrary[RefType.QualifiedCell] = Arbitrary {
    for
      sheet <- genSheetName
      ref <- genARef
    yield RefType.QualifiedCell(sheet, ref)
  }
  given Arbitrary[RefType.QualifiedRange] = Arbitrary {
    for
      sheet <- genSheetName
      range <- genCellRange
    yield RefType.QualifiedRange(sheet, range)
  }

  given Arbitrary[RefType] = Arbitrary {
    Gen.oneOf(
      genARef.map(RefType.Cell.apply),
      genCellRange.map(RefType.Range.apply),
      summon[Arbitrary[RefType.QualifiedCell]].arbitrary,
      summon[Arbitrary[RefType.QualifiedRange]].arbitrary
    )
  }

  // ========== Round-Trip Laws ==========

  property("Round-trip: RefType.Cell -> toA1 -> parse") {
    forAll { (cell: RefType.Cell) =>
      val a1 = cell.toA1
      RefType.parse(a1) match
        case Right(parsed) => assertEquals(parsed, cell)
        case Left(err) => fail(s"Failed to parse $a1: $err")
      true
    }
  }

  property("Round-trip: RefType.Range -> toA1 -> parse") {
    forAll { (range: RefType.Range) =>
      val a1 = range.toA1
      RefType.parse(a1) match
        case Right(parsed) => assertEquals(parsed, range)
        case Left(err) => fail(s"Failed to parse $a1: $err")
      true
    }
  }

  property("Round-trip: RefType.QualifiedCell -> toA1 -> parse") {
    forAll { (qcell: RefType.QualifiedCell) =>
      val a1 = qcell.toA1
      RefType.parse(a1) match
        case Right(parsed) => assertEquals(parsed, qcell)
        case Left(err) => fail(s"Failed to parse $a1: $err")
      true
    }
  }

  property("Round-trip: RefType.QualifiedRange -> toA1 -> parse") {
    forAll { (qrange: RefType.QualifiedRange) =>
      val a1 = qrange.toA1
      RefType.parse(a1) match
        case Right(parsed) => assertEquals(parsed, qrange)
        case Left(err) => fail(s"Failed to parse $a1: $err")
      true
    }
  }

  // ========== Parsing Tests ==========

  test("Parse simple cell: A1") {
    RefType.parse("A1") match
      case Right(RefType.Cell(ref)) => assertEquals(ref.toA1, "A1")
      case other => fail(s"Expected RefType.Cell, got $other")
  }

  test("Parse simple range: A1:B10") {
    RefType.parse("A1:B10") match
      case Right(RefType.Range(range)) => assertEquals(range.toA1, "A1:B10")
      case other => fail(s"Expected RefType.Range, got $other")
  }

  test("Parse qualified cell: Sales!A1") {
    RefType.parse("Sales!A1") match
      case Right(RefType.QualifiedCell(sheet, ref)) =>
        assertEquals(sheet.value, "Sales")
        assertEquals(ref.toA1, "A1")
      case other => fail(s"Expected RefType.QualifiedCell, got $other")
  }

  test("Parse qualified range: Sales!A1:B10") {
    RefType.parse("Sales!A1:B10") match
      case Right(RefType.QualifiedRange(sheet, range)) =>
        assertEquals(sheet.value, "Sales")
        assertEquals(range.toA1, "A1:B10")
      case other => fail(s"Expected RefType.QualifiedRange, got $other")
  }

  test("Parse quoted sheet: 'Q1 Sales'!A1") {
    RefType.parse("'Q1 Sales'!A1") match
      case Right(RefType.QualifiedCell(sheet, ref)) =>
        assertEquals(sheet.value, "Q1 Sales")
        assertEquals(ref.toA1, "A1")
      case other => fail(s"Expected RefType.QualifiedCell with quoted sheet, got $other")
  }

  test("Parse quoted sheet with range: 'Q1 Sales'!A1:B10") {
    RefType.parse("'Q1 Sales'!A1:B10") match
      case Right(RefType.QualifiedRange(sheet, range)) =>
        assertEquals(sheet.value, "Q1 Sales")
        assertEquals(range.toA1, "A1:B10")
      case other => fail(s"Expected RefType.QualifiedRange with quoted sheet, got $other")
  }

  test("Parse escaped quote: 'It''s Q1'!A1") {
    RefType.parse("'It''s Q1'!A1") match
      case Right(RefType.QualifiedCell(sheet, ref)) =>
        assertEquals(sheet.value, "It's Q1")
        assertEquals(ref.toA1, "A1")
      case other => fail(s"Expected RefType.QualifiedCell with escaped quote, got $other")
  }

  test("Parse multiple escaped quotes: 'It''s ''Q1'''!A1") {
    RefType.parse("'It''s ''Q1'''!A1") match
      case Right(RefType.QualifiedCell(sheet, ref)) =>
        assertEquals(sheet.value, "It's 'Q1'")
        assertEquals(ref.toA1, "A1")
      case other => fail(s"Expected RefType.QualifiedCell with multiple escaped quotes, got $other")
  }

  // ========== Edge Cases ==========

  test("Empty string fails") {
    assert(RefType.parse("").isLeft, "Empty string should fail")
  }

  test("Missing ref after bang fails") {
    assert(RefType.parse("Sales!").isLeft, "Missing ref after ! should fail")
  }

  test("Empty quoted sheet name fails") {
    assert(RefType.parse("''!A1").isLeft, "Empty quoted sheet name should fail")
  }

  test("Invalid sheet name chars fail") {
    assert(RefType.parse("Sales:Q1!A1").isLeft, "Sheet name with : should fail")
    assert(RefType.parse("Sales\\Q1!A1").isLeft, "Sheet name with \\ should fail")
    assert(RefType.parse("Sales/Q1!A1").isLeft, "Sheet name with / should fail")
  }

  test("Sheet name max length (31 chars)") {
    val name31 = "a" * 31
    val name32 = "a" * 32

    assert(RefType.parse(s"$name31!A1").isRight, "31-char sheet name should succeed")
    assert(RefType.parse(s"$name32!A1").isLeft, "32-char sheet name should fail")
  }

  // ========== toA1 Quoting Behavior ==========

  test("toA1 quotes sheet names with spaces") {
    val ref = RefType.QualifiedCell(SheetName.unsafe("Q1 Sales"), ARef.from1(1, 1))
    assertEquals(ref.toA1, "'Q1 Sales'!A1")
  }

  test("toA1 doesn't quote simple sheet names") {
    val ref = RefType.QualifiedCell(SheetName.unsafe("Sales"), ARef.from1(1, 1))
    assertEquals(ref.toA1, "Sales!A1")
  }

  test("toA1 quotes sheet names with single quotes") {
    val ref = RefType.QualifiedCell(SheetName.unsafe("It's Q1"), ARef.from1(1, 1))
    // Should produce escaped quotes
    assert(ref.toA1.contains("'"), "Should quote sheet name with apostrophe")
  }

  // ========== Pattern Matching ==========

  test("Pattern match on RefType variants") {
    val cell = RefType.Cell(ARef.from1(1, 1))
    val range = RefType.Range(CellRange.parse("A1:B10").toOption.get)
    val qcell = RefType.QualifiedCell(SheetName.unsafe("Sales"), ARef.from1(1, 1))
    val qrange =
      RefType.QualifiedRange(SheetName.unsafe("Sales"), CellRange.parse("A1:B10").toOption.get)

    cell match
      case RefType.Cell(_) => assert(true)
      case _ => fail("Should match Cell")

    range match
      case RefType.Range(_) => assert(true)
      case _ => fail("Should match Range")

    qcell match
      case RefType.QualifiedCell(_, _) => assert(true)
      case _ => fail("Should match QualifiedCell")

    qrange match
      case RefType.QualifiedRange(_, _) => assert(true)
      case _ => fail("Should match QualifiedRange")
  }

  // ========== Integration with Existing Parsers ==========

  test("RefType.parse delegates to ARef.parse for simple cells") {
    val expected = ARef.parse("XFD1048576").toOption.get
    RefType.parse("XFD1048576") match
      case Right(RefType.Cell(ref)) => assertEquals(ref, expected)
      case other => fail(s"Expected Cell with correct ref, got $other")
  }

  test("RefType.parse delegates to CellRange.parse for ranges") {
    val expected = CellRange.parse("A1:XFD1048576").toOption.get
    RefType.parse("A1:XFD1048576") match
      case Right(RefType.Range(range)) => assertEquals(range, expected)
      case other => fail(s"Expected Range with correct range, got $other")
  }

end RefTypeSpec
