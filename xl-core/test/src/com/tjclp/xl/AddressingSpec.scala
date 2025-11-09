package com.tjclp.xl

import munit.ScalaCheckSuite
import org.scalacheck.Prop.*
import Generators.given

/** Property tests for addressing types */
class AddressingSpec extends ScalaCheckSuite:

  property("Column.from0 and index0 are inverses") {
    forAll { (n: Int) =>
      if n >= 0 && n <= 16383 then
        val col = Column.from0(n)
        assertEquals(col.index0, n)
        true
      else
        true
    }
  }

  property("Column.from1 and index1 are inverses") {
    forAll { (n: Int) =>
      if n >= 1 && n <= 16384 then
        val col = Column.from1(n)
        assertEquals(col.index1, n)
        true
      else
        true
    }
  }

  property("Column.toLetter and fromLetter are inverses") {
    forAll { (col: Column) =>
      val letter = col.toLetter
      val parsed = Column.fromLetter(letter)
      parsed match
        case Right(c) => assertEquals(c, col)
        case Left(err) => fail(s"Failed to parse $letter: $err")
      true
    }
  }

  property("Row.from0 and index0 are inverses") {
    forAll { (n: Int) =>
      if n >= 0 && n <= 1048575 then
        val row = Row.from0(n)
        assertEquals(row.index0, n)
        true
      else
        true
    }
  }

  property("Row.from1 and index1 are inverses") {
    forAll { (n: Int) =>
      if n >= 1 && n <= 1048576 then
        val row = Row.from1(n)
        assertEquals(row.index1, n)
        true
      else
        true
    }
  }

  property("ARef packs and unpacks column and row correctly") {
    forAll { (col: Column, row: Row) =>
      val ref = ARef(col, row)
      assertEquals(ref.col, col)
      assertEquals(ref.row, row)
    }
  }

  property("ARef.toA1 and parse are inverses") {
    forAll { (ref: ARef) =>
      val a1 = ref.toA1
      val parsed = ARef.parse(a1)
      parsed match
        case Right(r) => assertEquals(r, ref)
        case Left(err) => fail(s"Failed to parse $a1: $err")
      true
    }
  }

  property("ARef.shift moves cell correctly") {
    forAll { (ref: ARef) =>
      val shifted = ref.shift(1, 2)
      assertEquals(shifted.col.index0, ref.col.index0 + 1)
      assertEquals(shifted.row.index0, ref.row.index0 + 2)
    }
  }

  property("CellRange.contains works correctly") {
    forAll { (range: CellRange, ref: ARef) =>
      val contains = range.contains(ref)
      val inCols = ref.col.index0 >= range.colStart.index0 && ref.col.index0 <= range.colEnd.index0
      val inRows = ref.row.index0 >= range.rowStart.index0 && ref.row.index0 <= range.rowEnd.index0
      assertEquals(contains, inCols && inRows)
    }
  }

  property("CellRange.size equals width * height") {
    forAll { (range: CellRange) =>
      assertEquals(range.size, range.width * range.height)
    }
  }

  property("CellRange.toA1 and parse are inverses") {
    forAll { (range: CellRange) =>
      val a1 = range.toA1
      val parsed = CellRange.parse(a1)
      parsed match
        case Right(r) =>
          assertEquals(r.start, range.start)
          assertEquals(r.end, range.end)
        case Left(err) => fail(s"Failed to parse $a1: $err")
      true
    }
  }

  property("CellRange.expand includes the reference") {
    forAll { (range: CellRange, ref: ARef) =>
      val expanded = range.expand(ref)
      assert(expanded.contains(ref))
      assert(expanded.contains(range.start))
      assert(expanded.contains(range.end))
    }
  }

  property("CellRange.intersects is symmetric") {
    forAll { (r1: CellRange, r2: CellRange) =>
      assertEquals(r1.intersects(r2), r2.intersects(r1))
    }
  }

  test("Column.fromLetter parses common examples") {
    assertEquals(Column.fromLetter("A"), Right(Column.from0(0)))
    assertEquals(Column.fromLetter("B"), Right(Column.from0(1)))
    assertEquals(Column.fromLetter("Z"), Right(Column.from0(25)))
    assertEquals(Column.fromLetter("AA"), Right(Column.from0(26)))
    assertEquals(Column.fromLetter("AB"), Right(Column.from0(27)))
    assert(Column.fromLetter("").isLeft)
    assert(Column.fromLetter("A1").isLeft)
  }

  test("ARef.parse handles common examples") {
    assertEquals(ARef.parse("A1"), Right(ARef.from1(1, 1)))
    assertEquals(ARef.parse("B2"), Right(ARef.from1(2, 2)))
    assertEquals(ARef.parse("Z99"), Right(ARef.from1(26, 99)))
    assertEquals(ARef.parse("AA100"), Right(ARef.from1(27, 100)))
    assert(ARef.parse("").isLeft)
    assert(ARef.parse("A").isLeft)
    assert(ARef.parse("1").isLeft)
    assert(ARef.parse("A0").isLeft)
  }

  test("CellRange.parse handles common examples") {
    assertEquals(
      CellRange.parse("A1:B2"),
      Right(CellRange(ARef.from1(1, 1), ARef.from1(2, 2)))
    )
    assertEquals(
      CellRange.parse("A1"),
      Right(CellRange(ARef.from1(1, 1), ARef.from1(1, 1)))
    )
    assert(CellRange.parse("").isLeft)
    assert(CellRange.parse("A1:").isLeft)
    assert(CellRange.parse(":B2").isLeft)
  }

  test("SheetName validation works") {
    assert(SheetName("Sheet1").isRight)
    assert(SheetName("My Sheet").isRight)
    assert(SheetName("").isLeft)
    assert(SheetName("A" * 32).isLeft) // Too long
    assert(SheetName("Sheet:1").isLeft) // Invalid char
    assert(SheetName("Sheet/1").isLeft) // Invalid char
    assert(SheetName("Sheet?1").isLeft) // Invalid char
  }
