package com.tjclp.xl.addressing

import com.tjclp.xl.Generators.given
import com.tjclp.xl.macros.ref
import munit.ScalaCheckSuite
import org.scalacheck.{Arbitrary, Gen}
import org.scalacheck.Prop.*

/**
 * GH-465: bounded navigation on `ARef` (`tryShift`/`tryDown`/`tryRight`/`clampShift`) and range
 * slicing on `CellRange` (`rows`/`columns`/`row`/`column`).
 *
 * `shift`/`down`/`up`/`left`/`right` stay total-but-unchecked (they can mint "A0" or column -1);
 * the bounded forms make the edge explicit: `None` past the grid, or a clamp onto it.
 */
class BoundedNavigationSpec extends ScalaCheckSuite:

  private val lastCell = ARef.from0(Column.MaxIndex0, Row.MaxIndex0) // XFD1048576

  private def inBounds(ref: ARef, dc: Int, dr: Int): Boolean =
    val c = ref.col.index0.toLong + dc
    val r = ref.row.index0.toLong + dr
    c >= 0 && c <= Column.MaxIndex0 && r >= 0 && r <= Row.MaxIndex0

  /** Offsets around the grid size, so both in-bounds and out-of-bounds results are frequent. */
  private val genColOffset: Gen[Int] = Gen.frequency(
    5 -> Gen.choose(-20000, 20000),
    1 -> Gen.oneOf(Int.MinValue, Int.MaxValue, 0, 1, -1)
  )
  private val genRowOffset: Gen[Int] = Gen.frequency(
    5 -> Gen.choose(-1200000, 1200000),
    1 -> Gen.oneOf(Int.MinValue, Int.MaxValue, 0, 1, -1)
  )

  /** Offsets that overrun BOTH axes on any start cell (|dc| > columns, |dr| > rows). */
  private val genOverrunOffsets: Gen[(Int, Int)] =
    for
      dc <- Gen.frequency(1 -> Gen.choose(16384, 100000), 1 -> Gen.choose(-100000, -16384))
      dr <- Gen.frequency(1 -> Gen.choose(1048576, 5000000), 1 -> Gen.choose(-5000000, -1048576))
    yield (dc, dr)

  // ========== tryShift / tryDown / tryRight ==========

  test("tryDown(1) at row 1048576 is None; tryDown(0) is the cell itself") {
    assertEquals(ARef.from1(1, 1048576).tryDown(1), None)
    assertEquals(lastCell.tryDown(1), None)
    assertEquals(ref"A1".tryDown(0), Some(ref"A1"))
    assertEquals(ref"A1".tryDown(1048575), Some(ARef.from1(1, 1048576)))
    assertEquals(ref"A1".tryDown(1048576), None)
  }

  test("tryShift(-1, 0) at column A is None; the other edges mirror") {
    assertEquals(ref"A1".tryShift(-1, 0), None)
    assertEquals(ref"A1".tryShift(0, -1), None) // would be "A0"
    assertEquals(ref"XFD1".tryRight(1), None)
    assertEquals(ref"XFD1".tryShift(1, 0), None)
    assertEquals(ref"A1".tryRight(16383), Some(ref"XFD1"))
    assertEquals(ref"A1".tryRight(16384), None)
    assertEquals(ref"B2".tryShift(-1, -1), Some(ref"A1"))
    assertEquals(lastCell.tryShift(1, 1), None)
    assertEquals(ref"C3".tryDown(-2), Some(ref"C1")) // negative steps walk back
    assertEquals(ref"C3".tryRight(-2), Some(ref"A3"))
  }

  test("tryShift never overflows on extreme offsets") {
    assertEquals(ref"A1".tryShift(Int.MaxValue, 0), None)
    assertEquals(ref"A1".tryShift(0, Int.MaxValue), None)
    assertEquals(ref"A1".tryShift(Int.MinValue, 0), None)
    assertEquals(lastCell.tryShift(Int.MinValue, Int.MinValue), None)
    assertEquals(lastCell.tryDown(Int.MaxValue), None)
    assertEquals(lastCell.tryRight(Int.MaxValue), None)
  }

  property("tryShift is Some exactly when the target is on the grid, and then agrees with shift") {
    forAll(Arbitrary.arbitrary[ARef], genColOffset, genRowOffset) { (ref, dc, dr) =>
      val shifted = ref.tryShift(dc, dr)
      assertEquals(shifted.isDefined, inBounds(ref, dc, dr), s"${ref.toA1} + ($dc, $dr)")
      shifted.foreach { target =>
        assertEquals(target, ref.shift(dc, dr))
        assertEquals(target.col.index0, ref.col.index0 + dc)
        assertEquals(target.row.index0, ref.row.index0 + dr)
      }
      true
    }
  }

  property("tryShift inverse law: tryShift(dc, dr).flatMap(_.tryShift(-dc, -dr)) == Some(ref)") {
    forAll(Arbitrary.arbitrary[ARef], genColOffset, genRowOffset) { (ref, dc, dr) =>
      val roundTrip = ref.tryShift(dc, dr).flatMap(_.tryShift(-dc, -dr))
      if inBounds(ref, dc, dr) then assertEquals(roundTrip, Some(ref), s"${ref.toA1} ($dc, $dr)")
      else assertEquals(roundTrip, None, s"${ref.toA1} ($dc, $dr)")
      true
    }
  }

  property("tryDown/tryRight are the single-axis tryShift") {
    forAll(Arbitrary.arbitrary[ARef], genColOffset, genRowOffset) { (ref, dc, dr) =>
      assertEquals(ref.tryDown(dr), ref.tryShift(0, dr))
      assertEquals(ref.tryRight(dc), ref.tryShift(dc, 0))
      true
    }
  }

  // ========== clampShift ==========

  test("clampShift pins to the grid edges") {
    assertEquals(ref"A1".clampShift(-1, -1), ref"A1")
    assertEquals(ref"A1".clampShift(-1, 0), ref"A1")
    assertEquals(ref"A1".clampShift(0, -1), ref"A1")
    assertEquals(lastCell.clampShift(1, 1), lastCell)
    assertEquals(ref"C3".clampShift(-10, 5), ref"A8") // column pinned, row shifted
    assertEquals(ref"C3".clampShift(2, -10), ref"E1") // row pinned, column shifted
    assertEquals(ref"A1".clampShift(Int.MaxValue, Int.MaxValue), lastCell)
    assertEquals(lastCell.clampShift(Int.MinValue, Int.MinValue), ref"A1")
  }

  property("clampShift always lands on the grid and agrees with shift/tryShift in bounds") {
    forAll(Arbitrary.arbitrary[ARef], genColOffset, genRowOffset) { (ref, dc, dr) =>
      val clamped = ref.clampShift(dc, dr)
      assert(clamped.col.index0 >= 0 && clamped.col.index0 <= Column.MaxIndex0, clamped.toA1)
      assert(clamped.row.index0 >= 0 && clamped.row.index0 <= Row.MaxIndex0, clamped.toA1)
      ref.tryShift(dc, dr).foreach(target => assertEquals(clamped, target))
      // each axis is either the exact target or the nearest edge, never something in between
      val wantCol = ref.col.index0.toLong + dc
      val wantRow = ref.row.index0.toLong + dr
      assertEquals(clamped.col.index0.toLong, math.max(0L, math.min(wantCol, Column.MaxIndex0)))
      assertEquals(clamped.row.index0.toLong, math.max(0L, math.min(wantRow, Row.MaxIndex0)))
      true
    }
  }

  property("clampShift is idempotent once both axes are pinned") {
    forAll(Arbitrary.arbitrary[ARef], genOverrunOffsets) { (ref, offsets) =>
      val (dc, dr) = offsets
      val once = ref.clampShift(dc, dr)
      assertEquals(once.clampShift(dc, dr), once, s"${ref.toA1} ($dc, $dr)")
      assertEquals(ref.tryShift(dc, dr), None)
      true
    }
  }

  // ========== CellRange rows / columns / row / column ==========

  test("rows slices top to bottom, one row high, spanning the range's columns") {
    assertEquals(ref"A1:B3".rows.toList, List(ref"A1:B1", ref"A2:B2", ref"A3:B3"))
    assertEquals(ref"A1:B3".rows.size, 3)
    assertEquals(ref"C5:C5".rows.toList, List(ref"C5:C5"))
  }

  test("columns slices left to right, one column wide, spanning the range's rows") {
    assertEquals(ref"A1:B3".columns.toList, List(ref"A1:A3", ref"B1:B3"))
    assertEquals(ref"A1:B3".columns.size, 2)
  }

  test("row(i) / column(i) are 0-based within the range and None outside it") {
    val range = ref"A1:B3"
    assertEquals(range.row(0), Some(ref"A1:B1"))
    assertEquals(range.row(2), Some(ref"A3:B3"))
    assertEquals(range.row(3), None)
    assertEquals(range.row(-1), None)
    assertEquals(range.column(0), Some(ref"A1:A3"))
    assertEquals(range.column(1), Some(ref"B1:B3"))
    assertEquals(range.column(2), None)
    assertEquals(range.column(-1), None)
    assertEquals(range.row(Int.MaxValue), None)
    assertEquals(range.column(Int.MinValue), None)
  }

  test("slices keep the range's anchors") {
    val anchored = CellRange.parse("$A$1:C3").toOption.getOrElse(fail("parse"))
    assertEquals(anchored.row(0).map(_.toA1Anchored), Some("$A$1:C1"))
    assertEquals(anchored.column(2).map(_.toA1Anchored), Some("$C$1:C3"))
    assertEquals(anchored.rows.toList.map(_.toA1Anchored), List("$A$1:C1", "$A$2:C2", "$A$3:C3"))
  }

  test("CellRange.empty has no rows or columns") {
    assertEquals(CellRange.empty.rows.toList, Nil)
    assertEquals(CellRange.empty.columns.toList, Nil)
    assertEquals(CellRange.empty.row(0), None)
    assertEquals(CellRange.empty.column(0), None)
  }

  test("full-column and full-row ranges slice lazily") {
    val fullColumn = CellRange.parse("A:A").toOption.getOrElse(fail("parse"))
    assertEquals(fullColumn.rows.take(2).toList, List(ref"A1:A1", ref"A2:A2"))
    assertEquals(fullColumn.columns.size, 1)
    val lastA = ARef.from0(0, Row.MaxIndex0) // A1048576
    assertEquals(fullColumn.row(Row.MaxIndex0), Some(CellRange(lastA, lastA)))
    assertEquals(fullColumn.row(Row.MaxIndex0 + 1), None)
    val fullRow = CellRange.parse("1:1").toOption.getOrElse(fail("parse"))
    assertEquals(fullRow.columns.take(2).toList, List(ref"A1:A1", ref"B1:B1"))
    assertEquals(fullRow.rows.size, 1)
    assertEquals(fullRow.column(Column.MaxIndex0), Some(CellRange(ref"XFD1", ref"XFD1")))
  }

  extension (range: CellRange) private def rowMajorCells: List[ARef] = range.cells.toList

  property("rows.size == height, columns.size == width, and the slices tile the range") {
    forAll { (range: CellRange) =>
      val rows = range.rows.toList
      val cols = range.columns.toList
      assertEquals(rows.size, range.height)
      assertEquals(cols.size, range.width)
      assert(
        rows.forall(r => r.height == 1 && r.colStart == range.colStart && r.colEnd == range.colEnd)
      )
      assert(
        cols.forall(c => c.width == 1 && c.rowStart == range.rowStart && c.rowEnd == range.rowEnd)
      )
      assertEquals(
        rows.flatMap(_.rowMajorCells),
        range.rowMajorCells
      ) // row slices in row-major order
      assertEquals(cols.flatMap(_.rowMajorCells).toSet, range.cells.toSet)
      true
    }
  }

  property("row(i) == rows.lift(i) and column(i) == columns.lift(i)") {
    forAll(Arbitrary.arbitrary[CellRange], Gen.choose(-2, 120)) { (range, i) =>
      assertEquals(range.row(i), range.rows.toVector.lift(i))
      assertEquals(range.column(i), range.columns.toVector.lift(i))
      true
    }
  }
