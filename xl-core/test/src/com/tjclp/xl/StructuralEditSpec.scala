package com.tjclp.xl

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.CellRange
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.sheets.FreezePane
import munit.ScalaCheckSuite
import org.scalacheck.Gen
import org.scalacheck.Prop.forAll

/**
 * GH-128/#129: pure structural row/column insert/delete on Sheet (cells, merges, props, freeze).
 */
class StructuralEditSpec extends ScalaCheckSuite:

  private def num(n: Int) = CellValue.Number(BigDecimal(n))

  test("insertRows shifts cells at/below down, leaves rows above unchanged") {
    val s = Sheet("S").put("A1" -> 1, "A3" -> 3) // rows 0 and 2
    val r = s.insertRows(at = 1, count = 1)
    assertEquals(r(ref"A1").value, num(1)) // above the cut: unchanged
    assertEquals(r(ref"A4").value, num(3)) // A3 -> A4
    assert(!r.contains(ref"A3"))
  }

  test("deleteRows removes the deleted row and shifts rows below up") {
    val s = Sheet("S").put("A1" -> 1, "A2" -> 2, "A3" -> 3)
    val r = s.deleteRows(at = 1, count = 1) // delete row index 1 (A2)
    assertEquals(r(ref"A1").value, num(1))
    assertEquals(r(ref"A2").value, num(3)) // A3 -> A2
    assert(!r.contains(ref"A3"))
  }

  test("insert-then-delete is identity on cells") {
    val s = Sheet("S").put("A1" -> 1, "B2" -> 2, "C3" -> 3)
    assertEquals(s.insertRows(1, 2).deleteRows(1, 2).cells, s.cells)
    assertEquals(s.insertColumns(1, 2).deleteColumns(1, 2).cells, s.cells)
  }

  test("insertColumns / deleteColumns mirror the row behavior") {
    val s = Sheet("S").put("A1" -> 1, "C1" -> 3) // cols 0 and 2
    val ins = s.insertColumns(at = 1, count = 1)
    assertEquals(ins(ref"A1").value, num(1))
    assertEquals(ins(ref"D1").value, num(3)) // C1 -> D1
    val del = s.deleteColumns(at = 0, count = 1) // delete column A
    assert(!del.contains(ref"A1"))
    assertEquals(del(ref"B1").value, num(3)) // C1 -> B1
  }

  test("deleteRows clamps a merged range spanning the cut") {
    val s = Sheet("S").copy(mergedRanges = Set(CellRange(ref"A1", ref"A4"))) // rows 0..3
    val r = s.deleteRows(at = 1, count = 2) // remove rows 1,2 -> A1:A2
    assertEquals(r.mergedRanges, Set(CellRange(ref"A1", ref"A2")))
  }

  test("deleteRows drops a merge fully inside the deletion") {
    val s = Sheet("S").copy(mergedRanges = Set(CellRange(ref"A2", ref"A3"))) // rows 1..2
    assertEquals(s.deleteRows(at = 1, count = 2).mergedRanges, Set.empty)
  }

  test("insertRows shifts the freeze-pane anchor") {
    val s = Sheet("S").copy(freezePane = Some(FreezePane.At(ref"B3")))
    assertEquals(s.insertRows(at = 0, count = 2).freezePane, Some(FreezePane.At(ref"B5")))
  }

  // ===== GH-428: insert-side clamp at the sheet edge =====

  test("GH-428: insertRows clamps a full-height merge at row 1048576") {
    val full = CellRange(ref"A1", ref"A1048576")
    val s = Sheet("S").copy(mergedRanges = Set(full))
    assertEquals(s.insertRows(at = 1, count = 2).mergedRanges, Set(full))
  }

  test("GH-428: insertColumns clamps a full-width merge at XFD") {
    val full = CellRange(ref"A1", ref"XFD1")
    val s = Sheet("S").copy(mergedRanges = Set(full))
    assertEquals(s.insertColumns(at = 1, count = 2).mergedRanges, Set(full))
  }

  test("GH-428: a merge near the edge clamps its end; one pushed fully past the edge drops") {
    val nearEdge = CellRange(ref"A1048570", ref"A1048576")
    val s = Sheet("S").copy(mergedRanges = Set(nearEdge))
    assertEquals(
      s.insertRows(at = 0, count = 3).mergedRanges,
      Set(CellRange(ref"A1048573", ref"A1048576"))
    )
    val last = CellRange(ref"A1048576", ref"A1048576")
    assertEquals(
      Sheet("S").copy(mergedRanges = Set(last)).insertRows(at = 0, count = 1).mergedRanges,
      Set.empty[CellRange]
    )
  }

  property("GH-428: shiftSpan results always stay within [0, axisMax] with start <= end") {
    val max = com.tjclp.xl.addressing.Row.MaxIndex0
    val genSpan = for
      s <- Gen.chooseNum(0, max)
      e <- Gen.chooseNum(s, max)
    yield (s, e)
    forAll(genSpan, Gen.chooseNum(0, max), Gen.chooseNum(1, 5000), Gen.oneOf(true, false)) {
      case ((s, e), at, count, deleting) =>
        Sheet.shiftSpan(s, e, at, count, deleting, max) match
          case None => true
          case Some((ns, ne)) => ns >= 0 && ns <= ne && ne <= max
    }
  }

  property("GH-428: PageSetup construction never throws for shifted repeat rows") {
    val max = com.tjclp.xl.addressing.Row.MaxIndex0
    val genSpan = for
      s <- Gen.chooseNum(0, max)
      e <- Gen.chooseNum(s, max)
    yield (s, e)
    forAll(genSpan, Gen.chooseNum(0, max), Gen.chooseNum(1, 5000), Gen.oneOf(true, false)) {
      case ((s, e), at, count, deleting) =>
        Sheet.shiftSpan(s, e, at, count, deleting, max).forall { case (ns, ne) =>
          // the clamp makes PageSetup's 1-based repeatRows invariant unreachable
          val ps = com.tjclp.xl.sheets.PageSetup(repeatRows = Some((ns + 1, ne + 1)))
          ps.repeatRows.contains((ns + 1, ne + 1))
        }
    }
  }
