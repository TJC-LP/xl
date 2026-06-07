package com.tjclp.xl

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.CellRange
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.sheets.FreezePane
import munit.FunSuite

/** GH-128/#129: pure structural row/column insert/delete on Sheet (cells, merges, props, freeze). */
class StructuralEditSpec extends FunSuite:

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
