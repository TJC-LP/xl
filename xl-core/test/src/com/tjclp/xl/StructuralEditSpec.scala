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

  // ===== GH-429: print areas, tables (autoFilter cases live below with the model) =====

  test("GH-429: insertRows extends the print area and shifts repeat rows") {
    val ps = com.tjclp.xl.sheets
      .PageSetup(printArea = Some(ref"A1:D10"), repeatRows = Some((3, 4)))
    val s = Sheet("S").copy(pageSetup = Some(ps))
    val r = s.insertRows(at = 1, count = 2) // inside the area, above the repeat span
    assertEquals(r.pageSetup.flatMap(_.printArea).map(_.toA1), Some("A1:D12"))
    assertEquals(r.pageSetup.flatMap(_.repeatRows), Some((5, 6)))
  }

  test("GH-429: deleteRows collapsing the print area / repeat rows clears them") {
    val ps = com.tjclp.xl.sheets
      .PageSetup(printArea = Some(ref"A3:D4"), repeatRows = Some((3, 4)))
    val s = Sheet("S").copy(pageSetup = Some(ps))
    val r = s.deleteRows(at = 2, count = 2) // rows 3-4 (0-based 2..3) vanish entirely
    assertEquals(r.pageSetup.flatMap(_.printArea), None)
    assertEquals(r.pageSetup.flatMap(_.repeatRows), None)
    // column edits leave repeatRows alone
    val c = s.insertColumns(at = 0, count = 2)
    assertEquals(c.pageSetup.flatMap(_.repeatRows), Some((3, 4)))
    assertEquals(c.pageSetup.flatMap(_.printArea).map(_.toA1), Some("C3:F4"))
  }

  test("GH-429: repeat rows clamp at the sheet edge on insert (PageSetup never throws)") {
    val max1 = com.tjclp.xl.addressing.Row.MaxIndex0 + 1 // 1-based last row
    val ps = com.tjclp.xl.sheets.PageSetup(repeatRows = Some((max1 - 1, max1)))
    val s = Sheet("S").copy(pageSetup = Some(ps))
    assertEquals(
      s.insertRows(at = 0, count = 1).pageSetup.flatMap(_.repeatRows),
      Some((max1, max1))
    )
    assertEquals(s.insertRows(at = 0, count = 5).pageSetup.flatMap(_.repeatRows), None)
  }

  test("GH-429: tables shift on row inserts and drop when their range collapses") {
    val table = com.tjclp.xl.tables.TableSpec
      .unsafeFromColumnNames("T1", "T1", ref"A2:C6", Vector("A", "B", "C"))
    val s = Sheet("S").withTable(table)
    val r = s.insertRows(at = 0, count = 2)
    assertEquals(r.tables.get("T1").map(_.range.toA1), Some("A4:C8"))
    assertEquals(r.tables.get("T1").map(_.columns), Some(table.columns))
    assert(r.tables.get("T1").exists(_.isValid))
    // deleting every table row drops the table entirely
    assertEquals(s.deleteRows(at = 1, count = 5).tables.get("T1"), None)
  }

  test("GH-429: a table shrunk below header+data minimum height is dropped") {
    val table = com.tjclp.xl.tables.TableSpec
      .unsafeFromColumnNames("T1", "T1", ref"A1:B3", Vector("A", "B"))
    val s = Sheet("S").withTable(table)
    // delete both data rows -> only the header would survive -> Excel would demand repair
    assertEquals(s.deleteRows(at = 1, count = 2).tables.get("T1"), None)
  }

  test("GH-429: an interior column insert splices fresh table columns and stamps header cells") {
    val table = com.tjclp.xl.tables.TableSpec
      .unsafeFromColumnNames("T1", "T1", ref"B1:D4", Vector("Alpha", "Column1", "Gamma"))
    val s = Sheet("S").withTable(table)
    val r = s.insertColumns(at = 2, count = 2) // inside B..D (0-based cols 1..3), at C
    val t = r.tables.get("T1").getOrElse(fail("table dropped"))
    assertEquals(t.range.toA1, "B1:F4")
    assert(t.isValid, s"column count must match range width: $t")
    // fresh names are case-insensitively unique against existing ones ("Column1" is taken)
    assertEquals(t.columns.map(_.name), Vector("Alpha", "Column2", "Column3", "Column1", "Gamma"))
    assertEquals(t.columns.map(_.id), Vector(1L, 4L, 5L, 2L, 3L))
    // Excel writes the synthesized names into the header row cells
    assertEquals(r(ref"C1").value, CellValue.Text("Column2"))
    assertEquals(r(ref"D1").value, CellValue.Text("Column3"))
  }

  test("GH-429: a column delete drops the deleted table columns; left inserts translate") {
    val table = com.tjclp.xl.tables.TableSpec
      .unsafeFromColumnNames("T1", "T1", ref"B1:D4", Vector("A", "B", "C"))
    val s = Sheet("S").withTable(table)
    val del = s.deleteColumns(at = 2, count = 1) // drop table column "B" (sheet col C)
    val t = del.tables.get("T1").getOrElse(fail("table dropped"))
    assertEquals(t.range.toA1, "B1:C4")
    assertEquals(t.columns.map(_.name), Vector("A", "C"))
    assert(t.isValid)
    val left = s.insertColumns(at = 0, count = 2) // pure translation
    assertEquals(left.tables.get("T1").map(_.range.toA1), Some("D1:F4"))
    assertEquals(left.tables.get("T1").map(_.columns), Some(table.columns))
  }

  test("GH-429: autoFilter Ranged shifts; a collapsed range becomes Remove; None stays None") {
    import com.tjclp.xl.sheets.AutoFilterState
    val s = Sheet("S").copy(autoFilter = Some(AutoFilterState.Ranged(ref"A3:D9")))
    assertEquals(
      s.insertRows(at = 0, count = 2).autoFilter,
      Some(AutoFilterState.Ranged(ref"A5:D11": CellRange))
    )
    assertEquals(
      s.deleteRows(at = 2, count = 7).autoFilter, // rows 3..9 vanish entirely
      Some(AutoFilterState.Remove)
    )
    assertEquals(Sheet("S").insertRows(at = 0, count = 1).autoFilter, None)
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
