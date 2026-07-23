package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import com.tjclp.xl.formula.eval.StructuralEditor
import munit.FunSuite

/**
 * GH-128 / GH-129: structural editing WITH formula rewriting (the xl-evaluator layer over the pure
 * xl-core cell shift). Covers conditional shifting, #REF! on deletion, range shrink, the
 * insert↔delete identity law, and cross-sheet reference rewriting.
 */
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial"))
class StructuralFormulaSpec extends FunSuite:

  private val S = SheetName.unsafe("S")

  private def sheetNamed(wb: Workbook, n: String): Sheet =
    wb.sheets.find(_.name == SheetName.unsafe(n)).get

  private def formulaCell(s: String): CellValue = CellValue.Formula(s, None)

  test("insert rows shifts local refs at/after the insertion point") {
    val s = new Sheet(name = S)
      .put(ref"B1", formulaCell("=A5"))
      .put(ref"B2", formulaCell("=A1"))
    val r = StructuralEditor.insertRows(Workbook(Vector(s)), S, at = 2, count = 1)
    val s2 = sheetNamed(r, "S")
    assertEquals(s2(ref"B1").value, formulaCell("A6")) // A5 (row 4 >= 2) -> A6
    assertEquals(s2(ref"B2").value, formulaCell("A1")) // A1 (row 0 < 2) unchanged
  }

  test("insert rows moves the formula cell itself and still rewrites refs") {
    val s = new Sheet(name = S).put(ref"B5", formulaCell("=A1"))
    val r = StructuralEditor.insertRows(Workbook(Vector(s)), S, at = 2, count = 1)
    val s2 = sheetNamed(r, "S")
    assertEquals(s2(ref"B6").value, formulaCell("A1")) // cell B5 (row 4) -> B6; ref A1 unchanged
    assert(!s2.contains(ref"B5"))
  }

  test("delete rows: ref into the deleted band becomes #REF!") {
    val s = new Sheet(name = S).put(ref"B1", formulaCell("=A3"))
    val r = StructuralEditor.deleteRows(Workbook(Vector(s)), S, at = 2, count = 1)
    // A3 = row index 2 = the deleted row
    assertEquals(sheetNamed(r, "S")(ref"B1").value, CellValue.Error(CellError.Ref))
  }

  test("delete rows: refs after the band shift up") {
    val s = new Sheet(name = S).put(ref"B1", formulaCell("=A5"))
    val r = StructuralEditor.deleteRows(Workbook(Vector(s)), S, at = 2, count = 1)
    assertEquals(sheetNamed(r, "S")(ref"B1").value, formulaCell("A4")) // A5 (row 4) -> A4
  }

  test("delete rows: a range straddling the deletion shrinks") {
    val s = new Sheet(name = S).put(ref"B1", formulaCell("=SUM(A1:A10)"))
    val r = StructuralEditor.deleteRows(Workbook(Vector(s)), S, at = 2, count = 1)
    assertEquals(sheetNamed(r, "S")(ref"B1").value, formulaCell("SUM(A1:A9)"))
  }

  test("insert then delete the same band is identity on formula refs") {
    val s = new Sheet(name = S).put(ref"B1", formulaCell("=A5"))
    val inserted = StructuralEditor.insertRows(Workbook(Vector(s)), S, at = 2, count = 2)
    val restored = StructuralEditor.deleteRows(inserted, S, at = 2, count = 2)
    assertEquals(sheetNamed(restored, "S")(ref"B1").value, formulaCell("A5"))
  }

  test("insert columns shifts column refs at/after the insertion point") {
    val s = new Sheet(name = S).put(ref"A1", formulaCell("=C1"))
    val r = StructuralEditor.insertColumns(Workbook(Vector(s)), S, at = 1, count = 1)
    // C1 (col 2 >= 1) -> D1; cell A1 (col 0 < 1) stays put
    assertEquals(sheetNamed(r, "S")(ref"A1").value, formulaCell("D1"))
  }

  test("delete columns: ref into the deleted band becomes #REF!") {
    val s = new Sheet(name = S).put(ref"A1", formulaCell("=C1"))
    val r = StructuralEditor.deleteColumns(Workbook(Vector(s)), S, at = 2, count = 1)
    // C1 = col index 2 = the deleted column
    assertEquals(sheetNamed(r, "S")(ref"A1").value, CellValue.Error(CellError.Ref))
  }

  test("cross-sheet references to the edited sheet are rewritten") {
    val data = new Sheet(name = SheetName.unsafe("Data")).put(ref"A5", CellValue.Number(99))
    val report =
      new Sheet(name = SheetName.unsafe("Report")).put(ref"B1", formulaCell("=Data!A5"))
    val r =
      StructuralEditor.insertRows(Workbook(Vector(data, report)), SheetName.unsafe("Data"), 2, 1)
    assertEquals(sheetNamed(r, "Report")(ref"B1").value, formulaCell("Data!A6"))
  }

  test("references to OTHER sheets are left untouched") {
    val data = new Sheet(name = SheetName.unsafe("Data"))
    val report = new Sheet(name = SheetName.unsafe("Report")).put(ref"B1", formulaCell("=Data!A5"))
    // Edit Report (not Data); the cross-ref to Data must not move.
    val r =
      StructuralEditor.insertRows(Workbook(Vector(data, report)), SheetName.unsafe("Report"), 2, 1)
    assertEquals(sheetNamed(r, "Report")(ref"B1").value, formulaCell("Data!A5"))
  }

  // ===== GH-428: insert-side clamp at the sheet edge =====

  test("GH-428: a formula ref pushed past the last row degrades to #REF!") {
    val s = new Sheet(name = S)
      .put(ref"B1", formulaCell("A1048576"))
      .put(ref"C1", formulaCell("XFD1"))
    val rows = StructuralEditor.insertRows(Workbook(Vector(s)), S, at = 0, count = 1)
    // the formula cell itself moved to B2; its ref had no home past the edge
    assertEquals(sheetNamed(rows, "S")(ref"B2").value, CellValue.Error(CellError.Ref))
    val cols = StructuralEditor.insertColumns(Workbook(Vector(s)), S, at = 0, count = 1)
    assertEquals(sheetNamed(cols, "S")(ref"D1").value, CellValue.Error(CellError.Ref))
  }

  test("GH-428: a range end clamps at the edge; a range start pushed past it voids the formula") {
    val s = new Sheet(name = S)
      .put(ref"B1", formulaCell("SUM(A10:A1048576)"))
      .put(ref"C1", formulaCell("SUM(A1048570:A1048576)"))
    val r = StructuralEditor.insertRows(Workbook(Vector(s)), S, at = 2, count = 2)
    val s2 = sheetNamed(r, "S")
    assertEquals(s2(ref"B1").value, formulaCell("SUM(A12:A1048576)")) // end pinned at the max
    val voided = StructuralEditor.insertRows(Workbook(Vector(s)), S, at = 0, count = 10)
    // the formula cell itself moved to C11; its range START passed the edge -> #REF!
    assertEquals(sheetNamed(voided, "S")(ref"C11").value, CellValue.Error(CellError.Ref))
  }

  // ===== GH-427: equals-free rewrite + structural cache invalidation =====

  test("GH-427: rewrite prints the model's equals-free form (no '=' lands in <f>)") {
    val s = new Sheet(name = S).put(ref"B1", formulaCell("A5*2")) // file-canonical form
    val r = StructuralEditor.insertRows(Workbook(Vector(s)), S, at = 2, count = 1)
    assertEquals(sheetNamed(r, "S")(ref"B1").value, formulaCell("A6*2"))
  }

  test("structural edits invalidate caches even when every formula reference survives") {
    val cached = Some(CellValue.Number(BigDecimal(4)))
    val s = new Sheet(name = S)
      .put(ref"B1", CellValue.Formula("A5*2", cached))
      .put(ref"B2", CellValue.Formula("A1*3", cached)) // refs above the cut: position unchanged
    val r = StructuralEditor.insertRows(Workbook(Vector(s)), S, at = 2, count = 1)
    val s2 = sheetNamed(r, "S")
    assertEquals(s2(ref"B1").value, CellValue.Formula("A6*2", None))
    assertEquals(s2(ref"B2").value, CellValue.Formula("A1*3", None))
  }

  test("deleting a row invalidates the old cached result of a shrinking SUM range") {
    val s = new Sheet(name = S)
      .put(ref"A1", CellValue.Number(BigDecimal(1)))
      .put(ref"A2", CellValue.Number(BigDecimal(2)))
      .put(ref"A3", CellValue.Number(BigDecimal(3)))
      .put(
        ref"C1",
        CellValue.Formula("SUM(A1:A3)", Some(CellValue.Number(BigDecimal(6))))
      )
    val r = StructuralEditor.deleteRows(Workbook(Vector(s)), S, at = 1, count = 1)
    assertEquals(
      sheetNamed(r, "S")(ref"C1").value,
      CellValue.Formula("SUM(A1:A2)", None)
    )
  }

  test("moving a position-sensitive formula invalidates its cached result") {
    val s = new Sheet(name = S)
      .put(ref"B2", CellValue.Formula("ROW()", Some(CellValue.Number(BigDecimal(2)))))
    val r = StructuralEditor.insertRows(Workbook(Vector(s)), S, at = 0, count = 1)
    assertEquals(sheetNamed(r, "S")(ref"B3").value, CellValue.Formula("ROW()", None))
  }

  test("GH-427: a fully-deleted reference still degrades to #REF! (cache dropped)") {
    val s = new Sheet(name = S)
      .put(ref"B1", CellValue.Formula("A3", Some(CellValue.Number(BigDecimal(7)))))
    val r = StructuralEditor.deleteRows(Workbook(Vector(s)), S, at = 2, count = 1)
    assertEquals(sheetNamed(r, "S")(ref"B1").value, CellValue.Error(CellError.Ref))
  }

  // ===== GH-274: INDIRECT and structural edits =====

  test("GH-274: insert row freezes INDIRECT text but shifts INDIRECT ref arguments") {
    val s = new Sheet(name = S)
      .put(ref"B1", formulaCell("=INDIRECT(\"A5\")")) // text is data — frozen (Excel parity)
      .put(ref"C1", formulaCell("=INDIRECT(A5)")) // the A5 ARGUMENT is a real ref — shifts
    val r = StructuralEditor.insertRows(Workbook(Vector(s)), S, at = 2, count = 1)
    val s2 = sheetNamed(r, "S")
    assertEquals(s2(ref"B1").value, formulaCell("INDIRECT(\"A5\")"))
    assertEquals(s2(ref"C1").value, formulaCell("INDIRECT(A6)"))
  }
