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
    assertEquals(s2(ref"B1").value, formulaCell("=A6")) // A5 (row 4 >= 2) -> A6
    assertEquals(s2(ref"B2").value, formulaCell("=A1")) // A1 (row 0 < 2) unchanged
  }

  test("insert rows moves the formula cell itself and still rewrites refs") {
    val s = new Sheet(name = S).put(ref"B5", formulaCell("=A1"))
    val r = StructuralEditor.insertRows(Workbook(Vector(s)), S, at = 2, count = 1)
    val s2 = sheetNamed(r, "S")
    assertEquals(s2(ref"B6").value, formulaCell("=A1")) // cell B5 (row 4) -> B6; ref A1 unchanged
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
    assertEquals(sheetNamed(r, "S")(ref"B1").value, formulaCell("=A4")) // A5 (row 4) -> A4
  }

  test("delete rows: a range straddling the deletion shrinks") {
    val s = new Sheet(name = S).put(ref"B1", formulaCell("=SUM(A1:A10)"))
    val r = StructuralEditor.deleteRows(Workbook(Vector(s)), S, at = 2, count = 1)
    assertEquals(sheetNamed(r, "S")(ref"B1").value, formulaCell("=SUM(A1:A9)"))
  }

  test("insert then delete the same band is identity on formula refs") {
    val s = new Sheet(name = S).put(ref"B1", formulaCell("=A5"))
    val inserted = StructuralEditor.insertRows(Workbook(Vector(s)), S, at = 2, count = 2)
    val restored = StructuralEditor.deleteRows(inserted, S, at = 2, count = 2)
    assertEquals(sheetNamed(restored, "S")(ref"B1").value, formulaCell("=A5"))
  }

  test("insert columns shifts column refs at/after the insertion point") {
    val s = new Sheet(name = S).put(ref"A1", formulaCell("=C1"))
    val r = StructuralEditor.insertColumns(Workbook(Vector(s)), S, at = 1, count = 1)
    // C1 (col 2 >= 1) -> D1; cell A1 (col 0 < 1) stays put
    assertEquals(sheetNamed(r, "S")(ref"A1").value, formulaCell("=D1"))
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
    val r = StructuralEditor.insertRows(Workbook(Vector(data, report)), SheetName.unsafe("Data"), 2, 1)
    assertEquals(sheetNamed(r, "Report")(ref"B1").value, formulaCell("=Data!A6"))
  }

  test("references to OTHER sheets are left untouched") {
    val data = new Sheet(name = SheetName.unsafe("Data"))
    val report = new Sheet(name = SheetName.unsafe("Report")).put(ref"B1", formulaCell("=Data!A5"))
    // Edit Report (not Data); the cross-ref to Data must not move.
    val r = StructuralEditor.insertRows(Workbook(Vector(data, report)), SheetName.unsafe("Report"), 2, 1)
    assertEquals(sheetNamed(r, "Report")(ref"B1").value, formulaCell("=Data!A5"))
  }
