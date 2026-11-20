package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cells.{CellValue, Comment}
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import com.tjclp.xl.formula.eval.StructuralEditor
import com.tjclp.xl.unsafe.*
import munit.FunSuite

/**
 * GH-472: a structural INSERT that would shift any populated position past the sheet edge (row
 * 1048576 / column XFD) is REFUSED with a typed `Left(XLError.OutOfBounds)` — never silently
 * written. Ranges have clamped since GH-428; data cells, comments, row/column properties, and
 * drawing anchors carry absolute positions and cannot clamp without destroying data, so the edit
 * itself must be rejected (Excel refuses files with cells past the edge outright).
 */
class StructuralBoundsSpec extends FunSuite:

  private val S = SheetName.unsafe("S")

  private def sheetNamed(wb: Workbook, n: String): Sheet =
    wb.sheets.find(_.name == SheetName.unsafe(n)).getOrElse(fail(s"missing sheet $n"))

  private def outOfBounds(result: XLResult[Workbook]): XLError =
    result match
      case Left(e: XLError.OutOfBounds) => e
      case Left(other) => fail(s"expected XLError.OutOfBounds, got $other")
      case Right(_) => fail("expected the structural edit to be refused")

  test("insertRows refuses when a populated cell would shift past row 1048576") {
    val s = new Sheet(name = S)
      .put(ref"A1048570", CellValue.Number(1))
      .put(ref"B1048572", CellValue.Number(2))
    val wb = Workbook(Vector(s))
    val err = outOfBounds(StructuralEditor.insertRowsChecked(wb, S, at = 4, count = 20))
    // the farthest populated row is named in the error
    assert(err.message.contains("1048572"), err.message)
    assert(err.message.contains("1048576"), err.message)
  }

  test("insertRows just under the cap still works (cell lands exactly on row 1048576)") {
    val s = new Sheet(name = S).put(ref"A1048556", CellValue.Number(7))
    val wb = StructuralEditor.insertRowsChecked(Workbook(Vector(s)), S, at = 4, count = 20).unsafe
    val s2 = sheetNamed(wb, "S")
    assertEquals(s2(ref"A1048576").value, CellValue.Number(7))
    assert(!s2.contains(ref"A1048556"))
  }

  test("insertColumns refuses when a populated cell would shift past column XFD") {
    val s = new Sheet(name = S).put(ref"XFC1", CellValue.Number(1))
    val wb = Workbook(Vector(s))
    val err = outOfBounds(StructuralEditor.insertColumnsChecked(wb, S, at = 2, count = 20))
    assert(err.message.contains("XFC"), err.message)
    assert(err.message.contains("XFD"), err.message)
  }

  test("insertColumns just under the cap still works (cell lands exactly on XFD)") {
    val s = new Sheet(name = S).put(ref"XFC1", CellValue.Number(9))
    val wb = StructuralEditor.insertColumnsChecked(Workbook(Vector(s)), S, at = 2, count = 1).unsafe
    val s2 = sheetNamed(wb, "S")
    assertEquals(s2(ref"XFD1").value, CellValue.Number(9))
  }

  test("insertRows refuses on a comment past the edge even with no cell there") {
    val s = new Sheet(name = S)
      .put(ref"A1", CellValue.Number(1))
      .comment(ref"C1048572", Comment.plainText("note"))
    val wb = Workbook(Vector(s))
    val err = outOfBounds(StructuralEditor.insertRowsChecked(wb, S, at = 4, count = 20))
    assert(err.message.contains("1048572"), err.message)
  }

  test("insertRows below every populated position does not refuse") {
    val s = new Sheet(name = S).put(ref"A1048570", CellValue.Number(1))
    // at is past the populated row: nothing shifts, nothing overflows
    val wb =
      StructuralEditor.insertRowsChecked(Workbook(Vector(s)), S, at = 1048571, count = 20).unsafe
    assertEquals(sheetNamed(wb, "S")(ref"A1048570").value, CellValue.Number(1))
  }

  test("deleteRows near the edge never refuses") {
    val s = new Sheet(name = S).put(ref"A1048576", CellValue.Number(1))
    val wb = StructuralEditor.deleteRowsChecked(Workbook(Vector(s)), S, at = 0, count = 5).unsafe
    assertEquals(sheetNamed(wb, "S")(ref"A1048571").value, CellValue.Number(1))
  }
