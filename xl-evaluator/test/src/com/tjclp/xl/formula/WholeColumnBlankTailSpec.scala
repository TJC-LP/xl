package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * GH-711 / GH-192: SUMPRODUCT, COUNTIF and COUNTIFS read a whole column within the sheet's used
 * extent, and add back the rows trimmed past it, every one blank, as Excel counts them. Each
 * expected value is LibreOffice 24.2's, except where a comment says otherwise.
 */
class WholeColumnBlankTailSpec extends FunSuite:

  private def num(n: Double): CellValue = CellValue.Number(BigDecimal(n))

  /** The review's book: A1=4, A2=5, B1=1; every other cell blank. */
  private val sheet: Sheet =
    Sheet(SheetName.unsafe("T"))
      .put(ref"A1", num(4))
      .put(ref"A2", num(5))
      .put(ref"B1", num(1))

  private def eval(formula: String, on: Sheet = sheet): XLResult[CellValue] =
    on.evaluateFormula(formula, Clock.system, Some(Workbook(Vector(on))), None)

  private def evalIn(wb: Workbook, on: Sheet)(formula: String): XLResult[CellValue] =
    on.evaluateFormula(formula, Clock.system, Some(wb), None)

  test("a selector's whole-column condition counts the blank rows past the used range") {
    // 1048576 rows, two filled: the trimmed condition must not lose the 1048574 blank ones
    assertEquals(eval("=SUMPRODUCT(IF(A:A=\"\",1,0))"), Right(num(1048574)))
    assertEquals(eval("=SUMPRODUCT(IF(ISBLANK(A:A),1,0))"), Right(num(1048574)))
    assertEquals(eval("=SUMPRODUCT(CHOOSE(1,--(A:A=\"\")))"), Right(num(1048574)))
    // the trimmed branches still line up with the condition (GH-711)
    assertEquals(eval("=SUMPRODUCT(IF(A:A>4,B:B,A:A))"), Right(num(4)))
    assertEquals(eval("=SUMPRODUCT(IF(A:A>4,B:B,A:A),A:A)"), Right(num(16)))
  }

  test("SUMPRODUCT counts every blank row of a whole column, as Excel does (GH-192)") {
    assertEquals(eval("=SUMPRODUCT(--(A:A=\"\"))"), Right(num(1048574)))
    assertEquals(eval("=SUMPRODUCT(--ISBLANK(A:A))"), Right(num(1048574)))
    assertEquals(eval("=SUMPRODUCT((A:A=\"\")*(B:B=\"\"))"), Right(num(1048574)))
    assertEquals(eval("=SUMPRODUCT(--(A:A=\"\"),--(B:B=\"\"))"), Right(num(1048574)))
    assertEquals(eval("=SUMPRODUCT((A:A=\"\")*2)"), Right(num(2097148)))
    // two columns: 2 * 1048576 cells, three filled
    assertEquals(eval("=SUMPRODUCT(--(A:B=\"\"))"), Right(num(2097149)))
    // a whole row: 16384 cells, two filled
    assertEquals(eval("=SUMPRODUCT(--(1:1=\"\"))"), Right(num(16382)))
  }

  test("a blank row that contributes nothing leaves the trimmed sum as it was") {
    assertEquals(eval("=SUMPRODUCT(--(A:A>0))"), Right(num(2)))
    assertEquals(eval("=SUMPRODUCT(A:A,B:B)"), Right(num(4)))
    assertEquals(eval("=SUMPRODUCT(A:A*2)"), Right(num(18)))
    assertEquals(eval("=SUMPRODUCT(IF(TRUE,B:B,A:A))"), Right(num(1)))
  }

  test("the rows above a used range that starts below row 1 count too") {
    val lower = Sheet(SheetName.unsafe("U"))
      .put(ref"A5", num(1))
      .put(ref"A6", num(2))
      .put(ref"A7", num(3))
    assertEquals(eval("=SUMPRODUCT(--(A:A=\"\"))", lower), Right(num(1048573)))
    assertEquals(eval("=SUMPRODUCT(A:A)", lower), Right(num(6)))
    assertEquals(eval("=COUNTIF(A:A,\"\")", lower), Right(num(1048573)))
  }

  test("an empty sheet is blank throughout") {
    val empty = Sheet(SheetName.unsafe("E"))
    assertEquals(eval("=SUMPRODUCT(--(A:A=\"\"))", empty), Right(num(1048576)))
    assertEquals(eval("=SUMPRODUCT(A:A)", empty), Right(num(0)))
    assertEquals(eval("=COUNTIF(A:A,\"\")", empty), Right(num(1048576)))
  }

  test("an error at a blank row is the result, as over the whole column in Excel") {
    // 1/blank is #DIV/0! in every trimmed row
    assertEquals(eval("=SUMPRODUCT(1/A:A)"), Right(CellValue.Error(CellError.Div0)))
    // ...and stays one when no row is trimmed off
    val bounded = eval("=SUMPRODUCT(1/A1:A3)")
    assertEquals(bounded, Right(CellValue.Error(CellError.Div0)))
  }

  test("the distinct-count idiom over whole columns stays exact: blanks match the criterion") {
    // A1:A3 = x, y, x. Each blank row adds 0/COUNTIF(A:A,"") = 0, COUNTIF counting every blank
    val texts = Sheet(SheetName.unsafe("W"))
      .put(ref"A1", CellValue.Text("x"))
      .put(ref"A2", CellValue.Text("y"))
      .put(ref"A3", CellValue.Text("x"))
    assertEquals(eval("=SUMPRODUCT((A:A<>\"\")/COUNTIF(A:A,A:A&\"\"))", texts), Right(num(2)))
    // Excel's whole-column 1/COUNTIF(A:A,A:A) divides by COUNTIF(A:A,<blank>) at every blank
    // row: no zeros, so #DIV/0! (the bounded range is the distinct count)
    assertEquals(
      eval("=SUMPRODUCT(1/COUNTIF(A:A,A:A))", texts),
      Right(CellValue.Error(CellError.Div0))
    )
    assertEquals(eval("=SUMPRODUCT(1/COUNTIF(A1:A3,A1:A3))", texts), Right(num(2)))
  }

  test("COUNTIF and COUNTIFS count the blank cells past the used range for a blank criterion") {
    assertEquals(eval("=COUNTIF(A:A,\"\")"), Right(num(1048574)))
    assertEquals(eval("=COUNTIF(A:A,\"=\")"), Right(num(1048574)))
    assertEquals(eval("=COUNTIF(A:A,\"<>\")"), Right(num(2)))
    assertEquals(eval("=COUNTIF(A:A,0)"), Right(num(0)))
    assertEquals(eval("=COUNTIF(A:A,\"<5\")"), Right(num(1)))
    assertEquals(eval("=COUNTIF(A:B,\"\")"), Right(num(2097149)))
    assertEquals(eval("=COUNTIFS(A:A,\"\",B:B,\"\")"), Right(num(1048574)))
    assertEquals(eval("=COUNTIFS(A:A,\"\",B:B,\"<>\")"), Right(num(0)))
    // a whole row (Excel's rule; LibreOffice 24.2 answers 0 here, its whole-row COUNTIF reads no
    // blank cell, while it counts the whole-column ones above)
    assertEquals(eval("=COUNTIF(1:1,\"\")"), Right(num(16382)))
  }

  test("the trimmed blank rows are blank on every sheet the arguments read") {
    val other = Sheet(SheetName.unsafe("O"))
      .put(ref"A1", num(1))
      .put(ref"A2", num(2))
      .put(ref"A3", num(3))
      .put(ref"A4", num(4))
    val wb = Workbook(Vector(sheet, other))
    val e = evalIn(wb, sheet)
    // shared extent rows 1-4: A3, A4 here are blank inside it, rows 5 on are blank in both
    assertEquals(e("=SUMPRODUCT(--(A:A=\"\"),--(O!A:A>0))"), Right(num(2)))
    assertEquals(e("=SUMPRODUCT(--(A:A=\"\"),--(O!A:A=\"\"))"), Right(num(1048572)))
  }
