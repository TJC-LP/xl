package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/** Cross-PR regression probes for #718, #719 and #720. */
class ReviewArrayRefsIntegrationSpec extends FunSuite:
  private def number(n: Int): CellValue = CellValue.Number(BigDecimal(n))
  private val sheet = (1 to 3).foldLeft(Sheet(SheetName.unsafe("Data"))) { (s, i) =>
    s.put(ARef.from0(0, i - 1), number(i))
      .put(ARef.from0(1, i - 1), number(i * 10))
  }

  private def array(formula: String): XLResult[CellValue] =
    sheet.evaluateFormula(formula, Clock.system, Some(Workbook(sheet)), None)

  private def plain(formula: String): XLResult[CellValue] =
    val placed = sheet.put(ref"H9", CellValue.Formula(formula, None))
    placed.evaluateCell(ref"H9", Clock.system, Some(Workbook(placed)))

  test("computed scalar INDEX position retains its reference in array mode") {
    assertEquals(array("=SUM(A1:INDEX(A1:A3,COUNTA(A1:A3)))"), Right(number(6)))
  }

  test("ROW accepts an INDEX reference with a computed scalar position") {
    assertEquals(array("=ROW(INDEX(B1:B3,SUM(1,1)))"), Right(number(2)))
  }

  test("ROW accepts an INDEX reference with a one-element array position") {
    assertEquals(array("=ROW(INDEX(B1:B3,{2}))"), Right(number(2)))
  }

  test("ROW accepts an XLOOKUP reference with a computed scalar key") {
    assertEquals(array("=ROW(XLOOKUP(SUM(1,1),A1:A3,B1:B3))"), Right(number(2)))
  }

  test("COUNTBLANK sees all rows of an INDEX reference with a computed scalar zero") {
    assertEquals(array("=COUNTBLANK(INDEX(A:A,SUM(0)))"), Right(number(1048573)))
  }

  test("a computed range propagates INDIRECT's reference error") {
    assertEquals(plain("SUM(A1:INDIRECT(\"nowhere!!\"))"), Right(CellValue.Error(CellError.Ref)))
  }

  test("a computed range propagates OFFSET's reference error") {
    assertEquals(plain("SUM(A1:OFFSET(A1,-1,0))"), Right(CellValue.Error(CellError.Ref)))
  }

  test("a computed range propagates an explicit XLOOKUP fallback error") {
    assertEquals(plain("SUM(A1:XLOOKUP(9,A1:A3,B1:B3,#N/A))"), Right(CellValue.Error(CellError.NA)))
  }

  test("SUMPRODUCT bounds ROW of a computed whole-column reference consistently") {
    assertEquals(array("=SUMPRODUCT(A:A,ROW(INDEX(A:A,0)))"), Right(number(14)))
  }

  test("SUMPRODUCT bounds COLUMN of a computed whole-row reference consistently") {
    assertEquals(array("=SUMPRODUCT(1:1,COLUMN(INDEX(1:1,0)))"), Right(number(21)))
  }

  test("SUMPRODUCT INDEX whole-column result aligns with a written column after leading blanks") {
    val offset = Sheet(SheetName.unsafe("Offset"))
      .put(ref"A2", number(2))
      .put(ref"A3", number(3))
    assertEquals(
      offset.evaluateFormula("=SUMPRODUCT(INDEX(A:A,0),A:A)"),
      Right(number(13))
    )
  }

  test("SUMPRODUCT INDEX whole-column result agrees with an empty written column") {
    assertEquals(
      Sheet(SheetName.unsafe("Empty")).evaluateFormula("=SUMPRODUCT(INDEX(A:A,0),A:A)"),
      Right(number(0))
    )
  }

  test("ROW keeps the reference of an INDEX result bound by LET in array mode") {
    assertEquals(array("=LET(chosen,INDEX(B1:B3,2),ROW(chosen))"), Right(number(2)))
  }
