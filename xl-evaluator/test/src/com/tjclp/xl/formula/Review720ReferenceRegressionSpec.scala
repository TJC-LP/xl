package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/** Regression coverage for PR #720 reference evaluation. */
class Review720ReferenceRegressionSpec extends FunSuite:
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

  test("a computed range propagates INDIRECT's reference error") {
    assertEquals(plain("SUM(A1:INDIRECT(\"nowhere!!\"))"), Right(CellValue.Error(CellError.Ref)))
  }

  test("a computed range propagates OFFSET's reference error") {
    assertEquals(plain("SUM(A1:OFFSET(A1,-1,0))"), Right(CellValue.Error(CellError.Ref)))
  }

  test("a computed range propagates an explicit XLOOKUP fallback error") {
    assertEquals(plain("SUM(A1:XLOOKUP(9,A1:A3,B1:B3,#N/A))"), Right(CellValue.Error(CellError.NA)))
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

  test("computed whole axes share both ends of the cross-sheet SUMPRODUCT extent") {
    val offset = Sheet(SheetName.unsafe("Offset")).put(ref"A3", number(3))
    val other = Sheet(SheetName.unsafe("Other"))
      .put(ref"B2", number(10))
      .put(ref"B3", number(20))
      .put(ref"B5", number(30))
    assertEquals(
      offset.evaluateFormula(
        "=SUMPRODUCT(A:A,INDEX(Other!B:B,0))",
        Clock.system,
        Some(Workbook(offset, other))
      ),
      Right(number(60))
    )
    val row = Sheet(SheetName.unsafe("Row")).put(ref"B1", number(2)).put(ref"C1", number(3))
    assertEquals(row.evaluateFormula("=SUMPRODUCT(INDEX(1:1,0),1:1)"), Right(number(13)))
    assertEquals(
      Sheet(SheetName.unsafe("Empty")).evaluateFormula("=SUMPRODUCT(INDEX(1:1,0),1:1)"),
      Right(number(0))
    )
  }

  test("a computed whole axis without static range inputs keeps its materialization semantics") {
    val book = Workbook(sheet).withDefinedName("fullCol", "Data!$A:$A")
    assertEquals(
      sheet.evaluateFormula("=SUMPRODUCT(INDEX(fullCol,0))", Clock.system, Some(book)),
      Right(number(6))
    )
  }
