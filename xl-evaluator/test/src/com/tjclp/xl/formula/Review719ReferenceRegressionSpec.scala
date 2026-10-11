package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/** Regression coverage for PR #719 reference evaluation. */
class Review719ReferenceRegressionSpec extends FunSuite:
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

  test("SUMPRODUCT bounds ROW of a computed whole-column reference consistently") {
    assertEquals(array("=SUMPRODUCT(A:A,ROW(INDEX(A:A,0)))"), Right(number(14)))
  }

  test("SUMPRODUCT bounds COLUMN of a computed whole-row reference consistently") {
    assertEquals(array("=SUMPRODUCT(1:1,COLUMN(INDEX(1:1,0)))"), Right(number(21)))
  }

  test("ROW keeps the reference of an INDEX result bound by LET in array mode") {
    assertEquals(array("=LET(chosen,INDEX(B1:B3,2),ROW(chosen))"), Right(number(2)))
  }

  test("computed positions preserve leading blanks, empty extents, and the INDEX origin") {
    val offset = Sheet(SheetName.unsafe("Offset"))
      .put(ref"A2", number(2))
      .put(ref"A3", number(3))
    assertEquals(
      offset.evaluateFormula("=SUMPRODUCT(A:A,ROW(INDEX(A:A,0)))"),
      Right(number(13))
    )
    assertEquals(
      offset.evaluateFormula("=SUMPRODUCT(ROW(INDEX(A:A,2)))"),
      Right(number(2))
    )
    assertEquals(
      Sheet(SheetName.unsafe("Empty")).evaluateFormula("=SUMPRODUCT(A:A,ROW(INDEX(A:A,0)))"),
      Right(number(0))
    )
  }

  test("computed positions use SUMPRODUCT's shared cross-sheet extent") {
    val other = Sheet(SheetName.unsafe("Other")).put(ref"A5", number(5))
    val wb = Workbook(sheet, other)
    assertEquals(
      sheet.evaluateFormula("=SUMPRODUCT(A:A,ROW(INDEX(Other!A:A,0)))", Clock.system, Some(wb)),
      Right(number(14))
    )
  }

  test("LET keeps references for both coordinates and shape inspection") {
    assertEquals(array("=LET(chosen,INDEX(A1:B3,2,0),COLUMN(chosen))"), Right(number(1)))
    assertEquals(array("=LET(chosen,INDEX(B:B,0),ROWS(chosen))"), Right(number(1048576)))
    assertEquals(array("=LET(chosen,INDEX(1:1,0),COLUMNS(chosen))"), Right(number(16384)))
  }
