package com.tjclp.xl.formula

import munit.FunSuite

import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.eval.SheetEvaluator.*
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.syntax.*

/**
 * GH-562: DATE(year, month, day) normalizes month and day overflow/underflow the way Excel does
 * instead of rejecting components outside 1-12 / 1-31.
 *
 * The month-spine idiom `=DATE(YEAR(x),MONTH(x)+1,1)` errored on every December and left the cell
 * and its dependents uncached. Excel's documented rule (support.microsoft.com, DATE function):
 * month and day are offsets from January 1 of the year — a month above 12 adds months, below 1
 * subtracts; days overflow and underflow the month the same way; a year from 0 to 1899 is offset by
 * 1900.
 */
class DateNormalizationSpec extends FunSuite:

  private val sheet = Sheet("Test")

  private def assertDate(formula: String, y: Int, m: Int, d: Int)(implicit
    loc: munit.Location
  ): Unit =
    assertEquals(sheet.evaluateFormula(s"=YEAR($formula)"), Right(CellValue.Number(y)), formula)
    assertEquals(sheet.evaluateFormula(s"=MONTH($formula)"), Right(CellValue.Number(m)), formula)
    assertEquals(sheet.evaluateFormula(s"=DAY($formula)"), Right(CellValue.Number(d)), formula)

  test("GH-562: month 13 rolls into January of the next year (the December month-spine step)") {
    assertDate("DATE(2026,13,1)", 2027, 1, 1)
    // serial 46388 per the issue's Excel oracle
    assertEquals(sheet.evaluateFormula("=DATE(2026,13,1)+0"), Right(CellValue.Number(46388)))
  }

  test("GH-562: Microsoft's documented examples") {
    assertDate("DATE(2008,14,2)", 2009, 2, 2) // month overflow
    assertDate("DATE(2008,-3,2)", 2007, 9, 2) // month underflow
    assertDate("DATE(2008,1,35)", 2008, 2, 4) // day overflow
    assertDate("DATE(2008,1,-15)", 2007, 12, 16) // day underflow
  }

  test("GH-562: month 0 and day 0 step back one unit") {
    assertDate("DATE(2026,0,1)", 2025, 12, 1)
    assertDate("DATE(2026,3,0)", 2026, 2, 28)
    assertDate("DATE(2024,3,0)", 2024, 2, 29) // leap year
  }

  test("GH-562: in-range components are unchanged") {
    assertDate("DATE(2025,11,21)", 2025, 11, 21)
    assertDate("DATE(2024,2,29)", 2024, 2, 29)
  }

  test("GH-562: the month-spine idiom evaluates across a year boundary") {
    val dated =
      sheet.put(ref"A1", CellValue.DateTime(java.time.LocalDate.of(2026, 12, 15).atStartOfDay()))
    assertEquals(
      dated.evaluateFormula("=DATE(YEAR(A1),MONTH(A1)+1,1)=DATE(2027,1,1)"),
      Right(CellValue.Bool(true))
    )
  }

  test("GH-562: years 0..1899 are offset by 1900, as in Excel") {
    assertDate("DATE(108,1,2)", 2008, 1, 2)
    assertDate("DATE(0,1,1)", 1900, 1, 1)
  }

  test("GH-562: results outside the 1900 date system are #NUM!") {
    assertEquals(sheet.evaluateFormula("=DATE(-1,1,1)"), Right(CellValue.Error(CellError.Num)))
    assertEquals(sheet.evaluateFormula("=DATE(10000,1,1)"), Right(CellValue.Error(CellError.Num)))
    assertEquals(sheet.evaluateFormula("=DATE(1900,1,-5)"), Right(CellValue.Error(CellError.Num)))
    assertEquals(sheet.evaluateFormula("=DATE(9999,12,32)"), Right(CellValue.Error(CellError.Num)))
  }
