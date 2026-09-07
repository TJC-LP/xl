package com.tjclp.xl.formula

import munit.FunSuite

import java.time.LocalDate

import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.eval.SheetEvaluator.*
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.syntax.*

/**
 * GH-561: a date in a TEXT position renders as its Excel serial number, never as ISO text.
 *
 * Dates are numbers in Excel's value model; only TEXT() formats them. `">="&DATE(2026,1,1)` is
 * therefore `">=46023"` — the criteria idiom COUNTIFS/SUMIFS/AVERAGEIFS/MAXIFS/MINIFS rows depend
 * on. The ISO rendering (`">=2026-01-01"`) matched no numeric cell, so every such row cached a
 * wrong number silently and `--strict` passed.
 */
class DateSerialTextSpec extends FunSuite:

  // C1:C3 = 46000, 46100, 46200 (the issue's repro); D1:D3 = 1, 10, 100
  private val sheet = Sheet("Test")
    .put(ref"C1", CellValue.Number(46000))
    .put(ref"C2", CellValue.Number(46100))
    .put(ref"C3", CellValue.Number(46200))
    .put(ref"D1", CellValue.Number(1))
    .put(ref"D2", CellValue.Number(10))
    .put(ref"D3", CellValue.Number(100))
    .put(ref"A1", CellValue.DateTime(LocalDate.of(2026, 1, 1).atStartOfDay()))
    .put(ref"A2", CellValue.DateTime(LocalDate.of(2026, 1, 1).atTime(12, 0)))
    .put(ref"A3", CellValue.Formula("DATE(2026,1,1)", None))

  private def assertScalar(formula: String, expected: CellValue)(implicit
    loc: munit.Location
  ): Unit =
    assertEquals(sheet.evaluateFormula(formula), Right(expected), formula)

  test("GH-561: & on a DATE() result yields the serial, matching Excel") {
    assertScalar("=\">=\"&DATE(2026,1,1)", CellValue.Text(">=46023"))
    assertScalar("=\"x\"&DATE(2026,1,1)&\"y\"", CellValue.Text("x46023y"))
  }

  test("GH-561: & on a date CELL (typed or an uncached =DATE formula) yields the serial") {
    assertScalar("=\">=\"&A1", CellValue.Text(">=46023"))
    assertScalar("=\">=\"&A3", CellValue.Text(">=46023"))
  }

  test("GH-561: a time fraction survives, rounded to Excel's 15 significant digits") {
    assertScalar("=A2&\"\"", CellValue.Text("46023.5"))
  }

  test("GH-561: the COUNTIFS/SUMIFS date-criteria idiom matches numeric date cells") {
    assertScalar("=COUNTIFS(C1:C3,\">=\"&DATE(2026,1,1))", CellValue.Number(2))
    assertScalar("=SUMIFS(D1:D3,C1:C3,\">=\"&DATE(2026,1,1))", CellValue.Number(110))
    assertScalar("=COUNTIF(C1:C3,\"<\"&A1)", CellValue.Number(1))
    assertScalar("=AVERAGEIFS(D1:D3,C1:C3,\">=\"&A3)", CellValue.Number(55))
    assertScalar("=MAXIFS(D1:D3,C1:C3,\"<\"&DATE(2026,1,1))", CellValue.Number(1))
    assertScalar("=MINIFS(D1:D3,C1:C3,\">=\"&DATE(2026,1,1))", CellValue.Number(10))
  }

  test("GH-561: text functions see the serial too (LEN of a date cell is the serial's length)") {
    assertScalar("=LEN(A1)", CellValue.Number(5))
    assertScalar("=CONCATENATE(\"d=\",DATE(2026,1,1))", CellValue.Text("d=46023"))
  }

  test("GH-561: TEXT() is still the formatting path") {
    assertScalar("=TEXT(DATE(2026,1,1),\"yyyy-mm-dd\")", CellValue.Text("2026-01-01"))
    assertScalar("=TEXT(A1,\"yyyy-mm-dd\")", CellValue.Text("2026-01-01"))
  }

  test("GH-561: dates still compare as numbers (unchanged)") {
    assertScalar("=DATE(2026,1,1)=46023", CellValue.Bool(true))
  }
