package com.tjclp.xl.formula

import munit.FunSuite

import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.CriteriaMatcher.*
import com.tjclp.xl.formula.eval.SheetEvaluator.*
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.syntax.*

/**
 * GH-565: a criteria string that spells an Excel error literal (`"#N/A"`, `"#DIV/0!"`, ...) is that
 * error VALUE, not text.
 *
 * Excel parses criteria like formula text, so `COUNTIF(range,"#N/A")` counts cells holding the #N/A
 * error and never cells holding the text "#N/A". On an Excel-authored production workbook whose
 * data column held the text, xl's `SUMIFS(...,"#N/A")` rows diverged from Excel's cached results.
 */
class CriteriaErrorLiteralSpec extends FunSuite:

  // A1:A5 = #N/A, #DIV/0!, text "#N/A", 1, blank; B1:B5 = 10, 20, 30, 40, blank
  private val sheet = Sheet("Test")
    .put(ref"A1", CellValue.Error(CellError.NA))
    .put(ref"A2", CellValue.Error(CellError.Div0))
    .put(ref"A3", CellValue.Text("#N/A"))
    .put(ref"A4", CellValue.Number(1))
    .put(ref"B1", CellValue.Number(10))
    .put(ref"B2", CellValue.Number(20))
    .put(ref"B3", CellValue.Number(30))
    .put(ref"B4", CellValue.Number(40))

  private def assertScalar(formula: String, expected: CellValue)(implicit
    loc: munit.Location
  ): Unit =
    assertEquals(sheet.evaluateFormula(formula), Right(expected), formula)

  // ===== CriteriaMatcher unit level =====

  test("GH-565: every error literal parses to the error value, checked before the wildcard rule") {
    assertEquals(
      parse(ExprValue.Text("#N/A")),
      Exact(ExprValue.Cell(CellValue.Error(CellError.NA)))
    )
    assertEquals(
      parse(ExprValue.Text("#DIV/0!")),
      Exact(ExprValue.Cell(CellValue.Error(CellError.Div0)))
    )
    // "#NAME?" contains '?' but is an error literal, not a one-char wildcard pattern
    assertEquals(
      parse(ExprValue.Text("#NAME?")),
      Exact(ExprValue.Cell(CellValue.Error(CellError.Name)))
    )
    assertEquals(
      parse(ExprValue.Text("=#REF!")),
      Exact(ExprValue.Cell(CellValue.Error(CellError.Ref)))
    )
    assertEquals(parse(ExprValue.Text("<>#N/A")), NotError(CellError.NA))
  }

  test("GH-565: an error cell matches its own literal and nothing else") {
    val na = Exact(ExprValue.Cell(CellValue.Error(CellError.NA)))
    assert(matches(CellValue.Error(CellError.NA), na))
    assert(!matches(CellValue.Error(CellError.Div0), na))
    assert(!matches(CellValue.Text("#N/A"), na))
    assert(!matches(CellValue.Error(CellError.NA), Exact(ExprValue.Text("#N/A"))))
    assert(matches(CellValue.Formula("NA()", Some(CellValue.Error(CellError.NA))), na))
  }

  test("GH-565: <>#N/A matches everything but that error, blanks and other errors included") {
    val notNa = NotError(CellError.NA)
    assert(!matches(CellValue.Error(CellError.NA), notNa))
    assert(matches(CellValue.Error(CellError.Div0), notNa))
    assert(matches(CellValue.Text("#N/A"), notNa))
    assert(matches(CellValue.Number(1), notNa))
    assert(matches(CellValue.Empty, notNa))
  }

  // ===== Evaluation level =====

  test("GH-565: COUNTIF with an error literal counts error cells, not the text (the issue repro)") {
    val textOnly = Sheet("Test")
      .put(ref"H1", CellValue.Text("#N/A"))
      .put(ref"H2", CellValue.Text("x"))
      .put(ref"H3", CellValue.Text("#N/A"))
    assertEquals(textOnly.evaluateFormula("=COUNTIF(H1:H3,\"#N/A\")"), Right(CellValue.Number(0)))
    assertScalar("=COUNTIF(A1:A5,\"#N/A\")", CellValue.Number(1))
    assertScalar("=COUNTIF(A1:A5,\"#DIV/0!\")", CellValue.Number(1))
    assertScalar("=COUNTIF(A1:A5,\"#NAME?\")", CellValue.Number(0))
  }

  test("GH-565: SUMIFS/COUNTIFS with error-literal criteria") {
    assertScalar("=SUMIFS(B1:B5,A1:A5,\"#N/A\")", CellValue.Number(10))
    assertScalar("=SUMIFS(B1:B5,A1:A5,\"<>#N/A\")", CellValue.Number(90))
    assertScalar("=COUNTIFS(A1:A5,\"<>#N/A\")", CellValue.Number(4))
  }

  test("GH-565: a reference to an error cell as the criterion matches that error") {
    assertScalar("=COUNTIF(A1:A5,A1)", CellValue.Number(1))
    assertScalar("=SUMIF(A1:A5,A2,B1:B5)", CellValue.Number(20))
  }

  test("GH-565: plain text criteria are unchanged") {
    assertScalar("=COUNTIF(A1:A5,\"x\")", CellValue.Number(0))
    assertScalar("=COUNTIF(A1:A5,1)", CellValue.Number(1))
  }
