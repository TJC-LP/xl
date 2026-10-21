package com.tjclp.xl.formula

import munit.FunSuite

import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.eval.SheetEvaluator.*
import com.tjclp.xl.formula.eval.WorkbookEvaluator.*
import com.tjclp.xl.richtext.{RichText, TextRun}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.syntax.*
import com.tjclp.xl.workbooks.Workbook

/**
 * COUNTBLANK against results desktop Excel (Microsoft 365 for Mac) calculated on this grid (#731).
 *
 * Data!A1:A14 = 1, blank, the constant "", `=""`, rich text "xy", " ", 0, FALSE, `=NA()`, `=A2`,
 * `=IF(1,"","x")`, "abc", blank, TRUE. B1:B5 = the constant "", `=""`, 5, blank, "t". C1:C2 =
 * `=""`, the constant "".
 */
class CountBlankExcelParitySpec extends FunSuite:

  /** The Data sheet, its formulas with their cached values or, `cached = false`, without them. */
  private def data(cached: Boolean): Sheet =
    def formula(expression: String, value: CellValue): CellValue =
      CellValue.Formula(expression, Option.when(cached)(value))
    Sheet("Data")
      .put(ref"A1", CellValue.Number(1))
      .put(ref"A3", CellValue.Text(""))
      .put(ref"A4", formula("\"\"", CellValue.Text("")))
      .put(ref"A5", CellValue.RichText(RichText(TextRun("x"), TextRun("y"))))
      .put(ref"A6", CellValue.Text(" "))
      .put(ref"A7", CellValue.Number(0))
      .put(ref"A8", CellValue.Bool(false))
      .put(ref"A9", formula("NA()", CellValue.Error(CellError.NA)))
      .put(ref"A10", formula("A2", CellValue.Number(0)))
      .put(ref"A11", formula("IF(1,\"\",\"x\")", CellValue.Text("")))
      .put(ref"A12", CellValue.Text("abc"))
      .put(ref"A14", CellValue.Bool(true))
      .put(ref"B1", CellValue.Text(""))
      .put(ref"B2", formula("\"\"", CellValue.Text("")))
      .put(ref"B3", CellValue.Number(5))
      .put(ref"B5", CellValue.Text("t"))
      .put(ref"C1", formula("\"\"", CellValue.Text("")))
      .put(ref"C2", CellValue.Text(""))

  private def workbook(cached: Boolean): Workbook =
    Workbook(Vector(Sheet("Calc"), data(cached)))

  private def check(formula: String, expected: CellValue, cached: Boolean = true)(implicit
    loc: munit.Location
  ): Unit =
    assertEquals(
      workbook(cached).evaluateFormula(formula, SheetName.unsafe("Calc")),
      Right(expected),
      s"$formula (cached = $cached)"
    )

  private def num(n: Long): CellValue = CellValue.Number(BigDecimal(n))

  private val value = CellValue.Error(CellError.Value)

  // ===== Item 1: COUNTBLANK counts a cell holding empty text =====

  test("#731: COUNTBLANK counts the constant \"\" and a formula's \"\" as blank") {
    for cached <- List(true, false) do
      check("=COUNTBLANK(Data!A1:A14)", num(5), cached)
      check("=COUNTBLANK(Data!A3)", num(1), cached)
      check("=COUNTBLANK(Data!A4)", num(1), cached)
      check("=COUNTBLANK(Data!A11)", num(1), cached)
      check("=COUNTBLANK(Data!A1:B5)", num(6), cached)
  }

  test("#731: COUNTBLANK does not count a space, 0, FALSE, an error, a formula's 0 or rich text") {
    for cached <- List(true, false) do
      check("=COUNTBLANK(Data!A6)", num(0), cached)
      check("=COUNTBLANK(Data!A7:A8)", num(0), cached)
      check("=COUNTBLANK(Data!A9)", num(0), cached)
      check("=COUNTBLANK(Data!A10)", num(0), cached)
      check("=COUNTBLANK(Data!A5)", num(0), cached)
  }

  test("#731: a whole column or row adds its empty-text cells to the blanks past the used range") {
    check("=COUNTBLANK(Data!C:C)", num(1048576))
    check("=COUNTBLANK(Data!A:A)", num(1048567))
    check("=COUNTBLANK(Data!1:2)", num(32767))
  }

  test("#731: empty text reached through an intersection or a returned reference is blank") {
    check("=COUNTBLANK(Data!A3:A4 Data!A4:A5)", num(1))
    check("=COUNTBLANK(IF(TRUE,Data!A1:A14,Data!B1:B5))", num(5))
    check("=COUNTBLANK(IF(TRUE,Data!A3,Data!A1))", num(1))
    check("=COUNTBLANK(CHOOSE(1,Data!A1:A14,Data!B1:B5))", num(5))
    check("=COUNTBLANK(INDEX(Data!A1:B14,0,1))", num(5))
    check("=COUNTBLANK(OFFSET(Data!A1,0,0,14,1))", num(5))
    check("=COUNTBLANK(INDIRECT(\"Data!A1:A14\"))", num(5))
    check("=COUNTBLANK(LET(x,Data!A1:A14,x))", num(5))
    check("=COUNTBLANK(XLOOKUP(1,Data!A1:A14,Data!B1:B14))", num(1))
  }

  test("#731: COUNTA keeps counting \"\" cells") {
    check("=COUNTA(Data!A1:A14)", num(12))
    check("=COUNTA(Data!A3,Data!A4)", num(2))
    check("=COUNTA(Data!C:C)", num(2))
    check("=COUNTIF(Data!A1:A14,\"\")", num(5))
  }

  test("#731: rich text with no characters is blank, like the text \"\"") {
    // Excel will not load a cell whose rich text has only empty runs, so this has no Excel result
    val sheet = Sheet("Data").put(ref"A1", CellValue.RichText(RichText(TextRun(""))))
    assertEquals(sheet.evaluateFormula("=COUNTBLANK(A1:A2)"), Right(num(2)))
  }

  test("#731: a sheet recalculation counts the \"\" of formulas it has not cached yet") {
    val sheet = data(cached = false).put(ref"E1", CellValue.Formula("COUNTBLANK(A1:A14)", None))
    assertEquals(sheet.evaluateWithDependencyCheck().map(_.get(ref"E1")), Right(Some(num(5))))
  }

  // ===== Item 3: an argument that is not a reference is #VALUE! =====

  test("#731: COUNTBLANK of an array is #VALUE!") {
    check("=COUNTBLANK(IF({TRUE,FALSE},Data!A1:A3,Data!B1:B3))", value)
    check("=SUM(COUNTBLANK(IF({TRUE,FALSE},Data!A1:A3,Data!B1:B3)))", value)
    check("=COUNTBLANK(IF({TRUE;FALSE},Data!A2:A3,Data!B2:B3))", value)
    check("=COUNTBLANK(IF(Data!A1:A14=\"\",\"\",Data!A1:A14))", value)
    check("=COUNTBLANK(IF({TRUE},Data!A1:A14,Data!B1:B14))", value)
    check("=COUNTBLANK(IF(TRUE,Data!A1:A14&\"\",\"\"))", value)
    check("=COUNTBLANK(IF(TRUE,{1,\"\"},{2,3}))", value)
    check("=COUNTBLANK(IF({TRUE,FALSE},Data!A3,Data!A1))", value)
  }

  test("#731: COUNTBLANK of a scalar value is #VALUE!") {
    check("=COUNTBLANK(IF(TRUE,\"\",\"x\"))", value)
    check("=COUNTBLANK(IF(TRUE,5,6))", value)
    // Excel will not load these (its parser requires a reference); #VALUE! follows the same rule
    check("=COUNTBLANK(5)", value)
    check("=COUNTBLANK(\"\")", value)
    check("=COUNTBLANK({1,\"\"})", value)
    check("=COUNTBLANK(Data!A1:A3+0)", value)
    check("=COUNTBLANK(Data!A1:A14&\"\")", value)
    check("=COUNTBLANK(TRANSPOSE(Data!A1:A3))", value)
  }
