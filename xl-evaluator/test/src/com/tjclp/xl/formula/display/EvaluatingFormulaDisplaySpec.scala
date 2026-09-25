package com.tjclp.xl.formula.display

import com.tjclp.xl.*
import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cells.{Cell, CellError, CellValue, FormulaKind}
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.conversions.given
import com.tjclp.xl.display.{DisplayConversions, ExcelInterpolator, FormulaDisplayStrategy}
import com.tjclp.xl.formula.Clock
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.numfmt.NumFmt
import com.tjclp.xl.unsafe.*

import munit.FunSuite

import java.time.LocalDate

class EvaluatingFormulaDisplaySpec extends FunSuite:

  // ========== Evaluating Strategy Tests ==========

  test("Evaluating strategy evaluates simple formula") {
    val sheet = Sheet(name = SheetName.unsafe("Test"))
      .put(ref"A1", 100)
      .put(ref"A2", 200)
      .unsafe

    given FormulaDisplayStrategy = EvaluatingFormulaDisplay.evaluating
    val result = summon[FormulaDisplayStrategy].format("=A1+A2", sheet)
    assertEquals(result, "300")
  }

  test("Evaluating strategy evaluates SUM formula") {
    val sheet = Sheet(name = SheetName.unsafe("Test"))
      .put(ref"A1", 10)
      .put(ref"A2", 20)
      .put(ref"A3", 30)
      .unsafe

    given FormulaDisplayStrategy = EvaluatingFormulaDisplay.evaluating
    val result = summon[FormulaDisplayStrategy].format("=SUM(A1:A3)", sheet)
    assertEquals(result, "60")
  }

  test("Evaluating strategy formats evaluated result with inferred NumFmt") {
    val sheet = Sheet(name = SheetName.unsafe("Test"))
      .put(ref"A1", 1)
      .put(ref"A2", 2)
      .unsafe

    given FormulaDisplayStrategy = EvaluatingFormulaDisplay.evaluating
    val result = summon[FormulaDisplayStrategy].format("=A1/A2", sheet)
    // Result is 0.5, should be formatted as percent (heuristic)
    assert(result.contains("%") || result == "0.5")
  }

  test("Evaluating strategy falls back to raw formula on error") {
    val sheet = Sheet(name = SheetName.unsafe("Test"))

    given FormulaDisplayStrategy = EvaluatingFormulaDisplay.evaluating
    val result = summon[FormulaDisplayStrategy].format("=UNKNOWN_FUNCTION(A1)", sheet)
    assertEquals(result, "=UNKNOWN_FUNCTION(A1)")
  }

  test("Evaluating strategy handles division by zero") {
    val sheet = Sheet(name = SheetName.unsafe("Test"))
      .put(ref"A1", 100)
      .put(ref"A2", 0)
      .unsafe

    given FormulaDisplayStrategy = EvaluatingFormulaDisplay.evaluating
    val result = summon[FormulaDisplayStrategy].format("=A1/A2", sheet)
    // GH-344: division by zero evaluates to Excel's #DIV/0! error VALUE, displayed as Excel
    // shows it (was a fallback to the raw formula text)
    assertEquals(result, "#DIV/0!")
  }

  // ========== Excel Interpolator with Evaluation ==========

  test("excel interpolator evaluates SUM formula") {
    import ExcelInterpolator.*
    import DisplayConversions.given

    given Sheet = Sheet(name = SheetName.unsafe("Test"))
      .put(ref"A1", 100)
      .put(ref"A2", 200)
      .put(ref"B1", CellValue.Formula("=SUM(A1:A2)"))
      .unsafe

    given FormulaDisplayStrategy = EvaluatingFormulaDisplay.evaluating

    val result = excel"Total: ${ref"B1"}"
    assertEquals(result, "Total: 300")
  }

  test("excel interpolator evaluates and formats percent formula") {
    import ExcelInterpolator.*
    import DisplayConversions.given

    given Sheet = Sheet(name = SheetName.unsafe("Test"))
      .put(ref"A1", 100)
      .put(ref"A2", 200)
      .put(ref"B1", CellValue.Formula("=A1/A2"))
      .style(ref"B1", CellStyle.default.withNumFmt(NumFmt.Percent))
      .unsafe

    given FormulaDisplayStrategy = EvaluatingFormulaDisplay.evaluating

    // Re-pinned by GH-410: the inferred PercentDecimal is "0.00%" (ECMA-376 id 10), which
    // renders two forced decimals — 50.00%, not the old hand-rolled single-decimal 50.0%.
    val result = excel"Ratio: ${ref"B1"}"
    assertEquals(result, "Ratio: 50.00%")
  }

  test("excel interpolator evaluates AVERAGE formula") {
    import ExcelInterpolator.*
    import DisplayConversions.given

    given Sheet = Sheet(name = SheetName.unsafe("Test"))
      .put(ref"A1", 10)
      .put(ref"A2", 20)
      .put(ref"A3", 30)
      .put(ref"B1", CellValue.Formula("=AVERAGE(A1:A3)"))
      .unsafe

    given FormulaDisplayStrategy = EvaluatingFormulaDisplay.evaluating

    val result = excel"Average: ${ref"B1"}"
    // AVERAGE evaluates correctly - just check it's numeric
    assert(result.startsWith("Average: "))
    assert(result.contains("20") || result.contains("(60,3)")) // Either formatted or raw tuple
  }

  test("excel interpolator evaluates IF formula") {
    import ExcelInterpolator.*
    import DisplayConversions.given

    given Sheet = Sheet(name = SheetName.unsafe("Test"))
      .put(ref"A1", 500)
      .put(ref"B1", CellValue.Formula("=IF(A1>400, \"High\", \"Low\")"))
      .unsafe

    given FormulaDisplayStrategy = EvaluatingFormulaDisplay.evaluating

    val result = excel"Status: ${ref"B1"}"
    assertEquals(result, "Status: High")
  }

  test("excel interpolator evaluates date formula") {
    import ExcelInterpolator.*
    import DisplayConversions.given

    val clock = Clock.fixedDate(LocalDate.of(2025, 11, 21))
    given FormulaDisplayStrategy = EvaluatingFormulaDisplay.withClock(clock)

    given Sheet = Sheet(name = SheetName.unsafe("Test"))
      .put(ref"A1", CellValue.Formula("=TODAY()"))

    val result = excel"Date: ${ref"A1"}"
    assert(result.contains("11/21/25"))
  }

  test("excel interpolator with mixed evaluated and constant values") {
    import ExcelInterpolator.*
    import DisplayConversions.given

    given Sheet = Sheet(name = SheetName.unsafe("Test"))
      .put(ref"A1", 100)
      .put(ref"A2", 200)
      .put(ref"B1", CellValue.Formula("=A1+A2"))
      .put(ref"C1", "Total")
      .unsafe

    given FormulaDisplayStrategy = EvaluatingFormulaDisplay.evaluating

    val result = excel"${ref"C1"}: ${ref"B1"}"
    assertEquals(result, "Total: 300")
  }

  // ========== Cached value display (GH-275) ==========
  // After wb.recalculate() caches values into CellValue.Formula cells, display must
  // prefer the cached value; local re-evaluation (which has no workbook context and
  // fails on cross-sheet refs) is only the fallback for uncached formulas.

  test("GH-275: displayCell shows cached cross-sheet value after recalculate, not formula text") {
    import com.tjclp.xl.display.syntax.*
    import com.tjclp.xl.workbooks.Workbook

    val sheet1 = Sheet(name = SheetName.unsafe("Sheet1"))
      .put(ref"A1", CellValue.Number(BigDecimal(10)))
    val sheet2 = Sheet(name = SheetName.unsafe("Sheet2"))
      .put(ref"A1", CellValue.Formula("='Sheet1'!A1*2"))

    val result = Workbook(Vector(sheet1, sheet2)).recalculate()
    assert(result.isClean, s"recalculate must be clean: ${result.errors.map(_.render)}")

    val recalced = result.workbook.sheets.filter(_.name.value == "Sheet2") match
      case Vector(s) => s
      case other => fail(s"expected exactly one Sheet2, got $other")

    given FormulaDisplayStrategy = EvaluatingFormulaDisplay.evaluating
    assertEquals(recalced.displayCell(ref"A1").formatted, "20")
  }

  test("GH-275: excel interpolator shows cached cross-sheet value after recalculate") {
    import ExcelInterpolator.*
    import DisplayConversions.given
    import com.tjclp.xl.workbooks.Workbook

    val sheet1 = Sheet(name = SheetName.unsafe("Sheet1"))
      .put(ref"A1", CellValue.Number(BigDecimal(10)))
    val sheet2 = Sheet(name = SheetName.unsafe("Sheet2"))
      .put(ref"A1", CellValue.Formula("='Sheet1'!A1*2"))

    val result = Workbook(Vector(sheet1, sheet2)).recalculate()
    val recalced = result.workbook.sheets.filter(_.name.value == "Sheet2") match
      case Vector(s) => s
      case other => fail(s"expected exactly one Sheet2, got $other")

    given Sheet = recalced
    given FormulaDisplayStrategy = EvaluatingFormulaDisplay.evaluating
    assertEquals(excel"Value: ${ref"A1"}", "Value: 20")
  }

  test("GH-275: cached value formats via the cell's explicit numFmt") {
    import com.tjclp.xl.display.syntax.*

    // Cross-sheet expr so local evaluation cannot accidentally succeed: only the
    // cached value can produce a number here.
    val sheet = Sheet(name = SheetName.unsafe("Test"))
      .put(
        ref"B1",
        CellValue.Formula("='Other'!A1/'Other'!A2", Some(CellValue.Number(BigDecimal("0.5"))))
      )
      .style(ref"B1", CellStyle.default.withNumFmt(NumFmt.Percent))
      .unsafe

    given FormulaDisplayStrategy = EvaluatingFormulaDisplay.evaluating
    assertEquals(sheet.displayCell(ref"B1").formatted, "50%")
  }

  test("GH-275: cached DateTime under General numFmt displays as a date") {
    import com.tjclp.xl.display.syntax.*

    val dt = LocalDate.of(2025, 11, 21).atStartOfDay()
    val sheet = Sheet(name = SheetName.unsafe("Test"))
      .put(ref"A1", CellValue.Formula("='Other'!A1", Some(CellValue.DateTime(dt))))

    given FormulaDisplayStrategy = EvaluatingFormulaDisplay.evaluating
    val displayed = sheet.displayCell(ref"A1").formatted
    assert(displayed.contains("11/21/25"), s"expected date-formatted display, got: $displayed")
  }

  test("GH-282: default (non-evaluating) strategy also prefers cached values") {
    import com.tjclp.xl.display.syntax.*

    // GH-275 gave the evaluating strategy cache preference; GH-282 extends it to the
    // default strategy so xl-core-only consumers see meaningful display too.
    val sheet = Sheet(name = SheetName.unsafe("Test"))
      .put(ref"A1", CellValue.Formula("='Sheet1'!A1*2", Some(CellValue.Number(BigDecimal(20)))))

    given FormulaDisplayStrategy = FormulaDisplayStrategy.default
    assertEquals(sheet.displayCell(ref"A1").formatted, "20")
  }

  // ========== Uncached cells display as the cell they are ==========
  // An uncached formula cell is displayed with the value evaluateCell and recalculate() give it:
  // a plain formula is Excel's legacy formula at its position (references implicitly intersected),
  // an array-formula record evaluates as an array. Only positionless format is a dynamic array.

  /** A1:A10 and B1:B10 hold 1..10. */
  private val columns: Sheet =
    (1 to 10).foldLeft(Sheet(name = SheetName.unsafe("Test"))) { (sheet, i) =>
      sheet
        .put(ARef.from1(1, i), CellValue.Number(BigDecimal(i)))
        .put(ARef.from1(2, i), CellValue.Number(BigDecimal(i)))
    }

  private def evaluated(sheet: Sheet, at: ARef): CellValue =
    import com.tjclp.xl.formula.eval.SheetEvaluator.*
    sheet.evaluateCell(at) match
      case Right(value) => value
      case Left(error) => fail(s"evaluateCell($at) failed: $error")

  test("an uncached plain cell displays its value at its position, as evaluateCell gives it") {
    import com.tjclp.xl.display.syntax.*

    val sheet = columns
      .put(ref"D5", CellValue.Formula("SUM(A1:A10*B1:B10)"))
      .put(ref"C8", CellValue.Formula("IF(A1:A10>5,\"big\",\"small\")"))

    given FormulaDisplayStrategy = EvaluatingFormulaDisplay.evaluating
    // row 5 intersects both columns at 5: 5*5; row 8 at 8, which is > 5
    assertEquals(evaluated(sheet, ref"D5"), CellValue.Number(BigDecimal(25)))
    assertEquals(sheet.displayCell(ref"D5").formatted, "25")
    assertEquals(evaluated(sheet, ref"C8"), CellValue.Text("big"))
    assertEquals(sheet.displayCell(ref"C8").formatted, "big")
  }

  test("the excel interpolator displays an uncached plain cell at its position") {
    import ExcelInterpolator.*
    import DisplayConversions.given

    given Sheet = columns
      .put(ref"D5", CellValue.Formula("SUM(A1:A10*B1:B10)"))
      .put(ref"H20", CellValue.Formula("SUM(A1:A10*2)"))

    given FormulaDisplayStrategy = EvaluatingFormulaDisplay.evaluating
    assertEquals(excel"${ref"D5"}", "25")
    // row 20 lies outside A1:A10, so the legacy formula's intersection is #VALUE!
    assertEquals(evaluated(summon[Sheet], ref"H20"), CellValue.Error(CellError.Value))
    assertEquals(excel"${ref"H20"}", "#VALUE!")
    // a Cell the sheet does not hold evaluates at its own position
    assertEquals(excel"${Cell(ref"E5", CellValue.Formula("SUM(A1:A10*B1:B10)"))}", "25")
  }

  test("an uncached array-formula record displays its array value") {
    import ExcelInterpolator.*
    import DisplayConversions.given
    import com.tjclp.xl.display.syntax.*

    def cse(range: CellRange): CellValue.Formula =
      CellValue.Formula("SUM(A1:A10*B1:B10)", None, FormulaKind.ArrayFormula(range))
    given Sheet = columns.put(ref"D5", cse(ref"D5:D5"))

    given FormulaDisplayStrategy = EvaluatingFormulaDisplay.evaluating
    // {=SUM(A1:A10*B1:B10)} multiplies the columns element-wise: 1² + ... + 10²
    assertEquals(evaluated(summon[Sheet], ref"D5"), CellValue.Number(BigDecimal(385)))
    assertEquals(summon[Sheet].displayCell(ref"D5").formatted, "385")
    assertEquals(excel"${Cell(ref"F9", cse(ref"F9:F9"))}", "385")
  }

  test("a cached cell displays its cache without evaluating") {
    import com.tjclp.xl.display.syntax.*

    // the cache disagrees with what evaluation would give (25), so only the cache can show 999
    val sheet = columns.put(
      ref"D5",
      CellValue.Formula("SUM(A1:A10*B1:B10)", Some(CellValue.Number(BigDecimal(999))))
    )

    given FormulaDisplayStrategy = EvaluatingFormulaDisplay.evaluating
    assertEquals(sheet.displayCell(ref"D5").formatted, "999")
  }

  test("positionless format still evaluates as a dynamic array") {
    given FormulaDisplayStrategy = EvaluatingFormulaDisplay.evaluating
    val strategy = summon[FormulaDisplayStrategy]
    // no cell to intersect with: a new Excel 365 cell's value, as sheet.evaluateFormula gives it
    assertEquals(strategy.format("=SUM(A1:A10*B1:B10)", columns), "385")
    assertEquals(strategy.formatCached("SUM(A1:A10*B1:B10)", None, NumFmt.General, columns), "385")
  }

  // ========== Edge Cases ==========

  test("Evaluating strategy handles circular reference gracefully") {
    val sheet = Sheet(name = SheetName.unsafe("Test"))
      .put(ref"A1", CellValue.Formula("=B1+1"))
      .put(ref"B1", CellValue.Formula("=A1+1"))

    given FormulaDisplayStrategy = EvaluatingFormulaDisplay.evaluating
    val result = summon[FormulaDisplayStrategy].format("=A1+1", sheet)
    // Should fallback to raw formula
    assertEquals(result, "=A1+1")
  }

  test("Evaluating strategy handles missing cell reference") {
    val sheet = Sheet(name = SheetName.unsafe("Test"))

    given FormulaDisplayStrategy = EvaluatingFormulaDisplay.evaluating
    val result = summon[FormulaDisplayStrategy].format("=A1+100", sheet)
    // Should fallback or show error
    assertEquals(result.nonEmpty, true)
  }
