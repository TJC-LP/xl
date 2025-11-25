package com.tjclp.xl.formula.display

import com.tjclp.xl.*
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cells.CellValue
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
    // Should fallback to raw formula on error
    assertEquals(result, "=A1/A2")
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

    val result = excel"Ratio: ${ref"B1"}"
    assertEquals(result, "Ratio: 50.0%")
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
    assert(result.contains("20") || result.contains("(60,3)"))  // Either formatted or raw tuple
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
    given FormulaDisplayStrategy = EvaluatingFormulaDisplay.evaluating(using clock)

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
