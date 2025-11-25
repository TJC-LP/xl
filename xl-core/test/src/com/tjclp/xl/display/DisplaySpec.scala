package com.tjclp.xl.display

import com.tjclp.xl.*
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{Cell, CellValue, CellError}
import com.tjclp.xl.conversions.given  // For put(ARef, value)
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.numfmt.NumFmt
import com.tjclp.xl.unsafe.*  // For .unsafe extension

import munit.FunSuite

import java.time.LocalDateTime

class DisplaySpec extends FunSuite:

  // ========== NumFmtFormatter Tests ==========

  test("formatValue - Currency format") {
    val value = CellValue.Number(BigDecimal("1234.56"))
    val result = NumFmtFormatter.formatValue(value, NumFmt.Currency)
    assertEquals(result, "$1,234.56")
  }

  test("formatValue - Percent format") {
    val value = CellValue.Number(BigDecimal("0.15"))
    val result = NumFmtFormatter.formatValue(value, NumFmt.Percent)
    assertEquals(result, "15%")
  }

  test("formatValue - PercentDecimal format") {
    val value = CellValue.Number(BigDecimal("0.156"))
    val result = NumFmtFormatter.formatValue(value, NumFmt.PercentDecimal)
    assertEquals(result, "15.6%")
  }

  test("formatValue - ThousandsSeparator format") {
    val value = CellValue.Number(BigDecimal("1234567"))
    val result = NumFmtFormatter.formatValue(value, NumFmt.ThousandsSeparator)
    assertEquals(result, "1,234,567")
  }

  test("formatValue - ThousandsDecimal format") {
    val value = CellValue.Number(BigDecimal("1234.5"))
    val result = NumFmtFormatter.formatValue(value, NumFmt.ThousandsDecimal)
    assertEquals(result, "1,234.50")
  }

  test("formatValue - Decimal format") {
    val value = CellValue.Number(BigDecimal("123.456"))
    val result = NumFmtFormatter.formatValue(value, NumFmt.Decimal)
    assert(result.startsWith("123.4"))  // At least 2 decimal places
  }

  test("formatValue - Integer format") {
    val value = CellValue.Number(BigDecimal("123.7"))
    val result = NumFmtFormatter.formatValue(value, NumFmt.Integer)
    assertEquals(result, "124")  // Rounded
  }

  test("formatValue - General format (whole number)") {
    val value = CellValue.Number(BigDecimal("100"))
    val result = NumFmtFormatter.formatValue(value, NumFmt.General)
    assertEquals(result, "100")
  }

  test("formatValue - General format (decimal)") {
    val value = CellValue.Number(BigDecimal("123.45"))
    val result = NumFmtFormatter.formatValue(value, NumFmt.General)
    assertEquals(result, "123.45")
  }

  test("formatValue - Text value") {
    val value = CellValue.Text("Hello World")
    val result = NumFmtFormatter.formatValue(value, NumFmt.General)
    assertEquals(result, "Hello World")
  }

  test("formatValue - Boolean true") {
    val value = CellValue.Bool(true)
    val result = NumFmtFormatter.formatValue(value, NumFmt.General)
    assertEquals(result, "TRUE")
  }

  test("formatValue - Boolean false") {
    val value = CellValue.Bool(false)
    val result = NumFmtFormatter.formatValue(value, NumFmt.General)
    assertEquals(result, "FALSE")
  }

  test("formatValue - Empty cell") {
    val value = CellValue.Empty
    val result = NumFmtFormatter.formatValue(value, NumFmt.General)
    assertEquals(result, "")
  }

  test("formatValue - Error cell") {
    val value = CellValue.Error(CellError.Div0)
    val result = NumFmtFormatter.formatValue(value, NumFmt.General)
    assertEquals(result, "#DIV/0!")
  }

  test("formatValue - Error NA") {
    val value = CellValue.Error(CellError.NA)
    val result = NumFmtFormatter.formatValue(value, NumFmt.General)
    assertEquals(result, "#N/A")
  }

  test("formatValue - DateTime with DateTime format") {
    val dt = LocalDateTime.of(2025, 11, 21, 14, 30)
    val value = CellValue.DateTime(dt)
    val result = NumFmtFormatter.formatValue(value, NumFmt.DateTime)
    assertEquals(result, "11/21/25 14:30")
  }

  test("formatValue - DateTime with Date format") {
    val dt = LocalDateTime.of(2025, 11, 21, 14, 30)
    val value = CellValue.DateTime(dt)
    val result = NumFmtFormatter.formatValue(value, NumFmt.Date)
    assertEquals(result, "11/21/25")
  }

  // ========== DisplayWrapper Tests ==========

  test("DisplayWrapper toString returns formatted string") {
    val wrapper = DisplayWrapper("$1,000.00")
    assertEquals(wrapper.toString, "$1,000.00")
  }

  test("DisplayWrapper in string interpolation") {
    val wrapper = DisplayWrapper("60%")
    val result = s"Value: $wrapper"
    assertEquals(result, "Value: 60%")
  }

  // ========== FormulaDisplayStrategy Tests ==========

  test("Default strategy shows raw formula") {
    val sheet = Sheet(name = SheetName.unsafe("Test"))
    val strategy = FormulaDisplayStrategy.default
    val result = strategy.format("=SUM(A1:A10)", sheet)
    assertEquals(result, "=SUM(A1:A10)")
  }

  test("Default strategy handles formula without = prefix") {
    val sheet = Sheet(name = SheetName.unsafe("Test"))
    val strategy = FormulaDisplayStrategy.default
    val result = strategy.format("SUM(A1:A10)", sheet)
    assertEquals(result, "=SUM(A1:A10)")
  }

  // ========== DisplayConversions Tests ==========

  test("ARef conversion with given Sheet") {
    import DisplayConversions.given

    given Sheet = Sheet(name = SheetName.unsafe("Test"))
      .put(ref"A1", BigDecimal("1000"))
      .style(ref"A1", CellStyle.default.withNumFmt(NumFmt.Currency))
      .unsafe

    given FormulaDisplayStrategy = FormulaDisplayStrategy.default

    val conv = summon[Conversion[ARef, DisplayWrapper]]
    val result = conv.apply(ref"A1")
    assertEquals(result.formatted, "$1,000.00")
  }

  test("Cell conversion with given Sheet") {
    import DisplayConversions.given

    val mySheet = Sheet(name = SheetName.unsafe("Test"))
      .put(ref"A1", BigDecimal("0.6"))
      .style(ref"A1", CellStyle.default.withNumFmt(NumFmt.Percent))
      .unsafe

    given Sheet = mySheet
    given FormulaDisplayStrategy = FormulaDisplayStrategy.default

    val cell = mySheet(ref"A1")
    val conv = summon[Conversion[Cell, DisplayWrapper]]
    val result = conv.apply(cell)
    assertEquals(result.formatted, "60%")
  }

  // ========== Excel Interpolator Tests ==========

  test("excel interpolator with currency") {
    import ExcelInterpolator.*
    import DisplayConversions.given

    given Sheet = Sheet(name = SheetName.unsafe("Test"))
      .put(ref"A1", BigDecimal("1000000"))
      .style(ref"A1", CellStyle.default.withNumFmt(NumFmt.Currency))
      .unsafe

    given FormulaDisplayStrategy = FormulaDisplayStrategy.default

    val result = excel"Revenue: ${ref"A1"}"
    assertEquals(result, "Revenue: $1,000,000.00")
  }

  test("excel interpolator with percent") {
    import ExcelInterpolator.*
    import DisplayConversions.given

    given Sheet = Sheet(name = SheetName.unsafe("Test"))
      .put(ref"B1", BigDecimal("0.75"))
      .style(ref"B1", CellStyle.default.withNumFmt(NumFmt.Percent))
      .unsafe

    given FormulaDisplayStrategy = FormulaDisplayStrategy.default

    val result = excel"Margin: ${ref"B1"}"
    assertEquals(result, "Margin: 75%")
  }

  test("excel interpolator with raw formula (default strategy)") {
    import ExcelInterpolator.*
    import DisplayConversions.given

    given Sheet = Sheet(name = SheetName.unsafe("Test"))
      .put(ref"C1", CellValue.Formula("=A1+B1"))

    given FormulaDisplayStrategy = FormulaDisplayStrategy.default

    val result = excel"Formula: ${ref"C1"}"
    assertEquals(result, "Formula: =A1+B1")
  }

  test("excel interpolator with mixed values") {
    import ExcelInterpolator.*
    import DisplayConversions.given

    given Sheet = Sheet(name = SheetName.unsafe("Test"))
      .put(ref"A1", BigDecimal("1000"))
      .put(ref"A2", "Product")
      .style(ref"A1", CellStyle.default.withNumFmt(NumFmt.Currency))
      .unsafe

    given FormulaDisplayStrategy = FormulaDisplayStrategy.default

    val quantity = 5
    val result = excel"${ref"A2"}: $quantity units @ ${ref"A1"} each"
    assertEquals(result, "Product: 5 units @ $1,000.00 each")
  }

  // ========== Syntax Extension Tests ==========

  test("sheet.display() returns formatted value") {
    import com.tjclp.xl.display.syntax.*
    import DisplayConversions.given

    val mySheet = Sheet(name = SheetName.unsafe("Test"))
      .put(ref"A1", BigDecimal("0.85"))
      .style(ref"A1", CellStyle.default.withNumFmt(NumFmt.Percent))
      .unsafe

    given FormulaDisplayStrategy = FormulaDisplayStrategy.default

    val result = mySheet.displayCell(ref"A1")
    assertEquals(result.formatted, "85%")
  }

  test("sheet.displayFormula() shows raw formula") {
    import com.tjclp.xl.display.syntax.*
    import DisplayConversions.given

    val mySheet = Sheet(name = SheetName.unsafe("Test"))
      .put(ref"B1", CellValue.Formula("=SUM(A1:A10)"))

    given FormulaDisplayStrategy = FormulaDisplayStrategy.default

    val result = mySheet.displayFormula(ref"B1")
    assertEquals(result, "=SUM(A1:A10)")
  }

  test("sheet.displayFormula() shows formatted value for non-formulas") {
    import com.tjclp.xl.display.syntax.*
    import DisplayConversions.given

    val mySheet = Sheet(name = SheetName.unsafe("Test"))
      .put(ref"A1", BigDecimal("100"))

    given FormulaDisplayStrategy = FormulaDisplayStrategy.default

    val result = mySheet.displayFormula(ref"A1")
    assertEquals(result, "100")
  }
