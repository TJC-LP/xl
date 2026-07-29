package com.tjclp.xl.display

import com.tjclp.xl.*
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{Cell, CellValue, CellError}
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.conversions.given // For put(ARef, value)
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.numfmt.NumFmt
import com.tjclp.xl.unsafe.* // For .unsafe extension

import munit.ScalaCheckSuite
import org.scalacheck.Gen
import org.scalacheck.Prop.forAll

import java.time.LocalDateTime

class DisplaySpec extends ScalaCheckSuite:

  // ========== NumFmtFormatter Tests ==========

  /** Every built-in (non-Custom) NumFmt variant. */
  private val builtInVariants: List[NumFmt] = List(
    NumFmt.General,
    NumFmt.Integer,
    NumFmt.Decimal,
    NumFmt.ThousandsSeparator,
    NumFmt.ThousandsDecimal,
    NumFmt.Currency,
    NumFmt.Percent,
    NumFmt.PercentDecimal,
    NumFmt.Scientific,
    NumFmt.Fraction,
    NumFmt.Date,
    NumFmt.DateTime,
    NumFmt.Time,
    NumFmt.Text
  )

  private val genNumeric: Gen[BigDecimal] = Gen
    .oneOf(
      Gen.chooseNum(-1e6, 1e6), // covers negative serials and general magnitudes
      Gen.chooseNum(-1.0, 1.0), // percent/fraction territory
      Gen.chooseNum(0.0, 60000.0) // Excel date-serial range for the calendar variants
    )
    .map(d => BigDecimal(java.lang.Double.toString(d)))

  private val genDateTime: Gen[LocalDateTime] = for
    day <- Gen.chooseNum(0, 59999)
    hour <- Gen.chooseNum(0, 23)
    minute <- Gen.chooseNum(0, 59)
    second <- Gen.chooseNum(0, 59)
  yield LocalDateTime.of(1900, 1, 1, hour, minute, second).plusDays(day.toLong)

  property("built-in enum arms render numbers exactly like their format codes (GH-410)") {
    // The enum arm and a file-declared Custom carrying NumFmt.formatCode(fmt) are the same
    // format per ECMA-376; both must render through FormatCodeParser identically.
    forAll(genNumeric) { (n: BigDecimal) =>
      builtInVariants.foreach { fmt =>
        assertEquals(
          NumFmtFormatter.formatValue(CellValue.Number(n), fmt),
          NumFmtFormatter.formatValue(CellValue.Number(n), NumFmt.Custom(NumFmt.formatCode(fmt))),
          s"enum arm diverged from its own format code for $fmt on $n"
        )
      }
      true
    }
  }

  property("built-in enum arms render DateTime values exactly like their format codes (GH-410)") {
    forAll(genDateTime) { (dt: LocalDateTime) =>
      builtInVariants.foreach { fmt =>
        assertEquals(
          NumFmtFormatter.formatValue(CellValue.DateTime(dt), fmt),
          NumFmtFormatter
            .formatValue(CellValue.DateTime(dt), NumFmt.Custom(NumFmt.formatCode(fmt))),
          s"enum arm diverged from its own format code for $fmt on $dt"
        )
      }
      true
    }
  }

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
    // Re-pinned by GH-410: PercentDecimal is "0.00%" (ECMA-376 id 10) — two forced decimal
    // placeholders, so Excel renders 15.60%, not the old enum arm's single-decimal 15.6%.
    val value = CellValue.Number(BigDecimal("0.156"))
    val result = NumFmtFormatter.formatValue(value, NumFmt.PercentDecimal)
    assertEquals(result, "15.60%")
  }

  test("formatValue - Custom 0.00% renders identically to PercentDecimal (GH-410)") {
    // GH-404 kept file-declared codes verbatim, which made a declared "0.00%" render correctly
    // ("15.60%") while the PercentDecimal enum arm hand-rolled "15.6%". GH-410 closed that
    // delta: the enum arms now render through FormatCodeParser on NumFmt.formatCode.
    val value = CellValue.Number(BigDecimal("0.156"))
    val result = NumFmtFormatter.formatValue(value, NumFmt.Custom("0.00%"))
    assertEquals(result, "15.60%")
    assertEquals(result, NumFmtFormatter.formatValue(value, NumFmt.PercentDecimal))
  }

  test("formatValue - Scientific format matches its 0.00E+00 code (GH-410)") {
    // ECMA-376 id 11 is "0.00E+00": one mantissa integer digit, two decimals, signed
    // two-digit exponent.
    def sci(s: String): String =
      NumFmtFormatter.formatValue(CellValue.Number(BigDecimal(s)), NumFmt.Scientific)
    assertEquals(sci("1234.5"), "1.23E+03")
    assertEquals(sci("0"), "0.00E+00")
    assertEquals(sci("0.0001234"), "1.23E-04")
    assertEquals(sci("-1234.5"), "-1.23E+03")
  }

  test("formatValue - Decimal rounds HALF_UP on the exact decimal value (GH-410)") {
    // "0.00" of exactly 2.675 is 2.68 in decimal HALF_UP; the old arm detoured through
    // Double where 2.675 is stored as 2.67499..., yielding 2.67.
    val result = NumFmtFormatter.formatValue(CellValue.Number(BigDecimal("2.675")), NumFmt.Decimal)
    assertEquals(result, "2.68")
  }

  test("formatValue - DateTime serial renders 24-hour like its m/d/yy h:mm code (GH-410)") {
    // ECMA-376 §18.8.31: 'h' is the 12-hour clock only when the code contains AM/PM;
    // "m/d/yy h:mm" (id 22) has none, so 14:30 renders as 14:30. The old arm used the
    // Java pattern 'h' (clock-hour 1-12) and showed 2:30.
    val serial = BigDecimal("45982.6041666666666667") // 2025-11-21 14:30
    val result = NumFmtFormatter.formatValue(CellValue.Number(serial), NumFmt.DateTime)
    assertEquals(result, "11/21/25 14:30")
  }

  test("formatValue - Time serial renders 24-hour (GH-410)") {
    val result = NumFmtFormatter.formatValue(CellValue.Number(BigDecimal("0.75")), NumFmt.Time)
    assertEquals(result, "18:00:00")
  }

  test("formatValue - Date format fills ###### for out-of-range serials (GH-410)") {
    // Excel renders negative (pre-1900) serials under date formats as ######; the old arm
    // fabricated an 1899 calendar date.
    val result = NumFmtFormatter.formatValue(CellValue.Number(BigDecimal(-5)), NumFmt.Date)
    assertEquals(result, "######")
  }

  test("formatValue - Text format renders numbers like General (GH-410)") {
    // "@" (id 49) has no numeric section: Excel shows the number in General format. The old
    // arm leaked BigDecimal.toString ("1E+3").
    val result = NumFmtFormatter.formatValue(CellValue.Number(BigDecimal("1E+3")), NumFmt.Text)
    assertEquals(result, "1000")
  }

  test("formatValue - Custom thousands/currency codes match the enum arms (GH-404 parity)") {
    // File-declared codes render via FormatCodeParser (Custom) while the programmatic enum arms
    // use java DecimalFormat; the enum arms round HALF_UP (Excel's half-away-from-zero display
    // rounding — 1234.5 under "#,##0" is "1,235"), keeping both paths in parity (GH-404)
    val values =
      List(BigDecimal("1234567"), BigDecimal("1234.5"), BigDecimal("0.5"), BigDecimal("2.5"))
    values.foreach { n =>
      assertEquals(
        NumFmtFormatter.formatValue(CellValue.Number(n), NumFmt.Custom("#,##0")),
        NumFmtFormatter.formatValue(CellValue.Number(n), NumFmt.ThousandsSeparator),
        s"#,##0 diverged for $n"
      )
      assertEquals(
        NumFmtFormatter.formatValue(CellValue.Number(n), NumFmt.Custom("#,##0.00")),
        NumFmtFormatter.formatValue(CellValue.Number(n), NumFmt.ThousandsDecimal),
        s"#,##0.00 diverged for $n"
      )
      assertEquals(
        NumFmtFormatter.formatValue(CellValue.Number(n), NumFmt.Custom("$#,##0.00")),
        NumFmtFormatter.formatValue(CellValue.Number(n), NumFmt.Currency),
        s"$$#,##0.00 diverged for $n"
      )
    }
    assertEquals(
      NumFmtFormatter.formatValue(CellValue.Number(BigDecimal("1234.5")), NumFmt.Custom("#,##0")),
      "1,235"
    )
  }

  test("formatValue - Custom General renders like the General format (GH-404 delta pin)") {
    // LibreOffice-produced files declare <numFmt numFmtId="164" formatCode="General"/>; since
    // GH-404 those resolve to Custom("General") and must render exactly like NumFmt.General.
    val values = List(BigDecimal("100"), BigDecimal("123.45"), BigDecimal("0.156"))
    values.foreach { n =>
      assertEquals(
        NumFmtFormatter.formatValue(CellValue.Number(n), NumFmt.Custom("General")),
        NumFmtFormatter.formatValue(CellValue.Number(n), NumFmt.General),
        s"Custom(General) diverged from General for $n"
      )
    }
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
    assert(result.startsWith("123.4")) // At least 2 decimal places
  }

  test("formatValue - Integer format") {
    val value = CellValue.Number(BigDecimal("123.7"))
    val result = NumFmtFormatter.formatValue(value, NumFmt.Integer)
    assertEquals(result, "124") // Rounded
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

  test("formatValue - General format (zero)") {
    val value = CellValue.Number(BigDecimal("0"))
    val result = NumFmtFormatter.formatValue(value, NumFmt.General)
    assertEquals(result, "0")
  }

  test("formatValue - General format (9 sig digits stays plain)") {
    val value = CellValue.Number(BigDecimal("0.000123456789"))
    val result = NumFmtFormatter.formatValue(value, NumFmt.General)
    assertEquals(result, "0.000123456789")
  }

  test("formatValue - General format (10 sig digits negative stays plain)") {
    val value = CellValue.Number(BigDecimal("-99999999.99"))
    val result = NumFmtFormatter.formatValue(value, NumFmt.General)
    assertEquals(result, "-99999999.99")
  }

  test("formatValue - General format (exactly 11 sig digits stays plain)") {
    val value = CellValue.Number(BigDecimal("12345678.901"))
    val result = NumFmtFormatter.formatValue(value, NumFmt.General)
    assertEquals(result, "12345678.901")
  }

  test("formatValue - General format (12 sig digits medium rounds to plain)") {
    val value = CellValue.Number(BigDecimal("123456789.012"))
    val result = NumFmtFormatter.formatValue(value, NumFmt.General)
    assertEquals(result, "123456789.01")
  }

  test("formatValue - General format (13 sig digits very small triggers scientific)") {
    val value = CellValue.Number(BigDecimal("0.0000123456789012"))
    val result = NumFmtFormatter.formatValue(value, NumFmt.General)
    assert(result.contains("E"), s"Expected scientific notation, got: $result")
  }

  test("formatValue - General format (15 sig digits very large triggers scientific)") {
    val value = CellValue.Number(BigDecimal("9999999999999.12"))
    val result = NumFmtFormatter.formatValue(value, NumFmt.General)
    assert(result.contains("E"), s"Expected scientific notation, got: $result")
  }

  test("formatValue - General format (0.0001 at threshold stays plain)") {
    val value = CellValue.Number(BigDecimal("0.0001"))
    val result = NumFmtFormatter.formatValue(value, NumFmt.General)
    assertEquals(result, "0.0001")
  }

  test("formatValue - General format (below 1e-4 with >11 sig digits triggers scientific)") {
    val value = CellValue.Number(BigDecimal("0.0000999999999999"))
    val result = NumFmtFormatter.formatValue(value, NumFmt.General)
    assert(result.contains("E"), s"Expected scientific notation, got: $result")
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

  test("Default strategy prefers cached value formatted via numFmt (GH-282)") {
    val sheet = Sheet(name = SheetName.unsafe("Test"))
    val strategy = FormulaDisplayStrategy.default
    val result = strategy.formatCached(
      "='Data'!A1*2",
      Some(CellValue.Number(BigDecimal("1234.5"))),
      NumFmt.Currency,
      sheet
    )
    assertEquals(result, "$1,234.50")
  }

  test("Default strategy falls back to formula text when uncached (GH-282)") {
    val sheet = Sheet(name = SheetName.unsafe("Test"))
    val strategy = FormulaDisplayStrategy.default
    val result = strategy.formatCached("='Data'!A1*2", None, NumFmt.Currency, sheet)
    assertEquals(result, "='Data'!A1*2")
  }

  test("Default strategy displays cached cross-sheet formula through cell display (GH-282)") {
    import DisplayConversions.given

    // Cross-sheet formula cannot be evaluated sheet-locally; the cached value (as written
    // by Excel or Workbook.recalculate()) is the only meaningful display for xl-core users.
    given Sheet = Sheet(name = SheetName.unsafe("Test"))
      .put(ref"C1", CellValue.Formula("='Data'!A1*2", Some(CellValue.Number(BigDecimal("0.6")))))
      .style(ref"C1", CellStyle.default.withNumFmt(NumFmt.Percent))
      .unsafe

    given FormulaDisplayStrategy = FormulaDisplayStrategy.default

    val conv = summon[Conversion[ARef, DisplayWrapper]]
    assertEquals(conv.apply(ref"C1").formatted, "60%")
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

  test(
    "excel interpolator routes zero through positive section of 2-section custom code (GH-254)"
  ) {
    import ExcelInterpolator.*
    import DisplayConversions.given

    given Sheet = Sheet(name = SheetName.unsafe("Test"))
      .put(ref"A1", BigDecimal("0"))
      .style(ref"A1", CellStyle.default.withNumFmt(NumFmt.Custom("0.0%_);(0.0%)")))
      .unsafe

    given FormulaDisplayStrategy = FormulaDisplayStrategy.default

    val result = excel"Rate: ${ref"A1"}"
    assertEquals(result.trim, "Rate: 0.0%") // trailing space from _) spacer
    assert(!result.contains("("), s"Zero must not use the negative section: $result")
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

  test("displayCell routes zero through positive section of 2-section custom code (GH-254)") {
    import com.tjclp.xl.display.syntax.*
    import DisplayConversions.given

    val houseCode = "\"$\"#,##0.0_);(\"$\"#,##0.0)"
    val mySheet = Sheet(name = SheetName.unsafe("Test"))
      .put(ref"A1", BigDecimal("0"))
      .style(ref"A1", CellStyle.default.withNumFmt(NumFmt.Custom(houseCode)))
      .unsafe

    given FormulaDisplayStrategy = FormulaDisplayStrategy.default

    val result = mySheet.displayCell(ref"A1")
    assertEquals(result.formatted.trim, "$0.0") // trailing space from _) spacer
    assert(!result.formatted.contains("("), s"Zero must not use the negative section: $result")
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
      .unsafe

    given FormulaDisplayStrategy = FormulaDisplayStrategy.default

    val result = mySheet.displayFormula(ref"A1")
    assertEquals(result, "100.00") // BigDecimal auto-applies Decimal format
  }
