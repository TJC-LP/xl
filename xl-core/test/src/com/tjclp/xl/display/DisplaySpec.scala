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

  test("formatValue - General keyword inside a custom code renders the value (GH-666)") {
    // Weaver dogfood: `General"A"` on 2021 showed "GeneralA"; Excel shows "2021A".
    val n = CellValue.Number(BigDecimal(2021))
    assertEquals(NumFmtFormatter.formatValue(n, NumFmt.Custom("General\"A\"")), "2021A")
    assertEquals(NumFmtFormatter.formatValue(n, NumFmt.Custom("General\"E\"")), "2021E")
    assertEquals(NumFmtFormatter.formatValue(n, NumFmt.Custom("\"FY\"General")), "FY2021")
    // The whole-code path (GH-404) is unchanged
    assertEquals(NumFmtFormatter.formatValue(n, NumFmt.Custom("General")), "2021")
    assertEquals(NumFmtFormatter.formatValue(n, NumFmt.Custom("general")), "2021")
  }

  test("formatValue - thousands-scaling commas divide by 1000 each (GH-666)") {
    // Weaver dogfood: `$#,##0.0,,"mm"` on 1,500,000 showed "$1,500,000.0mm"; Excel/LO "$1.5mm".
    val mm = NumFmt.Custom("$#,##0.0,,\"mm\"")
    assertEquals(NumFmtFormatter.formatValue(CellValue.Number(BigDecimal(1500000)), mm), "$1.5mm")
    val k = NumFmt.Custom("#,##0.0,\"k\"")
    assertEquals(NumFmtFormatter.formatValue(CellValue.Number(BigDecimal(1234567)), k), "1,234.6k")
    val plain = NumFmt.Custom("#,##0,")
    assertEquals(NumFmtFormatter.formatValue(CellValue.Number(BigDecimal(1234567)), plain), "1,235")
  }

  test("formatValue - GH-666 codes are display-only: the stored format code is byte-identical") {
    List("General\"A\"", "\"FY\"General", "$#,##0.0,,\"mm\"", "#,##0,", "#,##0.0,\"k\"", "0.0,,")
      .foreach { code =>
        val fmt = NumFmt.parse(code)
        assertEquals(fmt, NumFmt.Custom(code))
        assertEquals(NumFmt.formatCode(fmt), code)
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

  // ========== Cell-display General: Excel's 11-character rule (GH-672) ==========
  // A General cell shows at most 11 characters, the sign uncounted: plain while the number fits
  // (decimals rounded HALF_UP to the room left, trailing zeros stripped), E notation when the
  // integer part needs 12+ digits or the exponent is -5 or below (the %G small-number switch:
  // 0.0001 plain, 0.00001 is 1E-05). The E form keeps the significant digits that fit in 11
  // characters — 6 with a two-digit exponent (1.23457E+11), 5 with three (1.2346E+100). As in %G
  // the switch is decided on the rounded value: 0.0000999999999999 rounds to 0.0001, plain.
  // NOT the text-conversion rule (generalText: 15 digits / 20 characters, GH-665).

  private def general(s: String): String =
    NumFmtFormatter.formatValue(CellValue.Number(BigDecimal(s)), NumFmt.General)

  test("formatValue - General format (0.000123456789 rounds to 11 characters)") {
    assertEquals(general("0.000123456789"), "0.000123457")
  }

  test("formatValue - General format (10 sig digits negative stays plain)") {
    val value = CellValue.Number(BigDecimal("-99999999.99"))
    val result = NumFmtFormatter.formatValue(value, NumFmt.General)
    assertEquals(result, "-99999999.99")
  }

  test("formatValue - General format (12 characters round to 11: 12345678.901)") {
    assertEquals(general("12345678.901"), "12345678.9")
  }

  test("formatValue - General format (123456789.012 rounds to the integer, zeros stripped)") {
    assertEquals(general("123456789.012"), "123456789")
  }

  test("formatValue - General format (13 sig digits very small triggers scientific)") {
    assertEquals(general("0.0000123456789012"), "1.23457E-05")
  }

  test("formatValue - General format (15 sig digits very large triggers scientific)") {
    assertEquals(general("9999999999999.12"), "1E+13")
  }

  test("formatValue - General format (0.0001 at threshold stays plain)") {
    assertEquals(general("0.0001"), "0.0001")
  }

  test("formatValue - General format (below 1e-4 decides on the rounded value, like %G)") {
    // 9.99999999999E-05 rounds to six significant digits as 1.00000E-04, back in the plain range
    assertEquals(general("0.0000999999999999"), "0.0001")
    assertEquals(general("0.0000999994"), "9.99994E-05")
  }

  test("General cell display: the issue rows (GH-672)") {
    // Dogfood E3: 0.000012345 showed its full expansion, 123456789012 all twelve digits
    assertEquals(general("0.000012345"), "1.2345E-05")
    assertEquals(general("123456789012"), "1.23457E+11")
    assertEquals(general("-123456789012"), "-1.23457E+11")
    assertEquals(general("1234567890"), "1234567890")
    assertEquals(general("12345678901"), "12345678901")
    assertEquals(general("-12345678901"), "-12345678901")
  }

  test("General cell display: the 11-character boundary, both magnitudes (GH-672)") {
    List(
      "99999999999" -> "99999999999",
      "100000000000" -> "1E+11",
      "99999999999.4" -> "99999999999",
      "99999999999.5" -> "1E+11",
      "12345678901.5" -> "12345678902",
      "9999999999.5" -> "10000000000",
      "1234567890.12" -> "1234567890",
      "0.00001" -> "1E-05",
      "0.00005" -> "5E-05",
      "0.0001234567" -> "0.000123457",
      "0.3333333333333333" -> "0.333333333",
      "-0.3333333333333333" -> "-0.333333333",
      "0.6666666666666666" -> "0.666666667",
      "333.3333333333333" -> "333.3333333",
      "0.9999999999" -> "1",
      "0.5" -> "0.5",
      "2021" -> "2021",
      "45982.5" -> "45982.5",
      "1.50" -> "1.5",
      "1E+3" -> "1000",
      "1.5E+15" -> "1.5E+15",
      "1.234565E+20" -> "1.23457E+20",
      "1.23456789E+100" -> "1.2346E+100",
      "1.23456789E-100" -> "1.2346E-100",
      "9.999999E+99" -> "1E+100",
      "0" -> "0",
      "0.000" -> "0",
      "-0.0" -> "0"
    ).foreach { case (in, out) => assertEquals(general(in), out, s"General($in)") }
  }

  test("General cell display: extreme stored exponents stay bounded (GH-672)") {
    // A file can carry <v>1E+2147483647</v>: the integer branch asked for two billion digits
    assertEquals(general("1E+2147483647"), "1E+2147483647")
    assertEquals(general("-1E+2147483647"), "-1E+2147483647")
    assertEquals(general("1E-2147483647"), "1E-2147483647")
    // a ten-digit exponent leaves room for one significant digit (4.5 rounds HALF_UP to 5)
    assertEquals(general("4.5E-2147483646"), "5E-2147483646")
    assertEquals(general("4.5E+999999"), "4.5E+999999")
    // Scale Int.MinValue: the exponent (2147483656) is past Int, and rounding nine digits to one
    // must not push the scale past Int either
    val beyondInt =
      BigDecimal(new java.math.BigDecimal(java.math.BigInteger.valueOf(123456789L), Int.MinValue))
    assertEquals(
      NumFmtFormatter.formatValue(CellValue.Number(beyondInt), NumFmt.General),
      "1E+2147483656"
    )
    val nines =
      BigDecimal(new java.math.BigDecimal(java.math.BigInteger.valueOf(999999999L), Int.MaxValue))
    assertEquals(
      NumFmtFormatter.formatValue(CellValue.Number(nines), NumFmt.General),
      "1E-2147483638"
    )
  }

  test("General cell display: independent of the default Locale (GH-672)") {
    // The old E branch formatted a Double through f"%.6E": 1,234570E+11 under de_DE
    val saved = java.util.Locale.getDefault
    try
      java.util.Locale.setDefault(java.util.Locale.GERMANY)
      assertEquals(general("123456789012"), "1.23457E+11")
      assertEquals(general("0.3333333333333333"), "0.333333333")
    finally java.util.Locale.setDefault(saved)
  }

  test("General keyword inside a section shares the 11-character rule (GH-672)") {
    val units = NumFmt.Custom("General\" units\"")
    assertEquals(
      NumFmtFormatter.formatValue(CellValue.Number(BigDecimal("123456789012")), units),
      "1.23457E+11 units"
    )
  }

  /** Magnitudes across the whole Int scale range, both signs, up to 40 significant digits. */
  private val genGeneralDisplay: Gen[BigDecimal] =
    for
      digits <- Gen.choose(1, 40)
      unscaled <- Gen.listOfN(digits, Gen.choose(0, 9)).map(ds => BigInt(ds.mkString))
      scale <- Gen.oneOf(
        Gen.choose(-30, 30),
        Gen.choose(-400, 400),
        Gen.choose(Int.MinValue, Int.MinValue + 1000),
        Gen.choose(Int.MaxValue - 1000, Int.MaxValue)
      )
      negative <- Gen.oneOf(true, false)
    yield
      val magnitude = BigDecimal(new java.math.BigDecimal(unscaled.bigInteger, scale))
      if negative then -magnitude else magnitude

  property("General cell display: grammar, the 11-character bound and the sign law (GH-672)") {
    val grammar = "-?[0-9]+(\\.[0-9]+)?(E[+-][0-9]{2,})?".r
    forAll(genGeneralDisplay) { (n: BigDecimal) =>
      val text = NumFmtFormatter.formatValue(CellValue.Number(n), NumFmt.General)
      assert(grammar.matches(text), s"'$text' for $n")
      val unsigned = text.stripPrefix("-")
      val expDigits = unsigned.dropWhile(_ != 'E').drop(2).length
      // "E+" and up to seven exponent digits leave room for a one-digit mantissa within 11
      if expDigits <= 7 then assert(unsigned.length <= 11, s"'$text' (${unsigned.length}) for $n")
      else assertEquals(unsigned.length, 1 + 2 + expDigits, s"'$text' for $n")
      if n.signum == 0 then assertEquals(text, "0")
      else
        assertEquals(text.startsWith("-"), n.signum < 0, s"'$text' for $n")
        assertEquals(
          NumFmtFormatter.formatValue(CellValue.Number(-n), NumFmt.General),
          if n.signum < 0 then unsigned else "-" + text,
          s"$n"
        )
    }
  }

  property("General cell display: a plain rendering reads back within half a displayed unit") {
    // |x| in [1E-4, 1E11): the text is x rounded to the decimals the integer digits leave
    forAll(Gen.choose(-4, 10), Gen.choose(1L, 999999999999999L), Gen.oneOf(true, false)) {
      (exp: Int, unscaled: Long, negative: Boolean) =>
        val digits = unscaled.toString.length
        val magnitude = BigDecimal(java.math.BigDecimal.valueOf(unscaled, digits - 1 - exp))
        val n = if negative then -magnitude else magnitude
        val text = NumFmtFormatter.formatValue(CellValue.Number(n), NumFmt.General)
        if !text.contains('E') then
          val decimals = if exp >= 0 then math.max(0, 9 - exp) else 9
          val back = BigDecimal(text)
          assert(
            (back - n).abs <= BigDecimal(s"5E-${decimals + 1}"),
            s"'$text' for $n"
          )
    }
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

  test("Cell display hands the strategy the formula cell and its position") {
    import DisplayConversions.given
    import ExcelInterpolator.*
    import com.tjclp.xl.display.syntax.*

    // An evaluating strategy needs the position to evaluate an uncached formula as that cell
    given Sheet = Sheet(name = SheetName.unsafe("Test"))
      .put(ref"B2", CellValue.Formula("SUM(A1:A3)"))
      .style(ref"B2", CellStyle.default.withNumFmt(NumFmt.Percent))
      .unsafe

    given FormulaDisplayStrategy with
      def format(formula: String, sheet: Sheet): String = s"positionless:$formula"
      override def formatAt(
        formula: CellValue.Formula,
        numFmt: NumFmt,
        sheet: Sheet,
        at: ARef
      ): String = s"${at.toA1}:${formula.expression}:${numFmt == NumFmt.Percent}"

    assertEquals(summon[Sheet].displayCell(ref"B2").formatted, "B2:SUM(A1:A3):true")
    assertEquals(excel"${ref"B2"}", "B2:SUM(A1:A3):true")
    // a Cell the sheet does not hold is displayed at its own position
    assertEquals(excel"${Cell(ref"D4", CellValue.Formula("A1*2"))}", "D4:A1*2:false")
  }

  test("A strategy written without formatAt displays formula cells as before") {
    import com.tjclp.xl.display.syntax.*

    val sheet = Sheet(name = SheetName.unsafe("Test"))
      .put(ref"A1", CellValue.Formula("A2+1"))
      .put(ref"B1", CellValue.Formula("A2+1", Some(CellValue.Number(BigDecimal(3)))))

    // formatAt's default delegates to formatCached, whose default delegates to format
    given FormulaDisplayStrategy with
      def format(formula: String, sheet: Sheet): String = s"legacy:$formula"

    assertEquals(sheet.displayCell(ref"A1").formatted, "legacy:A2+1")
    assertEquals(sheet.displayCell(ref"B1").formatted, "legacy:A2+1")
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
