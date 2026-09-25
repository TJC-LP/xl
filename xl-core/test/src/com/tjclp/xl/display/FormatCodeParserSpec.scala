package com.tjclp.xl.display

import munit.FunSuite
import FormatCodeParser.*
import com.tjclp.xl.styles.numfmt.NumFmt

/**
 * Tests for Excel custom number format code parser.
 */
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial"))
class FormatCodeParserSpec extends FunSuite:

  // ========== Parsing Tests ==========

  test("parse: simple integer format #,##0") {
    val result = FormatCodeParser.parse("#,##0")
    assert(result.isRight, s"Parse failed: $result")
    val code = result.toOption.get
    assertEquals(code.negative, None)
    assertEquals(code.zero, None)
    assertEquals(code.text, None)
    assert(code.positive.pattern.hasThousands, "Should detect thousands separator")
  }

  test("parse: decimal format #,##0.00") {
    val result = FormatCodeParser.parse("#,##0.00")
    assert(result.isRight)
    val code = result.toOption.get
    assert(code.positive.pattern.hasThousands)

    val hasDecimal = code.positive.pattern.tokens.exists(_ == FormatToken.Decimal)
    assert(hasDecimal, "Should have decimal token")
  }

  test("parse: percent format 0%") {
    val result = FormatCodeParser.parse("0%")
    assert(result.isRight)
    val code = result.toOption.get
    assert(code.positive.pattern.hasPercent, "Should detect percent")
  }

  test("parse: currency with quoted literal \"$\"#,##0.00") {
    val result = FormatCodeParser.parse("\"$\"#,##0.00")
    assert(result.isRight)
    val code = result.toOption.get

    val hasLiteralDollar = code.positive.pattern.tokens.exists {
      case FormatToken.Literal("$") => true
      case _ => false
    }
    assert(hasLiteralDollar, "Should have literal dollar sign")
  }

  test("parse: two sections #,##0;(#,##0)") {
    val result = FormatCodeParser.parse("#,##0;(#,##0)")
    assert(result.isRight)
    val code = result.toOption.get
    assert(code.negative.isDefined, "Should have negative section")

    val hasOpenParen = code.negative.get.pattern.tokens.exists {
      case FormatToken.Literal("(") => true
      case _ => false
    }
    assert(hasOpenParen, "Negative section should have parenthesis")
  }

  test("parse: color condition [Red]0.00") {
    val result = FormatCodeParser.parse("[Red]0.00")
    assert(result.isRight)
    val code = result.toOption.get

    code.positive.condition match
      case Some(Condition.Color(name)) =>
        assertEquals(name, "Red")
      case other =>
        fail(s"Expected Color condition, got: $other")
  }

  test("parse: comparison condition [>100]\"High\"") {
    val result = FormatCodeParser.parse("[>100]\"High\"")
    assert(result.isRight)
    val code = result.toOption.get

    code.positive.condition match
      case Some(Condition.Compare(op, value)) =>
        assertEquals(op, ">")
        assertEquals(value, BigDecimal(100))
      case other =>
        fail(s"Expected Compare condition, got: $other")
  }

  test("parse: locale code [$-409]") {
    val result = FormatCodeParser.parse("[$-409]#,##0")
    assert(result.isRight)
    val code = result.toOption.get

    code.positive.condition match
      case Some(Condition.Locale(locale, symbol)) =>
        assertEquals(locale, "409")
        assertEquals(symbol, None)
      case other =>
        fail(s"Expected Locale condition, got: $other")
  }

  test("parse: spacer _)") {
    val result = FormatCodeParser.parse("#,##0_)")
    assert(result.isRight)
    val code = result.toOption.get

    val hasSpacer = code.positive.pattern.tokens.exists {
      case FormatToken.Spacer(')') => true
      case _ => false
    }
    assert(hasSpacer, "Should have spacer token")
  }

  test("parse: escaped character \\$") {
    val result = FormatCodeParser.parse("\\$#,##0")
    assert(result.isRight)
    val code = result.toOption.get

    val hasLiteralDollar = code.positive.pattern.tokens.exists {
      case FormatToken.Literal("$") => true
      case _ => false
    }
    assert(hasLiteralDollar, "Should have escaped literal dollar")
  }

  test("parse: date format m/d/yy") {
    val result = FormatCodeParser.parse("m/d/yy")
    assert(result.isRight)
    val code = result.toOption.get

    val dateParts = code.positive.pattern.tokens.collect { case FormatToken.DatePart(p) => p }
    assertEquals(dateParts, Vector("m", "d", "yy"))
  }

  test("parse: datetime format yyyy-mm-dd h:mm AM/PM") {
    val result = FormatCodeParser.parse("yyyy-mm-dd h:mm AM/PM")
    assert(result.isRight)
    val code = result.toOption.get

    val hasAmPm = code.positive.pattern.tokens.exists {
      case FormatToken.AmPm(_) => true
      case _ => false
    }
    assert(hasAmPm, "Should have AM/PM token")
  }

  test("parse: four sections format") {
    val result = FormatCodeParser.parse("#,##0;(#,##0);\"Zero\";@")
    assert(result.isRight)
    val code = result.toOption.get
    assert(code.positive.pattern.tokens.nonEmpty)
    assert(code.negative.isDefined)
    assert(code.zero.isDefined)
    assert(code.text.isDefined)
  }

  test("parse: empty format fails") {
    val result = FormatCodeParser.parse("")
    assert(result.isLeft)
  }

  // ========== Formatting Tests ==========

  test("applyFormat: simple integer") {
    val code = FormatCodeParser.parse("#,##0").toOption.get
    val (formatted, _) = FormatCodeParser.applyFormat(BigDecimal(1234567), code)
    assertEquals(formatted, "1,234,567")
  }

  test("applyFormat: decimal places") {
    val code = FormatCodeParser.parse("#,##0.00").toOption.get
    val (formatted, _) = FormatCodeParser.applyFormat(BigDecimal("1234.5"), code)
    assertEquals(formatted, "1,234.50")
  }

  test("applyFormat: negative with single section preserves default minus sign") {
    val code = FormatCodeParser.parse("#,##0.00").toOption.get
    val (formatted, _) = FormatCodeParser.applyFormat(BigDecimal("-1234.5"), code)
    assertEquals(formatted, "-1,234.50")
  }

  test("applyFormat: percent") {
    val code = FormatCodeParser.parse("0%").toOption.get
    val (formatted, _) = FormatCodeParser.applyFormat(BigDecimal("0.15"), code)
    assertEquals(formatted, "15%")
  }

  test("applyFormat: currency with literal") {
    val code = FormatCodeParser.parse("\"$\"#,##0.00").toOption.get
    val (formatted, _) = FormatCodeParser.applyFormat(BigDecimal("1234.56"), code)
    assertEquals(formatted, "$1,234.56")
  }

  test("applyFormat: negative with parentheses") {
    val code = FormatCodeParser.parse("#,##0;(#,##0)").toOption.get
    val (formatted, _) = FormatCodeParser.applyFormat(BigDecimal("-1234"), code)
    assertEquals(formatted, "(1,234)")
  }

  test("applyFormat: color condition returned") {
    val code = FormatCodeParser.parse("[Red]0.00").toOption.get
    val (_, color) = FormatCodeParser.applyFormat(BigDecimal("5"), code)
    assertEquals(color, Some("Red"))
  }

  test("applyFormat: zero value uses zero section") {
    val code = FormatCodeParser.parse("#,##0;(#,##0);\"Zero\"").toOption.get
    val (formatted, _) = FormatCodeParser.applyFormat(BigDecimal("0"), code)
    assertEquals(formatted, "Zero")
  }

  test("applyFormat: small decimals") {
    val code = FormatCodeParser.parse("0.0000").toOption.get
    val (formatted, _) = FormatCodeParser.applyFormat(BigDecimal("0.1234"), code)
    assertEquals(formatted, "0.1234")
  }

  test("applyFormat: rounding") {
    val code = FormatCodeParser.parse("0.00").toOption.get
    val (formatted, _) = FormatCodeParser.applyFormat(BigDecimal("1.235"), code)
    assertEquals(formatted, "1.24") // HALF_UP rounding
  }

  test("applyFormat: leading zeros with 0 placeholder") {
    val code = FormatCodeParser.parse("000").toOption.get
    val (formatted, _) = FormatCodeParser.applyFormat(BigDecimal("5"), code)
    assertEquals(formatted, "005")
  }

  test("applyFormat: suppress leading zeros with # placeholder") {
    val code = FormatCodeParser.parse("###").toOption.get
    val (formatted, _) = FormatCodeParser.applyFormat(BigDecimal("5"), code)
    assertEquals(formatted, "5")
  }

  // ========== Real-World Excel Format Codes ==========

  test("real format: accounting style $#,##0.00_);($#,##0.00)") {
    val code = FormatCodeParser.parse("\"$\"#,##0.00_);(\"$\"#,##0.00)").toOption.get

    val (pos, _) = FormatCodeParser.applyFormat(BigDecimal("1234.56"), code)
    assert(pos.contains("$"), s"Positive should have dollar: $pos")
    assert(pos.contains("1,234.56"), s"Positive should have formatted number: $pos")

    val (neg, _) = FormatCodeParser.applyFormat(BigDecimal("-1234.56"), code)
    assert(neg.contains("("), s"Negative should have parenthesis: $neg")
    assert(neg.contains("$"), s"Negative should have dollar: $neg")
  }

  test("real format: scientific 0.00E+00") {
    // Scientific format not fully implemented yet, but should parse
    val result = FormatCodeParser.parse("0.00E+00")
    assert(result.isRight, "Should parse scientific format")
  }

  test("real format: thousands with scaling #,##0,") {
    // Trailing comma scales by 1000; the interior comma still groups (GH-666)
    val code = FormatCodeParser.parse("#,##0,").toOption.get
    assertEquals(FormatCodeParser.applyFormat(BigDecimal("1234567"), code)._1, "1,235")
    assertEquals(FormatCodeParser.applyFormat(BigDecimal("1234567890"), code)._1, "1,234,568")
  }

  test("real format: mixed positive/negative/zero") {
    val code = FormatCodeParser.parse("#,##0.00;[Red](#,##0.00);\"—\"").toOption.get

    val (pos, posColor) = FormatCodeParser.applyFormat(BigDecimal("100"), code)
    assertEquals(posColor, None)
    assert(pos.contains("100.00"))

    val (neg, negColor) = FormatCodeParser.applyFormat(BigDecimal("-100"), code)
    assertEquals(negColor, Some("Red"))
    assert(neg.contains("("))

    val (zero, _) = FormatCodeParser.applyFormat(BigDecimal("0"), code)
    assertEquals(zero, "—")
  }

  // ========== Section Selection (Excel semantics, GH-254) ==========
  // Excel routes values to sections as follows:
  //   1 section:  all numbers use it
  //   2 sections: positive + zero use 1st, negative uses 2nd
  //   3 sections: positive → 1st, negative → 2nd, zero → 3rd
  //   4 sections: positive / negative / zero / text

  test("section selection: 1 section formats positive, negative, and zero") {
    val code = FormatCodeParser.parse("0.0").toOption.get
    assertEquals(FormatCodeParser.applyFormat(BigDecimal("1.5"), code)._1, "1.5")
    assertEquals(FormatCodeParser.applyFormat(BigDecimal("-1.5"), code)._1, "-1.5")
    assertEquals(FormatCodeParser.applyFormat(BigDecimal("0"), code)._1, "0.0")
  }

  test("section selection: 2 sections route zero to positive section") {
    val code = FormatCodeParser.parse("0.0;(0.0)").toOption.get
    assertEquals(FormatCodeParser.applyFormat(BigDecimal("1.5"), code)._1, "1.5")
    assertEquals(FormatCodeParser.applyFormat(BigDecimal("-1.5"), code)._1, "(1.5)")
    assertEquals(FormatCodeParser.applyFormat(BigDecimal("0"), code)._1, "0.0")
  }

  test("section selection: 3 sections route zero to explicit zero section") {
    val code = FormatCodeParser.parse("0.0;(0.0);\"-\"").toOption.get
    assertEquals(FormatCodeParser.applyFormat(BigDecimal("1.5"), code)._1, "1.5")
    assertEquals(FormatCodeParser.applyFormat(BigDecimal("-1.5"), code)._1, "(1.5)")
    assertEquals(FormatCodeParser.applyFormat(BigDecimal("0"), code)._1, "-")
  }

  test("section selection: 4 sections route pos/neg/zero/text independently") {
    val code = FormatCodeParser.parse("0.0;(0.0);\"zero\";@").toOption.get
    assertEquals(FormatCodeParser.applyFormat(BigDecimal("1.5"), code)._1, "1.5")
    assertEquals(FormatCodeParser.applyFormat(BigDecimal("-1.5"), code)._1, "(1.5)")
    assertEquals(FormatCodeParser.applyFormat(BigDecimal("0"), code)._1, "zero")
    assertEquals(FormatCodeParser.applyTextFormat("abc", code), "abc")
  }

  test("TJC house code: \"$\"#,##0.0_);(\"$\"#,##0.0) routes zero to positive (GH-254)") {
    val code = FormatCodeParser.parse("\"$\"#,##0.0_);(\"$\"#,##0.0)").toOption.get

    val (zero, _) = FormatCodeParser.applyFormat(BigDecimal("0"), code)
    assert(!zero.contains("("), s"Zero must not use the negative section: $zero")
    assertEquals(zero.trim, "$0.0") // trailing space from _) spacer

    val (pos, _) = FormatCodeParser.applyFormat(BigDecimal("1234.56"), code)
    assertEquals(pos.trim, "$1,234.6")

    val (neg, _) = FormatCodeParser.applyFormat(BigDecimal("-1234.56"), code)
    assertEquals(neg, "($1,234.6)")
  }

  test("TJC house code: 0.0%_);(0.0%) routes zero to positive (GH-254)") {
    val code = FormatCodeParser.parse("0.0%_);(0.0%)").toOption.get

    val (zero, _) = FormatCodeParser.applyFormat(BigDecimal("0"), code)
    assert(!zero.contains("("), s"Zero must not use the negative section: $zero")
    assertEquals(zero.trim, "0.0%") // trailing space from _) spacer

    val (pos, _) = FormatCodeParser.applyFormat(BigDecimal("0.125"), code)
    assertEquals(pos.trim, "12.5%")

    val (neg, _) = FormatCodeParser.applyFormat(BigDecimal("-0.125"), code)
    assertEquals(neg, "(12.5%)")
  }

  // ========== Empty Sections (hide-value idiom, GH-262) ==========
  // An EMPTY section means "display nothing" for that value class:
  //   "0.0;;" → positive shown, negative hidden, zero hidden
  //   "0.0;"  → negative hidden, zero routes to positive (2-section rule)

  test("parse: trailing empty sections preserved — 0.0;; yields 3 sections (GH-262)") {
    val code = FormatCodeParser.parse("0.0;;").toOption.get
    assert(code.negative.isDefined, "negative (2nd) section must exist")
    assert(code.zero.isDefined, "zero (3rd) section must exist")
    assertEquals(code.negative.get.pattern.tokens, Vector.empty[FormatToken])
    assertEquals(code.zero.get.pattern.tokens, Vector.empty[FormatToken])
    assertEquals(code.text, None)
  }

  test("hide-zero idiom: zero with 0.0;; renders empty string (GH-262)") {
    val code = FormatCodeParser.parse("0.0;;").toOption.get
    assertEquals(FormatCodeParser.applyFormat(BigDecimal("0"), code)._1, "")
    assertEquals(FormatCodeParser.applyFormat(BigDecimal("1.5"), code)._1, "1.5")
  }

  test("hide-negative idiom: negative with 0.0;; renders empty string (GH-262)") {
    val code = FormatCodeParser.parse("0.0;;").toOption.get
    assertEquals(FormatCodeParser.applyFormat(BigDecimal("-1.5"), code)._1, "")
  }

  test("trailing empty 2nd section: 0.0; hides negatives, zero routes to positive (GH-262)") {
    val code = FormatCodeParser.parse("0.0;").toOption.get
    assert(code.negative.isDefined, "negative (2nd) section must exist")
    assertEquals(code.negative.get.pattern.tokens, Vector.empty[FormatToken])
    assertEquals(FormatCodeParser.applyFormat(BigDecimal("1.5"), code)._1, "1.5")
    assertEquals(FormatCodeParser.applyFormat(BigDecimal("-1.5"), code)._1, "")
    // GH-254 zero routing must be intact: 2 sections → zero uses positive section
    assertEquals(FormatCodeParser.applyFormat(BigDecimal("0"), code)._1, "0.0")
  }

  test("interior empty zero section: 0;-0;;@ hides zero only (GH-262)") {
    val code = FormatCodeParser.parse("0;-0;;@").toOption.get
    assert(code.zero.isDefined, "zero (3rd) section must exist")
    assert(code.text.isDefined, "text (4th) section must exist")
    assertEquals(FormatCodeParser.applyFormat(BigDecimal("5"), code)._1, "5")
    assertEquals(FormatCodeParser.applyFormat(BigDecimal("-5"), code)._1, "-5")
    assertEquals(FormatCodeParser.applyFormat(BigDecimal("0"), code)._1, "")
    assertEquals(FormatCodeParser.applyTextFormat("abc", code), "abc")
  }

  test("NumFmtFormatter: Custom 0.0;; hides zero and negative, no General fallback (GH-262)") {
    val fmt = NumFmt.Custom("0.0;;")
    assertEquals(NumFmtFormatter.formatNumber(BigDecimal("0"), fmt), "")
    assertEquals(NumFmtFormatter.formatNumber(BigDecimal("-2.5"), fmt), "")
    assertEquals(NumFmtFormatter.formatNumber(BigDecimal("2.5"), fmt), "2.5")
  }

  // ========== Conditional sections (GH-285) ==========
  // Excel semantics per SheetJS/SSF choose_fmt (validated against Excel corpora):
  //   - compare conditions are honored on the first two sections only
  //   - the first matching condition wins; an unmatched value falls back to the
  //     third section when BOTH leading sections carry conditions, otherwise to the
  //     second (sections pad positionally: 1 section -> [s1,s1,s1], 2 -> [s1,s2,s1])
  //   - when conditions are present, sign/zero positional routing is suspended
  //   - multi-section formats render |value|; the minus sign must be written
  //     explicitly in the pattern; single-section formats keep the default minus
  //   - there is no ###### fallback: Excel's no-match-no-fallback behavior is
  //     undocumented (MS docs and major guides define no such case) and SSF's
  //     Excel-validated routing always resolves to a section

  test(
    "conditional routing: [>100]#,##0;[<=0]0.00;0.0 — first match wins, else fallback (GH-285)"
  ) {
    val code = FormatCodeParser.parse("[>100]#,##0;[<=0]0.00;0.0").toOption.get
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(250), code)._1, "250")
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(-5), code)._1, "5.00")
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(0), code)._1, "0.00")
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(50), code)._1, "50.0")
  }

  test(
    "conditional routing: [>100]0;0.0 — unconditioned second section is the else branch (GH-285)"
  ) {
    val code = FormatCodeParser.parse("[>100]0;0.0").toOption.get
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(250), code)._1, "250")
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(50), code)._1, "50.0")
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(0), code)._1, "0.0")
    // Multi-section: |value| is rendered; the sign must be explicit in the pattern
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(-50), code)._1, "50.0")
  }

  test("conditional routing: stacked color and condition [Red][<=100]0;[Blue][>100]0 (GH-285)") {
    val code = FormatCodeParser.parse("[Red][<=100]0;[Blue][>100]0").toOption.get
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(50), code), ("50", Some("Red")))
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(500), code), ("500", Some("Blue")))
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(-3), code), ("3", Some("Red")))
  }

  test(
    "conditional routing: both conditions unmatched, no third section → first pattern (GH-285)"
  ) {
    // SSF pads 2 sections to [s1, s2, s1], so the no-match fallback is the FIRST
    // section's pattern (its condition ignored). No ###### — see block comment.
    val code = FormatCodeParser.parse("[>100]0;[<0]0.00").toOption.get
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(50), code)._1, "50")
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(250), code)._1, "250")
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(-5), code)._1, "5.00")
  }

  test("conditional routing: single conditional section formats all values, sign kept (GH-285)") {
    val code = FormatCodeParser.parse("[>100]0").toOption.get
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(250), code)._1, "250")
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(50), code)._1, "50")
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(-50), code)._1, "-50")
  }

  test("conditional ops: =, <>, >= comparisons (GH-285)") {
    val eq = FormatCodeParser.parse("[=5]\"five\";0").toOption.get
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(5), eq)._1, "five")
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(4), eq)._1, "4")

    val ne = FormatCodeParser.parse("[<>0]0.0;\"zero\"").toOption.get
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(5), ne)._1, "5.0")
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(0), ne)._1, "zero")

    val ge = FormatCodeParser.parse("[>=10]\"big\";\"small\"").toOption.get
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(10), ge)._1, "big")
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(9), ge)._1, "small")
  }

  test("trailing @ among <4 sections is the text section, not the negative (GH-285)") {
    // SSF choose_fmt: "0.0;@" has ONE numeric section — negatives keep the default
    // minus instead of routing into the text section (which previously hid them).
    val code = FormatCodeParser.parse("0.0;@").toOption.get
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(-5), code)._1, "-5.0")
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(0), code)._1, "0.0")
    assertEquals(FormatCodeParser.applyTextFormat("abc", code), "abc")
  }

  test("trailing @ section with literals formats text values (GH-285)") {
    val code = FormatCodeParser.parse("0;0;\"val: \"@").toOption.get
    assertEquals(FormatCodeParser.applyTextFormat("abc", code), "val: abc")
  }

  // ========== Date display gaps (GH-283) ==========
  // 1. General on a date-typed value shows the Excel SERIAL NUMBER (dates ARE
  //    numbers; Excel's General does not pretty-print them) — not ISO text.
  // 2. Custom codes applied to dates route through section selection like any
  //    number (the serial is the routed value): ';;;' hides dates, numeric
  //    sections render the serial, conditional date codes pick sections by serial.
  // 3. Date-token sections fed an out-of-range serial (negative or >= 10000-01-01)
  //    render '######' like Excel's unrepresentable-date fill.

  test("formatDateTime: General shows the date serial, not ISO text (GH-283)") {
    val dt = java.time.LocalDateTime.of(2025, 11, 21, 0, 0, 0)
    assertEquals(NumFmtFormatter.formatDateTime(dt, NumFmt.General), "45982")
  }

  test("formatDateTime: General with a time component shows the fractional serial (GH-283)") {
    val noon = java.time.LocalDateTime.of(2025, 11, 21, 12, 0, 0)
    assertEquals(NumFmtFormatter.formatDateTime(noon, NumFmt.General), "45982.5")
  }

  test("formatValue: DateTime cell under General renders the serial (GH-283)") {
    import com.tjclp.xl.cells.CellValue
    val value = CellValue.DateTime(java.time.LocalDateTime.of(2025, 11, 21, 0, 0, 0))
    assertEquals(NumFmtFormatter.formatValue(value, NumFmt.General), "45982")
  }

  test("formatDateTime: numeric built-in formats apply to the serial (GH-283)") {
    val dt = java.time.LocalDateTime.of(2025, 11, 21, 0, 0, 0)
    assertEquals(NumFmtFormatter.formatDateTime(dt, NumFmt.Integer), "45982")
    assertEquals(NumFmtFormatter.formatDateTime(dt, NumFmt.ThousandsSeparator), "45,982")
  }

  test("formatDateTime: ';;;' hides a date value (GH-283)") {
    val dt = java.time.LocalDateTime.of(2025, 11, 21, 0, 0, 0)
    assertEquals(NumFmtFormatter.formatDateTime(dt, NumFmt.Custom(";;;")), "")
  }

  test("formatDateTime: custom NUMERIC code renders the serial, not ISO text (GH-283)") {
    val dt = java.time.LocalDateTime.of(2025, 11, 21, 0, 0, 0)
    assertEquals(NumFmtFormatter.formatDateTime(dt, NumFmt.Custom("0.00")), "45982.00")
  }

  test("formatDateTime: conditional date code routes sections by serial (GH-283/285)") {
    val fmt = NumFmt.Custom("[<45000]m/d/yy;yyyy")
    val recent = java.time.LocalDateTime.of(2025, 11, 21, 0, 0, 0) // serial 45982
    val old = java.time.LocalDateTime.of(2020, 1, 1, 0, 0, 0) // serial 43831
    assertEquals(NumFmtFormatter.formatDateTime(recent, fmt), "2025")
    assertEquals(NumFmtFormatter.formatDateTime(old, fmt), "1/1/20")
  }

  test("formatDateTime: 'm/d/yy;@' still renders via the date section (GH-283 regression)") {
    val dt = java.time.LocalDateTime.of(2025, 11, 21, 0, 0, 0)
    assertEquals(NumFmtFormatter.formatDateTime(dt, NumFmt.Custom("m/d/yy;@")), "11/21/25")
  }

  test("formatNumber: serial with a date code still renders the date (GH-283 regression)") {
    assertEquals(
      NumFmtFormatter.formatNumber(BigDecimal(45982), NumFmt.Custom("m/d/yy")),
      "11/21/25"
    )
    assertEquals(
      NumFmtFormatter.formatNumber(BigDecimal("0.5"), NumFmt.Custom("h:mm AM/PM")),
      "12:00 PM"
    )
  }

  test("formatNumber: out-of-range serial with a date code renders ###### (GH-283)") {
    assertEquals(NumFmtFormatter.formatNumber(BigDecimal(-5), NumFmt.Custom("m/d/yy")), "######")
    assertEquals(
      NumFmtFormatter.formatNumber(BigDecimal(3000000), NumFmt.Custom("m/d/yy")),
      "######"
    )
  }

  // ========== Date/Time Formatting Tests ==========

  test("applyDateFormat: simple date m/d/yy") {
    val code = FormatCodeParser.parse("m/d/yy").toOption.get
    val dt = java.time.LocalDateTime.of(2025, 11, 25, 16, 13, 0)
    val result = FormatCodeParser.applyDateFormat(dt, code)
    assertEquals(result, "11/25/25")
  }

  test("applyDateFormat: ISO-style yyyy-mm-dd") {
    val code = FormatCodeParser.parse("yyyy-mm-dd").toOption.get
    val dt = java.time.LocalDateTime.of(2025, 11, 25, 0, 0, 0)
    val result = FormatCodeParser.applyDateFormat(dt, code)
    assertEquals(result, "2025-11-25")
  }

  test("applyDateFormat: month abbreviation mmm-yy") {
    val code = FormatCodeParser.parse("mmm-yy").toOption.get
    val dt = java.time.LocalDateTime.of(2025, 11, 25, 0, 0, 0)
    val result = FormatCodeParser.applyDateFormat(dt, code)
    assertEquals(result, "Nov-25")
  }

  test("applyDateFormat: full month name mmmm d, yyyy") {
    val code = FormatCodeParser.parse("mmmm d, yyyy").toOption.get
    val dt = java.time.LocalDateTime.of(2025, 11, 25, 0, 0, 0)
    val result = FormatCodeParser.applyDateFormat(dt, code)
    assertEquals(result, "November 25, 2025")
  }

  test("applyDateFormat: time with AM/PM h:mm AM/PM") {
    val code = FormatCodeParser.parse("h:mm AM/PM").toOption.get
    val dt = java.time.LocalDateTime.of(2025, 11, 25, 16, 13, 0)
    val result = FormatCodeParser.applyDateFormat(dt, code)
    assertEquals(result, "4:13 PM")
  }

  test("applyDateFormat: time AM morning") {
    val code = FormatCodeParser.parse("h:mm AM/PM").toOption.get
    val dt = java.time.LocalDateTime.of(2025, 11, 25, 9, 30, 0)
    val result = FormatCodeParser.applyDateFormat(dt, code)
    assertEquals(result, "9:30 AM")
  }

  test("applyDateFormat: fiscal year yyyy\"A\"") {
    val code = FormatCodeParser.parse("yyyy\"A\"").toOption.get
    val dt = java.time.LocalDateTime.of(2025, 1, 1, 0, 0, 0)
    val result = FormatCodeParser.applyDateFormat(dt, code)
    assertEquals(result, "2025A")
  }

  test("applyDateFormat: full datetime yyyy-mm-dd h:mm AM/PM") {
    val code = FormatCodeParser.parse("yyyy-mm-dd h:mm AM/PM").toOption.get
    val dt = java.time.LocalDateTime.of(2025, 11, 25, 16, 13, 0)
    val result = FormatCodeParser.applyDateFormat(dt, code)
    assertEquals(result, "2025-11-25 4:13 PM")
  }

  test("applyDateFormat: month vs minute disambiguation") {
    // m after h = minute, m without h = month
    val timeCode = FormatCodeParser.parse("h:mm:ss").toOption.get
    val dateCode = FormatCodeParser.parse("m/d/yyyy").toOption.get
    val dt = java.time.LocalDateTime.of(2025, 3, 15, 14, 30, 45)

    // Re-pinned by GH-410 per ECMA-376 §18.8.31: 'h' uses the 12-hour clock only when the
    // code contains AM/PM; "h:mm:ss" has none, so 14:30:45 renders as 14:30:45 (was 2:30:45).
    val timeResult = FormatCodeParser.applyDateFormat(dt, timeCode)
    assertEquals(timeResult, "14:30:45") // mm = 30 (minute)

    val dateResult = FormatCodeParser.applyDateFormat(dt, dateCode)
    assertEquals(dateResult, "3/15/2025") // m = 3 (month)
  }

  test("applyDateFormat: bare h is the 24-hour clock, AM/PM switches to 12-hour (GH-410)") {
    // ECMA-376 §18.8.31: "If the format contains an AM or PM code, the hour is based on
    // the 12-hour clock; otherwise, it is based on the 24-hour clock."
    val bare = FormatCodeParser.parse("h:mm").toOption.get
    val meridiem = FormatCodeParser.parse("h:mm AM/PM").toOption.get
    val padded = FormatCodeParser.parse("hh:mm").toOption.get
    val afternoon = java.time.LocalDateTime.of(2025, 1, 1, 16, 5, 0)
    val morning = java.time.LocalDateTime.of(2025, 1, 1, 9, 5, 0)
    val midnight = java.time.LocalDateTime.of(2025, 1, 1, 0, 5, 0)

    assertEquals(FormatCodeParser.applyDateFormat(afternoon, bare), "16:05")
    assertEquals(FormatCodeParser.applyDateFormat(morning, bare), "9:05")
    assertEquals(FormatCodeParser.applyDateFormat(midnight, bare), "0:05")
    assertEquals(FormatCodeParser.applyDateFormat(afternoon, padded), "16:05")
    assertEquals(FormatCodeParser.applyDateFormat(midnight, padded), "00:05")
    assertEquals(FormatCodeParser.applyDateFormat(afternoon, meridiem), "4:05 PM")
    assertEquals(FormatCodeParser.applyDateFormat(midnight, meridiem), "12:05 AM")
  }

  // ========== Scientific Notation (GH-410) ==========

  private def formatWith(code: String, value: BigDecimal): String =
    FormatCodeParser.parse(code) match
      case Right(fmt) => FormatCodeParser.applyFormat(value, fmt)._1
      case Left(e) => fail(s"parse('$code') failed: $e")

  test("applyFormat: scientific 0.00E+00 basics (GH-410)") {
    assertEquals(formatWith("0.00E+00", BigDecimal("1234.5")), "1.23E+03")
    assertEquals(formatWith("0.00E+00", BigDecimal("123456.789")), "1.23E+05")
    assertEquals(formatWith("0.00E+00", BigDecimal("0")), "0.00E+00")
    assertEquals(formatWith("0.00E+00", BigDecimal("1")), "1.00E+00")
    assertEquals(formatWith("0.00E+00", BigDecimal("0.0001234")), "1.23E-04")
    assertEquals(formatWith("0.00E+00", BigDecimal("-1234.5")), "-1.23E+03")
  }

  test("applyFormat: scientific mantissa rounding renormalizes (GH-410)") {
    // 9.999 rounds to 10.00 at two decimals, which overflows one integer digit: Excel
    // renormalizes to 1.00E+01.
    assertEquals(formatWith("0.00E+00", BigDecimal("9.999")), "1.00E+01")
  }

  test("applyFormat: engineering ##0.0E+0 snaps exponent to placeholder multiples (GH-410)") {
    // Three integer placeholders make the exponent a multiple of 3 (id 48's code).
    assertEquals(formatWith("##0.0E+0", BigDecimal("123456")), "123.5E+3")
    assertEquals(formatWith("##0.0E+0", BigDecimal("1234.5")), "1.2E+3")
    assertEquals(formatWith("##0.0E+0", BigDecimal("0.01")), "10.0E-3")
  }

  test("applyFormat: E- shows the exponent sign only when negative (GH-410)") {
    assertEquals(formatWith("0.00E-00", BigDecimal("1234.5")), "1.23E03")
    assertEquals(formatWith("0.00E-00", BigDecimal("0.0001234")), "1.23E-04")
  }

  test("applyFormat: lowercase e is preserved (GH-410)") {
    assertEquals(formatWith("0.00e+00", BigDecimal("1234.5")), "1.23e+03")
  }

  test("applyFormat: exponent placeholders pad, never truncate (GH-410)") {
    assertEquals(formatWith("0.0E+0", BigDecimal("1.5E+12")), "1.5E+12")
    assertEquals(formatWith("0.00E+000", BigDecimal("1234.5")), "1.23E+003")
  }

  test("applyDateFormat: 12-hour noon/midnight") {
    val code = FormatCodeParser.parse("h:mm AM/PM").toOption.get

    // Noon (12:00 PM)
    val noon = java.time.LocalDateTime.of(2025, 1, 1, 12, 0, 0)
    assertEquals(FormatCodeParser.applyDateFormat(noon, code), "12:00 PM")

    // Midnight (12:00 AM)
    val midnight = java.time.LocalDateTime.of(2025, 1, 1, 0, 0, 0)
    assertEquals(FormatCodeParser.applyDateFormat(midnight, code), "12:00 AM")
  }

  test("applyDateFormat: padded day dd") {
    val code = FormatCodeParser.parse("yyyy-mm-dd").toOption.get
    val dt = java.time.LocalDateTime.of(2025, 1, 5, 0, 0, 0)
    val result = FormatCodeParser.applyDateFormat(dt, code)
    assertEquals(result, "2025-01-05")
  }

  test("applyDateFormat: day of week ddd/dddd") {
    val shortCode = FormatCodeParser.parse("ddd").toOption.get
    val longCode = FormatCodeParser.parse("dddd").toOption.get
    val dt = java.time.LocalDateTime.of(2025, 11, 25, 0, 0, 0) // Tuesday

    assertEquals(FormatCodeParser.applyDateFormat(dt, shortCode), "Tue")
    assertEquals(FormatCodeParser.applyDateFormat(dt, longCode), "Tuesday")
  }

  test("hasDateTokens: detects date patterns") {
    val dateCode = FormatCodeParser.parse("m/d/yy").toOption.get
    val numCode = FormatCodeParser.parse("#,##0.00").toOption.get

    assert(FormatCodeParser.hasDateTokens(dateCode), "Should detect date tokens")
    assert(
      !FormatCodeParser.hasDateTokens(numCode),
      "Should not detect date tokens in number format"
    )
  }

  // ========== Syndigo Real-World Formats ==========

  test("syndigo format: m/d/yy;@") {
    val code = FormatCodeParser.parse("m/d/yy;@").toOption.get
    val dt = java.time.LocalDateTime.of(2025, 1, 15, 0, 0, 0)
    val result = FormatCodeParser.applyDateFormat(dt, code)
    assertEquals(result, "1/15/25")
  }

  test("syndigo format: mmm-yy with locale (escaped hyphen)") {
    // Syndigo uses: [$-409]mmm\-yy;@
    val code = FormatCodeParser.parse("mmm\\-yy").toOption.get
    val dt = java.time.LocalDateTime.of(2025, 1, 15, 0, 0, 0)
    val result = FormatCodeParser.applyDateFormat(dt, code)
    assertEquals(result, "Jan-25")
  }

  test("syndigo format: mmmm d, yyyy (escaped spaces/comma)") {
    // Syndigo uses: [$-409]mmmm\ d\,\ yyyy;@
    val code = FormatCodeParser.parse("mmmm\\ d\\,\\ yyyy").toOption.get
    val dt = java.time.LocalDateTime.of(2025, 1, 15, 0, 0, 0)
    val result = FormatCodeParser.applyDateFormat(dt, code)
    assertEquals(result, "January 15, 2025")
  }

  // ========== Fractions (GH-243) ==========
  // Excel picks the last continued-fraction convergent whose denominator fits the
  // placeholder budget (1 digit for ?/?, 2 for ??/??), computed in IEEE-754 double
  // arithmetic over the FULL value. The algorithm is a verbatim port of SheetJS/SSF
  // `frac` (reverse-engineered from Excel); the "Excel corpus" tests below pin
  // expected strings taken directly from SSF's Excel-generated test/fraction.json.
  // Double noise is part of the spec: 12.3 → "12 1/3" but bare 0.3 → " 2/7".
  // Alignment: `?` placeholders pad with spaces (numerator right-aligns, denominator
  // left-aligns); a whole number blanks the fraction area to preserve width.

  test("fraction parse: # ?/? produces a Fraction token (GH-243)") {
    val code = FormatCodeParser.parse("# ?/?").toOption.get
    val fracTokens = code.positive.pattern.tokens.collect { case f: FormatToken.Fraction => f }
    assertEquals(fracTokens, Vector(FormatToken.Fraction("?", "?", None)))
  }

  test("fraction parse: fixed denominator # ?/8 captures the literal denominator (GH-243)") {
    val code = FormatCodeParser.parse("# ?/8").toOption.get
    val fracTokens = code.positive.pattern.tokens.collect { case f: FormatToken.Fraction => f }
    assertEquals(fracTokens, Vector(FormatToken.Fraction("?", "8", Some(8L))))
  }

  test("fraction parse: date separators are untouched — m/d/yy has no Fraction token (GH-243)") {
    val code = FormatCodeParser.parse("m/d/yy").toOption.get
    val fracTokens = code.positive.pattern.tokens.collect { case f: FormatToken.Fraction => f }
    assertEquals(fracTokens, Vector.empty)
  }

  test("NumFmt.Fraction: simple halves 0.5 → ' 1/2' (GH-243)") {
    assertEquals(NumFmtFormatter.formatNumber(BigDecimal("0.5"), NumFmt.Fraction), " 1/2")
  }

  test("NumFmt.Fraction: mixed number 1.5 → '1 1/2' (GH-243)") {
    assertEquals(NumFmtFormatter.formatNumber(BigDecimal("1.5"), NumFmt.Fraction), "1 1/2")
  }

  test("NumFmt.Fraction: 5.25 → '5 1/4' (MS docs example, GH-243)") {
    assertEquals(NumFmtFormatter.formatNumber(BigDecimal("5.25"), NumFmt.Fraction), "5 1/4")
  }

  test("NumFmt.Fraction: 3.14159 → '3 1/7' (GH-243)") {
    assertEquals(NumFmtFormatter.formatNumber(BigDecimal("3.14159"), NumFmt.Fraction), "3 1/7")
  }

  test(
    "NumFmt.Fraction: 0.3 → ' 2/7' — double of 0.3 sits below 3/10 so the CF path differs from 12.3 (GH-243)"
  ) {
    assertEquals(NumFmtFormatter.formatNumber(BigDecimal("0.3"), NumFmt.Fraction), " 2/7")
  }

  test("Excel corpus: # ?/? values from SSF fraction.json (GH-243)") {
    val cases = Vector(
      BigDecimal("1") -> "1    ",
      BigDecimal("-1.2") -> "-1 1/5",
      BigDecimal("12.3") -> "12 1/3",
      BigDecimal("-12.34") -> "-12 1/3",
      BigDecimal("123.45") -> "123 4/9",
      BigDecimal("-123.456") -> "-123 1/2",
      BigDecimal("1234.567") -> "1234 4/7",
      BigDecimal("-1234.5678") -> "-1234 4/7",
      BigDecimal("12345.6789") -> "12345 2/3",
      BigDecimal("-12345.67891") -> "-12345 2/3"
    )
    cases.foreach { case (value, expected) =>
      assertEquals(
        NumFmtFormatter.formatNumber(value, NumFmt.Custom("# ?/?")),
        expected,
        s"value $value"
      )
    }
  }

  test("Excel corpus: # ??/?? values from SSF fraction.json (GH-243)") {
    val cases = Vector(
      BigDecimal("1") -> "1      ",
      BigDecimal("-1.2") -> "-1  1/5 ",
      BigDecimal("12.3") -> "12  3/10",
      BigDecimal("-12.34") -> "-12 17/50",
      BigDecimal("123.45") -> "123  9/20",
      BigDecimal("-123.456") -> "-123 26/57",
      BigDecimal("1234.567") -> "1234 55/97",
      BigDecimal("-1234.5678") -> "-1234 46/81",
      BigDecimal("12345.6789") -> "12345 55/81",
      BigDecimal("-12345.67891") -> "-12345 55/81"
    )
    cases.foreach { case (value, expected) =>
      assertEquals(
        NumFmtFormatter.formatNumber(value, NumFmt.Custom("# ??/??")),
        expected,
        s"value $value"
      )
    }
  }

  test("Excel corpus: fixed denominators # ?/2 and # ?/4 from SSF fraction.json (GH-243)") {
    val halfCases = Vector(
      BigDecimal("1") -> "1    ",
      BigDecimal("-1.2") -> "-1    ",
      BigDecimal("12.3") -> "12 1/2",
      BigDecimal("-12.34") -> "-12 1/2"
    )
    halfCases.foreach { case (value, expected) =>
      assertEquals(
        NumFmtFormatter.formatNumber(value, NumFmt.Custom("# ?/2")),
        expected,
        s"value $value"
      )
    }
    val quarterCases = Vector(
      BigDecimal("-1.2") -> "-1 1/4",
      BigDecimal("123.45") -> "123 2/4"
    )
    quarterCases.foreach { case (value, expected) =>
      assertEquals(
        NumFmtFormatter.formatNumber(value, NumFmt.Custom("# ?/4")),
        expected,
        s"value $value"
      )
    }
  }

  test("NumFmt.Fraction: whole number blanks the fraction area: 2 → '2    ' (GH-243)") {
    assertEquals(NumFmtFormatter.formatNumber(BigDecimal("2"), NumFmt.Fraction), "2    ")
  }

  test("NumFmt.Fraction: zero → '0    ' (GH-243)") {
    assertEquals(NumFmtFormatter.formatNumber(BigDecimal("0"), NumFmt.Fraction), "0    ")
  }

  test("NumFmt.Fraction: negative mixed -1.5 → '-1 1/2' (GH-243)") {
    assertEquals(NumFmtFormatter.formatNumber(BigDecimal("-1.5"), NumFmt.Fraction), "-1 1/2")
  }

  test("NumFmt.Fraction: negative pure fraction -0.5 → '- 1/2' (GH-243)") {
    assertEquals(NumFmtFormatter.formatNumber(BigDecimal("-0.5"), NumFmt.Fraction), "- 1/2")
  }

  test("custom fraction: # ??/?? gives 5.3 → '5  3/10' (MS docs example, GH-243)") {
    val fmt = NumFmt.Custom("# ??/??")
    assertEquals(NumFmtFormatter.formatNumber(BigDecimal("5.3"), fmt), "5  3/10")
  }

  test("custom fraction: # ??/?? pads numerator left, denominator right: 0.5 → '  1/2 ' (GH-243)") {
    val fmt = NumFmt.Custom("# ??/??")
    assertEquals(NumFmtFormatter.formatNumber(BigDecimal("0.5"), fmt), "  1/2 ")
  }

  test("custom fraction: fixed denominator # ?/8 does not reduce: 0.5 → ' 4/8' (GH-243)") {
    val fmt = NumFmt.Custom("# ?/8")
    assertEquals(NumFmtFormatter.formatNumber(BigDecimal("0.5"), fmt), " 4/8")
  }

  test("custom fraction: fixed denominator rounds into the whole part: 0.96 → '1    ' (GH-243)") {
    val fmt = NumFmt.Custom("# ?/8")
    assertEquals(NumFmtFormatter.formatNumber(BigDecimal("0.96"), fmt), "1    ")
  }

  test("custom fraction: fixed denominator # ?/8: 0.3 → ' 2/8' (GH-243)") {
    val fmt = NumFmt.Custom("# ?/8")
    assertEquals(NumFmtFormatter.formatNumber(BigDecimal("0.3"), fmt), " 2/8")
  }

  test("custom fraction: improper ?/? without whole part: 1.5 → '3/2' (GH-243)") {
    val fmt = NumFmt.Custom("?/?")
    assertEquals(NumFmtFormatter.formatNumber(BigDecimal("1.5"), fmt), "3/2")
  }

  test("custom fraction: negative section applies: # ?/?;(# ?/?) → -1.5 → '(1 1/2)' (GH-243)") {
    val fmt = NumFmt.Custom("# ?/?;(# ?/?)")
    assertEquals(NumFmtFormatter.formatNumber(BigDecimal("-1.5"), fmt), "(1 1/2)")
  }

  // Totality: BigDecimal holds values beyond Double range (the evaluator can produce
  // >1E308). The convergent search runs on doubles, so out-of-range magnitudes must
  // bypass it and render the whole-number form (blank fraction area) — the same shape
  // huge-but-finite values already produce — instead of throwing.

  test("fraction totality: 1E400 (beyond Double range) renders the whole form, no throw (GH-243)") {
    val expected = "1" + "0" * 400 + "    "
    assertEquals(NumFmtFormatter.formatNumber(BigDecimal("1E400"), NumFmt.Fraction), expected)
    // Consistency pin: huge-but-finite values take the convergent path to the same shape.
    assertEquals(
      NumFmtFormatter.formatNumber(BigDecimal("1E20"), NumFmt.Fraction),
      "100000000000000000000    "
    )
  }

  test("fraction totality: custom # ?/? with 1E400 renders the whole form, no throw (GH-243)") {
    val expected = "1" + "0" * 400 + "    "
    assertEquals(
      NumFmtFormatter.formatNumber(BigDecimal("1E400"), NumFmt.Custom("# ?/?")),
      expected
    )
  }

  test("fraction totality: # ??/?? with -1E400 keeps the default minus, no throw (GH-243)") {
    val expected = "-1" + "0" * 400 + "      "
    assertEquals(
      NumFmtFormatter.formatNumber(BigDecimal("-1E400"), NumFmt.Custom("# ??/??")),
      expected
    )
  }

  // ========== English Month/Weekday Tables (exhaustive) ==========
  // These pin the exact strings the hardcoded English tables must produce — previously
  // TextStyle.getDisplayName(…, Locale.US), which Scala Native's locale data renders
  // differently (ADR-016). Every reachable (token, value) pair is asserted.

  test("applyDateFormat: mmm — all 12 English month abbreviations") {
    val expected =
      Vector("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    val code = FormatCodeParser.parse("mmm").toOption.get
    for month <- 1 to 12 do
      val dt = java.time.LocalDateTime.of(2025, month, 15, 0, 0, 0)
      assertEquals(FormatCodeParser.applyDateFormat(dt, code), expected(month - 1))
  }

  test("applyDateFormat: mmmm — all 12 English month names") {
    val expected = Vector(
      "January",
      "February",
      "March",
      "April",
      "May",
      "June",
      "July",
      "August",
      "September",
      "October",
      "November",
      "December"
    )
    val code = FormatCodeParser.parse("mmmm").toOption.get
    for month <- 1 to 12 do
      val dt = java.time.LocalDateTime.of(2025, month, 15, 0, 0, 0)
      assertEquals(FormatCodeParser.applyDateFormat(dt, code), expected(month - 1))
  }

  test("applyDateFormat: mmmmm — all 12 English month initials") {
    val expected = Vector("J", "F", "M", "A", "M", "J", "J", "A", "S", "O", "N", "D")
    val code = FormatCodeParser.parse("mmmmm").toOption.get
    for month <- 1 to 12 do
      val dt = java.time.LocalDateTime.of(2025, month, 15, 0, 0, 0)
      assertEquals(FormatCodeParser.applyDateFormat(dt, code), expected(month - 1))
  }

  test("applyDateFormat: ddd — all 7 English weekday abbreviations") {
    val expected = Vector("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")
    val code = FormatCodeParser.parse("ddd").toOption.get
    for offset <- 0 to 6 do
      // 2024-01-01 is a Monday
      val dt = java.time.LocalDateTime.of(2024, 1, 1, 0, 0, 0).plusDays(offset.toLong)
      assertEquals(FormatCodeParser.applyDateFormat(dt, code), expected(offset))
  }

  test("applyDateFormat: dddd — all 7 English weekday names") {
    val expected =
      Vector("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")
    val code = FormatCodeParser.parse("dddd").toOption.get
    for offset <- 0 to 6 do
      val dt = java.time.LocalDateTime.of(2024, 1, 1, 0, 0, 0).plusDays(offset.toLong)
      assertEquals(FormatCodeParser.applyDateFormat(dt, code), expected(offset))
  }

  // ========== isTextOnly (GH-501) ==========

  test("isTextOnly: a lone @ has no numeric section") {
    val code = FormatCodeParser.parse("@").toOption.get
    assert(FormatCodeParser.isTextOnly(code), "@ formats numbers as General text")
  }

  test("isTextOnly: any code with a numeric section is not text-only") {
    List(
      "0.00",
      "General",
      "0;0;\"val: \"@", // the trailing @ is the text arm of a numeric code (GH-285)
      "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)" // 4 sections keep all 4
    ).foreach { code =>
      val fmt = FormatCodeParser.parse(code).toOption.get
      assert(!FormatCodeParser.isTextOnly(fmt), s"'$code' routes numbers through a section")
    }
  }

  // ========== General keyword inside a section (GH-666) ==========
  // ECMA-376 §18.8.30: `General` is a keyword wherever it appears in a section, not seven
  // literal letters. Excel renders `General"A"` on 2021 as `2021A` (the FY-suffix idiom).

  test("parse: General keyword lexes as FormatToken.General, case-insensitively (GH-666)") {
    List("General\"A\"", "general\"A\"", "GENERAL\"A\"").foreach { code =>
      val fmt = FormatCodeParser.parse(code).toOption.get
      assertEquals(
        fmt.positive.pattern.tokens,
        Vector(FormatToken.General, FormatToken.Literal("A")),
        s"'$code' should lex General as one token"
      )
    }
  }

  test("applyFormat: General\"A\" on 2021 renders 2021A, not GeneralA (GH-666)") {
    val code = FormatCodeParser.parse("General\"A\"").toOption.get
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(2021), code)._1, "2021A")
    val e = FormatCodeParser.parse("General\"E\"").toOption.get
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(2021), e)._1, "2021E")
  }

  test("applyFormat: \"FY\"General and General\" units\" wrap the General rendering (GH-666)") {
    val fy = FormatCodeParser.parse("\"FY\"General").toOption.get
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(2021), fy)._1, "FY2021")
    val units = FormatCodeParser.parse("General\" units\"").toOption.get
    assertEquals(FormatCodeParser.applyFormat(BigDecimal("12.5"), units)._1, "12.5 units")
    // General inside a section keeps General's own rules: no forced decimals, trailing zeros gone
    assertEquals(FormatCodeParser.applyFormat(BigDecimal("1234.50"), units)._1, "1234.5 units")
  }

  test("applyFormat: General inside a single section keeps the default leading minus (GH-666)") {
    // Same mechanism as `"FY"0` on -2021 → -FY2021: the pattern renders |x|, applyFormat signs
    val suffix = FormatCodeParser.parse("General\"A\"").toOption.get
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(-2021), suffix)._1, "-2021A")
    val prefix = FormatCodeParser.parse("\"FY\"General").toOption.get
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(-2021), prefix)._1, "-FY2021")
    // Multi-section: the negative section owns its sign
    val paren = FormatCodeParser.parse("General;(General)").toOption.get
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(-2021), paren)._1, "(2021)")
  }

  test("parse: a quoted \"General\" and an escaped \\G stay literal text (GH-666)") {
    val quoted = FormatCodeParser.parse("\"General\"0").toOption.get
    assertEquals(
      quoted.positive.pattern.tokens,
      Vector(FormatToken.Literal("General"), FormatToken.Digit('0'))
    )
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(5), quoted)._1, "General5")
    val escaped = FormatCodeParser.parse("\\General0").toOption.get
    assert(
      !escaped.positive.pattern.tokens.contains(FormatToken.General),
      s"escaped G must not start the keyword: ${escaped.positive.pattern.tokens}"
    )
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(5), escaped)._1, "General5")
  }

  test("parse: date sections are untouched by the General keyword lexer (GH-666)") {
    val dt = java.time.LocalDateTime.of(2021, 3, 4, 13, 5, 6)
    List(
      "mmm d, yyyy" -> "Mar 4, 2021",
      "m/d/yy" -> "3/4/21",
      "dddd, mmmm d" -> "Thursday, March 4",
      "h:mm:ss AM/PM" -> "1:05:06 PM"
    ).foreach { case (code, expected) =>
      val fmt = FormatCodeParser.parse(code).toOption.get
      assert(!fmt.positive.pattern.tokens.contains(FormatToken.General), s"'$code' lexed General")
      assertEquals(FormatCodeParser.applyDateFormat(dt, fmt), expected, code)
    }
  }

  test("parse: whole-code General is one General token and renders like General (GH-666)") {
    val fmt = FormatCodeParser.parse("General").toOption.get
    assertEquals(fmt.positive.pattern.tokens, Vector(FormatToken.General))
    assertEquals(FormatCodeParser.applyFormat(BigDecimal("1234.5"), fmt)._1, "1234.5")
    assertEquals(FormatCodeParser.applyFormat(BigDecimal("-0.156"), fmt)._1, "-0.156")
    assert(!FormatCodeParser.isTextOnly(fmt))
  }

  test("applyTextFormat: General in the text section echoes the text (GH-666)") {
    val fmt = FormatCodeParser.parse("0;-0;0;General").toOption.get
    assertEquals(FormatCodeParser.applyTextFormat("abc", fmt), "abc")
  }

  // ========== Thousands-scaling commas (GH-666) ==========
  // Commas after the last digit placeholder of a section divide by 1000 each (Excel/LO:
  // `$#,##0.0,,"mm"` on 1,500,000 is `$1.5mm`); commas between digits keep grouping.

  test("parse: trailing commas lex as Scale, interior commas stay Thousands (GH-666)") {
    val fmt = FormatCodeParser.parse("#,##0.0,,\"mm\"").toOption.get
    val tokens = fmt.positive.pattern.tokens
    assertEquals(tokens.count(_ == FormatToken.Scale), 2)
    assertEquals(tokens.count(_ == FormatToken.Thousands), 1)
    assert(fmt.positive.pattern.hasThousands, "interior comma still groups")
    val bare = FormatCodeParser.parse("0,").toOption.get
    assertEquals(bare.positive.pattern.tokens, Vector(FormatToken.Digit('0'), FormatToken.Scale))
    assert(!bare.positive.pattern.hasThousands, "a lone scaling comma is not grouping")
  }

  test("applyFormat: $#,##0.0,,\"mm\" on 1,500,000 renders $1.5mm (GH-666)") {
    val code = FormatCodeParser.parse("$#,##0.0,,\"mm\"").toOption.get
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(1500000), code)._1, "$1.5mm")
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(1234567890), code)._1, "$1,234.6mm")
  }

  test("applyFormat: #,##0.0,\"k\" scales by one thousand then rounds and groups (GH-666)") {
    val code = FormatCodeParser.parse("#,##0.0,\"k\"").toOption.get
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(1234567), code)._1, "1,234.6k")
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(999950), code)._1, "1,000.0k")
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(0), code)._1, "0.0k")
  }

  test("applyFormat: 0, drops grouping and scales; 0.0,, on negatives keeps the sign (GH-666)") {
    val thousands = FormatCodeParser.parse("0,").toOption.get
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(1234567), thousands)._1, "1235")
    val millions = FormatCodeParser.parse("0.0,,").toOption.get
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(-1500000), millions)._1, "-1.5")
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(-1500000000), millions)._1, "-1500.0")
  }

  test("applyFormat: scaled multi-section code routes by the raw value, like Excel (GH-666)") {
    // Excel picks the section from the stored value (SSF choose_fmt): only a true zero fires
    // the dash section; 400 under a millions scale rounds to 0.0 in the positive section — the
    // same rule that makes Excel display "-0.00" for -0.001 under 0.00. LibreOffice 25.8
    // oracle (#672): 400 → 0.0, -400 → (0.0), 0 → -.
    val code = FormatCodeParser.parse("0.0,,;(0.0,,);\"-\"").toOption.get
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(0), code)._1, "-")
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(1500000), code)._1, "1.5")
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(-1500000), code)._1, "(1.5)")
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(400), code)._1, "0.0")
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(-400), code)._1, "(0.0)")
  }

  test(
    "applyFormat: scaling commas before %, padding, section end; percent alone unscaled (GH-666)"
  ) {
    val pct = FormatCodeParser.parse("0.0,%").toOption.get
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(1234), pct)._1, "123.4%")
    val padded = FormatCodeParser.parse("#,##0,_);(#,##0,)").toOption.get
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(1234567), padded)._1, "1,235 ")
    assertEquals(FormatCodeParser.applyFormat(BigDecimal(-1234567), padded)._1, "(1,235)")
    val plainPct = FormatCodeParser.parse("0.0%").toOption.get
    assertEquals(FormatCodeParser.applyFormat(BigDecimal("0.155"), plainPct)._1, "15.5%")
    val grouped = FormatCodeParser.parse("#,##0.00").toOption.get
    assertEquals(
      FormatCodeParser.applyFormat(BigDecimal("1234567.891"), grouped)._1,
      "1,234,567.89"
    )
  }

  // ========== Comma roles past GH-666 (#672, #681) ==========
  // A comma after a digit placeholder (or another scaling comma) groups when a digit placeholder
  // of the INTEGER part follows it and scales by 1000 otherwise; the decimal point ends the
  // integer part, so `#,##0,.0` scales (Excel: "a comma that follows a digit placeholder scales
  // the number by 1,000"). A comma after anything else is literal text. LibreOffice 25.8 is the
  // oracle for the literal rows; it refuses `0,.0`-shaped codes outright (renders them General),
  // so the before-the-point rows follow Excel's documented rule.

  private def fmt(code: String, n: BigDecimal): String =
    FormatCodeParser.applyFormat(n, FormatCodeParser.parse(code).toOption.get)._1

  test("parse: a comma before the decimal point scales; the interior comma still groups (#672)") {
    val tokens = FormatCodeParser.parse("#,##0,.0").toOption.get.positive.pattern.tokens
    assertEquals(tokens.count(_ == FormatToken.Scale), 1)
    assertEquals(tokens.count(_ == FormatToken.Thousands), 1)
    val twice = FormatCodeParser.parse("0,,.0").toOption.get.positive.pattern
    assertEquals(twice.tokens.count(_ == FormatToken.Scale), 2)
    assert(!twice.hasThousands, "scaling commas before the point are not grouping")
  }

  test("applyFormat: scaling commas before the decimal point divide by 1000 each (#672)") {
    assertEquals(fmt("#,##0,.0", BigDecimal(1234567)), "1,234.6")
    assertEquals(fmt("0,,.0", BigDecimal(123456789)), "123.5")
    assertEquals(fmt("0,,.0", BigDecimal(450000)), "0.5")
    assertEquals(fmt("0,.0", BigDecimal(1234567)), "1234.6")
    assertEquals(fmt("0,.0", BigDecimal(-1234567)), "-1234.6")
    assertEquals(fmt("0,.0", BigDecimal(999)), "1.0")
    assertEquals(fmt("#,##0,,.00", BigDecimal(123456789)), "123.46")
    assertEquals(fmt("#,##0,.00", BigDecimal("0.5")), "0.00")
    // before and after the point: both scale
    assertEquals(fmt("#,##0,.0,", BigDecimal(123456789)), "123.5")
    // a digit placeholder of the decimal part does not make a comma group
    assertEquals(fmt("0,.0;(0,.0)", BigDecimal(-2500)), "(2.5)")
  }

  test("applyFormat: a comma after a literal or a space is literal text (LO oracle, #672)") {
    // Was a grouping comma rendered nowhere: `0 ,` on 12345678 showed "12,345,678 "
    assertEquals(fmt("0 ,", BigDecimal(12345678)), "12345678 ,")
    assertEquals(fmt("0\" \",", BigDecimal(12345678)), "12345678 ,")
    assertEquals(fmt("0 ,", BigDecimal(1234)), "1234 ,")
    assertEquals(fmt("#,##0 ,", BigDecimal(1234567)), "1,234,567 ,")
    assertEquals(fmt("0,,\"x\",", BigDecimal(123456789)), "123x,")
    assertEquals(fmt("0%,", BigDecimal(12345)), "1234500%,")
    // a comma directly after a digit placeholder, then a literal, still scales
    assertEquals(fmt("0,\" \"", BigDecimal(12345678)), "12346 ")
    assertEquals(fmt("0,\"k\"", BigDecimal(12345)), "12k")
  }

  test("applyFormat: 0\"x\", keeps its comma as literal text (#681)") {
    // PR #679 review: the comma after a quoted literal was dropped (neither grouping nor scaling)
    assertEquals(fmt("0\"x\",", BigDecimal(12345678)), "12345678x,")
    assertEquals(fmt("0\"x\",", BigDecimal(1234)), "1234x,")
  }

  test("applyFormat: a scaling comma in a scientific or fraction section is inert (LO, #672)") {
    // LibreOffice 25.8: `0.0E+00,` and `0.0E+00,,` on 12345 are 1.2E+04; `# ?/?,` on 12.5 is
    // 12 1/2 — no division, no literal comma
    assertEquals(fmt("0.0E+00,", BigDecimal(12345)), "1.2E+04")
    assertEquals(fmt("0.0E+00,,", BigDecimal(12345)), "1.2E+04")
    assertEquals(fmt("# ?/?,", BigDecimal("12.5")), "12 1/2")
  }

  test("applyFormat: General mixed with digit placeholders renders the number once (#681)") {
    // PR #679 review: `General0` emitted the number twice (55, 1.52). Excel's behaviour on such a
    // code is undefined; LibreOffice 25.8 renders `General0` and `General.00` as the General value
    // alone (5, 1.5). xl's rule: the General keyword owns the number, and digit placeholders,
    // decimal points and commas beside it render nothing; literals still render.
    assertEquals(fmt("General0", BigDecimal(5)), "5")
    assertEquals(fmt("General0", BigDecimal("1.5")), "1.5")
    assertEquals(fmt("General.00", BigDecimal("1.5")), "1.5")
    assertEquals(fmt("0General", BigDecimal(5)), "5")
    assertEquals(fmt("\"a\"General0\"b\"", BigDecimal(5)), "a5b")
    assertEquals(fmt("General0", BigDecimal(-5)), "-5")
  }

  test("applyFormat: a comma after the General keyword is literal text (LO oracle, #681)") {
    // Was the number twice: General rendered 12345, then the comma emitted the grouped digits
    assertEquals(fmt("General,", BigDecimal(12345)), "12345,")
  }

  test("parse: commas in date and text sections render as commas (LO oracle, #672)") {
    // LibreOffice 25.8: `mmm d, yyyy` → Mar 4, 2021; "abc" under `0;-0;0;@,` and `@,` → abc,
    val dt = java.time.LocalDateTime.of(2021, 3, 4, 0, 0)
    val date = FormatCodeParser.parse("mmm d, yyyy").toOption.get
    assertEquals(FormatCodeParser.applyDateFormat(dt, date), "Mar 4, 2021")
    val text = FormatCodeParser.parse("0;-0;0;@,").toOption.get
    assertEquals(FormatCodeParser.applyTextFormat("abc", text), "abc,")
  }
