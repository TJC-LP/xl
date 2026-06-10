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
      case _                        => false
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
      case _                        => false
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
      case _                       => false
    }
    assert(hasSpacer, "Should have spacer token")
  }

  test("parse: escaped character \\$") {
    val result = FormatCodeParser.parse("\\$#,##0")
    assert(result.isRight)
    val code = result.toOption.get

    val hasLiteralDollar = code.positive.pattern.tokens.exists {
      case FormatToken.Literal("$") => true
      case _                        => false
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
      case _                   => false
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
    // Trailing comma scales by 1000
    val result = FormatCodeParser.parse("#,##0,")
    assert(result.isRight, "Should parse thousands scaling format")
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

    val timeResult = FormatCodeParser.applyDateFormat(dt, timeCode)
    assertEquals(timeResult, "2:30:45") // mm = 30 (minute)

    val dateResult = FormatCodeParser.applyDateFormat(dt, dateCode)
    assertEquals(dateResult, "3/15/2025") // m = 3 (month)
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
    assert(!FormatCodeParser.hasDateTokens(numCode), "Should not detect date tokens in number format")
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

  test("NumFmt.Fraction: 0.3 → ' 2/7' — double of 0.3 sits below 3/10 so the CF path differs from 12.3 (GH-243)") {
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
