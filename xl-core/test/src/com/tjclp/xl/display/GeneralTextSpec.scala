package com.tjclp.xl.display

import com.tjclp.xl.Generators.{genExcelDateTime, genRoundTripNumber, genWideBigDecimal}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.richtext.RichText

import munit.ScalaCheckSuite
import org.scalacheck.Gen
import org.scalacheck.Prop.forAll

import java.time.LocalDate

/**
 * GH-665: `NumFmtFormatter.generalText` — Excel's width-independent number → text conversion (the
 * rule behind `&`, CONCATENATE, TEXT(x,"General") and the formula bar).
 *
 * Pinned rows follow Apache POI's `NumberToTextConverter`, reverse-engineered against Excel. The
 * three rows public evidence pins least firmly are `1E20` → `1E+20`, `1E-16` → `0.0000000000000001`
 * and `-1E19` → `-10000000000000000000` (sign uncounted); a disagreement with a live Excel
 * spot-check changes one constant or one guard in `generalText`, not the design. Human Excel
 * spot-check: pending — record results here when run.
 *
 * `formatGeneral` (cell display, width-dependent) is a different rule and stays pinned by
 * DisplaySpec.
 */
class GeneralTextSpec extends ScalaCheckSuite:

  private def gt(s: String): String = NumFmtFormatter.generalText(BigDecimal(s))

  private def pin(rows: (String, String)*)(implicit loc: munit.Location): Unit =
    rows.foreach { case (in, out) => assertEquals(gt(in), out, s"generalText($in)") }

  // ===== Pinned table =====

  test("issue rows: stored scale never leaks into text") {
    pin(
      "2.0" -> "2",
      "25.00" -> "25",
      "1070.0" -> "1070",
      "16.00" -> "16",
      "1E+3" -> "1000",
      "0.3333333333333333333333333333333333" -> "0.333333333333333",
      "0.6666666666666666666666666666666667" -> "0.666666666666667"
    )
  }

  test("every zero is \"0\"") {
    pin("0" -> "0", "0.0" -> "0", "0E-10" -> "0", "-0.0" -> "0", "0E+10" -> "0", "0.000" -> "0")
  }

  test("15 significant digits, HALF_UP") {
    pin(
      "123.45678901234568" -> "123.456789012346",
      "1234567.8901234567" -> "1234567.89012346",
      "1234567890.123456789" -> "1234567890.12346",
      "0.1234567890123455" -> "0.123456789012346", // exact tie → up
      "3.141592653589793" -> "3.14159265358979",
      "9007199254740992" -> "9007199254740990"
    )
  }

  test("POI Excel-verified boundary pairs: the small side is decided by LENGTH, not exponent") {
    pin(
      "0.000012345678901234568" -> "1.23456789012346E-05",
      "0.00012345678901234567" -> "0.000123456789012346", // exactly 20 characters, plain
      "0.0012345678901234567" -> "0.00123456789012346",
      "0.0000123456789" -> "0.0000123456789",
      "0.00001" -> "0.00001",
      "0.000012345" -> "0.000012345",
      "0.0000001" -> "0.0000001",
      "1E-10" -> "0.0000000001",
      "2.5E-7" -> "0.00000025",
      "0.00000095367431640625" -> "9.5367431640625E-07",
      "1E-16" -> "0.0000000000000001"
    )
  }

  test("POI Excel-verified boundary pairs: the integer side switches after 20 digits") {
    pin(
      "1E19" -> "10000000000000000000",
      "1152921504606846976" -> "1152921504606850000", // 2^60
      "18446744073709551616" -> "18446744073709600000", // 2^64
      "12345678901234567890" -> "12345678901234600000",
      "1E20" -> "1E+20",
      "1E21" -> "1E+21",
      "123456789012345678901" -> "1.23456789012346E+20"
    )
  }

  test("carry: rounding that adds a digit is decided after the round") {
    pin(
      "999999999999999.9" -> "1000000000000000",
      "99999999999999999999" -> "1E+20"
    )
  }

  test("sign is never counted toward the 20 characters") {
    pin(
      "-1E19" -> "-10000000000000000000",
      "-0.00012345678901234567" -> "-0.000123456789012346",
      "-2.0" -> "-2"
    )
  }

  test("serials render like GH-561's dateSerialText did") {
    pin("46023" -> "46023", "46023.5" -> "46023.5", "2958465.99999999" -> "2958465.99999999")
  }

  test("extreme scales stay in E form: no Int overflow chooses the plain form (PR #679 review)") {
    // scale = Int.MaxValue: the Int sum `2 + (-exp - 1) + sig` wrapped negative, passed the
    // `<= 20` test and toPlainString tried to build two billion zeros (OutOfMemoryError)
    pin(
      "1E-2147483647" -> "1E-2147483647",
      "-1E-2147483647" -> "-1E-2147483647",
      "123E-2147483647" -> "1.23E-2147483645",
      "1E+2147483647" -> "1E+2147483647",
      "1.5E+2147483646" -> "1.5E+2147483646"
    )
    // scale = Int.MinValue: 1E+2147483648 cannot even be spelled as a literal
    val minScale = BigDecimal(new java.math.BigDecimal(java.math.BigInteger.ONE, Int.MinValue))
    assertEquals(NumFmtFormatter.generalText(minScale), "1E+2147483648")
    assertEquals(NumFmtFormatter.generalText(-minScale), "-1E+2147483648")
  }

  test("absurd magnitude stays bounded: E form, never a million zeros") {
    pin(
      "1E-100" -> "1E-100",
      "1E+1000000" -> "1E+1000000",
      "1E-1000000" -> "1E-1000000",
      "-1E+1000000" -> "-1E+1000000"
    )
  }

  // ===== Laws =====

  private val grammar = "-?[0-9]+(\\.[0-9]+)?(E[+-][0-9]{2,})?".r

  private def round15(n: BigDecimal): java.math.BigDecimal =
    n.bigDecimal.round(new java.math.MathContext(15, java.math.RoundingMode.HALF_UP))

  property("L1 scale invariance: trailing-zero rescaling never changes the text") {
    forAll(genWideBigDecimal, Gen.choose(0, 12)) { (n: BigDecimal, k: Int) =>
      val rescaled = BigDecimal(n.bigDecimal.setScale(n.scale + k))
      assertEquals(NumFmtFormatter.generalText(rescaled), NumFmtFormatter.generalText(n), s"$n +$k")
    }
  }

  property("L2 reconstruction: parsing the text gives the 15-digit rounding of the input") {
    forAll(genWideBigDecimal) { (n: BigDecimal) =>
      if n.signum != 0 then
        val back = new java.math.BigDecimal(NumFmtFormatter.generalText(n))
        assertEquals(back.compareTo(round15(n)), 0, s"generalText($n) = $back")
    }
  }

  property("L3 sign: generalText(-n) == \"-\" + generalText(n) for n != 0; every zero is \"0\"") {
    forAll(genWideBigDecimal) { (n: BigDecimal) =>
      if n.signum == 0 then assertEquals(NumFmtFormatter.generalText(n), "0")
      else
        assertEquals(
          NumFmtFormatter.generalText(-n.abs),
          "-" + NumFmtFormatter.generalText(n.abs),
          s"$n"
        )
    }
  }

  property("L4 grammar and bound: plain form ≤ 20 unsigned chars, E mantissa ≤ 16 chars") {
    forAll(genWideBigDecimal) { (n: BigDecimal) =>
      val s = NumFmtFormatter.generalText(n)
      assert(grammar.matches(s), s"grammar: $s")
      val unsigned = s.stripPrefix("-")
      if unsigned.contains('E') then
        val mantissa = unsigned.takeWhile(_ != 'E')
        assert(mantissa.length <= 16, s"mantissa too long: $s")
        assert(!mantissa.endsWith("0") || mantissa == "0", s"mantissa keeps trailing zero: $s")
      else assert(unsigned.length <= 20, s"plain form too long: $s")
    }
  }

  property("L5 idempotence: generalText(BigDecimal(generalText(n))) == generalText(n)") {
    forAll(genWideBigDecimal) { (n: BigDecimal) =>
      val once = NumFmtFormatter.generalText(n)
      assertEquals(NumFmtFormatter.generalText(BigDecimal(once)), once, s"$n")
    }
  }

  property("L6 whole numbers with ≤ 20 digits never contain E") {
    val genWhole: Gen[BigDecimal] =
      for
        len <- Gen.choose(1, 20)
        first <- Gen.choose(1, 9)
        rest <- Gen.listOfN(len - 1, Gen.choose(0, 9))
        negate <- Gen.oneOf(true, false)
      yield
        val v = BigDecimal((first :: rest).mkString)
        if negate then -v else v
    forAll(genWhole) { (n: BigDecimal) =>
      val s = NumFmtFormatter.generalText(n)
      assert(!s.contains('E'), s"$n rendered as $s")
      assertEquals(BigDecimal(s).compare(BigDecimal(round15(n))), 0, s"$n → $s")
    }
  }

  property("L7 GH-561 equivalence: Excel date serials render exactly as the old three-liner") {
    val genSerial: Gen[BigDecimal] =
      Gen.frequency(
        6 -> genExcelDateTime.map(dt => BigDecimal(CellValue.dateTimeToExcelSerial(dt))),
        2 -> Gen.choose(0L, 2958465L).map(BigDecimal(_)),
        1 -> Gen.oneOf(
          BigDecimal(0),
          BigDecimal("0.5"),
          BigDecimal("46023.5"),
          BigDecimal("2958465.99999999")
        )
      )
    forAll(genSerial) { (s: BigDecimal) =>
      val old = s.round(new java.math.MathContext(15)).bigDecimal.stripTrailingZeros.toPlainString
      assertEquals(NumFmtFormatter.generalText(s), old, s"serial $s")
    }
  }

  property("plain-range numbers agree with round-then-plain") {
    forAll(genRoundTripNumber) { (n: BigDecimal) =>
      val expected =
        if n.signum == 0 then "0"
        else
          val r = round15(n).stripTrailingZeros
          val plain = r.toPlainString
          if plain.stripPrefix("-").length <= 20 then plain else NumFmtFormatter.generalText(n)
      assertEquals(NumFmtFormatter.generalText(n), expected, s"$n")
    }
  }

  // ===== Whole-value form =====

  test("generalText(CellValue): every shape") {
    val gt = NumFmtFormatter.generalText(_: CellValue)
    assertEquals(gt(CellValue.Number(BigDecimal("2.0"))), "2")
    assertEquals(gt(CellValue.DateTime(LocalDate.of(2026, 1, 1).atStartOfDay())), "46023")
    assertEquals(gt(CellValue.DateTime(LocalDate.of(2026, 1, 1).atTime(12, 0))), "46023.5")
    assertEquals(gt(CellValue.Bool(true)), "TRUE")
    assertEquals(gt(CellValue.Bool(false)), "FALSE")
    assertEquals(gt(CellValue.Text("2.50")), "2.50")
    assertEquals(gt(CellValue.RichText(RichText.plain("rich"))), "rich")
    assertEquals(gt(CellValue.Empty), "")
    assertEquals(gt(CellValue.Error(CellError.Div0)), "#DIV/0!")
    assertEquals(
      gt(CellValue.Formula("A1*2", Some(CellValue.Number(BigDecimal("1070.0"))))),
      "1070"
    )
    assertEquals(gt(CellValue.Formula("A1*2", None)), "")
  }

  test("isGeneralCode is the whole-code keyword, case-insensitive") {
    assert(NumFmtFormatter.isGeneralCode("General"))
    assert(NumFmtFormatter.isGeneralCode("GENERAL"))
    assert(NumFmtFormatter.isGeneralCode("general"))
    assert(!NumFmtFormatter.isGeneralCode("0.00"))
    assert(!NumFmtFormatter.isGeneralCode("General;General"))
  }
