package com.tjclp.xl.display

import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.display.NumFmtFormatter.GeneralRule
import com.tjclp.xl.styles.numfmt.NumFmt

import munit.ScalaCheckSuite
import org.scalacheck.Gen
import org.scalacheck.Prop.{forAll, forAllNoShrink}

import java.math.{BigInteger, MathContext, RoundingMode}

/**
 * #689: custom number formats on stored values whose exponent sits at either end of the Int scale
 * range. A file can carry `<v>1E-2147483647</v>`; `0.00` asked BigDecimal.setScale for
 * 10^2147483645 and threw, so cell display broke the purity charter and TEXT surfaced an internal
 * evaluator defect. Every pattern kind (digits, percent, scaling commas, the General keyword,
 * scientific, fractions) now renders any stored exponent in bounded time and memory.
 */
class ExtremeExponentFormatSpec extends ScalaCheckSuite:

  private def fmt(code: String, n: BigDecimal, rule: GeneralRule = GeneralRule.CellDisplay) =
    NumFmtFormatter.formatValue(CellValue.Number(n), NumFmt.Custom(code), rule)

  /** unscaled × 10^-scale, for scales no literal can spell (Int.MinValue). */
  private def bd(unscaled: String, scale: Int): BigDecimal =
    BigDecimal(new java.math.BigDecimal(new BigInteger(unscaled), scale))

  private val tiny = BigDecimal("1E-2147483647")
  private val huge = BigDecimal("1E+2147483647")
  private val beyondInt = bd("1", Int.MinValue) // 1E+2147483648

  private val bothRules = List(GeneralRule.CellDisplay, GeneralRule.Text)

  // ===== The issue's examples =====

  test("#689: 0.00 on 1E-2147483647 rounds to 0.00 instead of throwing") {
    bothRules.foreach(rule => assertEquals(fmt("0.00", tiny, rule), "0.00", s"$rule"))
  }

  test("#689: General;-General and General% on 1E±2147483647 render the General form") {
    bothRules.foreach { rule =>
      assertEquals(fmt("General;-General", huge, rule), "1E+2147483647", s"$rule")
      assertEquals(fmt("General;-General", tiny, rule), "1E-2147483647", s"$rule")
      assertEquals(fmt("General;-General", -huge, rule), "-1E+2147483647", s"$rule")
      // percent moves the exponent past Int: the keyword renders it from a Long
      assertEquals(fmt("General%", huge, rule), "1E+2147483649%", s"$rule")
      assertEquals(fmt("General%", tiny, rule), "1E-2147483645%", s"$rule")
    }
  }

  // ===== Digit patterns =====

  test("#689: a value too small to show renders exactly like one that rounds to zero") {
    List("0", "0.00", "#,##0", "0%", "0.0,,", "0,", "\"$\"#,##0.00", "#,##0.00;(#,##0.00)")
      .foreach { code =>
        assertEquals(fmt(code, tiny), fmt(code, BigDecimal("0.0000001")), code)
        assertEquals(fmt(code, -tiny), fmt(code, BigDecimal("-0.0000001")), code)
        assertEquals(fmt(code, bd("12345678901234567890", Int.MaxValue)), fmt(code, 0), code)
      }
    assertEquals(fmt("0.00", tiny), "0.00")
    assertEquals(fmt("0%", tiny), "0%")
  }

  test("#689: a digit block past MaxDigitBlockLength renders in General form, literals kept") {
    assertEquals(fmt("0", huge), "1E+2147483647")
    assertEquals(fmt("0.00", -huge), "-1E+2147483647")
    assertEquals(fmt("#,##0", beyondInt), "1E+2147483648")
    assertEquals(fmt("\"$\"#,##0.00;(\"$\"#,##0.00)", -huge), "($1E+2147483647)")
    assertEquals(fmt("0%", huge), "1E+2147483649%")
    // the scaling comma divides before the bound is read
    assertEquals(fmt("0.0,,", huge), "1E+2147483641")
    // under TEXT the fallback is the text conversion: 15 significant digits
    assertEquals(
      fmt("0", bd("12345678901234567890", Int.MinValue), GeneralRule.Text),
      "1.23456789012346E+2147483667"
    )
  }

  test("#689: the bound is a cell's 32,767 characters, grouping commas included") {
    val max = FormatCodeParser.MaxDigitBlockLength
    assertEquals(max, 32767)
    // every number Excel can store still renders in full (GH-243's 1E400 pin is beyond even that)
    assertEquals(fmt("0", BigDecimal("1E+400")), "1" + "0" * 400)
    assertEquals(fmt("0", BigDecimal(10).pow(max - 1)), "1" + "0" * (max - 1))
    assertEquals(fmt("0", BigDecimal(10).pow(max)), s"1E+$max")
    // #,##0: 24,576 digits group to exactly 32,767 characters; one more digit passes the bound
    assertEquals(fmt("#,##0", BigDecimal(10).pow(24575)).length, max)
    assertEquals(fmt("#,##0", BigDecimal(10).pow(24576)), "1E+24576")
    // a rounding carry that adds the 32,768th digit falls back too (99…9.5 → 1E+32767); exact
    // java arithmetic, as Scala's subtraction would round the 0.5 away at 34 digits
    val nines = java.math.BigDecimal.TEN.pow(max).subtract(new java.math.BigDecimal("0.5"))
    assertEquals(fmt("0", BigDecimal(nines)), s"1E+$max")
    assertEquals(fmt("0", BigDecimal(nines.subtract(java.math.BigDecimal.ONE))).length, max)
  }

  test("#689: scaling commas and percent move the exponent exactly, past the Int range") {
    assertEquals(fmt("0,", tiny), "0")
    assertEquals(fmt("0,General", tiny), "1E-2147483650")
    assertEquals(fmt("0,General", beyondInt), "1E+2147483645")
    assertEquals(fmt("0%", beyondInt), "1E+2147483650%")
  }

  // ===== Scientific and fraction patterns =====

  test("#689: scientific patterns keep a Long exponent at both ends of the scale range") {
    assertEquals(fmt("0.00E+00", beyondInt), "1.00E+2147483648")
    assertEquals(fmt("0.00E+00", bd("12345678901234567890", Int.MinValue)), "1.23E+2147483667")
    assertEquals(fmt("0.00E+00", tiny), "1.00E-2147483647")
    // engineering notation: the Int arithmetic wrapped and printed 10.0E+2147483647
    assertEquals(fmt("##0.0E+0", beyondInt), "100.0E+2147483646")
    assertEquals(fmt("##0.0E+0", tiny), "100.0E-2147483649")
    assertEquals(fmt("0.00E+00%", huge), "1.00E+2147483649%")
  }

  test("#689: fraction patterns round tiny values to zero and huge ones to the General form") {
    assertEquals(fmt("?/8", tiny), "0/8")
    assertEquals(fmt("# ?/8", tiny), "0    ")
    assertEquals(fmt("# ?/?", tiny), "0    ")
    assertEquals(fmt("# ?/?", huge), "1E+2147483647    ")
    assertEquals(fmt("# ?/8", -huge), "-1E+2147483647    ")
    assertEquals(fmt("?/8", huge), "8E+2147483647/8")
    assertEquals(fmt("?/?", beyondInt), "1E+2147483648/1")
  }

  // ===== @ in a numeric section =====

  test("#689: @ in a numeric section renders the number as General, like a lone @") {
    // `@;@`: the trailing @ is the text section, so numbers take the first, an `@`. SSF (the
    // Excel-corpus-tested oracle) emits the value for it; the empty string dropped the number.
    // Excel itself unverified.
    assertEquals(fmt("@;@", BigDecimal("123456789012"), GeneralRule.Text), "123456789012")
    assertEquals(fmt("@;@", BigDecimal("123456789012")), "1.23457E+11")
    assertEquals(
      fmt("@;@", BigDecimal("123456789012"), GeneralRule.Text),
      fmt("@", BigDecimal("123456789012"), GeneralRule.Text)
    )
    assertEquals(fmt("@;@", BigDecimal(-5)), "-5")
    assertEquals(fmt("\"n=\"@;@", BigDecimal("2.50")), "n=2.5")
    // `@;0`: a lone `@` positive arm renders the number; negatives take the `0` arm, unsigned
    assertEquals(fmt("@;0", BigDecimal("123456789012"), GeneralRule.Text), "123456789012")
    assertEquals(fmt("@;0", BigDecimal("123456789012")), "1.23457E+11")
    assertEquals(fmt("@;0", BigDecimal("-5.4")), "5")
    // the text arm of a numeric code still never receives numbers
    assertEquals(fmt("0;@", BigDecimal(5)), "5")
    assertEquals(fmt("@;@", huge), "1E+2147483647")
  }

  test("#689: @ beside digit placeholders or General renders nothing; the number shows once") {
    // `@` stands in for the number only where nothing else in the section renders it. Beside
    // digits it renders nothing, as before #689: `0.00 @` keeps its digits (not ` 1.5`)
    assertEquals(fmt("0.00 @;0", BigDecimal("1.5")), "1.50 ")
    assertEquals(fmt("0.00 @;0", BigDecimal("1.5"), GeneralRule.Text), "1.50 ")
    assertEquals(fmt("0.00 @;-0.00 @;0", BigDecimal("-1.5")), "-1.50 ")
    assertEquals(fmt("#,##0 @;0", huge), "1E+2147483647 ")
    // beside the General keyword the value renders once (not `1.5 1.5`)
    assertEquals(fmt("General @;0", BigDecimal("1.5")), "1.5 ")
    assertEquals(fmt("@ General;0", BigDecimal("123456789012"), GeneralRule.Text), " 123456789012")
    // a decimal point alone is a digit pattern: it renders the number, `@` adds nothing
    assertEquals(fmt(".@;0", BigDecimal(2)), "2.")
  }

  // ===== Laws =====

  /** Formats the acceptance names, plus mixed sections, conditions, scaling and fractions. */
  private val codes = List(
    "0",
    "0.00",
    "#,##0",
    "0%",
    "0.00E+00",
    "##0.0E+0",
    "General%",
    "General;-General",
    "#,##0.00;(#,##0.00);\"zero\";@",
    "\"$\"#,##0_);[Red](\"$\"#,##0)",
    "[>1000]#,##0.0,\"K\";[<-1000]-#,##0.0,\"K\";0",
    "0.0,,\"mm\"",
    "0,General",
    "# ?/?",
    "?/8",
    "@;@",
    "yyyy-mm-dd;@"
  )

  /** Unscaled values up to 40 digits (zero included) across the whole Int scale range. */
  private val genExtreme: Gen[BigDecimal] =
    for
      digits <- Gen.choose(1, 40)
      unscaled <- Gen.listOfN(digits, Gen.choose(0, 9)).map(ds => BigInt(ds.mkString))
      scale <- Gen.oneOf(
        Gen.choose(Int.MinValue, Int.MinValue + 1000),
        Gen.choose(Int.MaxValue - 1000, Int.MaxValue),
        Gen.choose(-1000000, 1000000),
        Gen.choose(-33000, -24000),
        Gen.choose(-400, 400),
        Gen.choose(-20, 20)
      )
      negative <- Gen.oneOf(true, false)
    yield
      val magnitude = BigDecimal(new java.math.BigDecimal(unscaled.bigInteger, scale))
      if negative then -magnitude else magnitude

  // No shrinking over extreme scales: halving a value never moves its scale, and a failing case
  // re-run through ScalaCheck's Fractional shrinker for minutes instead of failing
  property("#689: every code renders every stored exponent, bounded in length and time") {
    forAllNoShrink(genExtreme) { (n: BigDecimal) =>
      for code <- codes; rule <- bothRules do
        val start = System.nanoTime
        val text = fmt(code, n, rule)
        val millis = (System.nanoTime - start) / 1000000
        // the digit block plus what the code itself spells (literals, decimals, padding)
        assert(
          text.length <= FormatCodeParser.MaxDigitBlockLength + 64,
          s"$code on $n: ${text.length} chars"
        )
        // generous on purpose: only a genuine hang trips it (a loaded CI runner must not)
        assert(millis < 10000, s"$code on $n under $rule took ${millis}ms")
    }
  }

  property("#689: trailing-zero rescaling never changes the rendering") {
    forAllNoShrink(genExtreme, Gen.choose(1, 12)) { (n: BigDecimal, k: Int) =>
      if n.scale.toLong + k <= Int.MaxValue then
        val rescaled = BigDecimal(n.bigDecimal.setScale(n.scale + k))
        for code <- codes do assertEquals(fmt(code, rescaled), fmt(code, n), s"$code on $n +$k")
    }
  }

  /** Values a digit pattern spells out in full: up to 40 digits, exponents within ±400. */
  private val genModerate: Gen[BigDecimal] =
    for
      digits <- Gen.choose(1, 40)
      unscaled <- Gen.listOfN(digits, Gen.choose(0, 9)).map(ds => BigInt(ds.mkString))
      scale <- Gen.choose(-400, 400)
      negative <- Gen.oneOf(true, false)
    yield
      val magnitude = BigDecimal(new java.math.BigDecimal(unscaled.bigInteger, scale))
      if negative then -magnitude else magnitude

  property("#689: 0 and 0.00 agree with BigDecimal HALF_UP rounding") {
    forAll(genModerate) { (n: BigDecimal) =>
      List(0, 2).foreach { d =>
        val rounded = n.bigDecimal.abs.setScale(d, RoundingMode.HALF_UP).toPlainString
        val expected = if n.signum < 0 then s"-$rounded" else rounded
        assertEquals(fmt(if d == 0 then "0" else "0.00", n), expected, s"$n")
      }
    }
  }

  property("#689: 0.00E+00 agrees with three-significant-digit HALF_UP rounding") {
    forAll(genModerate) { (n: BigDecimal) =>
      val expected =
        if n.signum == 0 then "0.00E+00"
        else
          val r = n.bigDecimal.abs.round(new MathContext(3, RoundingMode.HALF_UP))
          val e = r.precision - r.scale - 1
          val mantissa = r.movePointLeft(e).setScale(2, RoundingMode.UNNECESSARY).toPlainString
          val sign = if n.signum < 0 then "-" else ""
          f"$sign${mantissa}E${if e < 0 then "-" else "+"}${math.abs(e)}%02d"
      assertEquals(fmt("0.00E+00", n), expected, s"$n")
    }
  }
