package com.tjclp.xl

import munit.ScalaCheckSuite
import org.scalacheck.{Arbitrary, Gen, Prop}
import org.scalacheck.Prop.*
import com.tjclp.xl.styles.{*, given}
import com.tjclp.xl.styles.alignment.{Align, HAlign, VAlign}
import com.tjclp.xl.styles.border.{Border, BorderSide, BorderStyle}
import com.tjclp.xl.styles.color.{Color, ThemeSlot}
import com.tjclp.xl.styles.fill.{Fill, PatternType}
import com.tjclp.xl.styles.font.Font
import com.tjclp.xl.styles.numfmt.NumFmt
import com.tjclp.xl.styles.units.{Emu, Pt, Px, given}

/** Property tests for style system */
class StyleSpec extends ScalaCheckSuite:

  // ========== Generators ==========

  val genPt: Gen[Pt] = Gen.choose(0.1, 100.0).map(Pt.apply)
  val genPx: Gen[Px] = Gen.choose(0.1, 100.0).map(Px.apply)
  val genEmu: Gen[Emu] = Gen.choose(1000L, 1000000L).map(Emu.apply)

  val genThemeSlot: Gen[ThemeSlot] = Gen.oneOf(
    ThemeSlot.Dark1, ThemeSlot.Light1, ThemeSlot.Dark2, ThemeSlot.Light2,
    ThemeSlot.Accent1, ThemeSlot.Accent2, ThemeSlot.Accent3,
    ThemeSlot.Accent4, ThemeSlot.Accent5, ThemeSlot.Accent6
  )

  val genTint: Gen[Double] = Gen.choose(-1.0, 1.0)

  val genColor: Gen[Color] = Gen.oneOf(
    Gen.choose(Int.MinValue, Int.MaxValue).map(Color.Rgb.apply),
    for
      slot <- genThemeSlot
      tint <- genTint
    yield Color.Theme(slot, tint)
  )

  val genNumFmt: Gen[NumFmt] = Gen.oneOf(
    Gen.const(NumFmt.General), Gen.const(NumFmt.Integer), Gen.const(NumFmt.Decimal),
    Gen.const(NumFmt.ThousandsSeparator), Gen.const(NumFmt.ThousandsDecimal),
    Gen.const(NumFmt.Currency), Gen.const(NumFmt.Percent), Gen.const(NumFmt.PercentDecimal),
    Gen.const(NumFmt.Scientific), Gen.const(NumFmt.Fraction), Gen.const(NumFmt.Date),
    Gen.const(NumFmt.DateTime), Gen.const(NumFmt.Time), Gen.const(NumFmt.Text),
    Gen.alphaNumStr.map(NumFmt.Custom.apply)
  ).flatMap(identity)

  val genFont: Gen[Font] = for
    name <- Gen.oneOf("Calibri", "Arial", "Times New Roman", "Courier New")
    size <- Gen.choose(8.0, 24.0)
    bold <- Gen.oneOf(true, false)
    italic <- Gen.oneOf(true, false)
    underline <- Gen.oneOf(true, false)
    color <- Gen.option(genColor)
  yield Font(name, size, bold, italic, underline, color)

  val genPatternType: Gen[PatternType] = Gen.oneOf(
    PatternType.None, PatternType.Solid, PatternType.Gray125, PatternType.Gray0625,
    PatternType.DarkGray, PatternType.MediumGray, PatternType.LightGray
  )

  val genFill: Gen[Fill] = Gen.oneOf(
    Gen.const(Fill.None),
    genColor.map(Fill.Solid.apply),
    for
      fg <- genColor
      bg <- genColor
      pattern <- genPatternType
    yield Fill.Pattern(fg, bg, pattern)
  )

  val genBorderStyle: Gen[BorderStyle] = Gen.oneOf(
    BorderStyle.None, BorderStyle.Thin, BorderStyle.Medium, BorderStyle.Thick,
    BorderStyle.Dashed, BorderStyle.Dotted, BorderStyle.Double
  )

  val genBorderSide: Gen[BorderSide] = for
    style <- genBorderStyle
    color <- Gen.option(genColor)
  yield BorderSide(style, color)

  val genBorder: Gen[Border] = for
    left <- genBorderSide
    right <- genBorderSide
    top <- genBorderSide
    bottom <- genBorderSide
  yield Border(left, right, top, bottom)

  val genHAlign: Gen[HAlign] = Gen.oneOf(
    HAlign.Left, HAlign.Center, HAlign.Right, HAlign.Justify, HAlign.Fill
  )

  val genVAlign: Gen[VAlign] = Gen.oneOf(
    VAlign.Top, VAlign.Middle, VAlign.Bottom, VAlign.Justify
  )

  val genAlign: Gen[Align] = for
    h <- genHAlign
    v <- genVAlign
    wrap <- Gen.oneOf(true, false)
    indent <- Gen.choose(0, 10)
  yield Align(h, v, wrap, indent)

  val genCellStyle: Gen[CellStyle] = for
    font <- genFont
    fill <- genFill
    border <- genBorder
    numFmt <- genNumFmt
    align <- genAlign
  yield CellStyle(font = font, fill = fill, border = border, numFmt = numFmt, align = align)

  given Arbitrary[Pt] = Arbitrary(genPt)
  given Arbitrary[Px] = Arbitrary(genPx)
  given Arbitrary[Emu] = Arbitrary(genEmu)
  given Arbitrary[Color] = Arbitrary(genColor)
  given Arbitrary[NumFmt] = Arbitrary(genNumFmt)
  given Arbitrary[Font] = Arbitrary(genFont)
  given Arbitrary[Fill] = Arbitrary(genFill)
  given Arbitrary[Border] = Arbitrary(genBorder)
  given Arbitrary[Align] = Arbitrary(genAlign)
  given Arbitrary[CellStyle] = Arbitrary(genCellStyle)

  // ========== Unit Conversion Tests ==========

  property("Pt to Px and back preserves value") {
    forAll { (pt: Pt) =>
      val roundTrip = pt.toPx.toPt.value
      val diff = math.abs(roundTrip - pt.value)
      assert(diff < 0.0001, s"Round trip failed: $pt -> ${pt.toPx} -> ${pt.toPx.toPt}")
      true
    }
  }

  property("Pt to Emu and back preserves value") {
    forAll { (pt: Pt) =>
      val roundTrip = pt.toEmu.toPt.value
      val diff = math.abs(roundTrip - pt.value)
      assert(diff < 0.0001, s"Round trip failed: $pt -> ${pt.toEmu} -> ${pt.toEmu.toPt}")
      true
    }
  }

  property("Px to Emu and back preserves value") {
    forAll { (px: Px) =>
      val roundTrip = px.toEmu.toPx.value
      val diff = math.abs(roundTrip - px.value)
      // EMU uses Long, so there's truncation; tolerance accounts for rounding
      assert(diff < 0.001, s"Round trip failed: $px -> ${px.toEmu} -> ${px.toEmu.toPx}")
      true
    }
  }

  test("Unit conversion constants are correct") {
    // 72 points = 1 inch
    val inch = Pt(72.0)
    assertEquals(inch.toPx.value, 96.0, 0.0001) // 96 DPI

    // 914400 EMUs = 1 inch
    assertEquals(Emu(914400L).toPt.value, 72.0, 0.0001)
    assertEquals(Emu(914400L).toPx.value, 96.0, 0.0001)
  }

  // ========== Color Tests ==========

  test("Color.fromRgb creates correct ARGB") {
    val color = Color.fromRgb(255, 128, 64, 200)
    color match
      case Color.Rgb(argb) =>
        assertEquals((argb >> 24) & 0xFF, 200) // Alpha
        assertEquals((argb >> 16) & 0xFF, 255) // Red
        assertEquals((argb >> 8) & 0xFF, 128)  // Green
        assertEquals(argb & 0xFF, 64)          // Blue
      case _ => fail("Should be Rgb")
  }

  test("Color.fromHex parses 6-digit hex") {
    val result = Color.fromHex("#FF8040")
    assert(result.isRight)
    result.foreach {
      case Color.Rgb(argb) =>
        assertEquals((argb >> 16) & 0xFF, 0xFF)
        assertEquals((argb >> 8) & 0xFF, 0x80)
        assertEquals(argb & 0xFF, 0x40)
      case _ => fail("Should be Rgb")
    }
  }

  test("Color.fromHex parses 8-digit hex with alpha") {
    val result = Color.fromHex("#C8FF8040")
    assert(result.isRight)
    result.foreach {
      case Color.Rgb(argb) =>
        assertEquals((argb >> 24) & 0xFF, 0xC8)
        assertEquals((argb >> 16) & 0xFF, 0xFF)
        assertEquals((argb >> 8) & 0xFF, 0x80)
        assertEquals(argb & 0xFF, 0x40)
      case _ => fail("Should be Rgb")
    }
  }

  test("Color.fromHex rejects invalid input") {
    assert(Color.fromHex("#12345").isLeft)
    assert(Color.fromHex("#GGGGGG").isLeft)
    assert(Color.fromHex("notahex").isLeft)
  }

  test("Color.validTint accepts valid range") {
    assert(Color.validTint(-1.0).isRight)
    assert(Color.validTint(0.0).isRight)
    assert(Color.validTint(1.0).isRight)
    assert(Color.validTint(0.5).isRight)
  }

  test("Color.validTint rejects invalid range") {
    assert(Color.validTint(-1.1).isLeft)
    assert(Color.validTint(1.1).isLeft)
    assert(Color.validTint(2.0).isLeft)
  }

  // ========== NumFmt Tests ==========

  test("NumFmt.builtInId returns correct IDs") {
    assertEquals(NumFmt.builtInId(NumFmt.General), Some(0))
    assertEquals(NumFmt.builtInId(NumFmt.Integer), Some(1))
    assertEquals(NumFmt.builtInId(NumFmt.Decimal), Some(2))
    assertEquals(NumFmt.builtInId(NumFmt.Percent), Some(9))
    assertEquals(NumFmt.builtInId(NumFmt.Text), Some(49))
    assertEquals(NumFmt.builtInId(NumFmt.Custom("custom")), None)
  }

  test("NumFmt.parse recognizes built-in formats") {
    assertEquals(NumFmt.parse("General"), NumFmt.General)
    assertEquals(NumFmt.parse("0"), NumFmt.Integer)
    assertEquals(NumFmt.parse("0.00"), NumFmt.Decimal)
    assertEquals(NumFmt.parse("0%"), NumFmt.Percent)
    assertEquals(NumFmt.parse("@"), NumFmt.Text)
  }

  test("NumFmt.parse falls back to Custom for unknown formats") {
    val result = NumFmt.parse("[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy")
    result match
      case NumFmt.Custom(code) => assert(code.nonEmpty)
      case _ => fail("Should be Custom")
  }

  // ========== Font Tests ==========

  test("Font.default has correct values") {
    assertEquals(Font.default.name, "Calibri")
    assertEquals(Font.default.sizePt, 11.0)
    assertEquals(Font.default.bold, false)
    assertEquals(Font.default.italic, false)
    assertEquals(Font.default.underline, false)
    assertEquals(Font.default.color, None)
  }

  test("Font requires positive size") {
    intercept[IllegalArgumentException] {
      Font(sizePt = 0.0)
    }
    intercept[IllegalArgumentException] {
      Font(sizePt = -1.0)
    }
  }

  test("Font requires non-empty name") {
    intercept[IllegalArgumentException] {
      Font(name = "")
    }
  }

  property("Font builder methods work correctly") {
    forAll { (font: Font, name: String, size: Double, color: Color) =>
      val validName = if name.isEmpty then "Arial" else name
      val validSize = if size <= 0 then 12.0 else size

      val updated = font
        .withName(validName)
        .withSize(validSize)
        .withBold(true)
        .withItalic(true)
        .withUnderline(true)
        .withColor(color)

      assertEquals(updated.name, validName)
      assertEquals(updated.sizePt, validSize)
      assertEquals(updated.bold, true)
      assertEquals(updated.italic, true)
      assertEquals(updated.underline, true)
      assertEquals(updated.color, Some(color))
      true
    }
  }

  // ========== Border Tests ==========

  test("Border.all creates uniform border") {
    val border = Border.all(BorderStyle.Thin, Some(Color.Rgb(0xFF000000)))
    assertEquals(border.left.style, BorderStyle.Thin)
    assertEquals(border.right.style, BorderStyle.Thin)
    assertEquals(border.top.style, BorderStyle.Thin)
    assertEquals(border.bottom.style, BorderStyle.Thin)
  }

  test("Border.none has no borders") {
    assertEquals(Border.none.left.style, BorderStyle.None)
    assertEquals(Border.none.right.style, BorderStyle.None)
    assertEquals(Border.none.top.style, BorderStyle.None)
    assertEquals(Border.none.bottom.style, BorderStyle.None)
  }

  // ========== Align Tests ==========

  test("Align requires non-negative indent") {
    intercept[IllegalArgumentException] {
      Align(indent = -1)
    }
  }

  test("Align.default has correct values") {
    assertEquals(Align.default.horizontal, HAlign.Left)
    assertEquals(Align.default.vertical, VAlign.Bottom)
    assertEquals(Align.default.wrapText, false)
    assertEquals(Align.default.indent, 0)
  }

  // ========== CellStyle Tests ==========

  test("CellStyle.default uses all default components") {
    val style = CellStyle.default
    assertEquals(style.font, Font.default)
    assertEquals(style.fill, Fill.default)
    assertEquals(style.border, Border.none)
    assertEquals(style.numFmt, NumFmt.General)
    assertEquals(style.align, Align.default)
  }

  property("CellStyle builder methods work correctly") {
    forAll { (style: CellStyle, font: Font, fill: Fill, border: Border, numFmt: NumFmt, align: Align) =>
      val updated = style
        .withFont(font)
        .withFill(fill)
        .withBorder(border)
        .withNumFmt(numFmt)
        .withAlign(align)

      assertEquals(updated.font, font)
      assertEquals(updated.fill, fill)
      assertEquals(updated.border, border)
      assertEquals(updated.numFmt, numFmt)
      assertEquals(updated.align, align)
      true
    }
  }

  // ========== Canonicalization Tests ==========

  test("CellStyle.canonicalKey is deterministic") {
    val style = CellStyle(
      font = Font("Arial", 12.0, bold = true),
      fill = Fill.Solid(Color.Rgb(0xFFFFFFFF)),
      border = Border.all(BorderStyle.Thin),
      numFmt = NumFmt.Percent,
      align = Align(HAlign.Center, VAlign.Middle)
    )

    val key1 = CellStyle.canonicalKey(style)
    val key2 = CellStyle.canonicalKey(style)
    assertEquals(key1, key2)
  }

  test("CellStyle.canonicalKey differs for different styles") {
    val style1 = CellStyle(font = Font("Arial", 12.0))
    val style2 = CellStyle(font = Font("Calibri", 12.0))

    val key1 = CellStyle.canonicalKey(style1)
    val key2 = CellStyle.canonicalKey(style2)
    assertNotEquals(key1, key2)
  }

  property("CellStyle.canonicalKey: equal styles have equal keys") {
    forAll { (style: CellStyle) =>
      val copy = style.copy()
      val key1 = CellStyle.canonicalKey(style)
      val key2 = CellStyle.canonicalKey(copy)
      assertEquals(key1, key2)
      true
    }
  }

  // ========== NumFmt ID Preservation Tests (Regression Prevention) ==========

  test("canonicalKey ignores numFmtId (ensures deduplication works)") {
    // Critical: styles with same visual properties but different numFmtId must deduplicate
    val style1 = CellStyle(
      numFmt = NumFmt.General,
      numFmtId = Some(0)
    )
    val style2 = CellStyle(
      numFmt = NumFmt.General,
      numFmtId = None
    )
    val style3 = CellStyle(
      numFmt = NumFmt.General,
      numFmtId = Some(39) // Different ID but will map to Custom format
    )

    assertEquals(
      style1.canonicalKey,
      style2.canonicalKey,
      "Styles with same numFmt but different numFmtId must have same canonicalKey"
    )

    // Note: style3 would have different key because numFmt would be different
    // (ID 39 maps to accounting format via builtInById)
  }

  test("NumFmt.fromId recognizes all critical built-in format IDs") {
    // Test accounting formats (the ones that caused the original bug)
    assert(NumFmt.fromId(39).isDefined, "ID 39 (accounting #,##0.00) must be recognized")
    assert(NumFmt.fromId(40).isDefined, "ID 40 (accounting with red negatives) must be recognized")
    assert(NumFmt.fromId(41).isDefined, "ID 41 (accounting with spaces) must be recognized")
    assert(NumFmt.fromId(42).isDefined, "ID 42 (accounting with $) must be recognized")
    assert(NumFmt.fromId(43).isDefined, "ID 43 (accounting .00 with spaces) must be recognized")
    assert(NumFmt.fromId(44).isDefined, "ID 44 (accounting .00 with $) must be recognized")

    // Test that unrecognized IDs return None (as expected)
    assert(NumFmt.fromId(200).isEmpty, "Custom ID 200 should not be built-in")
  }
