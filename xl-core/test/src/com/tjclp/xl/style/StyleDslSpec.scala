package com.tjclp.xl.style

import com.tjclp.xl.style.alignment.{Align, HAlign, VAlign}
import com.tjclp.xl.style.border.{Border, BorderStyle}
import com.tjclp.xl.style.color.Color
import com.tjclp.xl.style.fill.Fill
import com.tjclp.xl.style.font.Font
import com.tjclp.xl.style.numfmt.NumFmt
import com.tjclp.xl.style.patch.StylePatch
import com.tjclp.xl.style.patch.StylePatch.* // For applyPatch extension
import munit.ScalaCheckSuite
import org.scalacheck.Prop.*

class StyleDslSpec extends ScalaCheckSuite:

  // ========== Font Styling Tests ==========

  test("bold sets bold flag") {
    val style = CellStyle.default.bold
    assert(style.font.bold)
  }

  test("italic sets italic flag") {
    val style = CellStyle.default.italic
    assert(style.font.italic)
  }

  test("underline sets underline flag") {
    val style = CellStyle.default.underline
    assert(style.font.underline)
  }

  test("size sets font size") {
    val style = CellStyle.default.size(16.0)
    assertEquals(style.font.sizePt, 16.0)
  }

  test("fontFamily sets font name") {
    val style = CellStyle.default.fontFamily("Arial")
    assertEquals(style.font.name, "Arial")
  }

  // ========== Preset Font Color Tests ==========

  test("red sets font color to red") {
    val style = CellStyle.default.red
    assertEquals(style.font.color, Some(Color.fromRgb(255, 0, 0)))
  }

  test("green sets font color to green") {
    val style = CellStyle.default.green
    assertEquals(style.font.color, Some(Color.fromRgb(0, 255, 0)))
  }

  test("blue sets font color to blue") {
    val style = CellStyle.default.blue
    assertEquals(style.font.color, Some(Color.fromRgb(0, 0, 255)))
  }

  test("white sets font color to white") {
    val style = CellStyle.default.white
    assertEquals(style.font.color, Some(Color.fromRgb(255, 255, 255)))
  }

  test("black sets font color to black") {
    val style = CellStyle.default.black
    assertEquals(style.font.color, Some(Color.fromRgb(0, 0, 0)))
  }

  // ========== Custom Font Color Tests ==========

  test("rgb sets custom font color from components") {
    val style = CellStyle.default.rgb(68, 114, 196)
    assertEquals(style.font.color, Some(Color.fromRgb(68, 114, 196)))
  }

  test("hex sets font color from hex code") {
    val style = CellStyle.default.hex("#4472C4")
    assert(style.font.color.isDefined)
  }

  test("hex with invalid code silently ignored (runtime validation)") {
    // Use variable to bypass compile-time validation, test runtime path
    val invalidCode: String = "invalid"  // Runtime string
    val style = CellStyle.default.hex(invalidCode)
    assertEquals(style, CellStyle.default) // No change, silent fail
  }

  test("hex with short code silently ignored (runtime validation)") {
    // Short codes like #F00 are not supported (need #RRGGBB)
    val shortCode: String = "#F00"
    val style = CellStyle.default.hex(shortCode)
    assertEquals(style, CellStyle.default) // No change, silent fail
  }

  test("hex with literal validates at compile-time") {
    // String literals are validated at compile-time by macro
    val red = CellStyle.default.hex("#FF0000")
    assert(red.font.color.isDefined)
    assertEquals(red.font.color.getOrElse(fail("Color should be defined")), Color.fromRgb(255, 0, 0))

    val blue = CellStyle.default.hex("#0000FF")
    assert(blue.font.color.isDefined)

    val customWithAlpha = CellStyle.default.hex("#FF4472C4")
    assert(customWithAlpha.font.color.isDefined)
  }

  test("bgHex with literal validates at compile-time") {
    // String literals are validated at compile-time by macro
    val lightGray = CellStyle.default.bgHex("#F5F5F5")
    assert(lightGray.fill match { case _: Fill.Solid => true; case _ => false })

    val darkBlue = CellStyle.default.bgHex("#003366")
    assert(darkBlue.fill match { case _: Fill.Solid => true; case _ => false })
  }

  // ========== Preset Background Color Tests ==========

  test("bgRed sets background to red") {
    val style = CellStyle.default.bgRed
    assertEquals(style.fill, Fill.Solid(Color.fromRgb(255, 0, 0)))
  }

  test("bgBlue sets background to blue") {
    val style = CellStyle.default.bgBlue
    assertEquals(style.fill, Fill.Solid(Color.fromRgb(68, 114, 196)))
  }

  test("bgGray sets background to gray") {
    val style = CellStyle.default.bgGray
    assertEquals(style.fill, Fill.Solid(Color.fromRgb(200, 200, 200)))
  }

  test("bgNone clears background") {
    val style = CellStyle.default.bgBlue.bgNone
    assertEquals(style.fill, Fill.None)
  }

  // ========== Custom Background Color Tests ==========

  test("bgRgb sets custom background from components") {
    val style = CellStyle.default.bgRgb(240, 240, 240)
    assertEquals(style.fill, Fill.Solid(Color.fromRgb(240, 240, 240)))
  }

  test("bgHex sets background from hex code") {
    val style = CellStyle.default.bgHex("#F0F0F0")
    style.fill match
      case Fill.Solid(Color.Rgb(argb)) =>
        assert((argb & 0x00FFFFFF) == 0x00F0F0F0)
      case _ => fail("Expected Solid fill with RGB color")
  }

  // ========== Alignment Tests ==========

  test("center sets horizontal center alignment") {
    val style = CellStyle.default.center
    assertEquals(style.align.horizontal, HAlign.Center)
  }

  test("left sets horizontal left alignment") {
    val style = CellStyle.default.left
    assertEquals(style.align.horizontal, HAlign.Left)
  }

  test("right sets horizontal right alignment") {
    val style = CellStyle.default.right
    assertEquals(style.align.horizontal, HAlign.Right)
  }

  test("top sets vertical top alignment") {
    val style = CellStyle.default.top
    assertEquals(style.align.vertical, VAlign.Top)
  }

  test("middle sets vertical middle alignment") {
    val style = CellStyle.default.middle
    assertEquals(style.align.vertical, VAlign.Middle)
  }

  test("bottom sets vertical bottom alignment") {
    val style = CellStyle.default.bottom
    assertEquals(style.align.vertical, VAlign.Bottom)
  }

  test("wrap enables text wrapping") {
    val style = CellStyle.default.wrap
    assert(style.align.wrapText)
  }

  // ========== Border Tests ==========

  test("bordered sets all borders to thin") {
    val style = CellStyle.default.bordered
    assertEquals(style.border, Border.all(BorderStyle.Thin))
  }

  test("borderedMedium sets all borders to medium") {
    val style = CellStyle.default.borderedMedium
    assertEquals(style.border, Border.all(BorderStyle.Medium))
  }

  test("borderedThick sets all borders to thick") {
    val style = CellStyle.default.borderedThick
    assertEquals(style.border, Border.all(BorderStyle.Thick))
  }

  // ========== Number Format Tests ==========

  test("currency sets currency format") {
    val style = CellStyle.default.currency
    assertEquals(style.numFmt, NumFmt.Currency)
  }

  test("percent sets percent format") {
    val style = CellStyle.default.percent
    assertEquals(style.numFmt, NumFmt.Percent)
  }

  test("decimal sets decimal format") {
    val style = CellStyle.default.decimal
    assertEquals(style.numFmt, NumFmt.Decimal)
  }

  test("dateFormat sets date format") {
    val style = CellStyle.default.dateFormat
    assertEquals(style.numFmt, NumFmt.Date)
  }

  test("dateTime sets datetime format") {
    val style = CellStyle.default.dateTime
    assertEquals(style.numFmt, NumFmt.DateTime)
  }

  // ========== Chaining Tests ==========

  test("chaining preserves all modifications") {
    val style = CellStyle.default.bold.red.center.wrap
    assert(style.font.bold)
    assertEquals(style.font.color, Some(Color.fromRgb(255, 0, 0)))
    assertEquals(style.align.horizontal, HAlign.Center)
    assert(style.align.wrapText)
  }

  test("complex chaining with custom colors") {
    val style = CellStyle.default
      .bold.size(14.0)
      .rgb(68, 114, 196)
      .bgRgb(240, 240, 240)
      .center.middle
      .bordered

    assert(style.font.bold)
    assertEquals(style.font.sizePt, 14.0)
    assert(style.font.color.isDefined)
    assertEquals(style.align.horizontal, HAlign.Center)
    assertEquals(style.align.vertical, VAlign.Middle)
    assertEquals(style.border, Border.all(BorderStyle.Thin))
  }

  test("chaining with hex codes") {
    val style = CellStyle.default
      .hex("#FF6B35")
      .bgHex("#F7F7F7")
      .bold.center

    assert(style.font.bold)
    assertEquals(style.align.horizontal, HAlign.Center)
    assert(style.font.color.isDefined)
  }

  // ========== Idempotence Tests ==========

  test("bold is idempotent") {
    val once = CellStyle.default.bold
    val twice = once.bold
    assertEquals(once.font.bold, twice.font.bold)
  }

  test("red is idempotent") {
    val once = CellStyle.default.red
    val twice = once.red
    assertEquals(once.font.color, twice.font.color)
  }

  test("center is idempotent") {
    val once = CellStyle.default.center
    val twice = once.center
    assertEquals(once.align.horizontal, twice.align.horizontal)
  }

  // ========== Prebuilt Style Tests ==========

  test("Style.header has expected properties") {
    val style = Style.header
    assert(style.font.bold)
    assertEquals(style.font.sizePt, 14.0)
    assertEquals(style.align.horizontal, HAlign.Center)
    assertEquals(style.align.vertical, VAlign.Middle)
    assertEquals(style.font.color, Some(Color.fromRgb(255, 255, 255)))
  }

  test("Style.currencyCell has currency format and right alignment") {
    val style = Style.currencyCell
    assertEquals(style.numFmt, NumFmt.Currency)
    assertEquals(style.align.horizontal, HAlign.Right)
  }

  test("Style.errorText is red, bold, italic") {
    val style = Style.errorText
    assert(style.font.bold)
    assert(style.font.italic)
    assertEquals(style.font.color, Some(Color.fromRgb(255, 0, 0)))
  }

  test("Style.successText is green and bold") {
    val style = Style.successText
    assert(style.font.bold)
    assertEquals(style.font.color, Some(Color.fromRgb(0, 255, 0)))
  }

  // ========== Real-World Use Case Tests ==========

  test("financial report header style") {
    val header = CellStyle.default.bold.size(16.0).white.bgBlue.center.middle.bordered
    assert(header.font.bold)
    assertEquals(header.font.sizePt, 16.0)
    assertEquals(header.font.color, Some(Color.fromRgb(255, 255, 255)))
    assertEquals(header.align.horizontal, HAlign.Center)
    assertEquals(header.align.vertical, VAlign.Middle)
  }

  test("brand color customization") {
    // Example: TJC brand colors
    val tjcBlue = CellStyle.default.hex("#003366")
    val tjcHeader = CellStyle.default.white.bgHex("#003366").bold.center

    assert(tjcHeader.font.bold)
    assertEquals(tjcHeader.align.horizontal, HAlign.Center)
    assertEquals(tjcHeader.font.color, Some(Color.fromRgb(255, 255, 255)))
  }

  test("data table with mixed styling") {
    val textStyle = CellStyle.default.left
    val numberStyle = CellStyle.default.decimal.right
    val dateStyle = CellStyle.default.dateFormat.center

    assertEquals(textStyle.align.horizontal, HAlign.Left)
    assertEquals(numberStyle.numFmt, NumFmt.Decimal)
    assertEquals(numberStyle.align.horizontal, HAlign.Right)
    assertEquals(dateStyle.numFmt, NumFmt.Date)
    assertEquals(dateStyle.align.horizontal, HAlign.Center)
  }

  // ========== StylePatch ++ Operator Tests ==========

  test("StylePatch ++ composes without type ascription") {
    import com.tjclp.xl.style.patch.StylePatch

    val font = Font(bold = true, sizePt = 14.0)
    val fill = Fill.Solid(Color.fromRgb(255, 0, 0))

    // Should compile without `: StylePatch` ascription
    val patch = StylePatch.SetFont(font) ++ StylePatch.SetFill(fill)

    val result = CellStyle.default.applyPatch(patch)
    assert(result.font.bold)
    assertEquals(result.font.sizePt, 14.0)
    assertEquals(result.fill, fill)
  }

  test("StylePatch ++ chains multiple patches") {
    import com.tjclp.xl.style.patch.StylePatch

    val patch =
      StylePatch.SetFont(Font.default.withBold(true)) ++
      StylePatch.SetFill(Fill.Solid(Color.fromRgb(0, 0, 255))) ++
      StylePatch.SetAlign(Align.default.withHAlign(HAlign.Center))

    val result = CellStyle.default.applyPatch(patch)
    assert(result.font.bold)
    assertEquals(result.align.horizontal, HAlign.Center)
  }

  // ========== Integration Tests ==========

  test("DSL integrates with StylePatch") {
    import com.tjclp.xl.style.patch.StylePatch

    // Build style with DSL
    val dslStyle = CellStyle.default.bold.red.center

    // Build equivalent with StylePatch
    val patch =
      StylePatch.SetFont(Font.default.withBold(true).withColor(Color.fromRgb(255, 0, 0))) ++
      StylePatch.SetAlign(Align.default.withHAlign(HAlign.Center))
    val patchStyle = CellStyle.default.applyPatch(patch)

    assertEquals(dslStyle.font.bold, patchStyle.font.bold)
    assertEquals(dslStyle.font.color, patchStyle.font.color)
    assertEquals(dslStyle.align.horizontal, patchStyle.align.horizontal)
  }

  test("DSL can be composed with existing withX methods") {
    val style = CellStyle.default
      .bold.red // DSL
      .withBorder(Border.all(BorderStyle.Thick)) // Existing method
      .center // DSL

    assert(style.font.bold)
    assertEquals(style.font.color, Some(Color.fromRgb(255, 0, 0)))
    assertEquals(style.border, Border.all(BorderStyle.Thick))
    assertEquals(style.align.horizontal, HAlign.Center)
  }
