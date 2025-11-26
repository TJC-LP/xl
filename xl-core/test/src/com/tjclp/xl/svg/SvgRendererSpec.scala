package com.tjclp.xl.svg

import munit.FunSuite
import com.tjclp.xl.api.*
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.macros.ref
import com.tjclp.xl.sheets.styleSyntax.*
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.color.Color
import com.tjclp.xl.styles.fill.Fill
import com.tjclp.xl.unsafe.*

/** Tests for SVG export functionality */
class SvgRendererSpec extends FunSuite:

  test("toSvg: basic SVG structure") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Text("Hello"))

    val svg = sheet.toSvg(ref"A1:A1")
    assert(svg.contains("<svg"), "Should contain svg tag")
    assert(svg.contains("xmlns"), "Should have xmlns attribute")
    assert(svg.contains("Hello"), "Should contain cell text")
  }

  test("toSvg: ARGB color renders with fill-opacity") {
    // Create a color with alpha channel: #80FF0000 (50% opacity red)
    // fromRgb(r, g, b, a) where a=128 gives ~50% opacity
    val semiTransparentRed = Color.fromRgb(255, 0, 0, 128)
    val style = CellStyle.default.withFill(Fill.Solid(semiTransparentRed))

    val sheet = Sheet("Test")
      .put(ref"A1" -> "Test")
      .unsafe
      .withCellStyle(ref"A1", style)

    val svg = sheet.toSvg(ref"A1:A1")

    // Should have fill="#FF0000" (the RGB part)
    assert(svg.contains("""fill="#FF0000""""), s"Should have RGB fill, got: $svg")
    // Should have fill-opacity for the alpha channel
    assert(svg.contains("fill-opacity="), s"Should have fill-opacity attribute, got: $svg")
    // Opacity should be approximately 0.5 (128/255 ≈ 0.502)
    assert(svg.contains("0.50"), s"Opacity should be ~0.5, got: $svg")
  }

  test("toSvg: RGB color (no alpha) renders without fill-opacity") {
    // Create a color without alpha channel (default a=255)
    val solidRed = Color.fromRgb(255, 0, 0)
    val style = CellStyle.default.withFill(Fill.Solid(solidRed))

    val sheet = Sheet("Test")
      .put(ref"A1" -> "Test")
      .unsafe
      .withCellStyle(ref"A1", style)

    val svg = sheet.toSvg(ref"A1:A1")

    // Should have fill but NOT fill-opacity for fully opaque colors
    assert(svg.contains("""fill="#"""), "Should have fill attribute")
    // For 6-char hex colors, no fill-opacity should be added
    val fillOpacityCount = "fill-opacity".r.findAllIn(svg).length
    assertEquals(fillOpacityCount, 0, "Should not have fill-opacity for opaque colors")
  }

  test("toSvg: fully transparent color has fill-opacity 0") {
    // Create fully transparent color: a=0
    val transparent = Color.fromRgb(255, 0, 0, 0)
    val style = CellStyle.default.withFill(Fill.Solid(transparent))

    val sheet = Sheet("Test")
      .put(ref"A1" -> "Test")
      .unsafe
      .withCellStyle(ref"A1", style)

    val svg = sheet.toSvg(ref"A1:A1")

    assert(svg.contains("fill-opacity=\"0.0\""), s"Should have fill-opacity=0.0, got: $svg")
  }
