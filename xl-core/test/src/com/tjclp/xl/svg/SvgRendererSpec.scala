package com.tjclp.xl.svg

import munit.FunSuite
import com.tjclp.xl.api.*
import com.tjclp.xl.addressing.{Column, Row}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.dsl.{*, given}
import com.tjclp.xl.macros.{ref, col}
import com.tjclp.xl.patch.Patch
import com.tjclp.xl.sheets.{ColumnProperties, RowProperties}
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

  // ========== Row/Column Sizing Tests ==========

  test("toSvg: explicit column width is used") {
    // Column width 20 chars → (20 * 7 + 5) = 145 pixels
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Data")
      .setColumnProperties(Column.from0(0), ColumnProperties(width = Some(20.0)))

    val svg = sheet.toSvg(ref"A1:A1")

    // The cell rect should use the explicit width
    assert(svg.contains("""width="145""""), s"Column width should be 145px (20 chars), got: $svg")
  }

  test("toSvg: explicit row height is used") {
    // Row height 30 points → (30 * 4/3) = 40 pixels
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Data")
      .setRowProperties(Row.from0(0), RowProperties(height = Some(30.0)))

    val svg = sheet.toSvg(ref"A1:A1")

    // The cell rect should use the explicit height
    assert(svg.contains("""height="40""""), s"Row height should be 40px (30pt), got: $svg")
  }

  test("toSvg: hidden column renders as 0px") {
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Data")
      .setColumnProperties(Column.from0(0), ColumnProperties(hidden = true))

    val svg = sheet.toSvg(ref"A1:A1")

    // Hidden column should have 0 width
    assert(svg.contains("""width="0""""), s"Hidden column should be 0px wide, got: $svg")
  }

  test("toSvg: hidden row renders as 0px") {
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Data")
      .setRowProperties(Row.from0(0), RowProperties(hidden = true))

    val svg = sheet.toSvg(ref"A1:A1")

    // Hidden row should have 0 height
    assert(svg.contains("""height="0""""), s"Hidden row should be 0px tall, got: $svg")
  }

  test("toSvg: mixed explicit and default dimensions") {
    // A1, B1 with A=explicit width, B=default
    // 15 chars → (15 * 7 + 5) = 110 pixels
    val sheet = Sheet("Test")
      .put(ref"A1" -> "A", ref"B1" -> "B")
      .setColumnProperties(Column.from0(0), ColumnProperties(width = Some(15.0))) // 110px

    val svg = sheet.toSvg(ref"A1:B1")

    // First column should be explicit, second column should use content-based sizing
    assert(svg.contains("""width="110""""), s"Column A should be 110px, got: $svg")
    // Column B should use default MinCellWidth (60) or content-based
  }

  test("toSvg: viewBox reflects actual dimensions with sizing") {
    // Wide column + tall row should change viewBox
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Data")
      .setColumnProperties(Column.from0(0), ColumnProperties(width = Some(30.0))) // 215px
      .setRowProperties(Row.from0(0), RowProperties(height = Some(45.0))) // 60px

    val svg = sheet.toSvg(ref"A1:A1")

    // viewBox should include HeaderWidth (40) + column (215) = 255 total width
    // viewBox should include HeaderHeight (24) + row (60) = 84 total height
    assert(svg.contains("viewBox=\"0 0 255 84\""), s"viewBox should be 255x84, got: $svg")
  }

  test("toSvg: DSL-based sizing works") {
    // Use the row/column DSL
    val patch = col"A".width(15.0).toPatch ++ row(0).height(24.0).toPatch

    val sheet = Patch
      .applyPatch(Sheet("Test").put(ref"A1" -> "Data"), patch)
      .toSvg(ref"A1:A1")

    // Column A: 15 chars → 110px, Row 0: 24pt → 32px
    assert(sheet.contains("""width="110""""), s"DSL column width should apply, got: $sheet")
    assert(sheet.contains("""height="32""""), s"DSL row height should apply, got: $sheet")
  }
