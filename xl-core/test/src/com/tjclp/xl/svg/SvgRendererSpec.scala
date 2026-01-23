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
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial"))
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
    // Column width 20 chars → (20 * 8 + 5) = 165 pixels
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Data")
      .setColumnProperties(Column.from0(0), ColumnProperties(width = Some(20.0)))

    val svg = sheet.toSvg(ref"A1:A1")

    // The cell rect should use the explicit width
    assert(svg.contains("""width="165""""), s"Column width should be 165px (20 chars), got: $svg")
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
    // 15 chars → (15 * 8 + 5) = 125 pixels
    val sheet = Sheet("Test")
      .put(ref"A1" -> "A", ref"B1" -> "B")
      .setColumnProperties(Column.from0(0), ColumnProperties(width = Some(15.0))) // 125px

    val svg = sheet.toSvg(ref"A1:B1")

    // First column should be explicit, second column should use content-based sizing
    assert(svg.contains("""width="125""""), s"Column A should be 125px, got: $svg")
    // Column B should use default MinCellWidth (60) or content-based
  }

  test("toSvg: viewBox reflects actual dimensions with sizing") {
    // Wide column + tall row should change viewBox
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Data")
      .setColumnProperties(Column.from0(0), ColumnProperties(width = Some(30.0))) // 245px
      .setRowProperties(Row.from0(0), RowProperties(height = Some(45.0))) // 60px

    val svg = sheet.toSvg(ref"A1:A1")

    // With showLabels=false (default), viewBox is just column (245) x row (60)
    assert(svg.contains("viewBox=\"0 0 245 60\""), s"viewBox should be 245x60, got: $svg")
  }

  test("toSvg: viewBox includes headers when showLabels=true") {
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Data")
      .setColumnProperties(Column.from0(0), ColumnProperties(width = Some(30.0))) // 245px
      .setRowProperties(Row.from0(0), RowProperties(height = Some(45.0))) // 60px

    val svg = sheet.toSvg(ref"A1:A1", showLabels = true)

    // viewBox should include HeaderWidth (40) + column (245) = 285 total width
    // viewBox should include HeaderHeight (24) + row (60) = 84 total height
    assert(svg.contains("viewBox=\"0 0 285 84\""), s"viewBox should be 285x84, got: $svg")
  }

  test("toSvg: DSL-based sizing works") {
    // Use the row/column DSL
    val patch = col"A".width(15.0).toPatch ++ row(0).height(24.0).toPatch

    val sheet = Patch
      .applyPatch(Sheet("Test").put(ref"A1" -> "Data"), patch)
      .toSvg(ref"A1:A1")

    // Column A: 15 chars → 125px, Row 0: 24pt → 32px
    assert(sheet.contains("""width="125""""), s"DSL column width should apply, got: $sheet")
    assert(sheet.contains("""height="32""""), s"DSL row height should apply, got: $sheet")
  }

  // ========== Indentation Tests ==========

  test("toSvg: cell with indent shifts text x-position") {
    val indentStyle = CellStyle.default.withAlign(Align(indent = 1))
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Indented")
      .unsafe
      .withCellStyle(ref"A1", indentStyle)

    val svg = sheet.toSvg(ref"A1:A1")

    // With showLabels=false: CellPaddingX (6) + indent (21) = 27
    assert(svg.contains("""x="27""""), s"Text should be at x=27 with indent=1, got: $svg")
  }

  test("toSvg: cell with indent and showLabels=true") {
    val indentStyle = CellStyle.default.withAlign(Align(indent = 1))
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Indented")
      .unsafe
      .withCellStyle(ref"A1", indentStyle)

    val svg = sheet.toSvg(ref"A1:A1", showLabels = true)

    // With showLabels=true: HeaderWidth (40) + CellPaddingX (6) + indent (21) = 67
    assert(svg.contains("""x="67""""), s"Text should be at x=67 with indent=1 and labels, got: $svg")
  }

  test("toSvg: indent 0 uses default padding") {
    val noIndent = CellStyle.default.withAlign(Align(indent = 0))
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Normal")
      .unsafe
      .withCellStyle(ref"A1", noIndent)

    val svg = sheet.toSvg(ref"A1:A1")

    // With showLabels=false: CellPaddingX (6) = 6
    assert(svg.contains("""x="6""""), s"Text should be at x=6 with no indent, got: $svg")
  }

  test("toSvg: large indent value works") {
    val largeIndent = CellStyle.default.withAlign(Align(indent = 3))
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Indented")
      .unsafe
      .withCellStyle(ref"A1", largeIndent)

    val svg = sheet.toSvg(ref"A1:A1")

    // With showLabels=false: CellPaddingX (6) + indent (3 * 21 = 63) = 69
    assert(svg.contains("""x="69""""), s"Text should be at x=69 with indent=3, got: $svg")
  }

  test("toSvg: center alignment with indent shifts slightly right") {
    val centerIndent = CellStyle.default.withAlign(Align(horizontal = HAlign.Center, indent = 2))
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Centered")
      .unsafe
      .withCellStyle(ref"A1", centerIndent)
      .setColumnProperties(Column.from0(0), ColumnProperties(width = Some(10.0))) // 75px

    val svg = sheet.toSvg(ref"A1:A1")

    // With showLabels=false: colWidth/2 (37) + indentPx/2 (21) = 58
    assert(svg.contains("text-anchor=\"middle\""), s"Should have middle anchor, got: $svg")
  }

  test("toSvg: right alignment ignores indent") {
    val rightIndent = CellStyle.default.withAlign(Align(horizontal = HAlign.Right, indent = 2))
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Right")
      .unsafe
      .withCellStyle(ref"A1", rightIndent)
      .setColumnProperties(Column.from0(0), ColumnProperties(width = Some(10.0))) // 85px

    val svg = sheet.toSvg(ref"A1:A1")

    // With showLabels=false: colWidth (85) - CellPaddingX (6) = 79
    // Indent is ignored for right alignment (Excel behavior)
    assert(svg.contains("""x="79""""), s"Right-aligned text should ignore indent, got: $svg")
    assert(svg.contains("text-anchor=\"end\""), s"Should have end anchor, got: $svg")
  }

  // ========== Vertical Alignment Tests ==========

  test("toSvg: VAlign.Top positions text near top of cell") {
    val topStyle = CellStyle.default.withAlign(Align(vertical = VAlign.Top))
    val sheet = Sheet("Test")
      .put(ref"A1" -> "TopText")
      .unsafe
      .withCellStyle(ref"A1", topStyle)
      .setRowProperties(Row.from0(0), RowProperties(height = Some(45.0))) // 60px

    val svg = sheet.toSvg(ref"A1:A1")

    // Find the text element containing our cell text
    val yMatch = """<text[^>]*y="(\d+)"[^>]*>TopText</text>""".r.findFirstMatchIn(svg)
    yMatch match
      case Some(m) =>
        val textY = m.group(1).toInt
        // With showLabels=false: cellY=0, Top-aligned: baselineOffset + padding ≈ 13
        assert(textY < 25, s"Top-aligned text y=$textY should be < 25 (near top of cell)")
      case None => fail(s"Could not find cell text y position in: $svg")
  }

  test("toSvg: VAlign.Bottom positions text near bottom of cell") {
    val bottomStyle = CellStyle.default.withAlign(Align(vertical = VAlign.Bottom))
    val sheet = Sheet("Test")
      .put(ref"A1" -> "BottomText")
      .unsafe
      .withCellStyle(ref"A1", bottomStyle)
      .setRowProperties(Row.from0(0), RowProperties(height = Some(45.0))) // 60px

    val svg = sheet.toSvg(ref"A1:A1")

    val yMatch = """<text[^>]*y="(\d+)"[^>]*>BottomText</text>""".r.findFirstMatchIn(svg)
    yMatch match
      case Some(m) =>
        val textY = m.group(1).toInt
        // With showLabels=false: cellY=0, height=60, bottom = 60, text should be near 56 (60 - 4 padding)
        assert(textY > 45, s"Bottom-aligned text y=$textY should be > 45 (near bottom of cell)")
      case None => fail(s"Could not find cell text y position in: $svg")
  }

  test("toSvg: VAlign.Middle positions text in center") {
    val middleStyle = CellStyle.default.withAlign(Align(vertical = VAlign.Middle))
    val sheet = Sheet("Test")
      .put(ref"A1" -> "MiddleText")
      .unsafe
      .withCellStyle(ref"A1", middleStyle)
      .setRowProperties(Row.from0(0), RowProperties(height = Some(45.0))) // 60px

    val svg = sheet.toSvg(ref"A1:A1")

    val yMatch = """<text[^>]*y="(\d+)"[^>]*>MiddleText</text>""".r.findFirstMatchIn(svg)
    yMatch match
      case Some(m) =>
        val textY = m.group(1).toInt
        // With showLabels=false: cellY=0, height=60, middle should be around y=30
        assert(textY > 20 && textY < 40, s"Middle-aligned text y=$textY should be 20-40 (center)")
      case None => fail(s"Could not find cell text y position in: $svg")
  }

  test("toSvg: default alignment is Bottom") {
    // No explicit VAlign set - should default to Bottom (Excel default)
    val sheet = Sheet("Test")
      .put(ref"A1" -> "DefaultText")
      .setRowProperties(Row.from0(0), RowProperties(height = Some(45.0))) // 60px

    val svg = sheet.toSvg(ref"A1:A1")

    val yMatch = """<text[^>]*y="(\d+)"[^>]*>DefaultText</text>""".r.findFirstMatchIn(svg)
    yMatch match
      case Some(m) =>
        val textY = m.group(1).toInt
        // With showLabels=false: cellY=0, Default is Bottom: height=60, bottom ≈ 56
        assert(textY > 45, s"Default (Bottom) text y=$textY should be > 45")
      case None => fail(s"Could not find cell text y position in: $svg")
  }

  // ========== Merged Cells Tests ==========

  test("toSvg: merged cells render as single rect spanning columns") {
    val range = CellRange.parse("A1:B1").toOption.get
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Merged")
      .merge(range)
      .setColumnProperties(Column.from0(0), ColumnProperties(width = Some(10.0))) // 85px
      .setColumnProperties(Column.from0(1), ColumnProperties(width = Some(10.0))) // 85px

    val svg = sheet.toSvg(ref"A1:B1")

    // Merged rect should span both columns (85 + 85 = 170px)
    assert(svg.contains("""width="170""""), s"Merged rect should be 170px wide, got: $svg")
    // Count cell rect elements (not the svg element itself)
    val cellRectCount = """<rect[^>]*class="cell"[^>]*>""".r.findAllIn(svg).length
    assertEquals(cellRectCount, 1, s"Should have exactly 1 cell rect (merged), got: $svg")
  }

  test("toSvg: merged cells render as single rect spanning rows") {
    val range = CellRange.parse("A1:A2").toOption.get
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Merged")
      .merge(range)
      .setRowProperties(Row.from0(0), RowProperties(height = Some(30.0))) // 40px
      .setRowProperties(Row.from0(1), RowProperties(height = Some(30.0))) // 40px

    val svg = sheet.toSvg(ref"A1:A2")

    // Merged rect should span both rows (40 + 40 = 80px)
    assert(svg.contains("""height="80""""), s"Merged rect should be 80px tall, got: $svg")
  }

  test("toSvg: interior merge cells are not rendered") {
    val range = CellRange.parse("A1:B1").toOption.get
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Anchor")
      .put(ref"B1" -> "Hidden")
      .merge(range)
      .setColumnProperties(Column.from0(0), ColumnProperties(width = Some(10.0))) // 75px
      .setColumnProperties(Column.from0(1), ColumnProperties(width = Some(10.0))) // 75px

    val svg = sheet.toSvg(ref"A1:B1")

    assert(svg.contains("Anchor"), "Anchor cell text should appear")
    assert(!svg.contains("Hidden"), s"Interior cell text should not appear, got: $svg")
  }

  test("toSvg: merged cell text is centered in merged area") {
    val range = CellRange.parse("A1:C1").toOption.get
    val centerStyle = CellStyle.default.withAlign(Align(horizontal = HAlign.Center))
    val sheet = Sheet("Test")
      .put(ref"A1" -> "CenterMerged")
      .unsafe
      .withCellStyle(ref"A1", centerStyle)
      .merge(range)
      .setColumnProperties(Column.from0(0), ColumnProperties(width = Some(10.0))) // 75px
      .setColumnProperties(Column.from0(1), ColumnProperties(width = Some(10.0))) // 75px
      .setColumnProperties(Column.from0(2), ColumnProperties(width = Some(10.0))) // 75px

    val svg = sheet.toSvg(ref"A1:C1")

    // Text should have middle anchor for centered merged cell
    assert(svg.contains("text-anchor=\"middle\""), s"Should have middle anchor, got: $svg")
    assert(svg.contains("CenterMerged"), "Merged text should appear")
  }

  test("toSvg: unmerged cells not affected by merge logic") {
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Normal")
      .put(ref"B1" -> "Also Normal")
      .setColumnProperties(Column.from0(0), ColumnProperties(width = Some(10.0))) // 75px
      .setColumnProperties(Column.from0(1), ColumnProperties(width = Some(10.0))) // 75px

    val svg = sheet.toSvg(ref"A1:B1")

    // Both cells should appear with their text
    assert(svg.contains("Normal"), "First cell should appear")
    assert(svg.contains("Also Normal"), "Second cell should appear")
    // Should have exactly 2 text elements for cells (not counting headers)
    val cellTextCount = """class="cell-text"""".r.findAllIn(svg).length
    assertEquals(cellTextCount, 2, s"Should have 2 cell text elements, got: $svg")
  }

  // ==================== Text Wrapping Tests ====================

  test("toSvg: wrapText=true renders multiple tspan elements for long text") {
    // Create a narrow column with wrapping enabled
    val wrapStyle = CellStyle.default.withAlign(Align(wrapText = true))
    val sheet = Sheet("Test")
      .put(ref"A1" -> "This is a long sentence that should wrap")
      .unsafe
      .withCellStyle(ref"A1", wrapStyle)
      .setColumnProperties(Column.from0(0), ColumnProperties(width = Some(8.0))) // ~61px, narrow

    val svg = sheet.toSvg(ref"A1:A1")

    // With wrapping, should have multiple tspan elements
    val tspanCount = "<tspan".r.findAllIn(svg).length
    assert(tspanCount > 1, s"Wrapped text should have multiple tspans, got $tspanCount: $svg")
  }

  test("toSvg: wrapText=false renders single text element") {
    // Default alignment (no wrapText)
    val sheet = Sheet("Test")
      .put(ref"A1" -> "This is a long sentence that should NOT wrap")
      .setColumnProperties(Column.from0(0), ColumnProperties(width = Some(8.0))) // narrow

    val svg = sheet.toSvg(ref"A1:A1")

    // Without wrapping, should have a single text element (no tspan with y attribute)
    val tspanCount = """<tspan x=""".r.findAllIn(svg).length
    assertEquals(tspanCount, 0, s"Non-wrapped text should not have tspan elements, got: $svg")
  }

  test("toSvg: wrapped text tspans have increasing y values") {
    val wrapStyle = CellStyle.default.withAlign(Align(wrapText = true))
    val sheet = Sheet("Test")
      .put(ref"A1" -> "First Second Third Fourth Fifth")
      .unsafe
      .withCellStyle(ref"A1", wrapStyle)
      .setColumnProperties(Column.from0(0), ColumnProperties(width = Some(6.0))) // very narrow ~47px
      .setRowProperties(Row.from0(0), RowProperties(height = Some(60.0))) // tall row for wrapped text

    val svg = sheet.toSvg(ref"A1:A1")

    // Extract y values from tspans
    val yPattern = """<tspan[^>]*y="(\d+)"[^>]*>""".r
    val yValues = yPattern.findAllMatchIn(svg).map(_.group(1).toInt).toList

    assert(yValues.length > 1, s"Should have multiple tspans, got $yValues")
    // Each y should be greater than the previous (lines go down)
    yValues.sliding(2).foreach {
      case List(y1, y2) =>
        assert(y2 > y1, s"y values should increase: $y1 -> $y2, full: $yValues")
      case _ => ()
    }
  }

  test("toSvg: wrapped text with VAlign.Top positions first line near top") {
    val topWrapStyle = CellStyle.default.withAlign(Align(wrapText = true, vertical = VAlign.Top))
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Line one Line two")
      .unsafe
      .withCellStyle(ref"A1", topWrapStyle)
      .setColumnProperties(Column.from0(0), ColumnProperties(width = Some(6.0)))
      .setRowProperties(Row.from0(0), RowProperties(height = Some(80.0)))

    val svg = sheet.toSvg(ref"A1:A1")

    // First tspan should be near top of cell (with showLabels=false, cell starts at y=0)
    val yPattern = """<tspan[^>]*y="(\d+)"[^>]*>""".r
    val yValues = yPattern.findAllMatchIn(svg).map(_.group(1).toInt).toList

    yValues.headOption.foreach { firstY =>
      // Top-aligned text should be in upper portion of cell (within first ~20px of 80px tall cell)
      // Cell starts at y=0, so first line should be around baselineOffset + padding ≈ 13
      assert(firstY < 25, s"First line should be near top, got y=$firstY")
    }
  }

  test("toSvg: single short word does not wrap even with wrapText=true") {
    val wrapStyle = CellStyle.default.withAlign(Align(wrapText = true))
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Hi")
      .unsafe
      .withCellStyle(ref"A1", wrapStyle)
      .setColumnProperties(Column.from0(0), ColumnProperties(width = Some(10.0)))

    val svg = sheet.toSvg(ref"A1:A1")

    // Short text that fits should only have one tspan
    val tspanCount = "<tspan".r.findAllIn(svg).length
    assertEquals(tspanCount, 1, s"Short text should have one tspan: $svg")
    assert(svg.contains("Hi"), "Text content should be preserved")
  }

  // ========== Font Rendering Tests ==========

  test("toSvg: default font-size in CSS is 15px (11pt converted)") {
    val sheet = Sheet("Test").put(ref"A1" -> "Text")

    val svg = sheet.toSvg(ref"A1:A1")

    // Default CSS class should use 15px (11pt * 4/3 ≈ 14.67, rounded to 15)
    assert(
      svg.contains("font-size: 15px"),
      s"Default CSS should have font-size: 15px (11pt converted), got: $svg"
    )
  }

  test("toSvg: non-default font size converts pt to px") {
    import com.tjclp.xl.styles.font.Font
    // 14pt should become ~19px (14 * 4/3 = 18.67)
    val largeFont = CellStyle.default.withFont(Font(sizePt = 14.0))
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Large")
      .unsafe
      .withCellStyle(ref"A1", largeFont)

    val svg = sheet.toSvg(ref"A1:A1")

    // 14pt * 4/3 = 18.67 → 18px (truncated)
    assert(svg.contains("""font-size="18px""""), s"14pt should be 18px, got: $svg")
  }

  test("toSvg: small font size converts correctly") {
    import com.tjclp.xl.styles.font.Font
    // 9pt should become 12px (9 * 4/3 = 12)
    val smallFont = CellStyle.default.withFont(Font(sizePt = 9.0))
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Small")
      .unsafe
      .withCellStyle(ref"A1", smallFont)

    val svg = sheet.toSvg(ref"A1:A1")

    assert(svg.contains("""font-size="12px""""), s"9pt should be 12px, got: $svg")
  }

  test("toSvg: font-family with spaces is quoted") {
    import com.tjclp.xl.styles.font.Font
    val timesFont = CellStyle.default.withFont(Font(name = "Times New Roman"))
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Serif")
      .unsafe
      .withCellStyle(ref"A1", timesFont)

    val svg = sheet.toSvg(ref"A1:A1")

    // Font name should be quoted to handle spaces
    assert(
      svg.contains("""font-family="'Times New Roman'""""),
      s"Font with spaces should be quoted, got: $svg"
    )
  }

  test("toSvg: simple font-family is also quoted") {
    import com.tjclp.xl.styles.font.Font
    val arialFont = CellStyle.default.withFont(Font(name = "Arial"))
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Sans")
      .unsafe
      .withCellStyle(ref"A1", arialFont)

    val svg = sheet.toSvg(ref"A1:A1")

    // Even simple font names should be quoted for consistency
    assert(svg.contains("""font-family="'Arial'""""), s"Font should be quoted, got: $svg")
  }

  test("toSvg: rich text font size converts pt to px") {
    import com.tjclp.xl.richtext.{RichText, TextRun}
    import com.tjclp.xl.styles.font.Font
    // Rich text with 16pt font should become ~21px (16 * 4/3 = 21.33)
    val richText = RichText(TextRun("Big", Some(Font(sizePt = 16.0))))
    val sheet = Sheet("Test").put(ref"A1", CellValue.RichText(richText))

    val svg = sheet.toSvg(ref"A1:A1")

    // 16pt * 4/3 = 21.33 → 21px (truncated)
    assert(svg.contains("""font-size="21px""""), s"Rich text 16pt should be 21px, got: $svg")
  }

  test("toSvg: rich text font-family is quoted") {
    import com.tjclp.xl.richtext.{RichText, TextRun}
    import com.tjclp.xl.styles.font.Font
    val richText = RichText(TextRun("Styled", Some(Font(name = "Comic Sans MS"))))
    val sheet = Sheet("Test").put(ref"A1", CellValue.RichText(richText))

    val svg = sheet.toSvg(ref"A1:A1")

    assert(
      svg.contains("""font-family="'Comic Sans MS'""""),
      s"Rich text font should be quoted, got: $svg"
    )
  }

  // ========== Gridlines Tests ==========

  test("toSvg: gridlines hidden by default") {
    val sheet = Sheet("Test").put(ref"A1" -> "Data")

    val svg = sheet.toSvg(ref"A1:A1")

    // Without showGridlines, cell rects should not have stroke attributes
    assert(!svg.contains("""stroke="#D0D0D0""""), s"Should not have gridline stroke, got: $svg")
  }

  test("toSvg: gridlines shown when showGridlines=true") {
    val sheet = Sheet("Test").put(ref"A1" -> "Data")

    val svg = sheet.toSvg(ref"A1:A1", showGridlines = true)

    // With showGridlines, cell rects should have inline stroke attributes
    assert(
      svg.contains("""stroke="#D0D0D0"""") && svg.contains("""stroke-width="0.5""""),
      s"Cell rects should have inline gridline stroke attributes, got: $svg"
    )
  }

  // ========== Border Rendering Tests ==========

  test("toSvg: cell with bottom border renders as line element") {
    import com.tjclp.xl.styles.border.{Border, BorderSide, BorderStyle}
    val borderStyle = CellStyle.default.withBorder(
      Border.none.withBottom(BorderSide(BorderStyle.Thin, Some(Color.fromRgb(0, 0, 0))))
    )
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Underlined")
      .unsafe
      .withCellStyle(ref"A1", borderStyle)

    val svg = sheet.toSvg(ref"A1:A1")

    // Border should be rendered as <line> element, not rect stroke
    assert(svg.contains("<line"), s"Border should render as <line> element, got: $svg")
    assert(svg.contains("""stroke="#000000""""), s"Border should have black stroke, got: $svg")
    assert(svg.contains("""stroke-width="1""""), s"Thin border should have width 1, got: $svg")
  }

  test("toSvg: cell with medium border has stroke-width 2") {
    import com.tjclp.xl.styles.border.{Border, BorderSide, BorderStyle}
    val borderStyle = CellStyle.default.withBorder(
      Border.none.withBottom(BorderSide(BorderStyle.Medium, Some(Color.fromRgb(0, 0, 0))))
    )
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Medium border")
      .unsafe
      .withCellStyle(ref"A1", borderStyle)

    val svg = sheet.toSvg(ref"A1:A1")

    assert(svg.contains("""stroke-width="2""""), s"Medium border should have width 2, got: $svg")
  }

  test("toSvg: cell with thick border has stroke-width 3") {
    import com.tjclp.xl.styles.border.{Border, BorderSide, BorderStyle}
    val borderStyle = CellStyle.default.withBorder(
      Border.none.withBottom(BorderSide(BorderStyle.Thick, Some(Color.fromRgb(0, 0, 0))))
    )
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Thick border")
      .unsafe
      .withCellStyle(ref"A1", borderStyle)

    val svg = sheet.toSvg(ref"A1:A1")

    assert(svg.contains("""stroke-width="3""""), s"Thick border should have width 3, got: $svg")
  }

  test("toSvg: dashed border has stroke-dasharray") {
    import com.tjclp.xl.styles.border.{Border, BorderSide, BorderStyle}
    val borderStyle = CellStyle.default.withBorder(
      Border.none.withBottom(BorderSide(BorderStyle.Dashed, Some(Color.fromRgb(0, 0, 0))))
    )
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Dashed")
      .unsafe
      .withCellStyle(ref"A1", borderStyle)

    val svg = sheet.toSvg(ref"A1:A1")

    assert(svg.contains("stroke-dasharray"), s"Dashed border should have dasharray, got: $svg")
    assert(svg.contains(""""4,2""""), s"Dashed pattern should be 4,2, got: $svg")
  }

  test("toSvg: dotted border has stroke-dasharray with small pattern") {
    import com.tjclp.xl.styles.border.{Border, BorderSide, BorderStyle}
    val borderStyle = CellStyle.default.withBorder(
      Border.none.withBottom(BorderSide(BorderStyle.Dotted, Some(Color.fromRgb(0, 0, 0))))
    )
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Dotted")
      .unsafe
      .withCellStyle(ref"A1", borderStyle)

    val svg = sheet.toSvg(ref"A1:A1")

    assert(svg.contains("stroke-dasharray"), s"Dotted border should have dasharray, got: $svg")
    assert(svg.contains(""""2,2""""), s"Dotted pattern should be 2,2, got: $svg")
  }

  test("toSvg: double border renders as two parallel lines") {
    import com.tjclp.xl.styles.border.{Border, BorderSide, BorderStyle}
    val borderStyle = CellStyle.default.withBorder(
      Border.none.withBottom(BorderSide(BorderStyle.Double, Some(Color.fromRgb(0, 0, 0))))
    )
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Double")
      .unsafe
      .withCellStyle(ref"A1", borderStyle)

    val svg = sheet.toSvg(ref"A1:A1")

    // Double border should have two <line> elements
    val lineCount = "<line".r.findAllIn(svg).length
    assert(lineCount >= 2, s"Double border should have at least 2 lines, got $lineCount: $svg")
  }

  test("toSvg: all four borders render independently") {
    import com.tjclp.xl.styles.border.{Border, BorderSide, BorderStyle}
    val allBorders = CellStyle.default.withBorder(
      Border(
        left = BorderSide(BorderStyle.Thin, Some(Color.fromRgb(0, 0, 0))),
        right = BorderSide(BorderStyle.Thin, Some(Color.fromRgb(0, 0, 0))),
        top = BorderSide(BorderStyle.Thin, Some(Color.fromRgb(0, 0, 0))),
        bottom = BorderSide(BorderStyle.Thin, Some(Color.fromRgb(0, 0, 0)))
      )
    )
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Boxed")
      .unsafe
      .withCellStyle(ref"A1", allBorders)

    val svg = sheet.toSvg(ref"A1:A1")

    // Should have 4 line elements (one per side)
    val lineCount = "<line".r.findAllIn(svg).length
    assertEquals(lineCount, 4, s"Box border should have 4 lines, got: $svg")
  }

  test("toSvg: border color from theme is resolved") {
    import com.tjclp.xl.styles.border.{Border, BorderSide, BorderStyle}
    import com.tjclp.xl.styles.color.{ThemePalette, ThemeSlot}
    // Use a theme color (accent1 = blue in default theme)
    val blueColor = Color.Theme(ThemeSlot.Accent1, 0.0)
    val borderStyle = CellStyle.default.withBorder(
      Border.none.withBottom(BorderSide(BorderStyle.Thin, Some(blueColor)))
    )
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Blue border")
      .unsafe
      .withCellStyle(ref"A1", borderStyle)

    val svg = sheet.toSvg(ref"A1:A1", theme = ThemePalette.office)

    // Theme color should be resolved to actual hex color (accent1 in office theme is blue)
    assert(svg.contains("stroke=\"#"), s"Theme color should resolve to hex, got: $svg")
    // Should NOT contain "theme" in the output
    assert(!svg.contains("theme("), s"Should not have raw theme reference, got: $svg")
  }

  test("toSvg: cell without border has no border lines") {
    val sheet = Sheet("Test").put(ref"A1" -> "No border")

    val svg = sheet.toSvg(ref"A1:A1")

    // Should have cell-borders group but it should be empty (or minimal)
    assert(svg.contains("""class="cell-borders""""), s"Should have borders group, got: $svg")
    // Should not have any <line> elements in borders group (only in cell rect as stroke if gridlines)
    val bordersSection = svg.split("cell-borders")(1).split("</g>")(0)
    val lineCount = "<line".r.findAllIn(bordersSection).length
    assertEquals(lineCount, 0, s"Cell without border should have no border lines, got: $svg")
  }

  test("toSvg: borders rendered in separate layer above backgrounds") {
    import com.tjclp.xl.styles.border.{Border, BorderSide, BorderStyle}
    val borderStyle = CellStyle.default.withBorder(
      Border.none.withBottom(BorderSide(BorderStyle.Thin, Some(Color.fromRgb(0, 0, 0))))
    )
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Test")
      .unsafe
      .withCellStyle(ref"A1", borderStyle)

    val svg = sheet.toSvg(ref"A1:A1")

    // Verify rendering order: cells -> cell-borders -> cell-text-layer
    val cellsIdx = svg.indexOf("""class="cells"""")
    val bordersIdx = svg.indexOf("""class="cell-borders"""")
    val textIdx = svg.indexOf("""class="cell-text-layer"""")

    assert(cellsIdx < bordersIdx, "Cells should come before borders")
    assert(bordersIdx < textIdx, "Borders should come before text")
  }

  // ========== Pattern Fill Tests (v0.3.0 Regression #84) ==========

  test("toSvg: pattern fill uses background color") {
    import com.tjclp.xl.styles.fill.PatternType
    // Create a LightGray pattern fill with light gray background
    val patternFill = Fill.Pattern(
      foreground = Color.fromRgb(0, 0, 0), // Black foreground
      background = Color.fromRgb(200, 200, 200), // Light gray background
      pattern = PatternType.LightGray
    )
    val style = CellStyle.default.withFill(patternFill)

    val sheet = Sheet("Test")
      .put(ref"A1" -> "Pattern")
      .unsafe
      .withCellStyle(ref"A1", style)

    val svg = sheet.toSvg(ref"A1:A1")

    // Should use the background color (#C8C8C8 = 200,200,200)
    assert(svg.contains("""fill="#C8C8C8""""), s"Pattern fill should use background color, got: $svg")
  }

  test("toSvg: pattern fill does not render as white") {
    import com.tjclp.xl.styles.fill.PatternType
    // Create a pattern fill with a distinct background color
    val patternFill = Fill.Pattern(
      foreground = Color.fromRgb(0, 0, 0),
      background = Color.fromRgb(0, 128, 255), // Blue background
      pattern = PatternType.Gray125
    )
    val style = CellStyle.default.withFill(patternFill)

    val sheet = Sheet("Test")
      .put(ref"A1" -> "Blue Pattern")
      .unsafe
      .withCellStyle(ref"A1", style)

    val svg = sheet.toSvg(ref"A1:A1")

    // Should not render as white (the old bug behavior)
    assert(!svg.contains("""fill="#FFFFFF"""") || svg.contains("""fill="#0080FF""""),
      s"Pattern fill should not be white, should be blue #0080FF, got: $svg")
    // Should use the background color
    assert(svg.contains("""fill="#0080FF""""), s"Pattern fill should use blue background, got: $svg")
  }

  test("toSvg: solid fill still works correctly") {
    val solidFill = Fill.Solid(Color.fromRgb(255, 0, 0)) // Red
    val style = CellStyle.default.withFill(solidFill)

    val sheet = Sheet("Test")
      .put(ref"A1" -> "Solid Red")
      .unsafe
      .withCellStyle(ref"A1", style)

    val svg = sheet.toSvg(ref"A1:A1")

    // Solid fill should still work as before
    assert(svg.contains("""fill="#FF0000""""), s"Solid fill should be red, got: $svg")
  }

  test("toSvg: Fill.None renders default background") {
    val sheet = Sheet("Test")
      .put(ref"A1" -> "No fill")

    val svg = sheet.toSvg(ref"A1:A1")

    // Default cells should have white background
    assert(svg.contains("""fill="#FFFFFF""""), s"No fill should render as white, got: $svg")
  }

  // ========== Text Clipping Tests (GH-146) ==========

  test("toSvg: text clipping prevents overflow beyond cell boundary") {
    // Test case from GH-146: text should be clipped when adjacent cell has content
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Sales Report") // 12 chars, ~80-90px
      .put(ref"B1" -> "Revenue")

    val svg = sheet.toSvg(ref"A1:B1")

    // Should have defs section
    assert(svg.contains("<defs>"), "Should have defs section")
    assert(svg.contains("</defs>"), "Should close defs section")

    // Should have clip paths for both cells
    assert(svg.contains("""<clipPath id="clip-0-0">"""), "Should have clip path for A1")
    assert(svg.contains("""<clipPath id="clip-72-0">"""), "Should have clip path for B1")

    // Text elements should reference their clip paths
    assert(svg.contains("""clip-path="url(#clip-0-0)">Sales Report</text>"""),
      s"A1 text should have clip-path attribute, got: $svg")
    assert(svg.contains("""clip-path="url(#clip-72-0)">Revenue</text>"""),
      s"B1 text should have clip-path attribute, got: $svg")

    // Clip path for A1 should be cell width only (72px) since B1 is not empty
    assert(svg.contains("""<clipPath id="clip-0-0"><rect x="0" y="0" width="72" height="20"/>"""),
      s"A1 clip should be 72px (single cell), got: $svg")
  }

  test("toSvg: text can overflow into empty cells with expanded clip region") {
    // Long text in A1, B1 and C1 are empty → should allow overflow with expanded clip
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Long text here") // ~85-90px, overflows into B1

    val svg = sheet.toSvg(ref"A1:C1")

    // Should have clip path for A1
    assert(svg.contains("""<clipPath id="clip-0-0">"""), "Should have clip path for A1")

    // Clip region should be expanded to allow overflow into empty B1
    // Width should be >72px (at least 2 cells = 144px)
    val clipPattern = """<clipPath id="clip-0-0"><rect x="0" y="0" width="(\d+)" """.r
    val widthOpt = clipPattern.findFirstMatchIn(svg).map(_.group(1).toInt)
    assert(widthOpt.exists(_ >= 144),
      s"Clip region should be ≥144px for overflow into empty cells, got: ${widthOpt.getOrElse(0)}")
  }

  test("toSvg: all cells have clip paths in defs section") {
    val sheet = Sheet("Test")
      .put(ref"A1" -> "A")
      .put(ref"B1" -> "B")
      .put(ref"A2" -> "C")

    val svg = sheet.toSvg(ref"A1:B2")

    // Should have 4 clip paths (A1, B1, A2, B2)
    val clipPathCount = """<clipPath id=""".r.findAllIn(svg).length
    assertEquals(clipPathCount, 4, s"Should have 4 clip paths for 4 cells, got: $clipPathCount")

    // All clip paths should be inside defs
    val defsStart = svg.indexOf("<defs>")
    val defsEnd = svg.indexOf("</defs>")
    assert(defsStart < defsEnd, "Defs section should be properly closed")

    // All clip paths should appear before </defs>
    val firstClipPath = svg.indexOf("""<clipPath""")
    val lastClipPath = svg.lastIndexOf("""</clipPath>""")
    assert(firstClipPath > defsStart && firstClipPath < defsEnd,
      "Clip paths should be inside defs section")
    assert(lastClipPath > defsStart && lastClipPath < defsEnd,
      "All clip paths should be inside defs section")
  }

  test("toSvg: defs section appears before cells group") {
    val sheet = Sheet("Test").put(ref"A1" -> "Test")
    val svg = sheet.toSvg(ref"A1:A1")

    val defsIdx = svg.indexOf("<defs>")
    val cellsIdx = svg.indexOf("""<g class="cells">""")

    assert(defsIdx < cellsIdx,
      s"Defs section should appear before cells group, defs at $defsIdx, cells at $cellsIdx")
  }
