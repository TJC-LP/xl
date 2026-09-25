package com.tjclp.xl.render

import com.tjclp.xl.addressing.{CellRange, Column, Row}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.macros.ref
import com.tjclp.xl.render.syntax.*
import com.tjclp.xl.richtext.{RichText, TextRun}
import com.tjclp.xl.sheets.{ColumnProperties, PageSetup, RowProperties, Sheet}
import com.tjclp.xl.sheets.styleSyntax.*
import com.tjclp.xl.styles.{CellStyle, StyleRegistry}
import com.tjclp.xl.styles.alignment.{Align, HAlign, VAlign}
import com.tjclp.xl.styles.color.Color
import com.tjclp.xl.styles.fill.Fill
import com.tjclp.xl.styles.font.Font
import com.tjclp.xl.styles.units.StyleId
import com.tjclp.xl.unsafe.*

import munit.FunSuite

/**
 * Layout the two renderers share with Excel: where too-wide text spills (into neighbours that hold
 * no value, whatever their formatting, in the direction its alignment points) and how tall a row
 * with no explicit height is (the height its content needs).
 */
class RenderLayoutSpec extends FunSuite:

  /** Wider than any window below (4-5 default columns of 72px), whatever the font backend. */
  private val long = "Quarterly Revenue Summary by Region and Product Line for the full year"

  private val yellow = CellStyle.default.withFill(Fill.Solid(Color.fromRgb(255, 255, 0)))
  private def aligned(h: HAlign) = CellStyle.default.withAlign(Align(horizontal = h))
  private def sized(pt: Double) = CellStyle.default.withFont(Font.default.withSize(pt))

  override def beforeAll(): Unit =
    assert(
      RenderUtils.measureTextWidth(long, None) > 5 * RenderUtils.DefaultColumnWidthPx,
      "precondition: the long label is wider than every window in this suite"
    )

  // ========== SVG probes ==========

  /** The clip rect (x, width) of a cell's text. */
  private def clipRect(svg: String, a1: String): Option[(Int, Int)] =
    s"""<clipPath id="clip-$a1"><rect x="(-?\\d+)" y="-?\\d+" width="(\\d+)"""".r
      .findFirstMatchIn(svg)
      .map(m => (m.group(1).toInt, m.group(2).toInt))

  /** The (width, fill) of the cell background rect drawn at (x, y). */
  private def cellRect(svg: String, x: Int, y: Int): Option[(Int, String)] =
    s"""<rect x="$x" y="$y" width="(\\d+)" height="\\d+" fill="(#[0-9A-F]{6})"[^>]*class="cell"/>""".r
      .findFirstMatchIn(svg)
      .map(m => (m.group(1).toInt, m.group(2)))

  /** The text x, anchor and content drawn under a cell's clip. */
  private def textUnder(svg: String, a1: String): Option[(Int, String, String)] =
    s"""<text x="(-?\\d+)"[^>]*text-anchor="([a-z]+)"[^>]*clip-path="url\\(#clip-$a1\\)"[^>]*>([^<]*)</text>""".r
      .findFirstMatchIn(svg)
      .map(m => (m.group(1).toInt, m.group(2), m.group(3)))

  /** The baseline y of a cell's single-line text. */
  private def baselineUnder(svg: String, a1: String): Option[Int] =
    s"""<text x="-?\\d+" y="(-?\\d+)"[^>]*clip-path="url\\(#clip-$a1\\)"""".r
      .findFirstMatchIn(svg)
      .map(_.group(1).toInt)

  // ========== Overflow: SVG ==========

  test("SVG: text spills over empty neighbours that carry a fill, each keeping its own fill") {
    val sheet = Sheet("T")
      .put(ref"A1" -> long)
      .unsafe
      .withCellStyle(ref"B1", yellow)
      .withCellStyle(ref"C1", yellow)
      .withCellStyle(ref"D1", yellow)

    val svg = sheet.toSvg(ref"A1:D1")
    assertEquals(clipRect(svg, "A1"), Some((0, 288)), s"the clip spans A1:D1: $svg")
    assertEquals(cellRect(svg, 0, 0), Some((72, "#FFFFFF")), s"A1's fill covers A1 only: $svg")
    List(72, 144, 216).foreach { x =>
      assertEquals(cellRect(svg, x, 0), Some((72, "#FFFF00")), s"neighbour at x=$x: $svg")
    }
    // Text is drawn after every background and border, so it lies over the yellow fills
    val textLayer = svg.indexOf("""<g class="cell-text-layer">""")
    assert(textLayer > svg.lastIndexOf("""fill="#FFFF00""""), s"text above fills: $svg")
    assertEquals(textUnder(svg, "A1").map(_._3), Some(long))
  }

  test("SVG: a neighbour holding any value blocks the spill, even an empty-string result") {
    List[CellValue](
      CellValue.Formula("\"\"", Some(CellValue.Text(""))),
      CellValue.Text(""),
      CellValue.Number(BigDecimal(0))
    ).foreach { blocker =>
      val sheet = Sheet("T").put(ref"A1" -> long).put(ref"B1", blocker)
      val svg = sheet.toSvg(ref"A1:D1")
      assertEquals(clipRect(svg, "A1"), Some((0, 72)), s"$blocker blocks: $svg")
    }
  }

  test("SVG: a merged neighbour blocks the spill, empty or not") {
    val merge = CellRange.parse("B1:C1").toOption.getOrElse(fail("range"))
    val sheet = Sheet("T").put(ref"A1" -> long).merge(merge)
    val svg = sheet.toSvg(ref"A1:D1")
    assertEquals(clipRect(svg, "A1"), Some((0, 72)), s"the merge blocks: $svg")
  }

  test("SVG: right-aligned text spills LEFT into empty cells, keeping its right anchor") {
    val sheet = Sheet("T")
      .put(ref"C1" -> long)
      .put(ref"D1" -> "x")
      .unsafe
      .withCellStyle(ref"C1", aligned(HAlign.Right))
      .withCellStyle(ref"A1", yellow)

    val svg = sheet.toSvg(ref"A1:D1")
    assertEquals(clipRect(svg, "C1"), Some((0, 216)), s"the clip spans A1:C1: $svg")
    assertEquals(
      textUnder(svg, "C1").map(t => (t._1, t._2)),
      Some((144 + 72 - RenderUtils.CellPaddingX, "end")),
      s"anchored at C1's own right edge: $svg"
    )
    assertEquals(cellRect(svg, 0, 0), Some((72, "#FFFF00")), s"A1 keeps its fill: $svg")
    assertEquals(cellRect(svg, 144, 0), Some((72, "#FFFFFF")), s"C1 is not widened: $svg")
  }

  test("SVG: centred text spills both ways, staying centred on its own cell") {
    val sheet = Sheet("T")
      .put(ref"C1" -> long)
      .unsafe
      .withCellStyle(ref"C1", aligned(HAlign.Center))

    val svg = sheet.toSvg(ref"A1:E1")
    assertEquals(clipRect(svg, "C1"), Some((0, 360)), s"the clip spans A1:E1: $svg")
    assertEquals(textUnder(svg, "C1").map(t => (t._1, t._2)), Some((180, "middle")), svg)
  }

  test("SVG: centred text is clipped on a side whose neighbour blocks, never re-centred") {
    val sheet = Sheet("T")
      .put(ref"B1" -> "x")
      .put(ref"C1" -> long)
      .unsafe
      .withCellStyle(ref"C1", aligned(HAlign.Center))

    val svg = sheet.toSvg(ref"A1:E1")
    assertEquals(clipRect(svg, "C1"), Some((144, 216)), s"the clip spans C1:E1 only: $svg")
    assertEquals(textUnder(svg, "C1").map(t => (t._1, t._2)), Some((180, "middle")), svg)
  }

  test("SVG: the render window's edge bounds a leftward spill") {
    val sheet = Sheet("T")
      .put(ref"D1" -> long)
      .unsafe
      .withCellStyle(ref"D1", aligned(HAlign.Right))

    val svg = sheet.toSvg(ref"B1:D1")
    assertEquals(clipRect(svg, "D1"), Some((0, 216)), s"B1 is the leftmost cell drawn: $svg")
  }

  test("SVG: numbers, wrapped text and merged anchors never spill, styled neighbour or not") {
    val wrap = CellStyle.default.withAlign(Align(wrapText = true))
    val merge = CellRange.parse("A3:B3").toOption.getOrElse(fail("range"))
    val sheet = Sheet("T")
      .put(ref"A1" -> 1234567890123.0)
      .put(ref"A2" -> long)
      .put(ref"A3" -> long)
      .unsafe
      .withCellStyle(ref"A2", wrap)
      .withCellStyle(ref"B1", yellow)
      .merge(merge)

    val svg = sheet.toSvg(ref"A1:D3")
    assertEquals(clipRect(svg, "A1").map(_._2), Some(72), s"a number hashes in place: $svg")
    assert(textUnder(svg, "A1").exists(_._3.forall(_ == '#')), s"#### for the number: $svg")
    assertEquals(clipRect(svg, "A2").map(_._2), Some(72), s"wrapped text wraps: $svg")
    assertEquals(clipRect(svg, "A3").map(_._2), Some(144), s"a merge clips to itself: $svg")
  }

  test("SVG: gridlines inside a spill span are suppressed, its outline kept") {
    val sheet = Sheet("T").put(ref"A1" -> long)
    val svg = sheet.toSvg(ref"A1:C1", showGridlines = true)
    val stroked = """<rect [^>]*stroke="#D0D0D0"[^>]*/>""".r.findAllIn(svg).toList
    assertEquals(
      stroked.map(r => """x="(\d+)"""".r.findFirstMatchIn(r).map(_.group(1)).getOrElse("?")),
      List("0"),
      s"one outline for the span, no gridline between A1, B1 and C1: $svg"
    )
    assert(stroked.forall(_.contains("""width="216"""")), s"the outline covers A1:C1: $stroked")
  }

  // ========== Overflow: HTML ==========

  /** The style of the first spilled text box. */
  private def spillStyle(html: String): Option[String] =
    """<div class="xl-overflow" style="([^"]*)">""".r.findFirstMatchIn(html).map(_.group(1))

  private def px(style: String, prop: String): Option[Int] =
    s"""(?:^|; )$prop: (-?\\d+)px""".r.findFirstMatchIn(style).map(_.group(1).toInt)

  private def tds(html: String): List[String] = """<td[^>]*>""".r.findAllIn(html).toList

  test("HTML: text spills over a filled empty neighbour, which keeps its own cell and fill") {
    val sheet = Sheet("T")
      .put(ref"A1" -> long)
      .unsafe
      .withCellStyle(ref"B1", yellow)

    val html = sheet.toHtml(ref"A1:D1")
    assert(!html.contains("colspan"), s"one <td> per cell, never a spanning one: $html")
    val cells = tds(html)
    assertEquals(cells.size, 4, s"A1, B1, C1, D1: $html")
    assert(cells(0).contains("width: 72px"), s"A1 stays one column wide: ${cells(0)}")
    assert(!cells(0).contains("overflow: hidden"), s"A1 lets its text out: ${cells(0)}")
    assert(cells(1).contains("background-color: #FFFF00"), s"B1 keeps its fill: ${cells(1)}")
    val style = spillStyle(html).getOrElse(fail(s"no spilled text box: $html"))
    assertEquals(px(style, "margin-left"), Some(0), style)
    assertEquals(px(style, "width"), Some(288), style)
    assert(style.contains("justify-content: flex-start"), style)
    assert(html.contains(long), html)
  }

  test("HTML: right-aligned text spills left; centred text spills both ways, centred on its cell") {
    val right = Sheet("T")
      .put(ref"C1" -> long)
      .unsafe
      .withCellStyle(ref"C1", aligned(HAlign.Right))
    val rightStyle = spillStyle(right.toHtml(ref"A1:D1")).getOrElse(fail("no right spill"))
    // D1 is empty too, but right-aligned text never spills to its right
    assertEquals((px(rightStyle, "margin-left"), px(rightStyle, "width")), (Some(-144), Some(216)))
    assert(rightStyle.contains("justify-content: flex-end"), rightStyle)

    val centre = Sheet("T")
      .put(ref"C1" -> long)
      .unsafe
      .withCellStyle(ref"C1", aligned(HAlign.Center))
    val centreStyle = spillStyle(centre.toHtml(ref"A1:E1")).getOrElse(fail("no centre spill"))
    assertEquals(
      (px(centreStyle, "margin-left"), px(centreStyle, "width")),
      (Some(-144), Some(360))
    )
    assert(centreStyle.contains("justify-content: center"), centreStyle)

    // Blocked on the left: the box covers C1:E1 and pads its right side so the text's centre
    // stays on C1's centre
    val blocked = centre.put(ref"B1" -> "x")
    val blockedStyle = spillStyle(blocked.toHtml(ref"A1:E1")).getOrElse(fail("no spill"))
    assertEquals(px(blockedStyle, "margin-left"), Some(0), blockedStyle)
    assertEquals(px(blockedStyle, "width"), Some(216), blockedStyle)
    assertEquals(px(blockedStyle, "padding-right"), Some(144), blockedStyle)
  }

  test("HTML: a value-bearing neighbour blocks the spill; numbers never spill") {
    val sheet = Sheet("T")
      .put(ref"A1" -> long)
      .put(ref"B1" -> "x")
      .put(ref"A2" -> 1234567890123.0)
    val html = sheet.toHtml(ref"A1:C2")
    assertEquals(spillStyle(html), None, s"nothing spills: $html")
  }

  /** The style of every merged cell's text box, in document order. */
  private def mergeBoxStyles(html: String): List[String] =
    """<div class="xl-merge" style="([^"]*)">""".r.findAllMatchIn(html).map(_.group(1)).toList

  test("HTML: a merged cell clips its text to the merge, never drawing over a neighbour") {
    // The fixed table layout keeps a merge exactly as wide as its columns, so a merged <td> that
    // did not clip would draw its text over the next cell's value
    def merged(a1: String) = CellRange.parse(a1).toOption.getOrElse(fail(a1))
    val sheet = Sheet("T")
      .put(ref"A1" -> long)
      .put(ref"C1" -> "C1")
      .put(ref"A3" -> long)
      .put(ref"C3" -> "C3")
      .put(ref"A5" -> long)
      .unsafe
      .withCellStyle(ref"A3", aligned(HAlign.Center))
      .withCellStyle(ref"A5", aligned(HAlign.Right))
      .merge(merged("A1:B1"))
      .merge(merged("A3:B3"))
      .merge(merged("A5:B5"))

    val html = sheet.toHtml(ref"A1:E5")
    val spanning = tds(html).filter(_.contains("colspan"))
    assertEquals(spanning.size, 3, html)
    spanning.foreach { td =>
      assert(td.contains("width: 144px"), td)
      assert(td.contains("overflow: hidden"), s"a merge clips to itself: $td")
    }
    assertEquals(spillStyle(html), None, s"merged text never spills, blocked or not: $html")
    // Anchored as SVG anchors it: centred text overflows (and is clipped) on both sides,
    // right-aligned text on its left
    val boxes = mergeBoxStyles(html)
    assertEquals(boxes.map(px(_, "width")), List(Some(144), Some(144), Some(144)), html)
    assertEquals(boxes.map(px(_, "margin-left")), List(Some(0), Some(0), Some(0)), html)
    assert(boxes.forall(_.contains("overflow: hidden")), boxes.toString)
    assertEquals(
      boxes.map("justify-content: ([a-z-]+)".r.findFirstMatchIn(_).map(_.group(1))),
      List(Some("flex-start"), Some("center"), Some("flex-end"))
    )

    val svg = sheet.toSvg(ref"A1:E5")
    List("A1", "A3", "A5").foreach { a1 =>
      assertEquals(clipRect(svg, a1), Some((0, 144)), s"SVG clips $a1 to its merge: $svg")
    }
  }

  /** The widths the `<colgroup>` declares, in order. */
  private def colgroupWidths(html: String): List[Int] =
    """<col style="width: (\d+)px">""".r.findAllMatchIn(html).map(_.group(1).toInt).toList

  /** The width the `<table>` tag's own style declares. */
  private def tableWidth(html: String): Option[Int] =
    """<table style="[^"]*; width: (\d+)px""".r.findFirstMatchIn(html).map(_.group(1).toInt)

  test("HTML: the table is exactly as wide as its columns, so browsers keep the fixed layout") {
    // CSS 2.1 §17.5.2.1: with an auto width a browser lays the table out automatically, and the
    // spilled text box widens its own column instead of lying over the neighbours
    val sheet = Sheet("T")
      .put(ref"A1" -> long)
      .setColumnProperties(Column.from0(1), ColumnProperties(width = Some(20.0)))
      .setColumnProperties(Column.from0(2), ColumnProperties(hidden = true))
      .copy(pageSetup = Some(PageSetup(scale = 50)))
    List(
      sheet.toHtml(ref"A1:E1"),
      sheet.toHtml(ref"A1:E1", showLabels = true),
      sheet.toHtml(ref"A1:E1", applyPrintScale = true),
      sheet.toHtml(ref"A1:E1", includeStyles = false, applyPrintScale = true, showLabels = true)
    ).foreach { html =>
      val cols = colgroupWidths(html)
      assert(cols.nonEmpty, html)
      assertEquals(tableWidth(html), Some(cols.sum), html)
    }
    val plain = Sheet("T").put(ref"A1" -> long)
    assertEquals(tableWidth(plain.toHtml(ref"A1:D1")), Some(4 * 72))
    assertEquals(
      tableWidth(plain.toHtml(ref"A1:D1", showLabels = true)),
      Some(RenderUtils.HeaderWidth + 4 * 72)
    )
  }

  test("text in a hidden column never spills: no box, no outline, no gridline gap, no text") {
    val sheet = Sheet("T")
      .put(ref"A1" -> long)
      .setColumnProperties(Column.from0(0), ColumnProperties(hidden = true))
    val window = ref"A1:E1"
    val widths = RenderUtils.calculateColumnWidths(sheet, window)
    val a1 = sheet.cells.get(ref"A1").getOrElse(fail("A1"))
    assertEquals(
      RenderUtils.overflowSpan(ResolvedCell(a1, sheet), ref"A1", widths(0), widths, sheet, 0, 4),
      OverflowSpan.none
    )

    val html = sheet.toHtml(window)
    assertEquals(spillStyle(html), None, s"no spill box: $html")
    assert(!html.contains(long), s"the hidden text is not in the page: $html")
    val labelled = sheet.toHtml(window, showLabels = true)
    assert(!labelled.contains(long), s"nor with labels: $labelled")
    assert(!labelled.contains(">A</td>"), s"the hidden column's letter is not drawn: $labelled")
    assert(labelled.contains(">B</td>"), labelled)

    val svg = sheet.toSvg(window, showGridlines = true)
    assert(!svg.contains(long), s"the hidden text is not drawn: $svg")
    assert(!svg.contains("""fill="none""""), s"no span outline: $svg")
    val gridlined = """<rect x="(\d+)" y="0" width="(\d+)"[^>]*stroke="#D0D0D0"""".r
      .findAllMatchIn(svg)
      .map(m => (m.group(1).toInt, m.group(2).toInt))
      .toList
    assertEquals(
      gridlined,
      List((0, 0), (0, 72), (72, 72), (144, 72), (216, 72)),
      s"every cell keeps its own gridlines: $svg"
    )
  }

  test("SVG and HTML take the same span from RenderUtils") {
    List(HAlign.General, HAlign.Left, HAlign.Right, HAlign.Center).foreach { h =>
      val sheet = Sheet("T")
        .put(ref"C1" -> long)
        .unsafe
        .withCellStyle(ref"C1", aligned(h))
        .withCellStyle(ref"B1", yellow)
      val (clipX, clipW) = clipRect(sheet.toSvg(ref"A1:E1"), "C1").getOrElse(fail(s"$h: clip"))
      val style = spillStyle(sheet.toHtml(ref"A1:E1")).getOrElse(fail(s"$h: no spill"))
      assertEquals(px(style, "margin-left"), Some(clipX - 144), s"$h: $style")
      assertEquals(px(style, "width"), Some(clipW), s"$h: $style")
    }
  }

  // ========== Row auto-height ==========

  test("autofit: Calibri 11/14/18/24pt rows take Excel's 20/25/31/42px") {
    val sheet = Sheet("T")
      .put(ref"A1" -> "eleven")
      .put(ref"A2" -> "fourteen")
      .put(ref"A3" -> "eighteen")
      .put(ref"A4" -> "twenty-four")
      .unsafe
      .withCellStyle(ref"A1", sized(11))
      .withCellStyle(ref"A2", sized(14))
      .withCellStyle(ref"A3", sized(18))
      .withCellStyle(ref"A4", sized(24))
    assertEquals(RenderUtils.calculateRowHeights(sheet, ref"A1:A4"), Vector(20, 25, 31, 42))
    assertEquals(RenderUtils.getRowHeight(sheet, Row.from0(2)), 31)
  }

  test("autofit: an explicit height wins exactly; hidden rows are 0; empty rows the default") {
    val sheet = Sheet("T")
      .put(ref"A1" -> "big")
      .put(ref"A2" -> "big")
      .unsafe
      .withCellStyle(ref"A1", sized(24))
      .withCellStyle(ref"A2", sized(24))
      .setRowProperties(Row.from0(0), RowProperties(height = Some(12.0)))
      .setRowProperties(Row.from0(1), RowProperties(hidden = true))
    assertEquals(RenderUtils.calculateRowHeights(sheet, ref"A1:A3"), Vector(16, 0, 20))
  }

  test("autofit: the largest font in the row wins, rich-text runs and unseen columns included") {
    val big = TextRun("Big", Some(Font.default.withSize(24)))
    val sheet = Sheet("T")
      .put(ref"A1" -> "small")
      .put(ref"H1" -> "outside the window")
      .put(ref"A2", CellValue.RichText(RichText(TextRun("small "), big)))
      .unsafe
      .withCellStyle(ref"H1", sized(18))
    assertEquals(RenderUtils.calculateRowHeights(sheet, ref"A1:C2"), Vector(31, 42))
  }

  test("autofit: merged cells do not drive the height; the sheet default is the floor") {
    val merge = CellRange.parse("A1:B2").toOption.getOrElse(fail("range"))
    val merged = Sheet("T")
      .put(ref"A1" -> "merged title")
      .unsafe
      .withCellStyle(ref"A1", sized(24))
      .merge(merge)
    assertEquals(RenderUtils.calculateRowHeights(merged, ref"A1:B2"), Vector(20, 20))

    val tallDefault = Sheet("T")
      .put(ref"A1" -> "title")
      .unsafe
      .withCellStyle(ref"A1", sized(18))
      .copy(defaultRowHeight = Some(30.0))
    assertEquals(RenderUtils.calculateRowHeights(tallDefault, ref"A1:A2"), Vector(40, 40))
  }

  test("autofit: text in the book's own base font keeps the book's default row height") {
    // An Arial 10 book: Excel's default row is 12.75pt (17px), and a styled Arial 10 cell sits in
    // it; only a larger font grows the row.
    val arial10 = CellStyle.default.withFont(Font("Arial", 10.0))
    val arial14 = CellStyle.default.withFont(Font("Arial", 14.0))
    val registry = StyleRegistry(Vector(arial10), Map(arial10.canonicalKey -> StyleId(0)))
    val sheet = Sheet("T")
      .copy(styleRegistry = registry, defaultRowHeight = Some(12.75))
      .put(ref"A1" -> "base font, styled")
      .put(ref"A2" -> "base font, unstyled")
      .put(ref"A3" -> "larger")
      .unsafe
      .withCellStyle(ref"A1", arial10)
      .withCellStyle(ref"A3", arial14)
    assertEquals(RenderUtils.calculateRowHeights(sheet, ref"A1:A3"), Vector(17, 17, 25))
  }

  test("autofit: with no stored default row, text in the base font still autofits") {
    // The 20px fallback row is Calibri 11's, not the book's: a book whose Normal style is 18pt
    // and that stores no default height sizes its 18pt text like any other
    val calibri18 = sized(18)
    val registry = StyleRegistry(Vector(calibri18), Map(calibri18.canonicalKey -> StyleId(0)))
    val sheet = Sheet("T")
      .copy(styleRegistry = registry)
      .put(ref"A1" -> "Quarterly gyp")
      .put(ref"A2" -> "bold")
      .unsafe
      .withCellStyle(ref"A1", calibri18.withAlign(Align(vertical = VAlign.Bottom)))
      .withCellStyle(ref"A2", calibri18.withFont(calibri18.font.withBold(true)))
    assertEquals(RenderUtils.calculateRowHeights(sheet, ref"A1:A2"), Vector(31, 31))

    val svg = sheet.toSvg(ref"A1:A1")
    val fontPx = 24 // 18pt
    val baseline = baselineUnder(svg, "A1").getOrElse(fail(s"no text: $svg"))
    assert(baseline + fontPx / 4 <= 31, s"descenders stay above the row's bottom: $baseline")
    assert(baseline - fontPx * 0.95 >= 0, s"ascenders stay below the row's top: $baseline")
  }

  test("autofit: wrapped text takes its drawn line count times the line height") {
    val wrap = CellStyle.default.withAlign(Align(wrapText = true))
    val sheet = Sheet("T")
      .put(ref"A1" -> "one two three four five six seven eight nine ten")
      .unsafe
      .withCellStyle(ref"A1", wrap)
      .setColumnProperties(Column.from0(0), ColumnProperties(width = Some(8.0)))

    val svg = sheet.toSvg(ref"A1:A1")
    val baselines =
      """<tspan x="-?\d+" y="(\d+)">""".r.findAllMatchIn(svg).map(_.group(1).toInt).toList
    assert(baselines.size > 1, s"the text wraps: $svg")
    val height = RenderUtils.calculateRowHeights(sheet, ref"A1:A1").headOption.getOrElse(0)
    assertEquals(height, baselines.size * 20, s"${baselines.size} lines of 11pt: $svg")
    assert(svg.contains(s"""height="$height""""), s"the SVG draws that height: $svg")
    assert(baselines.headOption.exists(_ >= 11), s"the first line clears the top: $baselines")
    assert(
      baselines.lastOption.exists(_ + 4 <= height),
      s"the last line clears the bottom: $baselines"
    )
  }

  test("SVG: the row header and the cells follow the autofitted height") {
    val sheet = Sheet("T").put(ref"A1" -> "Title").unsafe.withCellStyle(ref"A1", sized(18))
    val svg = sheet.toSvg(ref"A1:B2", showLabels = true)
    val header = RenderUtils.HeaderHeight
    assert(
      svg.contains(
        s"""<rect x="0" y="$header" width="${RenderUtils.HeaderWidth}" height="31" class="header"/>"""
      ),
      s"row 1's header is 31px: $svg"
    )
    assert(
      svg.contains(s"""<rect x="${RenderUtils.HeaderWidth}" y="$header" width="72" height="31" """),
      svg
    )
    assert(
      svg.contains(s"""y="${header + 31}" width="${RenderUtils.HeaderWidth}" height="20""""),
      svg
    )
  }

  test("SVG: an autofitted 18pt bottom-aligned line sits fully inside its row") {
    val title = sized(18).withAlign(Align(vertical = VAlign.Bottom))
    val sheet = Sheet("T").put(ref"A1" -> "Quarterly gyp").unsafe.withCellStyle(ref"A1", title)
    val svg = sheet.toSvg(ref"A1:A1")
    val fontPx = 24 // 18pt
    val baseline = baselineUnder(svg, "A1").getOrElse(fail(s"no text: $svg"))
    assert(baseline + fontPx / 4 <= 31, s"descenders stay above the row's bottom: $baseline")
    assert(baseline - fontPx * 0.95 >= 0, s"ascenders stay below the row's top: $baseline")
  }

  test("HTML: rows autofit to the same heights as SVG") {
    val sheet = Sheet("T")
      .put(ref"A1" -> "Title")
      .put(ref"A2" -> "body")
      .unsafe
      .withCellStyle(ref"A1", sized(18))
    val html = sheet.toHtml(ref"A1:B2", showLabels = true)
    val heights =
      """<tr style="height: (\d+)px">""".r.findAllMatchIn(html).map(_.group(1).toInt).toList
    assertEquals(heights, List(RenderUtils.HeaderHeight, 31, 20), html)
  }
