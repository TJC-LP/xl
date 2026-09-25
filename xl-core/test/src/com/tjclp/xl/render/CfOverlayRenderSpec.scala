package com.tjclp.xl.render

import java.util.Locale

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cf.{CfBar, CfOverlay, CfPaint}
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.macros.ref
import com.tjclp.xl.render.syntax.*
import com.tjclp.xl.richtext.{RichText, TextRun}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.sheets.styleSyntax.*
import com.tjclp.xl.styles.{CellStyle, Dxf, DxfFont}
import com.tjclp.xl.styles.alignment.Align
import com.tjclp.xl.styles.border.{Border, BorderSide, BorderStyle}
import com.tjclp.xl.styles.color.{Color, ThemePalette, ThemeSlot}
import com.tjclp.xl.styles.fill.Fill
import com.tjclp.xl.styles.font.{Font, Underline}
import com.tjclp.xl.styles.numfmt.NumFmt
import com.tjclp.xl.unsafe.*

import munit.ScalaCheckSuite
import org.scalacheck.Gen
import org.scalacheck.Prop.forAll

/**
 * GH-497: the renderers paint a precomputed conditional-formatting overlay — hand-built here, no
 * evaluator — through the overlay overloads, identically in SVG and HTML, and the CF-blind entry
 * points stay byte-identical.
 */
class CfOverlayRenderSpec extends ScalaCheckSuite:

  private val pink = Color.Rgb(0xffffc7ce)
  private val darkRed = Color.Rgb(0xff9c0006)
  private val barBlue = Color.Rgb(0xff638ec6)
  private val range = ref"A1:C3"

  private def overlay(entries: (ARef, CfPaint)*): CfOverlay = CfOverlay(entries.toMap)
  private def paint(dxf: Dxf): CfPaint = CfPaint(dxf, None)
  private def bar(fraction: Double, showValue: Boolean = true): CfPaint =
    CfPaint(Dxf(), Some(CfBar(fraction, barBlue, showValue)))

  /** A1 text, B2 = 50, C3 = 0.5; no cell carries a style. */
  private val plain: Sheet =
    Sheet("CF")
      .put(ref"A1", CellValue.Text("label"))
      .put(ref"B2", CellValue.Number(BigDecimal(50)))
      .put(ref"C3", CellValue.Number(BigDecimal("0.5")))

  /** Styled cells over every per-cell path: merge, borders, wrap, rich text, a spill. */
  private val styled: Sheet =
    val boxed = CellStyle.default.withBorder(Border.all(BorderStyle.Thin))
    val wrap = CellStyle.default.withAlign(Align(wrapText = true))
    val rich = RichText(Vector(TextRun("Bold", Some(Font.default.withBold(true))), TextRun(" run")))
    Sheet("Styled")
      .put(ref"A1" -> "a long label that spills over the empty cells to its right")
      .put(ref"A2" -> 1234.5)
      .put(ref"B2" -> "wrapped words in a narrow cell")
      .put(ref"C3", CellValue.RichText(rich))
      .put(ref"A4" -> "merged")
      .unsafe
      .withCellStyle(ref"A2", boxed)
      .withCellStyle(ref"B2", wrap)
      .merge(CellRange(ref"A4", ref"B4"))

  // ========== SVG probes ==========

  /** The fill of the cell background rect drawn at (x, y). */
  private def rectFill(svg: String, x: Int, y: Int): Option[String] =
    s"""<rect x="$x" y="$y" width="\\d+" height="\\d+" fill="(#[0-9A-F]{6})"[^>]*class="cell"/>""".r
      .findFirstMatchIn(svg)
      .map(_.group(1))

  /** The attributes of the `<text>` drawn under a cell's clip. */
  private def textAttrs(svg: String, a1: String): Option[String] =
    s"""<text ([^>]*clip-path="url\\(#clip-$a1\\)"[^>]*)>""".r.findFirstMatchIn(svg).map(_.group(1))

  /** The text drawn under a cell's clip. */
  private def textUnder(svg: String, a1: String): Option[String] =
    s"""<text [^>]*clip-path="url\\(#clip-$a1\\)"[^>]*>([^<]*)</text>""".r
      .findFirstMatchIn(svg)
      .map(_.group(1))

  private def bars(svg: String): List[String] =
    """<rect [^>]*class="cf-bar"/>""".r.findAllMatchIn(svg).map(_.matched).toList

  // ========== HTML probes ==========

  /** The style attribute of the td whose body is `body`. */
  private def tdStyle(html: String, body: String): Option[String] =
    s"""<td style="([^"]*)">${java.util.regex.Pattern.quote(body)}</td>""".r
      .findFirstMatchIn(html)
      .map(_.group(1))

  /** Every td's style attribute, in document order. */
  private def tdStyles(html: String): List[String] =
    """<td style="([^"]*)">""".r.findAllMatchIn(html).map(_.group(1)).toList

  // ========== Invariants ==========

  test("the CF-blind entry points are byte-identical to an empty overlay, SVG and HTML") {
    List(plain -> range, styled -> ref"A1:D5").foreach { (sheet, r) =>
      assertEquals(sheet.toSvg(r, CfOverlay.empty), sheet.toSvg(r))
      assertEquals(sheet.toHtml(r, CfOverlay.empty), sheet.toHtml(r))
      assertEquals(
        SvgRenderer.toSvg(sheet, r, true, ThemePalette.office, true, true, CfOverlay.empty),
        SvgRenderer.toSvg(sheet, r, showLabels = true, showGridlines = true)
      )
      assertEquals(
        HtmlRenderer
          .toHtml(sheet, r, true, false, ThemePalette.office, false, true, CfOverlay.empty),
        HtmlRenderer.toHtml(sheet, r, includeComments = false, showLabels = true)
      )
      assertEquals(
        sheet.toSvg(r, false, ThemePalette.office, false, true, CfOverlay.empty),
        sheet.toSvg(r, includeStyles = false, showGridlines = true)
      )
      assertEquals(
        sheet.toHtml(r, true, true, ThemePalette.office, false, false, CfOverlay.empty),
        sheet.toHtml(r)
      )
    }
  }

  test("includeStyles = false ignores the overlay entirely (CF is styling)") {
    val o = overlay(
      ref"B2" -> CfPaint(Dxf.fill(pink), Some(CfBar(0.5, barBlue, showValue = false))),
      ref"C1" -> paint(Dxf.fill(pink))
    )
    assertEquals(
      SvgRenderer.toSvg(plain, range, false, ThemePalette.office, false, false, o),
      plain.toSvg(range, includeStyles = false)
    )
    assertEquals(
      HtmlRenderer.toHtml(plain, range, false, true, ThemePalette.office, false, false, o),
      plain.toHtml(range, includeStyles = false)
    )
  }

  // ========== SVG ==========

  test("SVG: a dxf fill paints a present cell's rect and leaves its neighbours white") {
    val svg = plain.toSvg(range, overlay(ref"B2" -> paint(Dxf.fill(pink))))
    assertEquals(rectFill(svg, 72, 20), Some("#FFC7CE"), svg)
    assertEquals(rectFill(svg, 0, 20), Some("#FFFFFF"), svg)
    assertEquals(rectFill(svg, 144, 20), Some("#FFFFFF"), svg)
  }

  test("SVG: a dxf fill paints a cell the sheet does not hold") {
    val svg = plain.toSvg(range, overlay(ref"C1" -> paint(Dxf.fill(pink))))
    assertEquals(rectFill(svg, 144, 0), Some("#FFC7CE"), svg)
    assertEquals(textUnder(svg, "C1"), None, s"a painted empty cell draws no text: $svg")
  }

  test("SVG: a theme-colour dxf fill resolves through the renderer's theme") {
    val themed = Dxf.fill(Color.Theme(ThemeSlot.Accent1, 0.0))
    val svg = plain.toSvg(range, overlay(ref"B2" -> paint(themed)))
    val expected = Color.Theme(ThemeSlot.Accent1, 0.0).toResolvedHex(ThemePalette.office)
    assertEquals(rectFill(svg, 72, 20), Some(expected), svg)
  }

  test("SVG: a dxf font paints colour, weight and style over the base font") {
    val f = DxfFont(bold = Some(true), italic = Some(true), color = Some(darkRed))
    val svg = plain.toSvg(range, overlay(ref"B2" -> paint(Dxf.font(f))))
    val attrs = textAttrs(svg, "B2").getOrElse(fail(s"no B2 text: $svg"))
    assert(attrs.contains("""fill="#9C0006""""), attrs)
    assert(attrs.contains("""font-weight="bold""""), attrs)
    assert(attrs.contains("""font-style="italic""""), attrs)
  }

  test("SVG: a painted unstyled cell keeps the unstyled 15px text of its neighbours") {
    assert(plain.cells.values.forall(_.styleId.isEmpty), "precondition: nothing styled")
    val svg =
      plain.toSvg(range, overlay(ref"B2" -> paint(Dxf.font(DxfFont(color = Some(darkRed))))))
    assert(textAttrs(svg, "B2").exists(_.contains("""font-size="15px"""")), svg)
    assert(textAttrs(svg, "C3").exists(_.contains("""font-size="15px"""")), svg)
  }

  test("SVG: strike is drawn as line-through, with underline as 'underline line-through'") {
    val struck =
      plain.toSvg(range, overlay(ref"B2" -> paint(Dxf.font(DxfFont(strike = Some(true))))))
    assert(textAttrs(struck, "B2").exists(_.contains("""text-decoration="line-through"""")), struck)
    val both = plain.toSvg(
      range,
      overlay(
        ref"B2" -> paint(Dxf.font(DxfFont(strike = Some(true), underline = Some(Underline.Single))))
      )
    )
    assert(
      textAttrs(both, "B2").exists(_.contains("""text-decoration="underline line-through"""")),
      both
    )
  }

  test("SVG: strike on a rich-text cell decorates its whole <text>") {
    val svg = styled.toSvg(
      ref"A1:D5",
      overlay(ref"C3" -> paint(Dxf.font(DxfFont(strike = Some(true)))))
    )
    val richText = """<text y="[^"]*" class="cell-text"[^>]*clip-path="url\(#clip-C3\)"[^>]*>""".r
      .findFirstIn(svg)
      .getOrElse(fail(s"no rich text: $svg"))
    assert(richText.contains("""text-decoration="line-through""""), richText)
  }

  test("SVG: a dxf border side is drawn") {
    val svg = plain.toSvg(
      range,
      overlay(ref"B2" -> paint(Dxf(border = Some(Border(bottom = BorderSide(BorderStyle.Thick))))))
    )
    assert(
      svg.contains(
        """<line x1="72.0" y1="40.0" x2="144.0" y2="40.0" stroke="#000000" stroke-width="3"/>"""
      ),
      svg
    )
  }

  test("SVG: a dxf number format re-resolves the drawn text") {
    val svg = plain.toSvg(range, overlay(ref"C3" -> paint(Dxf(numFmt = Some(NumFmt.Percent)))))
    assertEquals(textUnder(svg, "C3"), Some("50%"), svg)
  }

  test("SVG and HTML: a dxf number format decides #### from the text it draws") {
    val wide = overlay(ref"C3" -> paint(Dxf(numFmt = Some(NumFmt.Custom("0.000000000000%")))))
    assertEquals(textUnder(plain.toSvg(range), "C3"), Some("0.5"), "fits unpainted")
    assert(textUnder(plain.toSvg(range, wide), "C3").exists(t => t.nonEmpty && t.forall(_ == '#')))
    assert("""<td style="[^"]*">#+</td>""".r.findFirstIn(plain.toHtml(range, wide)).isDefined)
  }

  test("SVG: a data bar is a gradient rect inset 2px, width round((w - 4) * fraction)") {
    val svg = plain.toSvg(range, overlay(ref"B2" -> bar(0.5), ref"C3" -> bar(0.9)))
    assertEquals(
      bars(svg),
      List(
        """<rect x="74" y="22" width="34" height="16" fill="url(#cf-bar-638EC6)" class="cf-bar"/>""",
        """<rect x="146" y="42" width="61" height="16" fill="url(#cf-bar-638EC6)" class="cf-bar"/>"""
      ),
      svg
    )
    val gradients = """<linearGradient id="([^"]+)"""".r.findAllMatchIn(svg).map(_.group(1)).toList
    assertEquals(gradients, List("cf-bar-638EC6"), s"one gradient per colour: $svg")
    assert(
      svg.contains(
        """<linearGradient id="cf-bar-638EC6" x1="0" y1="0" x2="1" y2="0"><stop offset="0" stop-color="#638EC6"/><stop offset="1" stop-color="#FFFFFF"/></linearGradient>"""
      ),
      svg
    )
    // above its cell's background, below the borders and the text
    val barAt = svg.indexOf("""class="cf-bar"""")
    assert(barAt > svg.indexOf("""<rect x="72" y="20""""), svg)
    assert(barAt < svg.indexOf("""<g class="cell-borders">"""), svg)
    assertEquals(textUnder(svg, "B2"), Some("50"), "the value still shows")
  }

  test("SVG: gradients are emitted once per colour, sorted by id") {
    val green = Color.Rgb(0xff63be7b)
    val svg = plain.toSvg(
      range,
      overlay(
        ref"B2" -> bar(0.5),
        ref"C3" -> CfPaint(Dxf(), Some(CfBar(0.5, green, showValue = true))),
        ref"A1" -> bar(0.1)
      )
    )
    val gradients = """<linearGradient id="([^"]+)"""".r.findAllMatchIn(svg).map(_.group(1)).toList
    assertEquals(gradients, List("cf-bar-638EC6", "cf-bar-63BE7B"), svg)
  }

  test("SVG: showValue = false hides the text only where a bar is drawn") {
    val svg = plain.toSvg(range, overlay(ref"B2" -> bar(0.5, showValue = false)))
    assertEquals(bars(svg).size, 1, svg)
    assertEquals(textUnder(svg, "B2"), None, svg)
    assertEquals(textUnder(svg, "C3"), Some("0.5"), svg)
  }

  test("SVG: a merged anchor's paint fills the whole merged rect") {
    val svg = styled.toSvg(ref"A1:D5", overlay(ref"A4" -> paint(Dxf.fill(pink))))
    assert(
      """<rect x="0" y="\d+" width="144" height="\d+" fill="#FFC7CE"""".r
        .findFirstIn(svg)
        .isDefined,
      svg
    )
  }

  test("SVG: a CF fill on an empty neighbour is painted under the spilling text") {
    val svg = styled.toSvg(ref"A1:D5", overlay(ref"B1" -> paint(Dxf.fill(pink))))
    assertEquals(rectFill(svg, 72, 0), Some("#FFC7CE"), svg)
    val clip = """<clipPath id="clip-A1"><rect x="0" y="0" width="(\d+)"""".r
      .findFirstMatchIn(svg)
      .map(_.group(1).toInt)
    assert(clip.exists(_ > 72), s"the paint does not block the spill: $svg")
  }

  // ========== HTML ==========

  test("HTML: a dxf fill and font paint the td") {
    val f = DxfFont(bold = Some(true), color = Some(darkRed))
    val html = plain.toHtml(range, overlay(ref"B2" -> paint(Dxf.fillAndFont(pink, f))))
    val style = tdStyle(html, "50").getOrElse(fail(s"no B2 td: $html"))
    assert(style.contains("background-color: #FFC7CE"), style)
    assert(style.contains("color: #9C0006"), style)
    assert(style.contains("font-weight: bold"), style)
  }

  test("HTML: strike is text-decoration line-through, with underline 'underline line-through'") {
    val struck =
      plain.toHtml(range, overlay(ref"B2" -> paint(Dxf.font(DxfFont(strike = Some(true))))))
    assert(tdStyle(struck, "50").exists(_.contains("text-decoration: line-through")), struck)
    val both = plain.toHtml(
      range,
      overlay(
        ref"B2" -> paint(Dxf.font(DxfFont(strike = Some(true), underline = Some(Underline.Single))))
      )
    )
    assert(tdStyle(both, "50").exists(_.contains("text-decoration: underline line-through")), both)
  }

  test("HTML: a painted empty td carries the fill and borders; an unpainted one is unchanged") {
    val o = overlay(
      ref"C1" -> paint(
        Dxf(fill = Some(Fill.Solid(pink)), border = Some(Border.all(BorderStyle.Thin)))
      )
    )
    val painted = tdStyles(plain.toHtml(range, o))
    val unpainted = tdStyles(plain.toHtml(range))
    // row 1: A1, B1, C1
    assert(painted(2).contains("background-color: #FFC7CE"), painted(2))
    assert(painted(2).contains("border-top: 1px solid #000000"), painted(2))
    assertEquals(painted.patch(2, Nil, 1), unpainted.patch(2, Nil, 1), "only C1 changes")
  }

  test("HTML: a data bar is a no-repeat gradient background, the SVG bar's width") {
    val html = plain.toHtml(range, overlay(ref"B2" -> bar(0.5)))
    val style = tdStyle(html, "50").getOrElse(fail(s"no B2 td: $html"))
    assert(
      style.contains(
        "background-image: linear-gradient(to right, #638EC6, #FFFFFF); background-size: 34px calc(100% - 4px); background-position: 2px center; background-repeat: no-repeat"
      ),
      style
    )
  }

  test("HTML: showValue = false renders the barred td empty") {
    val html = plain.toHtml(range, overlay(ref"B2" -> bar(0.5, showValue = false)))
    assert(!html.contains(">50</td>"), html)
    assert(html.contains("background-size: 34px"), html)
  }

  test("HTML: CSS is locale-independent") {
    val saved = Locale.getDefault
    try
      Locale.setDefault(Locale.GERMANY)
      val html = plain.toHtml(range, overlay(ref"B2" -> bar(0.45)))
      assert(html.contains("background-size: 31px calc(100% - 4px)"), html)
    finally Locale.setDefault(saved)
  }

  // ========== Parity ==========

  property("SVG and HTML agree on a paint's fill colour and bar width") {
    val genPaint = for
      rgb <- Gen.choose(0, 0xffffff)
      fraction <- Gen.choose(0.1, 0.9)
    yield CfPaint(Dxf.fill(Color.Rgb(0xff000000 | rgb)), Some(CfBar(fraction, barBlue, true)))
    forAll(genPaint) { p =>
      val o = overlay(ref"B2" -> p)
      val svg = plain.toSvg(range, o)
      val html = plain.toHtml(range, o)
      val hex = rectFill(svg, 72, 20).getOrElse(fail(svg))
      val svgWidth = """<rect x="74" y="22" width="(\d+)"""".r
        .findFirstMatchIn(svg)
        .map(_.group(1))
        .getOrElse(fail(svg))
      val style = tdStyle(html, "50").getOrElse(fail(html))
      assert(style.contains(s"background-color: $hex"), s"$hex / $style")
      assert(style.contains(s"background-size: ${svgWidth}px"), s"$svgWidth / $style")
    }
  }
