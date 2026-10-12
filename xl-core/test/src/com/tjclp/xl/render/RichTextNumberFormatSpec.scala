package com.tjclp.xl.render

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.display.NumFmtFormatter
import com.tjclp.xl.richtext.{RichText, TextRun}
import com.tjclp.xl.sheets.RowProperties
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.alignment.{Align, VAlign}
import com.tjclp.xl.styles.font.Font
import com.tjclp.xl.styles.numfmt.NumFmt

import munit.FunSuite

class RichTextNumberFormatSpec extends FunSuite:
  private val rich = RichText(Vector(TextRun("a").bold, TextRun("bc").italic))

  test("text sections preserve runs through repeated placeholders, literals and spacers") {
    val fmt = NumFmt.Custom("0;-0;0;\"T:\"@\"/\"@_)")
    val formatted = NumFmtFormatter.formatRichText(rich, fmt)
    assertEquals(formatted.toPlainText, "T:abc/abc ")
    assertEquals(formatted.runs.filter(_.font.isDefined), rich.runs ++ rich.runs)
    assertEquals(NumFmtFormatter.formatValue(CellValue.RichText(rich), fmt), formatted.toPlainText)
    assertEquals(
      RenderUtils.renderedContent(CellValue.RichText(rich), fmt).text,
      formatted.toPlainText
    )
  }

  test("literal-only and empty text sections can replace or hide rich text") {
    assertEquals(
      NumFmtFormatter.formatRichText(rich, NumFmt.Custom("0;0;0;\"hidden\"")).toPlainText,
      "hidden"
    )
    assertEquals(NumFmtFormatter.formatRichText(rich, NumFmt.Custom("0;0;0;")).toPlainText, "")
  }

  test("HTML and SVG render formatted rich text and retain run fonts") {
    val fmt = NumFmt.Custom("0;-0;0;\"T:\"@")
    val sheet = Sheet("S")
      .put(ref"A1", CellValue.RichText(rich))
      .withCellStyle(ref"A1", CellStyle.default.withNumFmt(fmt))
    val html = sheet.toHtml(ref"A1:C1")
    assert(html.contains("T:<b>a</b><i>bc</i>"), html)
    val svg = sheet.toSvg(ref"A1:C1")
    assert(svg.contains(">T:</tspan>"), svg)
    assert(svg.contains("font-weight=\"bold\""), svg)
    assert(svg.contains("font-style=\"italic\""), svg)
  }

  test("rich-text width includes format literals") {
    val value = CellValue.RichText(rich)
    val fmt = NumFmt.Custom("0;-0;0;\"a long prefix:\"@")
    assert(
      RenderUtils.measureCellValueWidth(value, fmt, None) >
        RenderUtils.measureCellValueWidth(value, NumFmt.General, None)
    )
  }

  private def baseline(svg: String): Option[String] =
    """<text [^>]*y="(-?\d+)"[^>]*class="cell-text""".r
      .findFirstMatchIn(svg)
      .map(_.group(1))

  test("replacement and hidden text sections do not retain the source run's large height") {
    val large = CellValue.RichText(RichText.single(TextRun("hidden").size(72)))
    List("0;0;0;\"OK\"", "0;0;0;").foreach { code =>
      val fmt = NumFmt.Custom(code)
      val style = CellStyle.default.withNumFmt(fmt).withAlign(Align(vertical = VAlign.Top))
      val actual = Sheet("S").put(ref"A1", large).withCellStyle(ref"A1", style)
      val expected = actual.put(ref"A1", CellValue.Text("hidden"))
      assertEquals(
        RenderUtils.autofitRowHeights(actual, 0, 0),
        RenderUtils.autofitRowHeights(expected, 0, 0)
      )
      val fixedActual = actual.setRowProperties(ref"A1".row, RowProperties(height = Some(15)))
      val fixedExpected = expected.setRowProperties(ref"A1".row, RowProperties(height = Some(15)))
      val svg = fixedActual.toSvg(ref"A1:A1")
      if code.endsWith(";") then
        assert(!svg.contains(">hidden</tspan>"), svg)
        assertEquals(RenderUtils.measureCellValueWidth(large, fmt, Some(style.font)), 0)
      else
        assertEquals(baseline(svg), baseline(fixedExpected.toSvg(ref"A1:A1")))
        assertEquals(baseline(svg), Some("15"))
        assert(svg.contains(">OK</tspan>"), svg)
    }
  }

  test("literal runs inherit the larger cell font for row height and SVG baseline") {
    val small = CellValue.RichText(RichText.single(TextRun("a").size(8)))
    val style = CellStyle.default
      .withNumFmt(NumFmt.Custom("0;0;0;\"prefix\"@"))
      .withFont(Font.default.withSize(36))
      .withAlign(Align(vertical = VAlign.Top))
    val actual = Sheet("S").put(ref"A1", small).withCellStyle(ref"A1", style)
    val expected = actual.put(ref"A1", CellValue.Text("a"))
    assertEquals(
      RenderUtils.autofitRowHeights(actual, 0, 0),
      RenderUtils.autofitRowHeights(expected, 0, 0)
    )
    assertEquals(baseline(actual.toSvg(ref"A1:A1")), baseline(expected.toSvg(ref"A1:A1")))
  }

  test("formatted literals are escaped in HTML and SVG without changing rich-text fonts") {
    val style = CellStyle.default.withNumFmt(NumFmt.Custom("0;0;0;\"<&\"@"))
    val sheet = Sheet("S").put(ref"A1", CellValue.RichText(rich)).withCellStyle(ref"A1", style)
    val html = sheet.toHtml(ref"A1:C1")
    val svg = sheet.toSvg(ref"A1:C1")
    assert(html.contains("&lt;&amp;<b>a</b><i>bc</i>"), html)
    assert(svg.contains(">&lt;&amp;</tspan>"), svg)
    assert(svg.contains("font-weight=\"bold\""), svg)
  }
