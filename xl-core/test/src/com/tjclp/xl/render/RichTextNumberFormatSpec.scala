package com.tjclp.xl.render

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.display.NumFmtFormatter
import com.tjclp.xl.richtext.{RichText, TextRun}
import com.tjclp.xl.styles.{CellStyle}
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
