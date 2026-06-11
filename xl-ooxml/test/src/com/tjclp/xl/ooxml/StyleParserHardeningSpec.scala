package com.tjclp.xl.ooxml

import scala.xml.XML

import munit.FunSuite

import com.tjclp.xl.ooxml.style.WorkbookStyles
import com.tjclp.xl.styles.font.Font

/**
 * GH-278: the DOM StyleParser must stay total on malformed numeric attributes.
 *
 * `Font` carries domain guards (`sizePt > 0`, nonEmpty name); the parser must filter values that
 * would violate them and fall back to defaults instead of throwing through `require` during read.
 * This is the DOM-side analog of the GH-264 streaming fix (StylePatcher).
 */
class StyleParserHardeningSpec extends FunSuite:

  private def stylesXml(fontXml: String): scala.xml.Elem =
    XML.loadString(
      s"""<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
         |  <fonts count="2">
         |    <font><sz val="11"/><name val="Calibri"/></font>
         |    $fontXml
         |  </fonts>
         |  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
         |  <borders count="1"><border/></borders>
         |  <cellXfs count="2">
         |    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
         |    <xf numFmtId="0" fontId="1" fillId="0" borderId="0"/>
         |  </cellXfs>
         |</styleSheet>""".stripMargin
    )

  private def secondFont(fontXml: String): Font =
    WorkbookStyles.fromXml(stylesXml(fontXml)) match
      case Right(styles) =>
        styles.fonts.lift(1).getOrElse(fail("expected two parsed fonts"))
      case Left(err) => fail(s"styles parse must not fail: $err")

  test("GH-278: negative font sz falls back to default size instead of throwing") {
    val font = secondFont("""<font><sz val="-4"/><name val="Arial"/></font>""")
    assertEquals(font.sizePt, Font.default.sizePt)
    assertEquals(font.name, "Arial", "fallback must only affect the size")
  }

  test("GH-278: zero font sz falls back to default size instead of throwing") {
    val font = secondFont("""<font><sz val="0"/><name val="Arial"/></font>""")
    assertEquals(font.sizePt, Font.default.sizePt)
  }

  test("GH-278: NaN font sz falls back to default size instead of throwing") {
    // "NaN".toDoubleOption parses; require(NaN > 0) would still throw
    val font = secondFont("""<font><sz val="NaN"/><name val="Arial"/></font>""")
    assertEquals(font.sizePt, Font.default.sizePt)
  }

  test("GH-278: unparseable font sz keeps default size (regression)") {
    val font = secondFont("""<font><sz val="big"/><name val="Arial"/></font>""")
    assertEquals(font.sizePt, Font.default.sizePt)
  }

  test("GH-278: valid font sz still parses exactly (regression)") {
    val font = secondFont("""<font><sz val="14.5"/><name val="Arial"/></font>""")
    assertEquals(font.sizePt, 14.5)
  }

  // ---- sweep: the same hole in rich-text run properties (<rPr>), reachable from
  // ---- both the DOM reader (parseTextRuns) and the SAX SST reader

  test("GH-278 sweep: rich-run <rPr> non-positive sz falls back to default size") {
    val rPr = XML.loadString("""<rPr><sz val="-2"/><name val="Arial"/></rPr>""")
    val font = XmlUtil.parseRunProperties(rPr)
    assertEquals(font.sizePt, Font.default.sizePt)
    assertEquals(font.name, "Arial")
  }

  test("GH-278 sweep: rich-run <rPr> NaN sz falls back to default size") {
    val rPr = XML.loadString("""<rPr><sz val="NaN"/></rPr>""")
    assertEquals(XmlUtil.parseRunProperties(rPr).sizePt, Font.default.sizePt)
  }

  test("GH-278 sweep: rich-run <rPr> empty name falls back to default name") {
    val rPr = XML.loadString("""<rPr><name val=""/></rPr>""")
    assertEquals(XmlUtil.parseRunProperties(rPr).name, Font.default.name)
  }

  test("GH-278 sweep: rich-run <rPr> valid sz and name still parse (regression)") {
    val rPr = XML.loadString("""<rPr><sz val="9.5"/><name val="Consolas"/></rPr>""")
    val font = XmlUtil.parseRunProperties(rPr)
    assertEquals(font.sizePt, 9.5)
    assertEquals(font.name, "Consolas")
  }
