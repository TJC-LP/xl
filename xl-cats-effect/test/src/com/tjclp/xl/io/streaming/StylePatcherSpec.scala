package com.tjclp.xl.io.streaming

import java.io.ByteArrayInputStream
import java.nio.charset.StandardCharsets
import java.util.zip.ZipInputStream

import munit.FunSuite
import com.tjclp.xl.api.*
import com.tjclp.xl.sheets.syntax.withCellStyle
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.{XlsxWriter, XmlSecurity}
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.border.{Border, BorderSide, BorderStyle}
import com.tjclp.xl.styles.color.{Color, ThemeSlot}
import com.tjclp.xl.styles.fill.{Fill, PatternType}
import com.tjclp.xl.styles.font.Font
import com.tjclp.xl.styles.numfmt.NumFmt

/**
 * StylePatcher numFmt-id parity with the canonical NumFmt table (GH-408).
 *
 * The streaming style path (addStyle) and the in-memory writer (StyleSerializer via
 * NumFmt.builtInId) must emit the same numFmtId for the same NumFmt value, and getStyle must
 * resolve ids the way the DOM StyleParser does (declared entries win verbatim, then NumFmt.fromId).
 * A hand-rolled id table here made Decimal/Percent/Currency render differently between the two
 * write paths.
 */
class StylePatcherSpec extends FunSuite:

  private val minimalStylesXml: String =
    """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      |<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      |<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
      |<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
      |<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
      |<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
      |<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
      |</styleSheet>""".stripMargin.replaceAll("\n", "")

  /** Every built-in NumFmt variant with its canonical OOXML id (ECMA-376 Part 1, §18.8.30). */
  private val builtInIds: List[(NumFmt, Int)] = List(
    NumFmt.General -> 0,
    NumFmt.Integer -> 1,
    NumFmt.Decimal -> 2,
    NumFmt.ThousandsSeparator -> 3,
    NumFmt.ThousandsDecimal -> 4,
    NumFmt.Currency -> 7,
    NumFmt.Percent -> 9,
    NumFmt.PercentDecimal -> 10,
    NumFmt.Scientific -> 11,
    NumFmt.Fraction -> 12,
    NumFmt.Date -> 14,
    NumFmt.Time -> 21,
    NumFmt.DateTime -> 22,
    NumFmt.Text -> 49
  )

  /** Patch a style with the given numFmt into styles.xml, returning (updated xml, new xf id). */
  private def patch(stylesXml: String, fmt: NumFmt): (String, Int) =
    StylePatcher.addStyle(stylesXml, CellStyle.default.withNumFmt(fmt)) match
      case Right(result) => result
      case Left(e) => fail(s"addStyle failed for $fmt: ${e.message}")

  /** The numFmtId attribute of cellXfs/xf at the given index. */
  private def numFmtIdOf(stylesXml: String, xfId: Int): Int =
    XmlSecurity.parseSafe(stylesXml, "styles.xml") match
      case Right(root) =>
        (root \ "cellXfs" \ "xf").lift(xfId) match
          case Some(xf) =>
            (xf \ "@numFmtId").text.toIntOption.getOrElse(fail(s"xf $xfId has no numFmtId"))
          case None => fail(s"no cellXfs/xf at index $xfId")
      case Left(e) => fail(s"parse failed: ${e.message}")

  test("addStyle: every builtin numFmt lands its canonical ECMA-376 id (GH-408)") {
    builtInIds.foreach { case (fmt, expectedId) =>
      val (updated, xfId) = patch(minimalStylesXml, fmt)
      assertEquals(numFmtIdOf(updated, xfId), expectedId, s"wrong numFmtId for $fmt")
    }
  }

  test("addStyle: emitted ids delegate to NumFmt.builtInId (one source of truth, GH-408)") {
    builtInIds.map(_._1).foreach { fmt =>
      val (updated, xfId) = patch(minimalStylesXml, fmt)
      assertEquals(
        Some(numFmtIdOf(updated, xfId)),
        NumFmt.builtInId(fmt),
        s"StylePatcher diverged from NumFmt.builtInId for $fmt"
      )
    }
  }

  test("addStyle then getStyle round-trips every builtin numFmt (GH-408)") {
    builtInIds.map(_._1).foreach { fmt =>
      val (updated, xfId) = patch(minimalStylesXml, fmt)
      StylePatcher.getStyle(updated, xfId) match
        case Right(Some(style)) => assertEquals(style.numFmt, fmt, s"round-trip lost $fmt")
        case Right(None) => fail(s"style $xfId not found after addStyle($fmt)")
        case Left(e) => fail(s"getStyle failed for $fmt: ${e.message}")
    }
  }

  test("addStyle then getStyle round-trips a custom format code (GH-408)") {
    val custom = NumFmt.Custom("0.000")
    val (updated, xfId) = patch(minimalStylesXml, custom)
    assertEquals(numFmtIdOf(updated, xfId), NumFmt.FirstCustomId, "custom ids start at 164")
    StylePatcher.getStyle(updated, xfId) match
      case Right(Some(style)) => assertEquals(style.numFmt, custom)
      case other => fail(s"expected round-tripped custom style, got $other")
  }

  test("addStyle: identical custom codes reuse one declared numFmt id (GH-408)") {
    val custom = NumFmt.Custom("0.000")
    val (afterFirst, _) = patch(minimalStylesXml, custom)
    val (afterSecond, secondXfId) = patch(afterFirst, custom)
    assertEquals(numFmtIdOf(afterSecond, secondXfId), NumFmt.FirstCustomId)
    XmlSecurity.parseSafe(afterSecond, "styles.xml") match
      case Right(root) =>
        assertEquals((root \ "numFmts" \ "numFmt").size, 1, "same code must not re-declare")
      case Left(e) => fail(s"parse failed: ${e.message}")
  }

  test("getStyle: declared numFmt entries win over the builtin table (DOM parity, GH-408)") {
    // The DOM StyleParser resolves numFmts.get(id).orElse(NumFmt.fromId(id)): a declared
    // <numFmt> keeps its code VERBATIM (GH-404) even when its id shadows a builtin slot.
    val declared =
      """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        |<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        |<numFmts count="1"><numFmt numFmtId="10" formatCode="0.0%"/></numFmts>
        |<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
        |<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
        |<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
        |<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
        |<cellXfs count="1"><xf numFmtId="10" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>
        |</styleSheet>""".stripMargin.replaceAll("\n", "")
    StylePatcher.getStyle(declared, 0) match
      case Right(Some(style)) => assertEquals(style.numFmt, NumFmt.Custom("0.0%"))
      case other => fail(s"expected declared code to win, got $other")
  }

  // ===== GH-566: texture fills through the streaming style path (`--stream style`) =====

  /**
   * openpyxl's fgColor-only `mediumGray` (fill 2, xf 1) and a two-colour `lightUp` hatch (fill 3,
   * xf 2) — the fills the streaming codec used to read back as solid/none and re-emit lowercase.
   */
  private val textureStylesXml: String =
    """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      |<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      |<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
      |<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="mediumGray"><fgColor rgb="FF808080"/></patternFill></fill><fill><patternFill patternType="lightUp"><fgColor rgb="FF0000FF"/><bgColor rgb="FFFFFF00"/></patternFill></fill></fills>
      |<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
      |<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
      |<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/><xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/></cellXfs>
      |</styleSheet>""".stripMargin.replaceAll("\n", "")

  private def styleOf(stylesXml: String, xfId: Int): CellStyle =
    StylePatcher.getStyle(stylesXml, xfId) match
      case Right(Some(style)) => style
      case other => fail(s"getStyle($xfId) gave $other")

  /** The `<patternFill>` of the fill the given xf points at. */
  private def patternFillOf(stylesXml: String, xfId: Int): scala.xml.Node =
    XmlSecurity.parseSafe(stylesXml, "styles.xml") match
      case Right(root) =>
        val xf = (root \ "cellXfs" \ "xf").lift(xfId).getOrElse(fail(s"no xf $xfId"))
        val fillId = (xf \ "@fillId").text.toIntOption.getOrElse(fail(s"xf $xfId has no fillId"))
        val fill = (root \ "fills" \ "fill").lift(fillId).getOrElse(fail(s"no fill $fillId"))
        (fill \ "patternFill").headOption.getOrElse(fail(s"fill $fillId has no patternFill"))
      case Left(e) => fail(s"parse failed: ${e.message}")

  test("GH-566: getStyle keeps a fgColor-only mediumGray texture as a pattern fill") {
    styleOf(textureStylesXml, 1).fill match
      case Fill.Pattern(_, _, PatternType.MediumGray) => ()
      case other => fail(s"streaming reader corrupted the texture: $other")
    styleOf(textureStylesXml, 2).fill match
      case Fill.Pattern(_, _, PatternType.LightUp) => ()
      case other => fail(s"streaming reader corrupted the two-colour texture: $other")
  }

  test("GH-566: a streamed bold overlay on textured cells keeps camelCase tokens and colours") {
    val bold = CellStyle.default.withFont(Font("Calibri", 11.0, bold = true))
    val mergedGray = StylePatcher.mergeStyles(styleOf(textureStylesXml, 1), bold)
    val (afterGray, grayXf) = StylePatcher.addStyle(textureStylesXml, mergedGray) match
      case Right(r) => r
      case Left(e) => fail(s"addStyle failed: ${e.message}")
    val gray = patternFillOf(afterGray, grayXf)
    assertEquals((gray \ "@patternType").text, "mediumGray", afterGray)
    assertEquals((gray \ "fgColor" \ "@rgb").text, "FF808080", afterGray)
    assert((gray \ "bgColor").isEmpty, s"an absent bgColor must stay absent: $afterGray")
    assert(!afterGray.contains("mediumgray"), s"schema-invalid lowercase token: $afterGray")

    val mergedUp = StylePatcher.mergeStyles(styleOf(afterGray, 2), bold)
    val (afterUp, upXf) = StylePatcher.addStyle(afterGray, mergedUp) match
      case Right(r) => r
      case Left(e) => fail(s"addStyle failed: ${e.message}")
    val up = patternFillOf(afterUp, upXf)
    assertEquals((up \ "@patternType").text, "lightUp", afterUp)
    assertEquals((up \ "fgColor" \ "@rgb").text, "FF0000FF", afterUp)
    assertEquals((up \ "bgColor" \ "@rgb").text, "FFFFFF00", afterUp)
    assert(!afterUp.contains("lightup"), s"schema-invalid lowercase token: $afterUp")
    // the overlay itself landed
    assert(styleOf(afterUp, upXf).font.bold)
  }

  // ===== Theme colours: the OOXML theme index, agreeing with the DOM writer for every slot =====

  /**
   * ECMA-376 §18.8.3 theme indices: 0=lt1, 1=dk1, 2=lt2, 3=dk2, 4..9=accent1..6 — the FIRST four
   * are the inverse of the ThemeSlot declaration order (Dark1, Light1, Dark2, Light2), so writing
   * `slot.ordinal` swapped Dark1/Light1 and Dark2/Light2 on the streaming path (`--stream style`)
   * while the DOM/SAX writers were right: a requested Dark1 fill rendered as Light1 (white).
   */
  private val ecmaThemeIndex: Map[ThemeSlot, Int] = Map(
    ThemeSlot.Light1 -> 0,
    ThemeSlot.Dark1 -> 1,
    ThemeSlot.Light2 -> 2,
    ThemeSlot.Dark2 -> 3,
    ThemeSlot.Accent1 -> 4,
    ThemeSlot.Accent2 -> 5,
    ThemeSlot.Accent3 -> 6,
    ThemeSlot.Accent4 -> 7,
    ThemeSlot.Accent5 -> 8,
    ThemeSlot.Accent6 -> 9
  )

  /** A style using `color` in the font, a solid fill, a textured fill's background and a border. */
  private def themedStyle(color: Color): CellStyle =
    CellStyle.default
      .withFont(Font("Calibri", 11.0, color = Some(color)))
      .withFill(Fill.Solid(color))
      .withBorder(Border(left = BorderSide(BorderStyle.Thin, Some(color))))

  private def parse(stylesXml: String): scala.xml.Elem =
    XmlSecurity.parseSafe(stylesXml, "styles.xml") match
      case Right(root) => root
      case Left(e) => fail(s"parse failed: ${e.message}")

  /** Every colour element of `label` under `section` that names a theme, as its attribute map. */
  private def themeColours(
    root: scala.xml.Elem,
    section: String,
    label: String
  ): Set[Map[String, String]] =
    (root \ section \\ label)
      .filter(c => (c \ "@theme").nonEmpty)
      .map(_.attributes.asAttrMap)
      .toSet

  /** The DOM writer's styles.xml for one cell carrying `style`. */
  private def domStylesXml(style: CellStyle): scala.xml.Elem =
    val wb = Workbook(Vector(Sheet("S").put(ref"A1", "x").withCellStyle(ref"A1", style)))
    val bytes = XlsxWriter.writeToBytes(wb).fold(e => fail(e.message), identity)
    val zin = new ZipInputStream(new ByteArrayInputStream(bytes))
    try
      val xml = Iterator
        .continually(Option(zin.getNextEntry))
        .takeWhile(_.isDefined)
        .flatten
        .collectFirst {
          case e if e.getName == "xl/styles.xml" =>
            new String(zin.readAllBytes(), StandardCharsets.UTF_8)
        }
        .getOrElse(fail("no styles.xml"))
      parse(xml)
    finally zin.close()

  test("theme colours stream by OOXML theme index (Dark1 = theme=\"1\"), never by enum ordinal") {
    ThemeSlot.values.foreach { slot =>
      val (updated, _) =
        StylePatcher.addStyle(minimalStylesXml, themedStyle(Color.Theme(slot, 0.0))) match
          case Right(r) => r
          case Left(e) => fail(s"addStyle failed for $slot: ${e.message}")
      val root = parse(updated)
      val expected = Set(Map("theme" -> ecmaThemeIndex(slot).toString))
      assertEquals(themeColours(root, "fills", "fgColor"), expected, s"$slot fill: $updated")
      assertEquals(themeColours(root, "fonts", "color"), expected, s"$slot font: $updated")
      assertEquals(themeColours(root, "borders", "color"), expected, s"$slot border: $updated")
      assert(
        !updated.contains("tint=\"0.0\""),
        s"a zero tint is omitted, as the DOM writer does: $updated"
      )
    }
  }

  test("theme colours: the streaming and DOM writers spell the same CellStyle identically") {
    ThemeSlot.values.foreach { slot =>
      val style = themedStyle(Color.Theme(slot, 0.0))
      val streamed = StylePatcher.addStyle(minimalStylesXml, style) match
        case Right((xml, _)) => parse(xml)
        case Left(e) => fail(s"addStyle failed for $slot: ${e.message}")
      val dom = domStylesXml(style)
      Seq(("fills", "fgColor"), ("fonts", "color"), ("borders", "color")).foreach {
        case (section, label) =>
          assertEquals(
            themeColours(streamed, section, label),
            themeColours(dom, section, label),
            s"$slot: $section/$label differ between --stream and the DOM writer"
          )
      }
    }
  }

  test("theme colours: a tint streams in Excel's form and the two writers agree") {
    val style = themedStyle(Color.Theme(ThemeSlot.Accent1, 0.7999816888943144))
    val streamed = StylePatcher.addStyle(minimalStylesXml, style) match
      case Right((xml, _)) => parse(xml)
      case Left(e) => fail(s"addStyle failed: ${e.message}")
    assertEquals(
      themeColours(streamed, "fills", "fgColor"),
      Set(Map("theme" -> "4", "tint" -> "0.79998168889431442"))
    )
    assertEquals(
      themeColours(streamed, "fills", "fgColor"),
      themeColours(domStylesXml(style), "fills", "fgColor")
    )
  }

  test(
    "theme colours: getStyle reads <fgColor theme=\"1\"/> back as Dark1 (the DOM parser's slot)"
  ) {
    ThemeSlot.values.foreach { slot =>
      val style = themedStyle(Color.Theme(slot, 0.0))
      val (updated, xfId) = StylePatcher.addStyle(minimalStylesXml, style) match
        case Right(r) => r
        case Left(e) => fail(s"addStyle failed for $slot: ${e.message}")
      assertEquals(styleOf(updated, xfId).fill, style.fill, s"$slot did not round-trip")
      assertEquals(styleOf(updated, xfId).font.color, Some(Color.Theme(slot, 0.0)))
    }
    // and an Excel-authored index resolves through the same table: theme="1" is dk1
    val excelDark1 = minimalStylesXml
      .replace(
        """<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>""",
        """<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor theme="1"/></patternFill></fill></fills>"""
      )
      .replace(
        """<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>""",
        """<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/></cellXfs>"""
      )
    assertEquals(styleOf(excelDark1, 1).fill, Fill.Solid(Color.Theme(ThemeSlot.Dark1, 0.0)))
  }
