package com.tjclp.xl.io.streaming

import munit.FunSuite
import com.tjclp.xl.ooxml.XmlSecurity
import com.tjclp.xl.styles.CellStyle
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
