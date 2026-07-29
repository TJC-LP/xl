package com.tjclp.xl.io.streaming

import munit.FunSuite
import com.tjclp.xl.ooxml.XmlSecurity
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.numfmt.NumFmt

/**
 * StylePatcher numFmt-id parity with the canonical NumFmt table (GH-408).
 *
 * The streaming style path (addStyle) and the in-memory writer (StyleSerializer via
 * NumFmt.builtInId) must emit the same numFmtId for the same NumFmt value, and getStyle must
 * resolve ids the way the DOM StyleParser does (declared entries win verbatim, then
 * NumFmt.fromId). A hand-rolled id table here made Decimal/Percent/Currency render differently
 * between the two write paths.
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
