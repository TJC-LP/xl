package com.tjclp.xl.ooxml

import java.io.ByteArrayOutputStream

import com.tjclp.xl.ooxml.style.{OoxmlStyles, StyleIndex, WorkbookStyles}
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.border.{Border, BorderSide, BorderStyle}
import com.tjclp.xl.styles.color.Color
import com.tjclp.xl.styles.fill.{Fill, PatternType}
import com.tjclp.xl.styles.font.Font
import com.tjclp.xl.styles.units.StyleId
import munit.FunSuite

/**
 * GH-287: the writer must emit canonical ECMA-376 camelCase enum tokens for ST_BorderStyle
 * (§18.18.3) and ST_PatternType (§18.18.55) — `dashDot`, `mediumDashDotDot`, `lightUp`, ... — not
 * `style.toString.toLowerCase` (`dashdot`, `mediumdashdotdot`, `lightup`), which is schema-invalid
 * and rejected by strict consumers.
 *
 * Asserts the exact emitted attribute string for EVERY enum case on BOTH writer paths (scala-xml
 * `toXml` and SAX `writeSax`), and that the lenient reader still accepts both casings.
 */
class StyleEnumCasingSpec extends FunSuite:

  /** Canonical ST_BorderStyle token per ECMA-376 §18.18.3 for every BorderStyle case. */
  private val borderTokens: Map[BorderStyle, String] = Map(
    BorderStyle.None -> "none",
    BorderStyle.Thin -> "thin",
    BorderStyle.Medium -> "medium",
    BorderStyle.Thick -> "thick",
    BorderStyle.Dashed -> "dashed",
    BorderStyle.Dotted -> "dotted",
    BorderStyle.Double -> "double",
    BorderStyle.Hair -> "hair",
    BorderStyle.DashDot -> "dashDot",
    BorderStyle.DashDotDot -> "dashDotDot",
    BorderStyle.SlantDashDot -> "slantDashDot",
    BorderStyle.MediumDashed -> "mediumDashed",
    BorderStyle.MediumDashDot -> "mediumDashDot",
    BorderStyle.MediumDashDotDot -> "mediumDashDotDot"
  )

  /** Canonical ST_PatternType token per ECMA-376 §18.18.55 for every PatternType case. */
  private val patternTokens: Map[PatternType, String] = Map(
    PatternType.None -> "none",
    PatternType.Solid -> "solid",
    PatternType.Gray125 -> "gray125",
    PatternType.Gray0625 -> "gray0625",
    PatternType.DarkGray -> "darkGray",
    PatternType.MediumGray -> "mediumGray",
    PatternType.LightGray -> "lightGray",
    PatternType.DarkHorizontal -> "darkHorizontal",
    PatternType.DarkVertical -> "darkVertical",
    PatternType.DarkDown -> "darkDown",
    PatternType.DarkUp -> "darkUp",
    PatternType.DarkGrid -> "darkGrid",
    PatternType.DarkTrellis -> "darkTrellis",
    PatternType.LightHorizontal -> "lightHorizontal",
    PatternType.LightVertical -> "lightVertical",
    PatternType.LightDown -> "lightDown",
    PatternType.LightUp -> "lightUp",
    PatternType.LightGrid -> "lightGrid",
    PatternType.LightTrellis -> "lightTrellis"
  )

  // Sanity: the token tables are total over the enums (a new case must get a token + this test).
  test("token tables cover every BorderStyle and PatternType case") {
    assertEquals(borderTokens.keySet, BorderStyle.values.toSet)
    assertEquals(patternTokens.keySet, PatternType.values.toSet)
  }

  /** Texture patterns: every PatternType that Fill.Pattern can carry (None/Solid have own cases) */
  private val texturePatterns: Vector[PatternType] =
    PatternType.values.toVector.filter {
      case PatternType.None | PatternType.Solid => false
      case _ => true
    }

  private def stylesWithAllEnums: OoxmlStyles =
    val borders = BorderStyle.values.toVector.map { bs =>
      Border(left = BorderSide(bs, Some(Color.Rgb(0xff000000))))
    }
    val fills: Vector[Fill] =
      Vector(Fill.None, Fill.Solid(Color.Rgb(0xffff0000))) ++
        texturePatterns.map(p => Fill.Pattern(Color.Rgb(0xff000000), Color.Rgb(0xffffffff), p))
    val index = StyleIndex(
      fonts = Vector(Font.default),
      fills = fills,
      borders = borders,
      numFmts = Vector.empty,
      cellStyles = Vector(CellStyle.default),
      styleToIndex = Map(CellStyle.default.canonicalKey -> StyleId(0))
    )
    OoxmlStyles(index)

  private def scalaXmlOutput: String = XmlUtil.compact(stylesWithAllEnums.toXml)

  private def saxOutput: String =
    val baos = new ByteArrayOutputStream()
    val writer = StaxSaxWriter.create(baos)
    stylesWithAllEnums.writeSax(writer)
    baos.toString("UTF-8")

  private def assertEmitsCanonicalTokens(xml: String, path: String): Unit =
    borderTokens.foreach { case (style, token) =>
      if style != BorderStyle.None then // None emits a bare side element with no style attribute
        assert(
          xml.contains(s"""style="$token""""),
          s"$path: expected border style=\"$token\" for $style. Got:\n$xml"
        )
    }
    patternTokens.foreach { case (pattern, token) =>
      assert(
        xml.contains(s"""patternType="$token""""),
        s"$path: expected patternType=\"$token\" for $pattern. Got:\n$xml"
      )
    }
    // No schema-invalid all-lowercase multi-word tokens anywhere
    val invalid = List(
      "dashdot",
      "dashdotdot",
      "slantdashdot",
      "mediumdashed",
      "mediumdashdot",
      "mediumdashdotdot",
      "darkgray",
      "mediumgray",
      "lightgray",
      "darkhorizontal",
      "darkvertical",
      "darkdown",
      "darkup",
      "darkgrid",
      "darktrellis",
      "lighthorizontal",
      "lightvertical",
      "lightdown",
      "lightup",
      "lightgrid",
      "lighttrellis"
    )
    invalid.foreach { token =>
      assert(
        !xml.contains(s"\"$token\""),
        s"$path: schema-invalid lowercase token \"$token\" still emitted. Got:\n$xml"
      )
    }

  test("GH-287: scala-xml writer emits canonical camelCase border/fill tokens") {
    assertEmitsCanonicalTokens(scalaXmlOutput, "toXml")
  }

  test("GH-287: SAX writer emits canonical camelCase border/fill tokens") {
    assertEmitsCanonicalTokens(saxOutput, "writeSax")
  }

  test("GH-287: lenient reader accepts BOTH camelCase and legacy lowercase tokens") {
    def stylesXml(borderToken: String, patternToken: String) =
      s"""<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
         |  <fonts count="1"><font><name val="Calibri"/><sz val="11.0"/></font></fonts>
         |  <fills count="1"><fill><patternFill patternType="$patternToken"><fgColor rgb="FF000000"/><bgColor rgb="FFFFFFFF"/></patternFill></fill></fills>
         |  <borders count="1"><border><left style="$borderToken"><color rgb="FF000000"/></left><right/><top/><bottom/></border></borders>
         |  <cellXfs count="1"><xf borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0"/></cellXfs>
         |</styleSheet>""".stripMargin

    List("mediumDashDotDot" -> "lightUp", "mediumdashdotdot" -> "lightup").foreach {
      case (borderToken, patternToken) =>
        val elem = XmlSecurity
          .parseSafe(stylesXml(borderToken, patternToken), "styles.xml")
          .fold(err => fail(s"XML parse failed for ($borderToken, $patternToken): $err"), identity)
        val parsed = WorkbookStyles
          .fromXml(elem)
          .fold(
            msg => fail(s"styles parse failed for ($borderToken, $patternToken): $msg"),
            identity
          )
        assertEquals(
          parsed.borders(0).left.style,
          BorderStyle.MediumDashDotDot,
          s"reader should accept border token '$borderToken'"
        )
        assertEquals(
          parsed.fills(0),
          Fill.Pattern(Color.Rgb(0xff000000), Color.Rgb(0xffffffff), PatternType.LightUp),
          s"reader should accept pattern token '$patternToken'"
        )
    }
  }
