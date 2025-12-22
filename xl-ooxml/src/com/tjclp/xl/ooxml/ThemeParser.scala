package com.tjclp.xl.ooxml

import scala.xml.*
import com.tjclp.xl.styles.color.ThemePalette

/**
 * Parser for OOXML theme (xl/theme/theme1.xml).
 *
 * Extracts color scheme (clrScheme) to build ThemePalette for resolving theme colors. Theme colors
 * can be specified as:
 *   - srgbClr: Direct RGB value (val="RRGGBB")
 *   - sysClr: System color with lastClr fallback (val="windowText" lastClr="000000")
 *
 * OOXML Structure:
 * {{{
 * <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
 *   <a:themeElements>
 *     <a:clrScheme>
 *       <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
 *       <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
 *       <a:dk2><a:srgbClr val="1F497D"/></a:dk2>
 *       ...
 *     </a:clrScheme>
 *   </a:themeElements>
 * </a:theme>
 * }}}
 */
object ThemeParser:

  /** DrawingML namespace for theme elements */
  private val nsDrawingML = "http://schemas.openxmlformats.org/drawingml/2006/main"

  /**
   * Parse theme XML and extract color palette.
   *
   * REQUIRES: themeXml is valid OOXML theme XML (xl/theme/theme1.xml) ENSURES:
   *   - Returns ThemePalette with all 12 theme colors
   *   - Falls back to Office theme defaults for missing colors
   *   - Uses lastClr attribute for system colors DETERMINISTIC: Yes
   *
   * @param themeXml
   *   XML string from xl/theme/theme1.xml
   * @return
   *   Either[String, ThemePalette] with error message or parsed palette
   */
  def parse(themeXml: String): Either[String, ThemePalette] =
    XmlSecurity.parseSafe(themeXml, "xl/theme/theme1.xml") match
      case Right(elem) => parseThemeElement(elem)
      case Left(err) => Left(s"Failed to parse theme XML: ${err.message}")

  /**
   * Parse theme from XML element.
   *
   * @param elem
   *   Root <theme> or <a:theme> element
   * @return
   *   Either[String, ThemePalette]
   */
  def parseThemeElement(elem: Elem): Either[String, ThemePalette] =
    // Find clrScheme element (may be prefixed with 'a:' or unprefixed)
    val clrScheme = findClrScheme(elem)

    clrScheme match
      case None => Left("Missing <clrScheme> element in theme")
      case Some(cs) =>
        // Extract colors, falling back to Office defaults if missing
        val dark1 = extractColor(cs, "dk1").getOrElse(0xff000000)
        val light1 = extractColor(cs, "lt1").getOrElse(0xffffffff)
        val dark2 = extractColor(cs, "dk2").getOrElse(0xff44546a)
        val light2 = extractColor(cs, "lt2").getOrElse(0xffe7e6e6)
        val accent1 = extractColor(cs, "accent1").getOrElse(0xff4472c4)
        val accent2 = extractColor(cs, "accent2").getOrElse(0xffed7d31)
        val accent3 = extractColor(cs, "accent3").getOrElse(0xffa5a5a5)
        val accent4 = extractColor(cs, "accent4").getOrElse(0xffffc000)
        val accent5 = extractColor(cs, "accent5").getOrElse(0xff5b9bd5)
        val accent6 = extractColor(cs, "accent6").getOrElse(0xff70ad47)
        val hyperlink = extractColor(cs, "hlink").getOrElse(0xff0563c1)
        val followedHyperlink = extractColor(cs, "folHlink").getOrElse(0xff954f72)

        Right(
          ThemePalette(
            dark1 = dark1,
            light1 = light1,
            dark2 = dark2,
            light2 = light2,
            accent1 = accent1,
            accent2 = accent2,
            accent3 = accent3,
            accent4 = accent4,
            accent5 = accent5,
            accent6 = accent6,
            hyperlink = hyperlink,
            followedHyperlink = followedHyperlink
          )
        )

  /**
   * Find clrScheme element in theme structure.
   *
   * Handles both prefixed (a:clrScheme) and unprefixed (clrScheme) element names, searching through
   * themeElements container.
   */
  private def findClrScheme(elem: Elem): Option[Elem] =
    // Search for themeElements first
    val themeElements = elem.child.collectFirst {
      case e: Elem if e.label == "themeElements" => e
    }

    // Then find clrScheme within themeElements (or directly in elem if no themeElements)
    val searchRoot = themeElements.getOrElse(elem)
    searchRoot.child.collectFirst { case e: Elem if e.label == "clrScheme" => e }

  /**
   * Extract a color from clrScheme by element name.
   *
   * Handles both srgbClr (direct RGB) and sysClr (system color with lastClr fallback).
   *
   * @param clrScheme
   *   The clrScheme element
   * @param colorName
   *   The color element name (e.g., "dk1", "accent1")
   * @return
   *   Optional ARGB color value (0xAARRGGBB format)
   */
  private def extractColor(clrScheme: Elem, colorName: String): Option[Int] =
    // Find the color container element (e.g., <a:dk1> or <dk1>)
    val colorContainer = clrScheme.child.collectFirst {
      case e: Elem if e.label == colorName => e
    }

    colorContainer.flatMap { container =>
      // Look for srgbClr or sysClr child
      container.child.collectFirst {
        case e: Elem if e.label == "srgbClr" =>
          parseRgbColor(e)
        case e: Elem if e.label == "sysClr" =>
          parseSysColor(e)
      }.flatten
    }

  /**
   * Parse srgbClr element.
   *
   * Format: <a:srgbClr val="4472C4"/>
   *
   * @return
   *   ARGB value with alpha=0xFF
   */
  private def parseRgbColor(elem: Elem): Option[Int] =
    elem.attributes.asAttrMap.get("val").flatMap { hex =>
      try
        val rgb = Integer.parseInt(hex, 16)
        Some(0xff000000 | rgb) // Add full alpha
      catch case _: NumberFormatException => None
    }

  /**
   * Parse sysClr element using lastClr attribute.
   *
   * Format: <a:sysClr val="windowText" lastClr="000000"/>
   *
   * The lastClr attribute contains the actual color value to use (snapshot of system color at file
   * creation time).
   *
   * @return
   *   ARGB value with alpha=0xFF
   */
  private def parseSysColor(elem: Elem): Option[Int] =
    val attrs = elem.attributes.asAttrMap
    // Prefer lastClr (actual color), fall back to mapping val to a default
    attrs
      .get("lastClr")
      .flatMap { hex =>
        try
          val rgb = Integer.parseInt(hex, 16)
          Some(0xff000000 | rgb)
        catch case _: NumberFormatException => None
      }
      .orElse {
        // Fallback: map system color names to defaults
        attrs.get("val").flatMap {
          case "windowText" => Some(0xff000000) // Black
          case "window" => Some(0xffffffff) // White
          case _ => None
        }
      }
