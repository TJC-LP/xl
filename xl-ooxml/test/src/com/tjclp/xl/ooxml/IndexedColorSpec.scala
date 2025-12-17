package com.tjclp.xl.ooxml

import munit.FunSuite

import scala.xml.Elem

import com.tjclp.xl.styles.color.Color

/**
 * Regression tests for indexed color support (#84).
 *
 * Excel files can use indexed colors (0-63 standard palette). This test ensures we parse them
 * correctly into RGB Color values.
 */
class IndexedColorSpec extends FunSuite:

  // Parse a color element via WorkbookStyles by embedding it in a font element
  private def parseColorFromXml(colorXml: Elem): Option[Color] =
    // Create a minimal styles.xml structure and parse it
    val stylesXml = <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <fonts count="1">
        <font>
          <sz val="11"/>
          <name val="Calibri"/>
          {colorXml}
        </font>
      </fonts>
      <fills count="1">
        <fill><patternFill patternType="none"/></fill>
      </fills>
      <borders count="1">
        <border><left/><right/><top/><bottom/><diagonal/></border>
      </borders>
      <cellXfs count="1">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
      </cellXfs>
    </styleSheet>

    WorkbookStyles
      .fromXml(stylesXml)
      .toOption
      .flatMap(_.cellStyles.headOption)
      .flatMap(_.font.color)

  test("indexed color 0 (black) parses to RGB 0x000000") {
    val xml = <color indexed="0"/>
    val color = parseColorFromXml(xml)
    color match
      case Some(Color.Rgb(argb)) =>
        val rgb = argb & 0xffffff
        assertEquals(rgb, 0x000000, "Index 0 should be black")
      case other =>
        fail(s"Expected Rgb color, got: $other")
  }

  test("indexed color 1 (white) parses to RGB 0xFFFFFF") {
    val xml = <color indexed="1"/>
    val color = parseColorFromXml(xml)
    color match
      case Some(Color.Rgb(argb)) =>
        val rgb = argb & 0xffffff
        assertEquals(rgb, 0xffffff, "Index 1 should be white")
      case other =>
        fail(s"Expected Rgb color, got: $other")
  }

  test("indexed color 2 (red) parses to RGB 0xFF0000") {
    val xml = <color indexed="2"/>
    val color = parseColorFromXml(xml)
    color match
      case Some(Color.Rgb(argb)) =>
        val rgb = argb & 0xffffff
        assertEquals(rgb, 0xff0000, "Index 2 should be red")
      case other =>
        fail(s"Expected Rgb color, got: $other")
  }

  test("indexed color 18 (navy) parses to RGB 0x000080") {
    val xml = <color indexed="18"/>
    val color = parseColorFromXml(xml)
    color match
      case Some(Color.Rgb(argb)) =>
        val rgb = argb & 0xffffff
        assertEquals(rgb, 0x000080, "Index 18 should be navy")
      case other =>
        fail(s"Expected Rgb color, got: $other")
  }

  test("indexed color 63 (gray 80%) parses to RGB 0x333333") {
    val xml = <color indexed="63"/>
    val color = parseColorFromXml(xml)
    color match
      case Some(Color.Rgb(argb)) =>
        val rgb = argb & 0xffffff
        assertEquals(rgb, 0x333333, "Index 63 should be gray 80%")
      case other =>
        fail(s"Expected Rgb color, got: $other")
  }

  test("indexed color 64+ returns None (custom colors not supported)") {
    val xml = <color indexed="64"/>
    val color = parseColorFromXml(xml)
    assertEquals(color, None, "Index 64+ should return None")
  }

  test("RGB color takes precedence over indexed") {
    // When both rgb and indexed are present, rgb should win
    val xml = <color rgb="FFFF0000" indexed="18"/>
    val color = parseColorFromXml(xml)
    color match
      case Some(Color.Rgb(argb)) =>
        val rgb = argb & 0xffffff
        assertEquals(rgb, 0xff0000, "RGB should take precedence")
      case other =>
        fail(s"Expected Rgb color, got: $other")
  }

  test("theme color takes precedence over indexed") {
    // When both theme and indexed are present, theme should win
    val xml = <color theme="1" indexed="18"/>
    val color = parseColorFromXml(xml)
    color match
      case Some(Color.Theme(_, _)) =>
        () // Theme color parsed correctly
      case other =>
        fail(s"Expected Theme color, got: $other")
  }
