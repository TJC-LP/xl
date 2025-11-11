package com.tjclp.xl.ooxml

import munit.FunSuite
import com.tjclp.xl.style.{Fill, Color, PatternType, CellStyle}
import scala.xml.Elem

/** Tests for OOXML Styles serialization (xl/styles.xml) */
class OoxmlStylesSpec extends FunSuite:

  test("defaultFills has None at index 0") {
    // Use empty StyleIndex to test defaultFills behavior
    val styles = OoxmlStyles(StyleIndex.empty)
    val xml = styles.toXml

    val fills = xml \ "fills" \ "fill"
    assert(fills.nonEmpty, "Should have fills in styles.xml")

    // First fill should be patternFill with patternType="none"
    val firstFill = fills(0)
    val patternFill = firstFill \ "patternFill"
    assert(patternFill.nonEmpty, "First fill should have patternFill element")

    val patternType = (patternFill \ "@patternType").text
    assertEquals(patternType, "none", "First default fill should have patternType='none'")
  }

  test("defaultFills has Gray125 at index 1") {
    val styles = OoxmlStyles(StyleIndex.empty)
    val xml = styles.toXml

    val fills = xml \ "fills" \ "fill"
    assert(fills.size >= 2, "Should have at least 2 default fills")

    // Second fill should be patternFill with patternType="gray125"
    val secondFill = fills(1)
    val patternFill = secondFill \ "patternFill"
    assert(patternFill.nonEmpty, "Second fill should have patternFill element")

    val patternType = (patternFill \ "@patternType").text
    assertEquals(patternType, "gray125", "Second default fill should have patternType='gray125'")
  }

  test("toXml emits gray125 pattern in fills section") {
    val styles = OoxmlStyles(StyleIndex.empty)
    val xml = styles.toXml

    // Verify fills section exists with correct count
    val fillsElem = xml \ "fills"
    assert(fillsElem.nonEmpty, "Should have fills element")

    val count = (fillsElem \ "@count").text.toInt
    assert(count >= 2, "Should have at least 2 fills (default fills)")

    // Verify gray125 pattern has foreground and background colors
    val fills = xml \ "fills" \ "fill"
    val gray125Fill = fills(1)
    val patternFill = gray125Fill \ "patternFill"

    // Check foreground color
    val fgColor = patternFill \ "fgColor"
    assert(fgColor.nonEmpty, "gray125 pattern should have fgColor")
    val fgRgb = (fgColor \ "@rgb").text
    assert(fgRgb.nonEmpty, "fgColor should have rgb attribute")

    // Check background color
    val bgColor = patternFill \ "bgColor"
    assert(bgColor.nonEmpty, "gray125 pattern should have bgColor")
    val bgRgb = (bgColor \ "@rgb").text
    assert(bgRgb.nonEmpty, "bgColor should have rgb attribute")

    // Verify pattern type is gray125
    val patternType = (patternFill \ "@patternType").text
    assertEquals(patternType, "gray125", "Pattern type should be gray125")
  }
