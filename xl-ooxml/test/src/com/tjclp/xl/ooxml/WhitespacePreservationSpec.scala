package com.tjclp.xl.ooxml

import munit.FunSuite
import com.tjclp.xl.{CellValue, RichText, TextRun}
import com.tjclp.xl.macros.cell

/** Tests for whitespace preservation in plain text cells (OOXML writer) */
class WhitespacePreservationSpec extends FunSuite:

  test("plain text with leading space preserves xml:space") {
    val ooxmlCell = OoxmlCell(cell"A1", CellValue.Text(" leading"), None, "inlineStr")
    val xml = ooxmlCell.toXml

    val xmlString = xml.toString
    assert(xmlString.contains("<is>"), "Should have inline string element")

    // Check for xml:space="preserve"
    assert(xmlString.contains("xml:space=\"preserve\""), "Should have xml:space='preserve' for leading space")

    // Verify text content contains leading space
    assert(xmlString.contains(" leading"), "Text should contain leading space")
  }

  test("plain text with trailing space preserves xml:space") {
    val ooxmlCell = OoxmlCell(cell"A1", CellValue.Text("trailing "), None, "inlineStr")
    val xml = ooxmlCell.toXml

    val xmlString = xml.toString
    assert(xmlString.contains("xml:space=\"preserve\""), "Should have xml:space='preserve' for trailing space")
    assert(xmlString.contains("trailing "), "Text should contain trailing space")
  }

  test("plain text with double spaces preserves xml:space") {
    val ooxmlCell = OoxmlCell(cell"A1", CellValue.Text("double  space"), None, "inlineStr")
    val xml = ooxmlCell.toXml

    val xmlString = xml.toString
    assert(xmlString.contains("xml:space=\"preserve\""), "Should have xml:space='preserve' for double spaces")
    assert(xmlString.contains("double  space"), "Text should contain double spaces")
  }

  test("plain text without spaces omits xml:space") {
    val ooxmlCell = OoxmlCell(cell"A1", CellValue.Text("normal"), None, "inlineStr")
    val xml = ooxmlCell.toXml

    val xmlString = xml.toString
    assert(!xmlString.contains("xml:space"), "Should not have xml:space for normal text")
    assert(xmlString.contains("normal"), "Text should be present")
  }

  test("RichText with spaces includes xml:space on run text elements") {
    val richText = RichText(Vector(
      TextRun(" leading", None),
      TextRun("trailing ", None)
    ))
    val ooxmlCell = OoxmlCell(cell"A1", CellValue.RichText(richText), None, "inlineStr")
    val xml = ooxmlCell.toXml

    val xmlString = xml.toString
    assert(xmlString.contains("xml:space=\"preserve\""), "RichText with spaces should have xml:space='preserve' on <t> elements")
    assert(xmlString.contains(" leading"), "Should contain leading space")
    assert(xmlString.contains("trailing "), "Should contain trailing space")
  }
