package com.tjclp.xl.ooxml

import munit.FunSuite
import com.tjclp.xl.*
import com.tjclp.xl.cell.CellValue
import com.tjclp.xl.RichText.{*, given}
import com.tjclp.xl.macros.cell

/** Tests for rich text OOXML serialization */
class RichTextXmlSpec extends FunSuite:

  test("OoxmlCell: plain text uses simple <is><t> structure") {
    val ooxmlCell = OoxmlCell(cell"A1", CellValue.Text("Hello"), None, "inlineStr")
    val xml = ooxmlCell.toXml

    val xmlString = xml.toString
    assert(xmlString.contains("<is>"), "Should have inline string element")
    assert(xmlString.contains("<t>Hello</t>"), "Should have text element")
    assert(!xmlString.contains("<r>"), "Plain text should not have runs")
  }

  test("OoxmlCell: rich text with bold generates <r> with <b/>") {
    val richText = RichText("Bold".bold)
    val ooxmlCell = OoxmlCell(cell"A1", CellValue.RichText(richText), None, "inlineStr")
    val xml = ooxmlCell.toXml

    val xmlString = xml.toString
    assert(xmlString.contains("<is>"), "Should have inline string element")
    assert(xmlString.contains("<r>"), "Should have text run element")
    assert(xmlString.contains("<rPr>"), "Should have run properties")
    assert(xmlString.contains("<b/>"), "Should have bold tag")
    assert(xmlString.contains("<t>Bold</t>"), "Should have text content")
  }

  test("OoxmlCell: rich text with italic generates <r> with <i/>") {
    val richText = RichText("Italic".italic)
    val ooxmlCell = OoxmlCell(cell"A1", CellValue.RichText(richText), None, "inlineStr")
    val xml = ooxmlCell.toXml

    val xmlString = xml.toString
    assert(xmlString.contains("<i/>"), "Should have italic tag")
  }

  test("OoxmlCell: rich text with underline generates <r> with <u/>") {
    val richText = RichText("Underline".underline)
    val ooxmlCell = OoxmlCell(cell"A1", CellValue.RichText(richText), None, "inlineStr")
    val xml = ooxmlCell.toXml

    val xmlString = xml.toString
    assert(xmlString.contains("<u/>"), "Should have underline tag")
  }

  test("OoxmlCell: rich text with color generates <color rgb=>") {
    val richText = RichText("Red".red)
    val ooxmlCell = OoxmlCell(cell"A1", CellValue.RichText(richText), None, "inlineStr")
    val xml = ooxmlCell.toXml

    val xmlString = xml.toString
    assert(xmlString.contains("<color"), "Should have color element")
    assert(xmlString.contains("rgb="), "Should have rgb attribute")
    assert(xmlString.contains("FF0000"), "Should have red color value")
  }

  test("OoxmlCell: rich text with font size generates <sz val=>") {
    val richText = RichText("Big".size(18.0))
    val ooxmlCell = OoxmlCell(cell"A1", CellValue.RichText(richText), None, "inlineStr")
    val xml = ooxmlCell.toXml

    val xmlString = xml.toString
    assert(xmlString.contains("<sz"), "Should have size element")
    assert(xmlString.contains("val=\"18.0\""), "Should have size value")
  }

  test("OoxmlCell: rich text with font family generates <name val=>") {
    val richText = RichText("Text".fontFamily("Calibri"))
    val ooxmlCell = OoxmlCell(cell"A1", CellValue.RichText(richText), None, "inlineStr")
    val xml = ooxmlCell.toXml

    val xmlString = xml.toString
    assert(xmlString.contains("<name"), "Should have name element")
    assert(xmlString.contains("val=\"Calibri\""), "Should have font name")
  }

  test("OoxmlCell: rich text with multiple runs generates multiple <r> elements") {
    val richText = "Bold".bold + " normal " + "Italic".italic
    val ooxmlCell = OoxmlCell(cell"A1", CellValue.RichText(richText), None, "inlineStr")
    val xml = ooxmlCell.toXml

    val xmlString = xml.toString
    // Count occurrences of <r> (simple but effective check)
    val rCount = "<r>".r.findAllIn(xmlString).length
    assertEquals(rCount, 3, "Should have 3 text runs")
  }

  test("OoxmlCell: rich text with mixed formatting") {
    val richText = "Error: ".red.bold + "File not found".underline
    val ooxmlCell = OoxmlCell(cell"A1", CellValue.RichText(richText), None, "inlineStr")
    val xml = ooxmlCell.toXml

    val xmlString = xml.toString
    assert(xmlString.contains("<b/>"), "Should have bold")
    assert(xmlString.contains("<u/>"), "Should have underline")
    assert(xmlString.contains("<color"), "Should have color")
  }

  test("OoxmlCell: rich text run without formatting has no <rPr>") {
    val richText = RichText.plain("Plain")  // Create plain RichText
    val ooxmlCell = OoxmlCell(cell"A1", CellValue.RichText(richText), None, "inlineStr")
    val xml = ooxmlCell.toXml

    val xmlString = xml.toString
    assert(xmlString.contains("<r>"), "Should have run element")
    assert(xmlString.contains("<t>Plain</t>"), "Should have text")
    // Plain run has no font, so no <rPr> should be present
    assert(!xmlString.contains("<rPr>"), "Plain run should not have run properties")
  }

  test("OoxmlCell: rich text preserves text content exactly") {
    val richText = "Special chars: <>&\"".bold
    val ooxmlCell = OoxmlCell(cell"A1", CellValue.RichText(richText), None, "inlineStr")
    val xml = ooxmlCell.toXml

    val xmlString = xml.toString
    // XML library should auto-escape these characters
    assert(xmlString.contains("&lt;"), "Should escape <")
    assert(xmlString.contains("&gt;"), "Should escape >")
    assert(xmlString.contains("&amp;"), "Should escape &")
  }
