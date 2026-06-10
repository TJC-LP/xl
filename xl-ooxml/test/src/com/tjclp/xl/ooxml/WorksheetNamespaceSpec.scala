package com.tjclp.xl.ooxml

import java.io.{ByteArrayOutputStream, StringReader}
import javax.xml.parsers.DocumentBuilderFactory

import munit.FunSuite

import scala.xml.XML

class WorksheetNamespaceSpec extends FunSuite:

  private val nsRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

  private def parseNamespaceAware(xml: String): org.w3c.dom.Document =
    val factory = DocumentBuilderFactory.newInstance()
    factory.setNamespaceAware(true)
    try factory.newDocumentBuilder().parse(new org.xml.sax.InputSource(new StringReader(xml)))
    catch
      case e: org.xml.sax.SAXException =>
        fail(s"output is not namespace-well-formed: ${e.getMessage}\n$xml")

  private def parseWorksheet(input: String): worksheet.OoxmlWorksheet =
    OoxmlWorksheet
      .fromXml(XML.loadString(input))
      .fold(err => fail(s"Failed to parse worksheet: $err"), identity)

  // GH-291: openpyxl declares xmlns:r locally on <drawing>, not on the worksheet root. The
  // regenerated worksheet must keep the prefix bound (hoisted onto the root, the standard Excel
  // layout) on BOTH emission paths or Excel reports the workbook as corrupt.
  private val openpyxlStyleInput =
    """
      |<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      |  <sheetData>
      |    <row r="1">
      |      <c r="A1" t="inlineStr"><is><t>Hello</t></is></c>
      |    </row>
      |  </sheetData>
      |  <drawing xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
      |           r:id="rId1"/>
      |</worksheet>
      |""".stripMargin

  test("GH-291: toXml keeps locally-declared xmlns:r on <drawing> bound (hoisted to root)") {
    val regenerated = parseWorksheet(openpyxlStyleInput).toXml

    assertEquals(Option(regenerated.scope.getURI("r")), Some(nsRel))

    val doc = parseNamespaceAware(XmlUtil.compact(regenerated))
    val drawings = doc.getElementsByTagNameNS("*", "drawing")
    assertEquals(drawings.getLength, 1)
    val id = drawings.item(0) match
      case e: org.w3c.dom.Element => e.getAttributeNS(nsRel, "id")
      case other => fail(s"unexpected node: $other")
    assertEquals(id, "rId1")
  }

  test("GH-291: writeSax keeps locally-declared xmlns:r on <drawing> bound (hoisted to root)") {
    val ws = parseWorksheet(openpyxlStyleInput)
    val out = new ByteArrayOutputStream()
    val saxWriter = StaxSaxWriter.create(out)
    ws.writeSax(saxWriter)
    val xml = out.toString("UTF-8")

    val doc = parseNamespaceAware(xml)
    // Bound at the ROOT (standard Excel layout, parity with the ScalaXml backend)
    assertEquals(Option(doc.getDocumentElement.lookupNamespaceURI("r")), Some(nsRel))
    val drawings = doc.getElementsByTagNameNS("*", "drawing")
    assertEquals(drawings.getLength, 1)
    val id = drawings.item(0) match
      case e: org.w3c.dom.Element => e.getAttributeNS(nsRel, "id")
      case other => fail(s"unexpected node: $other")
    assertEquals(id, "rId1")
  }

  test("GH-291: row-level x14ac:dyDescent stays bound when the root lacks the prefix") {
    // Rows re-emit x14ac:dyDescent as a raw qualified attribute; a source that bound the prefix
    // locally on <row> (legal XML) must not regenerate into an unbound prefix.
    val input =
      """
        |<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        |  <sheetData>
        |    <row r="1" x14ac:dyDescent="0.25"
        |         xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
        |      <c r="A1" t="inlineStr"><is><t>Hello</t></is></c>
        |    </row>
        |  </sheetData>
        |</worksheet>
        |""".stripMargin

    val regenerated = parseWorksheet(input).toXml
    val doc = parseNamespaceAware(XmlUtil.compact(regenerated))
    assertEquals(
      Option(doc.getDocumentElement.lookupNamespaceURI("x14ac")),
      Some("http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac")
    )
  }

  test("GH-291: locally-declared prefix inside extLst survives regeneration") {
    // x14 is NOT in any well-known fallback table: the binding must be carried from the source.
    val input =
      """
        |<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        |  <sheetData>
        |    <row r="1">
        |      <c r="A1" t="inlineStr"><is><t>Hello</t></is></c>
        |    </row>
        |  </sheetData>
        |  <extLst>
        |    <ext uri="{example}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
        |      <x14:sparklineGroups/>
        |    </ext>
        |  </extLst>
        |</worksheet>
        |""".stripMargin

    val regenerated = parseWorksheet(input).toXml
    val doc = parseNamespaceAware(XmlUtil.compact(regenerated))
    val groups =
      doc.getElementsByTagNameNS(
        "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
        "sparklineGroups"
      )
    assertEquals(groups.getLength, 1, "x14:sparklineGroups lost its namespace binding")
  }

  test("worksheet preserves namespaces and avoids pollution") {
    val input =
      """
        |<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        |           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        |           xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        |           xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
        |           mc:Ignorable="x14ac">
        |  <sheetPr xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        |           xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
        |    <tabColor rgb="FFFFFF"/>
        |  </sheetPr>
        |  <sheetData>
        |    <row r="1">
        |      <c r="A1" t="inlineStr"><is><t>Hello</t></is></c>
        |    </row>
        |  </sheetData>
        |</worksheet>
        |""".stripMargin

    val worksheet = OoxmlWorksheet
      .fromXml(XML.loadString(input))
      .fold(err => fail(s"Failed to parse worksheet: $err"), identity)

    val regenerated = worksheet.toXml

    val mcUri = regenerated.scope.getURI("mc")
    assertEquals(mcUri, "http://schemas.openxmlformats.org/markup-compatibility/2006")
    val ignorable = regenerated.attribute(
      "http://schemas.openxmlformats.org/markup-compatibility/2006",
      "Ignorable"
    )
    assertEquals(ignorable.map(_.text), Some("x14ac"))

    val xmlString = regenerated.toString
    val mcDeclCount = xmlString.split("xmlns:mc=").length - 1
    assertEquals(mcDeclCount, 1)

    val sheetPrXml = (regenerated \ "sheetPr").headOption
      .getOrElse(fail("sheetPr missing"))
      .toString
    assert(!sheetPrXml.contains("xmlns:mc"))
    assert(!sheetPrXml.contains("xmlns:x14ac"))
  }
