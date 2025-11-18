package com.tjclp.xl.ooxml

import munit.FunSuite

import scala.xml.XML

class WorksheetNamespaceSpec extends FunSuite:

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
