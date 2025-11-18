package com.tjclp.xl.ooxml

import munit.FunSuite

import scala.xml.XML

class WorkbookNamespaceSpec extends FunSuite:

  test("round-trip workbook preserves mc:Ignorable attribute and namespace bindings") {
    val input =
      """
        |<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        |          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        |          xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        |          xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
        |          xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
        |          mc:Ignorable="x15 xr">
        |  <fileVersion appName="xl" lastEdited="7"/>
        |  <sheets>
        |    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
        |  </sheets>
        |  <definedNames>
        |    <definedName name="Company_Name">Sheet1!$A$1</definedName>
        |  </definedNames>
        |</workbook>
        |""".stripMargin

    val elem = XML.loadString(input)
    val workbook = OoxmlWorkbook.fromXml(elem).fold(err => fail(err), identity)
    val regenerated = workbook.toXml

    val ignorable = regenerated.attribute(
      "http://schemas.openxmlformats.org/markup-compatibility/2006",
      "Ignorable"
    )
    assertEquals(ignorable.map(_.text), Some("x15 xr"))
    assertEquals(
      regenerated.scope.getURI("mc"),
      "http://schemas.openxmlformats.org/markup-compatibility/2006"
    )
    assertEquals(
      regenerated.scope.getURI("xr"),
      "http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
    )

    val xmlString = regenerated.toString
    val mcDeclCount = xmlString.split("xmlns:mc=").length - 1
    assertEquals(mcDeclCount, 1)
  }

  test("generated minimal workbook exposes default spreadsheet and relationship namespaces") {
    val xml = OoxmlWorkbook.minimal().toXml.toString
    assert(xml.contains("xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""))
    assert(
      xml.contains(
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""
      )
    )
  }
