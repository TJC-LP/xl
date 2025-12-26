package com.tjclp.xl.ooxml

import munit.FunSuite
import javax.xml.stream.XMLOutputFactory
import scala.xml.XML
import java.io.StringWriter
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.addressing.{ARef, Column, Row, SheetName}
import com.tjclp.xl.sheets.Sheet

class OoxmlWorksheetSaxSpec extends FunSuite:

  test("writeSax emits inline strings with xml:space preserve when needed") {
    val cell = Cell(ARef(Column.from0(0), Row.from0(0)), CellValue.Text("  spaced"))
    val sheet = Sheet(SheetName.unsafe("Sheet1"), Map(cell.ref -> cell))
    val ws = OoxmlWorksheet.fromDomain(sheet)

    val saxXml = writeWithSax(ws)
    val parsed = XML.loadString(saxXml)

    val tElem = (parsed \\ "t").headOption.getOrElse(fail("expected <t> element"))
    val spaceAttr = tElem.attributes.asAttrMap.get("xml:space").orElse(tElem.attributes.asAttrMap.get("space"))
    assertEquals(spaceAttr, Some("preserve"))
    assertEquals(tElem.text, "  spaced")
  }

  test("writeSax sheetData matches toXml for simple rows") {
    val cell1 = Cell(ARef(Column.from0(0), Row.from0(0)), CellValue.Number(BigDecimal(42)))
    val cell2 = Cell(ARef(Column.from0(1), Row.from0(1)), CellValue.Bool(true))
    val sheet = Sheet(SheetName.unsafe("Sheet1"), Map(cell1.ref -> cell1, cell2.ref -> cell2))
    val ws = OoxmlWorksheet.fromDomain(sheet)

    val saxXml = XML.loadString(writeWithSax(ws))
    val xmlTree = ws.toXml

    // Compare key structures
    val saxRows = saxXml \\ "row"
    val xmlRows = xmlTree \\ "row"
    assertEquals(saxRows.map(_ \ "@r").map(_.text), xmlRows.map(_ \ "@r").map(_.text))
    assertEquals(saxRows.flatMap(_ \ "c").map(_ \ "@r").map(_.text), xmlRows.flatMap(_ \ "c").map(_ \ "@r").map(_.text))
  }

  private def writeWithSax(ws: OoxmlWorksheet): String =
    val output = StringWriter()
    val xmlWriter = XMLOutputFactory.newInstance().createXMLStreamWriter(output)
    val saxWriter = new TestSaxWriter(xmlWriter)
    ws.writeSax(saxWriter)
    xmlWriter.flush()
    output.toString

private class TestSaxWriter(underlying: javax.xml.stream.XMLStreamWriter) extends SaxWriter:
  def startDocument(): Unit = underlying.writeStartDocument("UTF-8", "1.0")
  def endDocument(): Unit = underlying.writeEndDocument()
  def startElement(name: String): Unit = underlying.writeStartElement(name)
  def startElement(name: String, namespace: String): Unit =
    underlying.writeStartElement("", name, namespace)
  def writeAttribute(name: String, value: String): Unit =
    // Handle prefixed attributes like xml:space
    name.split(":", 2) match
      case Array(prefix, local) =>
        val ns = if prefix == "xml" then "http://www.w3.org/XML/1998/namespace" else ""
        underlying.writeAttribute(prefix, ns, local, value)
      case _ => underlying.writeAttribute(name, value)
  def writeCharacters(text: String): Unit = underlying.writeCharacters(text)
  def endElement(): Unit = underlying.writeEndElement()
  def flush(): Unit = underlying.flush()
