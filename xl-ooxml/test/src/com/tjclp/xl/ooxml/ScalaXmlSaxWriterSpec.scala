package com.tjclp.xl.ooxml

import munit.FunSuite
import scala.xml.XML
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.addressing.{ARef, Column, Row, SheetName}
import com.tjclp.xl.sheets.Sheet

class ScalaXmlSaxWriterSpec extends FunSuite:

  test("ScalaXmlSaxWriter builds Elem equivalent to toXml for simple sheet") {
    val cell = Cell(ARef(Column.from0(0), Row.from0(0)), CellValue.Number(BigDecimal(10)))
    val sheet = Sheet(SheetName.unsafe("Sheet1"), Map(cell.ref -> cell))
    val ws = OoxmlWorksheet.fromDomain(sheet)

    val writer = ScalaXmlSaxWriter()
    ws.writeSax(writer)
    val saxElem = writer.result.getOrElse(fail("expected worksheet root"))

    val domElem = ws.toXml

    assertEquals((saxElem \\ "c").map(_ \ "@r").map(_.text), (domElem \\ "c").map(_ \ "@r").map(_.text))
  }

  test("ScalaXmlSaxWriter preserves xml:space on inline strings") {
    val cell = Cell(ARef(Column.from0(0), Row.from0(0)), CellValue.Text("  spaced"))
    val sheet = Sheet(SheetName.unsafe("Sheet1"), Map(cell.ref -> cell))
    val ws = OoxmlWorksheet.fromDomain(sheet)

    val writer = ScalaXmlSaxWriter()
    ws.writeSax(writer)
    val saxElem = writer.result.getOrElse(fail("expected worksheet root"))

    val tElem = (saxElem \\ "t").headOption.getOrElse(fail("expected <t> element"))
    val spaceAttr = tElem.attributes.asAttrMap.get("xml:space").orElse(tElem.attributes.asAttrMap.get("space"))
    assertEquals(spaceAttr, Some("preserve"))
    assertEquals(tElem.text, "  spaced")
  }
