package com.tjclp.xl.io.streaming

import java.io.{ByteArrayInputStream, ByteArrayOutputStream}
import java.nio.charset.StandardCharsets

import munit.FunSuite

import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.XmlSecurity

/**
 * GH-649: `--stream put/putf/style` copies every non-sheetData element attribute by attribute
 * (TransformHandler → StaxSaxWriter.writeAttribute). SAX hands the value back with `&#10;` already
 * decoded to a newline, so the writer must re-escape it — the JDK XMLStreamWriter wrote it raw and
 * the next read normalized the data-validation prompt to `l1 l2`.
 */
class StreamingAttributeEscapingSpec extends FunSuite:

  private val sheetXml =
    """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      |<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      |<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
      |<dataValidations count="1"><dataValidation type="list" allowBlank="1" showInputMessage="1" promptTitle="Two&#10;lines" prompt="l1&#10;l2&#9;tab&#13;cr" sqref="A2:A5"><formula1>"a,b"</formula1></dataValidation></dataValidations>
      |</worksheet>""".stripMargin

  test("GH-649: a streamed put re-emits a copied prompt's line break as &#10; and it reads back") {
    val out = new ByteArrayOutputStream()
    StreamingTransform.transformWorksheet(
      new ByteArrayInputStream(sheetXml.getBytes(StandardCharsets.UTF_8)),
      out,
      Map(ref"B1" -> StreamingTransform.CellPatch.SetValue(CellValue.Number(BigDecimal(2))))
    )
    val xml = out.toString("UTF-8")
    assert(xml.contains("""prompt="l1&#10;l2&#9;tab&#13;cr""""), xml)
    assert(xml.contains("""promptTitle="Two&#10;lines""""), xml)
    assert(!xml.contains("prompt=\"l1\nl2"), s"raw newline inside an attribute value: $xml")
    val ws = XmlSecurity.parseSafe(xml, "sheet1.xml").fold(e => fail(e.message), identity)
    val dv = (ws \ "dataValidations" \ "dataValidation").headOption.getOrElse(fail("dv lost"))
    assertEquals(dv \@ "prompt", "l1\nl2\ttab\rcr")
    assertEquals(dv \@ "promptTitle", "Two\nlines")
    assert(xml.contains("""<c r="B1"><v>2</v></c>"""), s"the put itself must land: $xml")
  }
