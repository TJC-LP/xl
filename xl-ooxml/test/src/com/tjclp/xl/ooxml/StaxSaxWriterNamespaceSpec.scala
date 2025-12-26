package com.tjclp.xl.ooxml

import munit.FunSuite
import java.io.{ByteArrayInputStream, ByteArrayOutputStream}
import java.nio.charset.StandardCharsets
import javax.xml.parsers.DocumentBuilderFactory
import org.w3c.dom.Element

class StaxSaxWriterNamespaceSpec extends FunSuite:

  test("prefixed attributes use declared namespace bindings") {
    val output = new ByteArrayOutputStream()
    val writer = StaxSaxWriter.create(output)

    writer.startDocument()
    writer.startElement("root")
    writer.writeAttribute("xmlns:x14ac", XmlUtil.nsX14ac)
    writer.startElement("child")
    writer.writeAttribute("x14ac:dyDescent", "0.25")
    writer.endElement()
    writer.endElement()
    writer.endDocument()
    writer.flush()

    val xml = new String(output.toByteArray, StandardCharsets.UTF_8)
    assert(xml.contains(s"""xmlns:x14ac="${XmlUtil.nsX14ac}""""))
    assert(!xml.contains("""xmlns:x14ac="""""))

    val factory = DocumentBuilderFactory.newInstance()
    factory.setNamespaceAware(true)
    val doc =
      factory
        .newDocumentBuilder()
        .parse(new ByteArrayInputStream(xml.getBytes(StandardCharsets.UTF_8)))
    val child = doc.getElementsByTagName("child").item(0)
    assert(child != null)
    val attrs = child.getAttributes
    assert(attrs != null)
    val attr = attrs.getNamedItemNS(XmlUtil.nsX14ac, "dyDescent")
    assert(attr != null)
    assertEquals(attr.getNodeValue, "0.25")
  }
