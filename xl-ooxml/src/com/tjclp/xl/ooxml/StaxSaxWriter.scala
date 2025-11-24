package com.tjclp.xl.ooxml

import javax.xml.stream.{XMLOutputFactory, XMLStreamWriter}
import java.io.OutputStream

/** SAX writer interpreter backed by javax.xml.stream.XMLStreamWriter (StAX) */
class StaxSaxWriter(underlying: XMLStreamWriter) extends SaxWriter:
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

object StaxSaxWriter:
  def create(out: OutputStream): StaxSaxWriter =
    val factory = XMLOutputFactory.newInstance()
    factory.setProperty(XMLOutputFactory.IS_REPAIRING_NAMESPACES, false)
    new StaxSaxWriter(factory.createXMLStreamWriter(out, "UTF-8"))
