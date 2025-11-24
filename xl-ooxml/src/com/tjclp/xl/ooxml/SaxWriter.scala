package com.tjclp.xl.ooxml

/**
 * Pure abstraction for SAX-style XML emission.
 *
 * Keeps effectful XML writers (e.g., javax.xml.stream.XMLStreamWriter) out of the pure OOXML
 * module. Implementations can wrap imperative writers inside xl-cats-effect.
 */
trait SaxWriter:
  def startDocument(): Unit
  def endDocument(): Unit
  def startElement(name: String): Unit
  def startElement(name: String, namespace: String): Unit
  def writeAttribute(name: String, value: String): Unit
  def writeCharacters(text: String): Unit
  def endElement(): Unit
  def flush(): Unit

/** Types that can serialize themselves using a SaxWriter */
trait SaxSerializable:
  def writeSax(writer: SaxWriter): Unit

object SaxWriter:

  /**
   * Write attributes in deterministic order, then execute the body to emit element contents.
   *
   * Deterministic attribute ordering is required for stable byte-for-byte output and golden tests.
   */
  def withAttributes(writer: SaxWriter, attrs: (String, String)*)(body: => Unit): Unit =
    attrs.sortBy(_._1).foreach { case (name, value) =>
      writer.writeAttribute(name, value)
    }
    body
