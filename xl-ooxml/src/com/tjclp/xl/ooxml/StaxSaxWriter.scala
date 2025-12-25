package com.tjclp.xl.ooxml

import javax.xml.stream.{XMLOutputFactory, XMLStreamWriter}
import java.io.OutputStream

/**
 * SAX writer interpreter backed by javax.xml.stream.XMLStreamWriter (StAX)
 *
 * @note
 *   Not thread-safe. Create separate instances for concurrent use.
 */
class StaxSaxWriter(underlying: XMLStreamWriter) extends SaxWriter:
  private val xmlNamespace = "http://www.w3.org/XML/1998/namespace"
  private val knownNamespaces = Map(
    "r" -> XmlUtil.nsRelationships,
    "mc" -> "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "xr" -> "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
    "x14ac" -> "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
  )

  private val namespaceStack =
    new java.util.ArrayDeque[scala.collection.mutable.Map[String, Option[String]]]()
  private val namespaceBindings = scala.collection.mutable.Map.empty[String, String]

  private def pushScope(): Unit =
    namespaceStack.push(scala.collection.mutable.Map.empty[String, Option[String]])

  private def popScope(): Unit =
    val scope = namespaceStack.pop()
    scope.foreach { case (prefix, previous) =>
      previous match
        case Some(uri) => namespaceBindings.update(prefix, uri)
        case None => namespaceBindings.remove(prefix)
    }

  private def ensureScope(): scala.collection.mutable.Map[String, Option[String]] =
    if namespaceStack.isEmpty then pushScope()
    namespaceStack.peek()

  private def recordNamespace(prefix: String, uri: String): Unit =
    val scope = ensureScope()
    if !scope.contains(prefix) then scope.update(prefix, namespaceBindings.get(prefix))
    namespaceBindings.update(prefix, uri)

  private def lookupNamespace(prefix: String): Option[String] =
    if prefix == "xml" then Some(xmlNamespace)
    else namespaceBindings.get(prefix).orElse(knownNamespaces.get(prefix))

  private def writeNamespaceDecl(prefix: String, uri: String, force: Boolean): Unit =
    val current = namespaceBindings.get(prefix)
    val changed = current.forall(_ != uri)
    if force || changed then
      if changed then recordNamespace(prefix, uri)
      if prefix.isEmpty then underlying.writeDefaultNamespace(uri)
      else underlying.writeNamespace(prefix, uri)

  def startDocument(): Unit =
    namespaceBindings.clear()
    namespaceStack.clear()
    underlying.writeStartDocument("UTF-8", "1.0")

  def endDocument(): Unit = underlying.writeEndDocument()

  def startElement(name: String): Unit =
    pushScope()
    underlying.writeStartElement(name)
  def startElement(name: String, namespace: String): Unit =
    // Parse prefix from element name (e.g., "r:id" → prefix="r", local="id")
    // StAX API: writeStartElement(prefix, localName, namespaceURI)
    pushScope()
    name.split(":", 2) match
      case Array(prefix, local) =>
        underlying.writeStartElement(prefix, local, namespace)
      case _ =>
        underlying.writeStartElement("", name, namespace)
  def writeAttribute(name: String, value: String): Unit =
    name match
      case "xmlns" =>
        writeNamespaceDecl("", value, force = true)
      case _ if name.startsWith("xmlns:") =>
        val prefix = name.stripPrefix("xmlns:")
        writeNamespaceDecl(prefix, value, force = true)
      case _ =>
        // Handle prefixed attributes with proper namespace URIs
        name.split(":", 2) match
          case Array(prefix, local) =>
            val ns = lookupNamespace(prefix)
            if prefix != "xml" then
              ns.foreach(uri => writeNamespaceDecl(prefix, uri, force = false))
            underlying.writeAttribute(prefix, ns.getOrElse(""), local, value)
          case _ =>
            underlying.writeAttribute(name, value)
  def writeCharacters(text: String): Unit = underlying.writeCharacters(text)
  def endElement(): Unit =
    underlying.writeEndElement()
    if !namespaceStack.isEmpty then popScope()
  def flush(): Unit = underlying.flush()

object StaxSaxWriter:
  def create(out: OutputStream): StaxSaxWriter =
    val factory = XMLOutputFactory.newInstance()
    factory.setProperty(XMLOutputFactory.IS_REPAIRING_NAMESPACES, false)
    new StaxSaxWriter(factory.createXMLStreamWriter(out, "UTF-8"))
