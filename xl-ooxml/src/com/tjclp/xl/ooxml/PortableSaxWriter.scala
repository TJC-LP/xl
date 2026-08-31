package com.tjclp.xl.ooxml

import java.io.OutputStream
import com.tjclp.xl.xml.XmlTextWriter

/**
 * SAX writer interpreter backed by the portable [[com.tjclp.xl.xml.XmlTextWriter]] (ADR-016 wave
 * A2): byte-identical to [[StaxSaxWriter]] on the SaxWriter protocol, with no javax.xml.stream
 * dependency.
 *
 * Structure is a verbatim copy of StaxSaxWriter — same knownNamespaces table, same scope
 * push/pop/rebind-restore, same writeNamespaceDecl force/alreadyBound/changed logic, same
 * prefixed-name handling, same GH-237 sanitize in writeCharacters, same no-scope-push emptyElement
 * — so byte parity holds by construction and is enforced by the W1/W2 parity suites.
 *
 * flush() is watermark-idempotent: it writes only the bytes produced since the previous flush (the
 * Deflated worksheet path flushes twice — inside DirectSaxEmitter.emitWorksheet and again in
 * XlsxWriter — and must not double-write).
 *
 * @note
 *   Not thread-safe. Create separate instances for concurrent use.
 */
class PortableSaxWriter(underlying: XmlTextWriter, out: OutputStream) extends SaxWriter:
  private val xmlNamespace = "http://www.w3.org/XML/1998/namespace"
  private val knownNamespaces = Map(
    "r" -> XmlUtil.nsRelationships,
    "mc" -> "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "xr" -> "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
    "x14ac" -> "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
    "x14" -> "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
    "xm" -> "http://schemas.microsoft.com/office/excel/2006/main"
  )

  private val namespaceStack =
    new java.util.ArrayDeque[scala.collection.mutable.Map[String, Option[String]]]()
  private val namespaceBindings = scala.collection.mutable.Map.empty[String, String]

  /** Flush watermark: number of underlying bytes already written to `out`. */
  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  private var flushed = 0

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
    val alreadyBound = current.contains(uri)
    // Don't write duplicate declarations - if already bound to same URI, skip
    if alreadyBound then ()
    else
      val changed = current.forall(_ != uri)
      if force || changed then
        if changed then recordNamespace(prefix, uri)
        if prefix.isEmpty then underlying.writeAttribute("xmlns", uri)
        else underlying.writeAttribute(s"xmlns:$prefix", uri)

  def startDocument(): Unit =
    namespaceBindings.clear()
    namespaceStack.clear()
    underlying.writeStartDocument()

  def endDocument(): Unit = underlying.writeEndDocument()

  def startElement(name: String): Unit =
    pushScope()
    underlying.writeStartElement(name)
  def startElement(name: String, namespace: String): Unit =
    // Parse prefix from element name (e.g., "r:id" → prefix="r", local="id"); the portable writer
    // takes the qName as written and the xmlns declaration is emitted explicitly (StAX with
    // IS_REPAIRING_NAMESPACES=false behaves the same way).
    pushScope()
    underlying.writeStartElement(name)
    name.split(":", 2) match
      case Array(prefix, _) => writeNamespaceDecl(prefix, namespace, force = true)
      case _ => writeNamespaceDecl("", namespace, force = true)
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
          case Array(prefix, _) =>
            val ns = lookupNamespace(prefix)
            if prefix != "xml" then
              ns.foreach(uri => writeNamespaceDecl(prefix, uri, force = false))
            underlying.writeAttribute(name, value)
          case _ =>
            underlying.writeAttribute(name, value)
  // GH-237: strip XML-illegal control chars (identical policy to StaxSaxWriter so backends never
  // diverge). XmlUtil.sanitizeXmlText is a no-op fast path for clean text.
  def writeCharacters(text: String): Unit =
    underlying.writeCharacters(XmlUtil.sanitizeXmlText(text))
  def endElement(): Unit =
    underlying.writeEndElement()
    if !namespaceStack.isEmpty then popScope()
  // A true empty-element tag (`/>`, closed lazily by the next write). No scope push: an empty
  // element cannot nest children (GH-430).
  override def emptyElement(name: String, attrs: Seq[(String, String)]): Unit =
    underlying.writeEmptyElement(name)
    attrs.foreach { case (attrName, value) => writeAttribute(attrName, value) }
  def flush(): Unit =
    val bytes = underlying.toBytes
    if bytes.length > flushed then
      out.write(bytes, flushed, bytes.length - flushed)
      flushed = bytes.length
    out.flush()

object PortableSaxWriter:
  def create(out: OutputStream): PortableSaxWriter =
    new PortableSaxWriter(new XmlTextWriter, out)
