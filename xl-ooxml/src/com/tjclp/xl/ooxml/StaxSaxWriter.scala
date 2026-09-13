package com.tjclp.xl.ooxml

import java.io.OutputStream

/**
 * SAX writer interpreter that streams UTF-8 bytes straight to an `OutputStream`.
 *
 * Historically a thin layer over the JDK's non-repairing `javax.xml.stream.XMLStreamWriter`. That
 * writer emits TAB/LF/CR — and even XML-illegal C0 control characters — raw inside attribute
 * values, and always re-escapes `&`, so the character references attribute-value normalization
 * requires (`&#10;`) could not be pushed through it: a data-validation prompt with a line break, a
 * multi-line Name Manager comment or a table header with Alt+Enter came back as `l1 l2` (GH-649).
 * [[XmlTagWriter]] now owns the bytes and spells text and attribute values through
 * [[XmlUtil.escapeText]] / [[XmlUtil.escapeAttr]], the same tables the scala-xml backend uses —
 * parity by construction. The namespace bookkeeping (declaration placement, duplicate suppression,
 * well-known prefixes) is unchanged, and the bytes are identical to the JDK writer's except for the
 * newly escaped characters (pinned against the retired writer in AttributeEscapingSpec).
 *
 * @note
 *   Not thread-safe. Create separate instances for concurrent use.
 */
class StaxSaxWriter private (tag: XmlTagWriter) extends SaxWriter:
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
        tag.namespace(prefix, uri)

  def startDocument(): Unit =
    namespaceBindings.clear()
    namespaceStack.clear()
    tag.startDocument()

  def endDocument(): Unit = tag.endDocument()

  def startElement(name: String): Unit =
    pushScope()
    tag.startElement(name)
  def startElement(name: String, namespace: String): Unit =
    // The qualified name is written as given ("r:id" stays "r:id"); the declaration follows it
    // because a non-repairing writer never emits xmlns on its own
    pushScope()
    tag.startElement(name)
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
        // A prefixed attribute declares its namespace first when the prefix is known and unbound
        name.split(":", 2) match
          case Array(prefix, _) =>
            if prefix != "xml" then
              lookupNamespace(prefix).foreach(uri => writeNamespaceDecl(prefix, uri, force = false))
            tag.attribute(name, value)
          case _ =>
            tag.attribute(name, value)
  // XML-illegal control chars are dropped by the escaping (GH-237, XmlUtil.escapeText)
  def writeCharacters(text: String): Unit = tag.characters(text)
  def endElement(): Unit =
    tag.endElement()
    if !namespaceStack.isEmpty then popScope()
  // A true empty-element tag (`<name a="b"/>`, closed by the next write). No scope push: an empty
  // element cannot nest children.
  override def emptyElement(name: String, attrs: Seq[(String, String)]): Unit =
    tag.emptyElement(name)
    attrs.foreach { case (attrName, value) => writeAttribute(attrName, value) }
  def flush(): Unit = tag.flush()

object StaxSaxWriter:
  def create(out: OutputStream): StaxSaxWriter = new StaxSaxWriter(new XmlTagWriter(out))

/**
 * Minimal UTF-8 XML tag writer: start/empty/end tags, attributes, namespace declarations and text,
 * with the byte layout of the JDK's non-repairing `XMLStreamWriter` so the SaxStax backend's output
 * is unchanged apart from the escaping (GH-649): the declaration `<?xml version="1.0"
 * encoding="UTF-8"?>` without a standalone flag or trailing newline, `<a></a>` for a start/end
 * pair, `<a/>` for an empty element closed by the following call, attributes and namespace
 * declarations in call order, `"` verbatim in text. Each call hands its bytes to the stream before
 * returning — the JDK writer was unbuffered too, and callers read the target
 * `ByteArrayOutputStream` without an explicit flush. Unbalanced input is closed, never thrown on:
 * `endDocument` ends every open element, a surplus `endElement` is ignored, an attribute with no
 * open tag is dropped.
 */
@SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
private[ooxml] final class XmlTagWriter(out: OutputStream):
  private var buf = new Array[Byte](8192)
  private var pos = 0

  /** A start tag or empty-element tag is open: attributes and namespace declarations may follow. */
  private var tagOpen = false
  private var openIsEmpty = false
  private val open = scala.collection.mutable.ArrayBuffer.empty[String]
  private val scratch = new StringBuilder(256)

  def startDocument(): Unit =
    closeStartTag()
    chars("<?xml version=\"1.0\" encoding=\"UTF-8\"?>")
    emit()

  def endDocument(): Unit =
    closeStartTag()
    while open.nonEmpty do closeElement()
    emit()

  def startElement(name: String): Unit =
    closeStartTag()
    chars("<")
    chars(name)
    open += name
    tagOpen = true
    openIsEmpty = false
    emit()

  def emptyElement(name: String): Unit =
    closeStartTag()
    chars("<")
    chars(name)
    tagOpen = true
    openIsEmpty = true
    emit()

  def attribute(name: String, value: String): Unit =
    if tagOpen then
      chars(" ")
      chars(name)
      chars("=\"")
      escaped(value)
      chars("\"")
      emit()

  /** `xmlns="uri"` for the empty (or `xmlns`) prefix; the `xml` prefix is never declared. */
  def namespace(prefix: String, uri: String): Unit =
    if tagOpen && prefix != "xml" then
      chars(" xmlns")
      if prefix.nonEmpty && prefix != "xmlns" then
        chars(":")
        chars(prefix)
      chars("=\"")
      escaped(uri)
      chars("\"")
      emit()

  def characters(text: String): Unit =
    closeStartTag()
    scratch.clear()
    XmlUtil.escapeText(text, scratch)
    chars(scratch)
    emit()

  def endElement(): Unit =
    closeStartTag()
    if open.nonEmpty then closeElement()
    emit()

  def flush(): Unit =
    emit()
    out.flush()

  private def closeElement(): Unit =
    val name = open.remove(open.length - 1)
    chars("</")
    chars(name)
    chars(">")

  private def closeStartTag(): Unit =
    if tagOpen then
      if openIsEmpty then chars("/>") else chars(">")
      tagOpen = false
      openIsEmpty = false

  private def escaped(value: String): Unit =
    scratch.clear()
    XmlUtil.escapeAttr(value, scratch)
    chars(scratch)

  private def emit(): Unit =
    if pos > 0 then
      out.write(buf, 0, pos)
      pos = 0

  private def ensure(n: Int): Unit =
    if pos + n > buf.length then
      emit()
      if n > buf.length then buf = new Array[Byte](math.max(n, buf.length * 2))

  /** UTF-8 encode; a surrogate pair is one 4-byte sequence, an unpaired surrogate becomes '?'. */
  private def chars(cs: CharSequence): Unit =
    val n = cs.length
    ensure(n * 3)
    var i = 0
    while i < n do
      val c = cs.charAt(i)
      if c < 0x80 then
        buf(pos) = c.toByte
        pos += 1
      else if c < 0x800 then
        buf(pos) = (0xc0 | (c >> 6)).toByte
        buf(pos + 1) = (0x80 | (c & 0x3f)).toByte
        pos += 2
      else if Character.isHighSurrogate(c) && i + 1 < n && Character.isLowSurrogate(
          cs.charAt(i + 1)
        )
      then
        val cp = Character.toCodePoint(c, cs.charAt(i + 1))
        buf(pos) = (0xf0 | (cp >> 18)).toByte
        buf(pos + 1) = (0x80 | ((cp >> 12) & 0x3f)).toByte
        buf(pos + 2) = (0x80 | ((cp >> 6) & 0x3f)).toByte
        buf(pos + 3) = (0x80 | (cp & 0x3f)).toByte
        pos += 4
        i += 1
      else if Character.isSurrogate(c) then
        buf(pos) = '?'.toByte
        pos += 1
      else
        buf(pos) = (0xe0 | (c >> 12)).toByte
        buf(pos + 1) = (0x80 | ((c >> 6) & 0x3f)).toByte
        buf(pos + 2) = (0x80 | (c & 0x3f)).toByte
        pos += 3
      i += 1
