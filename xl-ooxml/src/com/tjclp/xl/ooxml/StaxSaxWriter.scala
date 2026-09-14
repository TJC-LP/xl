package com.tjclp.xl.ooxml

import java.io.{BufferedWriter, OutputStream, OutputStreamWriter, Writer}
import java.nio.charset.StandardCharsets

/**
 * Streaming SAX writer interpreter: the `XmlBackend.SaxStax` backend.
 *
 * Writes the document straight to a UTF-8 [[java.io.Writer]] in the exact format the JDK's
 * `javax.xml.stream.XMLStreamWriter` produced when it used to sit underneath this class — the same
 * declaration (`<?xml version="1.0" encoding="UTF-8"?>`, no trailing newline), start/end tags,
 * empty-element tags closed by the next write, attributes in call order — so every byte pin on this
 * backend still holds. The JDK writer had to go for one reason (GH-649): its `writeAttribute`
 * escapes unconditionally and writes TAB, LF and CR raw, where XML attribute-value normalization
 * turns them into spaces on the next parse; there is no way to hand it a `&#10;`. Text and
 * attribute values now go through the same [[XmlUtil.escapeText]] / [[XmlUtil.escapeAttr]] the
 * `compact` backend uses, so a multiline data-validation prompt is `prompt="l1&#10;l2"` on both
 * backends, as Excel writes it, and the two agree byte for byte. XML-illegal control characters are
 * dropped on the way (GH-237 parity).
 *
 * The name is the backend vocabulary (`WriterConfig.saxStax`, `XmlBackend.SaxStax`), kept.
 *
 * @note
 *   Not thread-safe. Create separate instances for concurrent use.
 */
@SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
class StaxSaxWriter(out: Writer) extends SaxWriter:
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

  /** Qualified names of the open (non-empty) elements, innermost first, for their end tags. */
  private val elementStack = new java.util.ArrayDeque[String]()

  /** A start tag's attribute list is still open (the `>` or `/>` not yet written). */
  private var tagOpen = false

  /** ... and it is an empty-element tag (`emptyElement`), to be closed with `/>`. */
  private var tagEmpty = false

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
        if prefix.isEmpty then rawAttribute("xmlns", uri)
        else rawAttribute(s"xmlns:$prefix", uri)

  // ----- byte emission -----

  /** Finish the open start tag, if any: `>` for an element with content, `/>` for an empty one. */
  private def closeOpenTag(): Unit =
    if tagOpen then
      out.write(if tagEmpty then "/>" else ">")
      tagOpen = false
      tagEmpty = false

  private def openTag(qName: String, empty: Boolean): Unit =
    closeOpenTag()
    out.write('<')
    out.write(qName)
    tagOpen = true
    tagEmpty = empty
    if !empty then elementStack.push(qName)

  /** ` name="value"` with the value escaped ([[XmlUtil.escapeAttr]]); only inside a start tag. */
  private def rawAttribute(name: String, value: String): Unit =
    if !tagOpen then
      throw new IllegalStateException(s"attribute '$name' written outside a start tag")
    out.write(' ')
    out.write(name)
    out.write("=\"")
    if value.forall(c => c >= ' ' && c != '&' && c != '<' && c != '>' && c != '"') then
      out.write(value)
    else out.write(XmlUtil.escapeAttr(value))
    out.write('"')

  private def writeText(text: String): Unit =
    if text.forall(c => XmlUtil.isLegalXmlChar(c) && c != '&' && c != '<' && c != '>') then
      out.write(text)
    else out.write(XmlUtil.escapeText(text))

  // ----- SaxWriter -----

  def startDocument(): Unit =
    namespaceBindings.clear()
    namespaceStack.clear()
    elementStack.clear()
    tagOpen = false
    tagEmpty = false
    out.write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>")

  /** Closes every element still open, as the StAX `writeEndDocument` did. */
  def endDocument(): Unit =
    closeOpenTag()
    while !elementStack.isEmpty do
      out.write("</")
      out.write(elementStack.pop())
      out.write('>')

  def startElement(name: String): Unit =
    pushScope()
    openTag(name, empty = false)

  def startElement(name: String, namespace: String): Unit =
    // A prefixed name ("x15ac:absPath") declares its prefix; an unprefixed one the default namespace
    pushScope()
    openTag(name, empty = false)
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
        // A prefixed attribute declares its prefix first when the binding is known and new here
        name.split(":", 2) match
          case Array(prefix, _) =>
            if prefix != "xml" then
              lookupNamespace(prefix).foreach(uri => writeNamespaceDecl(prefix, uri, force = false))
            rawAttribute(name, value)
          case _ =>
            rawAttribute(name, value)

  def writeCharacters(text: String): Unit =
    closeOpenTag()
    writeText(text)

  def endElement(): Unit =
    closeOpenTag()
    if !elementStack.isEmpty then
      out.write("</")
      out.write(elementStack.pop())
      out.write('>')
    if !namespaceStack.isEmpty then popScope()

  // A true empty-element tag (`<f t="dataTable" .../>`), closed by the next write. No scope push
  // and no element-stack entry: an empty element cannot nest children.
  override def emptyElement(name: String, attrs: Seq[(String, String)]): Unit =
    openTag(name, empty = true)
    attrs.foreach { case (attrName, value) => writeAttribute(attrName, value) }

  def flush(): Unit = out.flush()

object StaxSaxWriter:
  /** A writer over `out`, UTF-8, buffered; call [[StaxSaxWriter.flush]] when done. */
  def create(out: OutputStream): StaxSaxWriter =
    new StaxSaxWriter(new BufferedWriter(new OutputStreamWriter(out, StandardCharsets.UTF_8)))
