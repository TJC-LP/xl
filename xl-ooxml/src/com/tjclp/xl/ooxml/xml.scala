package com.tjclp.xl.ooxml

import scala.xml.*

import com.tjclp.xl.error.{XLError, XLResult}

/**
 * XML serialization and deserialization for OOXML parts.
 *
 * All serialization is deterministic with canonical attribute/element ordering for stable diffs and
 * golden tests.
 */

/** Trait for types that can be serialized to XML */
trait XmlWritable:
  /** Convert to XML element with deterministic formatting */
  def toXml: Elem

/** Trait for types that can be deserialized from XML */
trait XmlReadable[A]:
  /** Parse from XML element, returning error message on failure */
  def fromXml(elem: Elem): Either[String, A]

/** XML utilities for OOXML */
object XmlUtil:
  /** OOXML namespaces */
  val nsSpreadsheetML = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  val nsRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  val nsPackageRels = "http://schemas.openxmlformats.org/package/2006/relationships"
  val nsContentTypes = "http://schemas.openxmlformats.org/package/2006/content-types"
  val nsX14ac = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"

  /** Relationship type URIs */
  val relTypeOfficeDocument =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
  val relTypeWorksheet =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
  val relTypeStyles = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
  val relTypeSharedStrings =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"

  /** Content type URIs */
  val ctWorkbook = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
  val ctWorksheet = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
  val ctStyles = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
  val ctSharedStrings =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
  val ctRelationships = "application/vnd.openxmlformats-package.relationships+xml"

  /** Sort attributes by name for deterministic output */
  def sortAttributes(attrs: MetaData): MetaData =
    val sorted = attrs.asAttrMap.toSeq.sortBy(_._1)
    sorted.foldLeft(Null: MetaData) { case (acc, (key, value)) =>
      new UnprefixedAttribute(key, value, acc)
    }

  /** Create element with sorted attributes */
  def elem(label: String, attrs: (String, String)*)(children: Node*): Elem =
    val sortedAttrs = attrs.sortBy(_._1).foldLeft(Null: MetaData) { case (acc, (key, value)) =>
      new UnprefixedAttribute(key, value, acc)
    }
    Elem(null, label, sortedAttrs, TopScope, minimizeEmpty = true, children*)

  /**
   * Create XML element with attributes in EXACT order provided (no sorting).
   *
   * Use this for elements where Excel expects specific attribute order (e.g., row, cell).
   */
  def elemOrdered(label: String, attrs: (String, String)*)(children: Node*): Elem =
    val orderedAttrs = attrs.foldRight(Null: MetaData) { case ((key, value), acc) =>
      new UnprefixedAttribute(key, value, acc)
    }
    Elem(null, label, orderedAttrs, TopScope, minimizeEmpty = true, children*)

  /** Create element with namespace */
  def elemNS(prefix: String, label: String, ns: String, attrs: (String, String)*)(
    children: Node*
  ): Elem =
    val sortedAttrs = attrs.sortBy(_._1).foldLeft(Null: MetaData) { case (acc, (key, value)) =>
      new UnprefixedAttribute(key, value, acc)
    }
    Elem(
      prefix,
      label,
      sortedAttrs,
      NamespaceBinding(prefix, ns, TopScope),
      minimizeEmpty = true,
      children*
    )

  /** Pretty print XML with proper indentation */
  def prettyPrint(node: Node): String =
    val printer = new PrettyPrinter(80, 2)
    s"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${printer.format(node)}"""

  /** Compact XML without indentation (newline after declaration for Excel compatibility) */
  def compact(node: Node): String =
    val printer = new PrettyPrinter(0, 0)
    s"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${printer.format(node)}"""

  /** Get required attribute value */
  def getAttr(elem: Elem, name: String): Either[String, String] =
    elem.attribute(name).map(_.text).toRight(s"Missing required attribute: $name")

  /** Get optional attribute value */
  def getAttrOpt(elem: Elem, name: String): Option[String] =
    elem.attribute(name).map(_.text)

  /**
   * Get optional namespaced attribute value.
   *
   * Handles attributes with namespace prefixes (e.g., "x14ac:dyDescent"). Scala XML stores
   * namespaced attributes by URI, so we need to resolve the prefix.
   *
   * @param elem
   *   The element to search
   * @param prefixedName
   *   Attribute name with prefix (e.g., "x14ac:dyDescent")
   * @return
   *   Some(value) if attribute found, None otherwise
   */
  def getNamespacedAttrOpt(elem: Elem, prefixedName: String): Option[String] =
    prefixedName.split(':') match
      case Array(prefix, localName) =>
        // Resolve prefix to namespace URI using element's scope
        val nsUri = elem.scope.getURI(prefix)
        if nsUri != null then elem.attribute(nsUri, localName).map(_.text)
        else None // Prefix not found in scope
      case _ =>
        // No prefix - fallback to normal attribute lookup
        getAttrOpt(elem, prefixedName)

  /** Get required child element */
  def getChild(elem: Elem, label: String): Either[String, Elem] =
    (elem \ label).headOption match
      case Some(e: Elem) => Right(e)
      case _ => Left(s"Missing required child element: $label")

  /** Get all child elements with given label */
  def getChildren(elem: Elem, label: String): Seq[Elem] =
    (elem \ label).collect { case e: Elem => e }

  /**
   * Determine whether a text node requires xml:space="preserve".
   *
   * REQUIRES: s is a (possibly empty) String ENSURES:
   *   - Returns true when text has leading or trailing whitespace
   *   - Returns true when text contains consecutive spaces (" ")
   *   - Returns false for strings that Excel can safely trim
   * DETERMINISTIC: Yes (pure function)
   */
  def needsXmlSpacePreserve(s: String): Boolean =
    s.nonEmpty && (s.startsWith(" ") || s.endsWith(" ") || s.contains("  "))

  /**
   * Extract text content from XML element, preserving whitespace.
   *
   * Unlike `.text` which normalizes whitespace, this method preserves exact whitespace (including
   * leading/trailing/multiple spaces) by extracting raw text node content.
   *
   * REQUIRES: elem is valid Elem ENSURES:
   *   - Returns exact text content without normalization
   *   - Preserves leading/trailing whitespace
   *   - Preserves multiple consecutive spaces
   *   - Empty if element has no text children
   * DETERMINISTIC: Yes (pure function)
   *
   * @param elem
   *   XML element to extract text from
   * @return
   *   Raw text content preserving all whitespace
   */
  def getTextPreservingWhitespace(elem: Elem): String =
    elem.child.collect {
      case scala.xml.Text(data) => data
      case scala.xml.PCData(data) => data
    }.mkString

  /**
   * Recursively strip namespace declarations from XML element tree.
   *
   * Removes redundant xmlns attributes that cause Excel corruption when elements are re-embedded in
   * a parent that already declares the namespace.
   *
   * REQUIRES: elem is valid Elem ENSURES:
   *   - Returns Elem with TopScope (no namespace bindings)
   *   - Recursively processes all child Elems
   *   - Preserves all attributes except xmlns
   *   - Preserves text content and structure
   * DETERMINISTIC: Yes (pure transformation)
   *
   * @param elem
   *   Element to strip namespaces from
   * @return
   *   Element with TopScope applied recursively
   */
  def stripNamespaces(elem: Elem): Elem =
    val cleanedChildren = elem.child.map {
      case e: Elem => stripNamespaces(e)
      case other => other
    }
    elem.copy(scope = TopScope, child = cleanedChildren)

  /**
   * Parse run properties (<rPr>) to Font.
   *
   * REQUIRES: rPrElem is <rPr> element from OOXML ENSURES:
   *   - Returns Font with properties extracted from child elements
   *   - Missing properties use Font defaults (Calibri, 11pt, no formatting)
   * DETERMINISTIC: Yes (pure XML traversal)
   *
   * OOXML structure:
   *   - <b/> → bold
   *   - <i/> → italic
   *   - <u/> → underline
   *   - <color rgb="RRGGBB"/> → font color (hex without # prefix)
   *   - <sz val="14.0"/> → size in points
   *   - <name val="Arial"/> → font family
   *
   * @param rPrElem
   *   The <rPr> element to parse
   * @return
   *   Font with formatting properties (default Font if no properties)
   */
  def parseRunProperties(rPrElem: Elem): com.tjclp.xl.styles.font.Font =
    import com.tjclp.xl.styles.font.Font
    import com.tjclp.xl.styles.color.Color

    val bold = (rPrElem \ "b").nonEmpty
    val italic = (rPrElem \ "i").nonEmpty
    val underline = (rPrElem \ "u").nonEmpty

    val color =
      (rPrElem \ "color").headOption.collect { case elem: Elem => elem }.flatMap { colorElem =>
        getAttrOpt(colorElem, "rgb").flatMap { rgb =>
          // Add # prefix for Color.fromHex
          Color.fromHex(s"#$rgb").toOption
        }
      }

    val sizePt = (rPrElem \ "sz").headOption
      .collect { case elem: Elem => elem }
      .flatMap(e => getAttrOpt(e, "val"))
      .flatMap(_.toDoubleOption)
      .getOrElse(11.0)

    val name = (rPrElem \ "name").headOption
      .collect { case elem: Elem => elem }
      .flatMap(e => getAttrOpt(e, "val"))
      .getOrElse("Calibri")

    Font(
      name = name,
      sizePt = sizePt,
      bold = bold,
      italic = italic,
      underline = underline,
      color = color
    )

  /**
   * Parse text runs (<r> elements) to RichText.
   *
   * REQUIRES: runElems is sequence of <r> elements from <si> or <is> ENSURES:
   *   - Returns RichText with TextRun for each <r> element
   *   - Each run may have optional formatting from <rPr>
   *   - Runs without <rPr> use default formatting
   *   - Returns error if any <r> is missing required <t> element
   * DETERMINISTIC: Yes (stable iteration order)
   *
   * OOXML structure:
   *   - <r><t>text</t></r> → unformatted run
   *   - <r><rPr>...</rPr><t>text</t></r> → formatted run
   *
   * @param runElems
   *   Sequence of <r> elements
   * @return
   *   Either[String, RichText] with error if any run is malformed
   */
  def parseTextRuns(runElems: Seq[Node]): Either[String, com.tjclp.xl.richtext.RichText] =
    import com.tjclp.xl.richtext.{TextRun, RichText}

    val runs = runElems.collect { case e: Elem if e.label == "r" => e }.map { rElem =>
      // Extract optional <rPr> for formatting and preserve raw XML
      val rPrElemOpt = (rElem \ "rPr").headOption.collect { case elem: Elem => elem }
      val font = rPrElemOpt.map(parseRunProperties)
      val rawRPrXml = rPrElemOpt.map(elem => compact(elem)) // Preserve as XML string

      // Extract required <t> text (preserving whitespace)
      (rElem \ "t").headOption
        .collect { case elem: Elem => elem }
        .map(getTextPreservingWhitespace) match
        case Some(text) => Right(TextRun(text, font, rawRPrXml))
        case None => Left("Text run <r> missing <t> element")
    }

    val errors = runs.collect { case Left(err) => err }
    if errors.nonEmpty then Left(s"TextRun parse errors: ${errors.mkString(", ")}")
    else
      val textRuns = runs.collect { case Right(run) => run }.toVector
      Right(RichText(textRuns))

object XmlSecurity:
  /** Shared XXE-safe XML parser. */
  def parseSafe(xmlString: String, location: String): XLResult[Elem] =
    try
      val factory = javax.xml.parsers.SAXParserFactory.newInstance()
      factory.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true)
      factory.setFeature("http://xml.org/sax/features/external-general-entities", false)
      factory.setFeature("http://xml.org/sax/features/external-parameter-entities", false)
      factory.setFeature("http://apache.org/xml/features/nonvalidating/load-external-dtd", false)
      factory.setXIncludeAware(false)
      factory.setNamespaceAware(true)

      val loader = XML.withSAXParser(factory.newSAXParser())
      Right(loader.loadString(xmlString))
    catch
      case e: Exception => Left(XLError.ParseError(location, s"XML parse error: ${e.getMessage}"))
