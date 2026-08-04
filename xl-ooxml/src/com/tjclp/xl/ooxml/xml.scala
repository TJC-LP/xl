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
  // DrawingML namespaces (GH-221)
  val nsSpreadsheetDrawing =
    "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  val nsDrawingMain = "http://schemas.openxmlformats.org/drawingml/2006/main"
  // Chart namespace (GH-222)
  val nsChart = "http://schemas.openxmlformats.org/drawingml/2006/chart"

  /** Relationship type URIs */
  val relTypeOfficeDocument =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
  val relTypeWorksheet =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
  val relTypeStyles = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
  val relTypeSharedStrings =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
  val relTypeComments =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
  val relTypeVmlDrawing =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"
  val relTypeTable =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"
  val relTypeHyperlink =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
  val relTypeDrawing =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
  val relTypeImage =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
  val relTypeChart =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
  val relTypeTheme =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
  // docProps relationships (GH-242): core props is a PACKAGE relationship type, app is officeDocument
  val relTypeCoreProperties =
    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
  val relTypeExtendedProperties =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
  // Workbook-level reference targets checked by the structural lint (GH-397)
  val relTypeExternalLink =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink"
  // The externalLink part's own <externalBook r:id> target (GH-413)
  val relTypeExternalLinkPath =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath"
  val relTypeChartsheet =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet"
  val relTypeDialogsheet =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet"
  val relTypePivotCacheDefinition =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition"

  /** Content type URIs */
  val ctWorkbook = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
  val ctWorksheet = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
  val ctStyles = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
  val ctSharedStrings =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
  val ctComments = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"
  val ctVmlDrawing = "application/vnd.openxmlformats-officedocument.vmlDrawing"
  val ctTable = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"
  val ctDrawing = "application/vnd.openxmlformats-officedocument.drawing+xml"
  val ctChart = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
  val ctTheme = "application/vnd.openxmlformats-officedocument.theme+xml"
  val ctRelationships = "application/vnd.openxmlformats-package.relationships+xml"
  val ctCoreProperties = "application/vnd.openxmlformats-package.core-properties+xml"
  val ctExtendedProperties =
    "application/vnd.openxmlformats-officedocument.extended-properties+xml"

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
    // Use toString instead of PrettyPrinter to preserve all whitespace (including newlines)
    // PrettyPrinter normalizes whitespace in text nodes, which corrupts comments with newlines
    s"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${node.toString}"""

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
    s.nonEmpty && (s.startsWith(" ") || s.endsWith(" ") || s.contains("  ") ||
      s.contains("\n") || s.contains("\r") || s.contains("\t"))

  /**
   * True if `c` is a valid XML 1.0 character. Tab/LF/CR are allowed; the other C0 control chars
   * (U+0000–08, 0B, 0C, 0E–1F) are forbidden. Everything >= U+0020 is kept — including surrogate
   * code units, so astral characters (emoji) encoded as surrogate pairs survive intact.
   */
  def isLegalXmlChar(c: Char): Boolean =
    c == '\t' || c == '\n' || c == '\r' || c >= ' '

  /**
   * Strip XML 1.0-illegal control characters from text content (GH-237). Such characters occur in
   * DB/CSV-imported data; left raw, the StAX backend silently DROPS them while the ScalaXml backend
   * emits them raw — divergent, and a determinism/validity breach. Applying this identically on
   * both write paths makes output backend-independent. Fast path returns the input unchanged.
   */
  def sanitizeXmlText(s: String): String =
    if s.forall(isLegalXmlChar) then s else s.filter(isLegalXmlChar)

  /** True when `s` contains a literal `_xHHHH_` pattern (lowercase `x`, 4 hex digits) at `i`. */
  private def isXstringEscapeAt(s: String, i: Int): Boolean =
    def hex(c: Char): Boolean =
      (c >= '0' && c <= '9') || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F')
    i + 6 < s.length && s.charAt(i) == '_' && s.charAt(i + 1) == 'x' &&
    hex(s.charAt(i + 2)) && hex(s.charAt(i + 3)) && hex(s.charAt(i + 4)) &&
    hex(s.charAt(i + 5)) && s.charAt(i + 6) == '_'

  /**
   * ECMA-376 ST_Xstring escape (Part 1, §22.9.2.19) for `<t>`/`<v>` text content (GH-288).
   *
   * XML 1.0 parsers normalize raw CR and CRLF in element content to LF, so a raw `\r` cannot
   * round-trip; Excel stores it as the `_x000D_` escape. Text that LITERALLY contains an `_xHHHH_`
   * pattern must protect its leading underscore as `_x005F_` so decoding is unambiguous.
   *
   * Only CR needs escaping for fidelity: TAB and LF survive element content unchanged, and
   * XML-illegal control characters are deliberately STRIPPED by [[sanitizeXmlText]] (GH-237), not
   * escaped. Fast path returns the input unchanged.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  def escapeXstring(s: String): String =
    var needs = false
    var i = 0
    while !needs && i < s.length do
      val c = s.charAt(i)
      if c == '\r' || (c == '_' && isXstringEscapeAt(s, i)) then needs = true
      i += 1
    if !needs then s
    else
      val sb = new java.lang.StringBuilder(s.length + 16)
      var j = 0
      while j < s.length do
        val c = s.charAt(j)
        if c == '\r' then sb.append("_x000D_")
        else if c == '_' && isXstringEscapeAt(s, j) then sb.append("_x005F_")
        else sb.append(c)
        j += 1
      sb.toString

  /**
   * ST_Xstring escape for ATTRIBUTE values (GH-429): prompt/error message attributes on
   * `<dataValidation>`.
   *
   * Attribute-value normalization (XML 1.0 §3.3.3) turns raw TAB/LF/CR into spaces on re-parse — a
   * raw newline in an attr value survives serialization but comes back as a space — so all three
   * must be escaped for fidelity, not just CR as in element content ([[escapeXstring]]). Literal
   * `_xHHHH_` patterns protect their leading underscore as `_x005F_`.
   *
   * Law: `decodeXstring(escapeXstringAttr(s)) == s`.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  def escapeXstringAttr(s: String): String =
    var needs = false
    var i = 0
    while !needs && i < s.length do
      val c = s.charAt(i)
      if c == '\r' || c == '\n' || c == '\t' || (c == '_' && isXstringEscapeAt(s, i)) then
        needs = true
      i += 1
    if !needs then s
    else
      val sb = new java.lang.StringBuilder(s.length + 16)
      var j = 0
      while j < s.length do
        val c = s.charAt(j)
        if c == '\r' then sb.append("_x000D_")
        else if c == '\n' then sb.append("_x000A_")
        else if c == '\t' then sb.append("_x0009_")
        else if c == '_' && isXstringEscapeAt(s, j) then sb.append("_x005F_")
        else sb.append(c)
        j += 1
      sb.toString

  /**
   * Decode ECMA-376 `_xHHHH_` escapes (Part 1, §22.9.2.19) in `<t>`/`<v>` text content (GH-288).
   *
   * Single left-to-right pass: `_x000D_` → CR, `_x005F_` → `_` (so `_x005F_x000D_` decodes to the
   * literal text `_x000D_`). Hex digits are accepted in either case; the `x` must be lowercase per
   * the spec. Fast path returns the input unchanged.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  def decodeXstring(s: String): String =
    if s.indexOf("_x") < 0 then s
    else
      val sb = new java.lang.StringBuilder(s.length)
      var i = 0
      while i < s.length do
        if isXstringEscapeAt(s, i) then
          sb.append(Integer.parseInt(s.substring(i + 2, i + 6), 16).toChar)
          i += 7
        else
          sb.append(s.charAt(i))
          i += 1
      sb.toString

  /**
   * Format a BigDecimal for an OOXML `<v>` element as a plain decimal, never scientific notation
   * (GH-238). Excel never writes `1.0E+10` in `<v>` and stricter consumers reject it.
   */
  def plainNumber(n: BigDecimal): String = n.bigDecimal.toPlainString

  /**
   * Plain-decimal rendering of a Double serial (date/time). Routes through the Double's shortest
   * round-trip string so normal serials are byte-identical to the previous `toString` output while
   * scientific notation (if it ever arose) is expanded.
   */
  def plainNumber(d: Double): String = BigDecimal(d.toString).bigDecimal.toPlainString

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
    // Optimization: Early exit if scope already TopScope and all child Elems too (2-3% speedup)
    if elem.scope == TopScope && elem.child.forall {
        case e: Elem => e.scope == TopScope
        case _ => true
      }
    then elem
    else
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
   *   - <rFont val="Arial"/> → font family (CT_RPrElt spelling; legacy <name val=…/> read as
   *     fallback, GH-383)
   *
   * @param rPrElem
   *   The <rPr> element to parse
   * @return
   *   Font with formatting properties (default Font if no properties)
   */
  def parseRunProperties(rPrElem: Elem): com.tjclp.xl.styles.font.Font =
    import com.tjclp.xl.styles.font.{Font, Underline}
    import com.tjclp.xl.styles.color.Color

    val bold = (rPrElem \ "b").nonEmpty
    val italic = (rPrElem \ "i").nonEmpty
    // u@val per ST_UnderlineValues (GH-423): bare <u/> means single; unknown tokens read
    // lenient as Single (the pre-typed truthy behavior for any present <u>)
    val underline = (rPrElem \ "u").headOption match
      case Some(u) =>
        u.attribute("val")
          .map(attr => Underline.fromToken(attr.text).getOrElse(Underline.Single))
          .getOrElse(Underline.Single)
      case None => Underline.None

    val color =
      (rPrElem \ "color").headOption.collect { case elem: Elem => elem }.flatMap { colorElem =>
        getAttrOpt(colorElem, "rgb").flatMap { rgb =>
          // Add # prefix for Color.fromHex
          Color.fromHex(s"#$rgb").toOption
        }
      }

    // Font requires sizePt > 0 and a nonEmpty name; filter malformed values so
    // rich-text parsing stays total instead of throwing through require (GH-278)
    val sizePt = (rPrElem \ "sz").headOption
      .collect { case elem: Elem => elem }
      .flatMap(e => getAttrOpt(e, "val"))
      .flatMap(_.toDoubleOption)
      .filter(_ > 0)
      .getOrElse(11.0)

    // CT_RPrElt spells the font element <rFont val=…/> (real Excel files use it
    // exclusively); older xl versions wrote the CT_Font spelling <name val=…/>,
    // which stays readable as a fallback (GH-383).
    def fontName(label: String): Option[String] =
      (rPrElem \ label).headOption
        .collect { case elem: Elem => elem }
        .flatMap(e => getAttrOpt(e, "val"))
        .filter(_.nonEmpty)

    val name = fontName("rFont").orElse(fontName("name")).getOrElse("Calibri")

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

      // Extract required <t> text (preserving whitespace, decoding _xHHHH_ escapes — GH-288)
      (rElem \ "t").headOption
        .collect { case elem: Elem => elem }
        .map(e => decodeXstring(getTextPreservingWhitespace(e))) match
        case Some(text) => Right(TextRun(text, font, rawRPrXml))
        case None => Left("Text run <r> missing <t> element")
    }

    val errors = runs.collect { case Left(err) => err }
    if errors.nonEmpty then Left(s"TextRun parse errors: ${errors.mkString(", ")}")
    else
      val textRuns = runs.collect { case Right(run) => run }.toVector
      Right(RichText(textRuns))

object XmlSecurity:
  /**
   * JAXP entity-SIZE limit properties lifted (set to 0 = unlimited) on every parser built by
   * [[secureSaxParserFactory]] (GH-457): real workbooks accumulate tens of thousands of
   * definedNames (multi-MB single-line workbook.xml), and JAXP builds that default these limits to
   * 100KB — the GraalVM the native image is built from, where `-Djdk.xml.*` cannot reach the binary
   * — reject the document entity with JAXP00010003 on every verb. Lifting SIZE limits is safe
   * because doctype declarations are rejected outright (`disallow-doctype-decl`), so no general
   * entity can ever be defined and entity-expansion attacks remain structurally impossible;
   * `entityExpansionLimit` is deliberately left at its default. Each limit lists its legacy Oracle
   * URI first (recognized on every mainstream JDK), then the modern `jdk.xml.` name.
   */
  private val liftedEntitySizeLimits: List[List[String]] = List(
    List(
      "http://www.oracle.com/xml/jaxp/properties/maxGeneralEntitySizeLimit",
      "jdk.xml.maxGeneralEntitySizeLimit"
    ),
    List(
      "http://www.oracle.com/xml/jaxp/properties/totalEntitySizeLimit",
      "jdk.xml.totalEntitySizeLimit"
    )
  )

  /** Set each lifted limit under the first property name the parser recognizes (GH-457). */
  private def liftEntitySizeLimits(parser: javax.xml.parsers.SAXParser): Unit =
    liftedEntitySizeLimits.foreach { aliases =>
      val _ = aliases.exists { name =>
        try
          parser.setProperty(name, "0")
          true
        catch
          // Tolerate exotic parsers that don't know the JAXP limit properties: hardening
          // features above still apply; only the limit-lifting is best-effort.
          case _: org.xml.sax.SAXNotRecognizedException => false
          case _: org.xml.sax.SAXNotSupportedException => false
      }
    }

  /**
   * Delegating factory so every call site building parsers via
   * `secureSaxParserFactory().newSAXParser()` inherits the GH-457 limit-lifting automatically — the
   * JAXP limits are parser properties, not factory features, so they cannot be carried by the
   * underlying factory itself.
   */
  private final class SecureFactory(underlying: javax.xml.parsers.SAXParserFactory)
      extends javax.xml.parsers.SAXParserFactory:
    override def newSAXParser(): javax.xml.parsers.SAXParser =
      val parser = underlying.newSAXParser()
      liftEntitySizeLimits(parser)
      parser
    override def setFeature(name: String, value: Boolean): Unit = underlying.setFeature(name, value)
    override def getFeature(name: String): Boolean = underlying.getFeature(name)
    override def setNamespaceAware(awareness: Boolean): Unit =
      underlying.setNamespaceAware(awareness)
    override def isNamespaceAware(): Boolean = underlying.isNamespaceAware()
    override def setValidating(validating: Boolean): Unit = underlying.setValidating(validating)
    override def isValidating(): Boolean = underlying.isValidating()
    override def setXIncludeAware(state: Boolean): Unit = underlying.setXIncludeAware(state)
    override def isXIncludeAware(): Boolean = underlying.isXIncludeAware()
    override def setSchema(schema: javax.xml.validation.Schema): Unit =
      underlying.setSchema(schema)
    override def getSchema(): javax.xml.validation.Schema = underlying.getSchema()

  /**
   * SAXParserFactory carrying the XXE hardening posture shared by EVERY parse surface: doctype
   * declarations rejected, external general/parameter entities and external DTD loading disabled,
   * XInclude off, namespaces on. Streaming SAX sites (worksheet/SST/dimension readers, streaming
   * transforms) must build their parsers from this factory so the hardening cannot drift per-site
   * (GH-350) — pair it with [[stripLeadingDoctypeStream]] on the input so benign doctypes still
   * read. Parsers it creates additionally have the JAXP entity-SIZE limits lifted so name-bloated
   * multi-MB core parts read on the native image (GH-457, see [[liftedEntitySizeLimits]]).
   */
  def secureSaxParserFactory(): javax.xml.parsers.SAXParserFactory =
    val factory = javax.xml.parsers.SAXParserFactory.newInstance()
    factory.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true)
    factory.setFeature("http://xml.org/sax/features/external-general-entities", false)
    factory.setFeature("http://xml.org/sax/features/external-parameter-entities", false)
    factory.setFeature("http://apache.org/xml/features/nonvalidating/load-external-dtd", false)
    factory.setXIncludeAware(false)
    factory.setNamespaceAware(true)
    new SecureFactory(factory)

  /**
   * Thread-local pool of XXE-safe SAX parsers for performance.
   *
   * Optimization: Parser creation is expensive (factory + security features setup). Reuse parsers
   * per thread to avoid 10k+ instantiations during RichText parsing in hot READ path (5-8%
   * speedup).
   */
  private val parserPool: ThreadLocal[javax.xml.parsers.SAXParser] =
    ThreadLocal.withInitial(() => secureSaxParserFactory().newSAXParser())

  /**
   * Strip a benign leading <!DOCTYPE ...> declaration from the XML prolog (GH-350).
   *
   * Some third-party producers emit a DOCTYPE before the root element; the pooled parser hard-fails
   * on ANY doctype (`disallow-doctype-decl=true`). Rather than weakening that security posture,
   * remove the declaration before the parser ever sees it — the parser keeps
   * doctype/external-entity features disabled (defense in depth), and entity definitions in a
   * stripped internal subset are simply never honored (any reference to them fails as an undeclared
   * entity), which is safe for OOXML parts since they never rely on DTD entities.
   *
   * A leading UTF-8 BOM (which survives the bytes->String decode as U+FEFF and which Xerces REJECTS
   * in a character stream) is dropped first — third-party producers that emit DOCTYPEs are exactly
   * the ones that emit BOMs, and a BOM char is meaningless once already decoded.
   *
   * The scanner only walks the prolog (whitespace, the XML declaration / processing instructions,
   * comments); on the first non-prolog construct it stops, so element content can never be eaten.
   * Within the DOCTYPE it tracks quotes and `[`/`]` internal-subset nesting and skips
   * `<!-- ... -->` / `<? ... ?>` spans wholesale (a `]`, `>`, or quote inside a subset comment must
   * not confuse the state machine), handling both the `<!DOCTYPE x SYSTEM "...">` and
   * `<!DOCTYPE x [ ... ]>` forms. An unterminated DOCTYPE is left untouched for the parser to
   * report.
   */
  private[ooxml] def stripLeadingDoctype(xml: String): String =
    val body = if xml.nonEmpty && xml.charAt(0) == '\uFEFF' then xml.substring(1) else xml
    doctypeSpan(body) match
      case Some((start, end)) =>
        // Keep the stripped region's newlines so reported line numbers still match the source
        body.substring(0, start) + body.substring(start, end).filter(_ == '\n') +
          body.substring(end)
      case None => body

  /**
   * Locate the `[start, end)` span of a leading DOCTYPE declaration (see [[stripLeadingDoctype]]).
   */
  private def doctypeSpan(xml: String): Option[(Int, Int)] =
    @annotation.tailrec
    def doctypeEnd(i: Int, depth: Int, quote: Option[Char]): Option[Int] =
      if i >= xml.length then None
      else
        val c = xml.charAt(i)
        quote match
          case Some(q) => doctypeEnd(i + 1, depth, if c == q then None else quote)
          case None =>
            // Comments and PIs are legal in the internal subset; skip them wholesale so their
            // content (']', '>', quotes) can't corrupt the depth/quote state (GH-350).
            if xml.startsWith("<!--", i) then
              val end = xml.indexOf("-->", i + 4)
              if end < 0 then None else doctypeEnd(end + 3, depth, None)
            else if xml.startsWith("<?", i) then
              val end = xml.indexOf("?>", i + 2)
              if end < 0 then None else doctypeEnd(end + 2, depth, None)
            else
              c match
                case '"' | '\'' => doctypeEnd(i + 1, depth, Some(c))
                case '[' => doctypeEnd(i + 1, depth + 1, None)
                case ']' => doctypeEnd(i + 1, depth - 1, None)
                case '>' if depth == 0 => Some(i + 1)
                case _ => doctypeEnd(i + 1, depth, None)

    @annotation.tailrec
    def prologScan(i: Int): Option[(Int, Int)] =
      if i >= xml.length then None
      else if xml.charAt(i).isWhitespace then prologScan(i + 1)
      else if xml.startsWith("<?", i) then
        val end = xml.indexOf("?>", i + 2)
        if end < 0 then None else prologScan(end + 2)
      else if xml.startsWith("<!--", i) then
        val end = xml.indexOf("-->", i + 4)
        if end < 0 then None else prologScan(end + 3)
      else if xml.startsWith("<!DOCTYPE", i) then doctypeEnd(i + 9, 0, None).map(end => (i, end))
      else None // first element (or malformed input) — nothing to strip

    prologScan(0)

  /** Upper bound on the prolog + DOCTYPE region [[stripLeadingDoctypeStream]] will buffer/scan. */
  private val MaxPrologScanBytes = 16384

  /**
   * Stream-level [[stripLeadingDoctype]] for SAX paths that parse InputStreams directly (GH-350):
   * the streaming worksheet/SST/dimension readers and streaming transforms build their parsers from
   * [[secureSaxParserFactory]] (doctype disallowed), so a benign leading DOCTYPE must be removed
   * before the parser sees the bytes — same tolerance as the in-memory path.
   *
   * Buffers at most [[MaxPrologScanBytes]] from the head, removes a leading DOCTYPE (newlines kept
   * so reported line numbers still match), and returns head + untouched remainder as one stream.
   * The scan runs on an ISO-8859-1 view of the head (1 char = 1 byte, so span indexes map straight
   * back to byte offsets; UTF-8 continuation bytes can never alias the ASCII delimiters the scanner
   * tracks). A UTF-8 BOM is skipped for scanning but kept in the output. If the scan is
   * inconclusive within the buffer (no doctype, truncated doctype, non-UTF-8-family encoding), the
   * input passes through unchanged for the parser to report.
   */
  def stripLeadingDoctypeStream(in: java.io.InputStream): java.io.InputStream =
    val head = in.readNBytes(MaxPrologScanBytes)
    val bomLen =
      if head.length >= 3 && head(0) == 0xef.toByte && head(1) == 0xbb.toByte &&
        head(2) == 0xbf.toByte
      then 3
      else 0
    val headStr =
      new String(head, bomLen, head.length - bomLen, java.nio.charset.StandardCharsets.ISO_8859_1)
    val cleaned = doctypeSpan(headStr) match
      case Some((start, end)) =>
        val out = new java.io.ByteArrayOutputStream(head.length)
        out.write(head, 0, bomLen + start)
        (start until end).foreach(i => if headStr.charAt(i) == '\n' then out.write('\n'))
        out.write(head, bomLen + end, head.length - bomLen - end)
        out.toByteArray
      case None => head
    new java.io.SequenceInputStream(new java.io.ByteArrayInputStream(cleaned), in)

  /** Shared XXE-safe XML parser with pooling optimization. */
  def parseSafe(xmlString: String, location: String): XLResult[Elem] =
    try
      val loader = XML.withSAXParser(parserPool.get())
      Right(loader.loadString(stripLeadingDoctype(xmlString)))
    catch
      // GH-350: name the offending construct's position so a failed core part is diagnosable
      case e: org.xml.sax.SAXParseException =>
        Left(
          XLError.ParseError(
            location,
            s"XML parse error at line ${e.getLineNumber}, column ${e.getColumnNumber}: ${e.getMessage}"
          )
        )
      case e: Exception => Left(XLError.ParseError(location, s"XML parse error: ${e.getMessage}"))
