package com.tjclp.xl.ooxml

import java.io.InputStream

import org.xml.sax.{Attributes, InputSource, SAXException}
import org.xml.sax.helpers.DefaultHandler

/**
 * Whether a sheet part that is copied byte for byte names a shared-string position the source table
 * does not have (GH-781).
 *
 * A `t="s"` cell value is a zero-based index into the `<si>` items of `xl/sharedStrings.xml`
 * (ECMA-376 Part 1, §18.4.9 `sst`; §18.3.1.4 `c`, `t` attribute value `s`). An index at or past the
 * table's size names no entry, and the reader returns `#REF!` for it. Appending a string to the
 * table gives that position a value, so a copied part holding such an index would read the appended
 * string after the save. The writer therefore does not append when any copied part holds one (see
 * `XlsxWriter`'s shared-string plan).
 *
 * Cells that reference the table: a worksheet or macro sheet's `<c>`, and a revision log's
 * `<nc>`/`<oc>` new and old cell values (§18.11, both of type CT_Cell).
 *
 * The scan stops at the first such index and keeps at most [[MaxValueChars]] characters of one
 * value, so its memory does not depend on the part's size or on the index values in it.
 */
private[ooxml] object SharedStringBound:

  /**
   * Longest digit string kept from one `<v>`. A value with more digits is past any table: the
   * position is at most `Int.MaxValue` (10 digits), and the reader parses it as an `Int`.
   */
  val MaxValueChars: Int = 20

  /** Parts whose cells can reference the table: worksheets, macro sheets and revision logs. */
  def isCellBearingPart(path: String): Boolean =
    (path.startsWith("xl/worksheets/") || path.startsWith("xl/macrosheets/") ||
      path.startsWith("xl/revisions/revisionLog")) &&
      !path.contains("/_rels/") && path.endsWith(".xml")

  /**
   * True when the part holds a `t="s"` value that is a non-negative integer at or past `tableSize`,
   * or when the part is not readable XML (it may hold any index). Surrounding whitespace in the
   * value is ignored, so a reader that trims it is covered too. A value that is not an integer
   * names no position before or after an append, so it does not count.
   */
  def referencesPast(in: InputStream, tableSize: Int): Boolean =
    val handler = new BoundHandler(tableSize)
    try
      val parser = XmlSecurity.secureSaxParserFactory().newSAXParser()
      parser.parse(InputSource(XmlSecurity.stripLeadingDoctypeStream(in)), handler)
      false
    catch
      case Found => true
      case _: SAXException => true
      case _: java.io.IOException => true

  /** Thrown by the handler to stop the parse at the first index past the table. */
  private object Found extends SAXException("shared-string index past the table"):
    override def fillInStackTrace(): Throwable = this

  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private final class BoundHandler(tableSize: Int) extends DefaultHandler:
    private var inSharedCell = false
    private var inValue = false
    private var tooLong = false
    private val digits = new java.lang.StringBuilder

    private def label(localName: String, qName: String): String =
      if localName.nonEmpty then localName else qName.substring(qName.indexOf(':') + 1)

    override def startElement(
      uri: String,
      localName: String,
      qName: String,
      atts: Attributes
    ): Unit =
      label(localName, qName) match
        case "c" | "nc" | "oc" => inSharedCell = atts.getValue("t") == "s"
        case "v" if inSharedCell =>
          inValue = true
          tooLong = false
          digits.setLength(0)
        case _ => ()

    override def characters(ch: Array[Char], start: Int, length: Int): Unit =
      if inValue then
        var i = start
        while i < start + length do
          val c = ch(i)
          if !Character.isWhitespace(c) then
            if digits.length < MaxValueChars then digits.append(c) else tooLong = true
          i += 1

    override def endElement(uri: String, localName: String, qName: String): Unit =
      label(localName, qName) match
        case "v" if inValue =>
          inValue = false
          val text = digits.toString
          val allDigits = text.nonEmpty && text.forall(c => c >= '0' && c <= '9')
          if allDigits && (tooLong || text.toLongOption.exists(_ >= tableSize)) then throw Found
        case "c" | "nc" | "oc" => inSharedCell = false
        case _ => ()
