package com.tjclp.xl.ooxml

import java.io.InputStream

import scala.collection.immutable.BitSet
import scala.collection.mutable

import org.xml.sax.{Attributes, InputSource, SAXException}
import org.xml.sax.helpers.DefaultHandler

/**
 * Removing shared strings no cell references (GH-768).
 *
 * A `t="s"` cell stores a position in `xl/sharedStrings.xml`, so an entry can only leave the table
 * without disturbing a sheet part when no part that keeps its bytes references a later position.
 * The writer therefore:
 *   - drops an unreferenced entry when every part referencing a later position is regenerated on
 *     this write (the regenerated parts look their strings up again);
 *   - otherwise replaces it with an empty entry (`<si><t/></si>`), so the copied parts keep their
 *     positions and the removed text still leaves the package.
 *
 * Only cells with `t="s"` reference the table: a worksheet or macro sheet's `<c>` (ECMA-376 Part 1,
 * §18.3.1.4, `t` attribute value `s`; §18.4.9 `sst`) and a revision log's `<nc>`/`<oc>` new and old
 * cell values (§18.11, both of type CT_Cell).
 *
 * A `t="s"` value is a zero-based index into the table's `<si>` items (§18.4.9 `sst`), so an index
 * at or past the table's size, or one that is not a non-negative integer, names no entry. Such a
 * reference keeps nothing, so it never enters [[RefScan.indices]]; it is counted in
 * [[RefScan.outOfRange]] instead. A copied part with such a reference is malformed, so the write
 * keeps every entry of the table, as it does for a part it cannot read.
 */
private[ooxml] object SharedStringCompaction:

  /**
   * The shared-string references of one sheet part.
   *
   * @param indices
   *   the table positions its `t="s"` cells name, each below the table's size
   * @param count
   *   the number of `t="s"` sheet cells, `<c>` only (the table's `count` attribute counts them)
   * @param outOfRange
   *   the number of `t="s"` cell values (`<c>`, `<nc>`, `<oc>`) that name no entry: an index at or
   *   past the table's size, or a value that is not a non-negative integer
   */
  final case class RefScan(indices: BitSet, count: Int, outOfRange: Int) derives CanEqual:
    /** The highest position referenced, or -1 when the part references none. */
    def maxIndex: Int = indices.lastOption.getOrElse(-1)

  object RefScan:
    val empty: RefScan = RefScan(BitSet.empty, 0, 0)

  /** Parts whose cells can reference the table: worksheets, macro sheets and revision logs. */
  def isCellBearingPart(path: String): Boolean =
    (path.startsWith("xl/worksheets/") || path.startsWith("xl/macrosheets/") ||
      path.startsWith("xl/revisions/revisionLog")) &&
      !path.contains("/_rels/") && path.endsWith(".xml")

  /**
   * SAX scan of a sheet part for its `t="s"` cells, against a table of `tableSize` entries. Only
   * indices below `tableSize` enter [[RefScan.indices]], so the scan's memory is bounded by the
   * table, not by the values in the part. `Left` when the part is not well-formed XML; the caller
   * then keeps every entry (an unread part may reference any of them).
   */
  def scan(in: InputStream, tableSize: Int): Either[String, RefScan] =
    try
      val parser = XmlSecurity.secureSaxParserFactory().newSAXParser()
      val handler = new RefHandler(tableSize)
      parser.parse(InputSource(XmlSecurity.stripLeadingDoctypeStream(in)), handler)
      Right(handler.result())
    catch
      case e: SAXException => Left(s"Malformed XML: ${e.getMessage}")
      case e: java.io.IOException => Left(s"Read failed: ${e.getMessage}")

  /**
   * The table after removing `orphans`: an orphan at or below `pinnedMax` (the highest position a
   * copied part references) becomes an empty plain entry, one above it is dropped. Every other
   * entry keeps its order, so positions at or below `pinnedMax` are unchanged.
   */
  def withoutOrphans(
    entries: Vector[SSTEntry],
    orphans: BitSet,
    pinnedMax: Int
  ): Vector[SSTEntry] =
    entries.iterator.zipWithIndex.flatMap { (entry, idx) =>
      if !orphans.contains(idx) then Some(entry)
      else if idx <= pinnedMax then Some(Left(""): SSTEntry)
      else None
    }.toVector

  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  private final class RefHandler(tableSize: Int) extends DefaultHandler:
    private val indices = mutable.BitSet.empty
    private var count = 0
    private var outOfRange = 0
    private var inSharedCell = false
    private var inValue = false
    private val value = new java.lang.StringBuilder

    def result(): RefScan = RefScan(indices.toImmutable, count, outOfRange)

    private def label(localName: String, qName: String): String =
      if localName.nonEmpty then localName else qName.substring(qName.indexOf(':') + 1)

    override def startElement(
      uri: String,
      localName: String,
      qName: String,
      atts: Attributes
    ): Unit =
      label(localName, qName) match
        case name @ ("c" | "nc" | "oc") =>
          inSharedCell = atts.getValue("t") == "s"
          // the table's count attribute counts sheet cells; a revision's cell value is history
          if inSharedCell && name == "c" then count += 1
        case "v" if inSharedCell =>
          inValue = true
          value.setLength(0)
        case _ => ()

    override def characters(ch: Array[Char], start: Int, length: Int): Unit =
      if inValue then value.append(ch, start, length)

    override def endElement(uri: String, localName: String, qName: String): Unit =
      label(localName, qName) match
        case "v" if inValue =>
          inValue = false
          value.toString.trim.toIntOption.filter(i => i >= 0 && i < tableSize) match
            case Some(idx) => indices += idx
            case None => outOfRange += 1
        case "c" | "nc" | "oc" => inSharedCell = false
        case _ => ()
