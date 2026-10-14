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
 * Only `<c t="s">` cells reference the table: ECMA-376 Part 1, §18.3.1.4 (`c`, `t` attribute, value
 * `s`) and §18.4.9 (`sst`).
 */
private[ooxml] object SharedStringCompaction:

  /**
   * The shared-string references of one sheet part.
   *
   * @param indices
   *   the table positions its `t="s"` cells name
   * @param count
   *   the number of `t="s"` cells (the table's `count` attribute counts references)
   */
  final case class RefScan(indices: BitSet, count: Int) derives CanEqual:
    /** The highest position referenced, or -1 when the part references none. */
    def maxIndex: Int = indices.lastOption.getOrElse(-1)

  object RefScan:
    val empty: RefScan = RefScan(BitSet.empty, 0)

  /** Sheet-class parts whose cells can reference the table (worksheets and macro sheets). */
  def isCellBearingPart(path: String): Boolean =
    (path.startsWith("xl/worksheets/") || path.startsWith("xl/macrosheets/")) &&
      !path.contains("/_rels/") && path.endsWith(".xml")

  /**
   * SAX scan of a sheet part for its `t="s"` cells. `Left` when the part is not well-formed XML;
   * the caller then keeps every entry (an unread part may reference any of them).
   */
  def scan(in: InputStream): Either[String, RefScan] =
    try
      val parser = XmlSecurity.secureSaxParserFactory().newSAXParser()
      val handler = new RefHandler
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
  private final class RefHandler extends DefaultHandler:
    private val indices = mutable.BitSet.empty
    private var count = 0
    private var inSharedCell = false
    private var inValue = false
    private val value = new java.lang.StringBuilder

    def result(): RefScan = RefScan(indices.toImmutable, count)

    private def label(localName: String, qName: String): String =
      if localName.nonEmpty then localName else qName.substring(qName.indexOf(':') + 1)

    override def startElement(
      uri: String,
      localName: String,
      qName: String,
      atts: Attributes
    ): Unit =
      label(localName, qName) match
        case "c" =>
          inSharedCell = atts.getValue("t") == "s"
          if inSharedCell then count += 1
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
          value.toString.trim.toIntOption.filter(_ >= 0).foreach(indices += _)
        case "c" => inSharedCell = false
        case _ => ()
