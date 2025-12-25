package com.tjclp.xl.ooxml.worksheet

import scala.xml.Elem

import com.tjclp.xl.ooxml.SaxWriter
import com.tjclp.xl.ooxml.XmlUtil.elemOrdered

/**
 * Row in worksheet with full attribute preservation for surgical modification.
 *
 * Preserves all OOXML row attributes to maintain byte-level fidelity during regeneration.
 */
case class OoxmlRow(
  rowIndex: Int, // 1-based
  cells: Seq[OoxmlCell],
  // Row-level attributes (all optional)
  spans: Option[String] = None, // "2:16" (cell coverage optimization hint)
  style: Option[Int] = None, // s="7" (row-level style ID)
  height: Option[Double] = None, // ht="24.95" (custom row height in points)
  customHeight: Boolean = false, // customHeight="1"
  customFormat: Boolean = false, // customFormat="1"
  hidden: Boolean = false, // hidden="1"
  outlineLevel: Option[Int] = None, // outlineLevel="1" (grouping level)
  collapsed: Boolean = false, // collapsed="1" (outline collapsed)
  thickBot: Boolean = false, // thickBot="1" (thick bottom border)
  thickTop: Boolean = false, // thickTop="1" (thick top border)
  dyDescent: Option[Double] = None // x14ac:dyDescent="0.25" (font descent adjustment)
):
  def writeSax(writer: SaxWriter): Unit =
    writer.startElement("row")

    val attrs = Seq.newBuilder[(String, String)]
    attrs += ("r" -> rowIndex.toString)
    spans.foreach(s => attrs += ("spans" -> s))
    style.foreach(s => attrs += ("s" -> s.toString))
    if customFormat then attrs += ("customFormat" -> "1")
    height.foreach(h => attrs += ("ht" -> h.toString))
    if customHeight then attrs += ("customHeight" -> "1")
    if hidden then attrs += ("hidden" -> "1")
    outlineLevel.foreach(l => attrs += ("outlineLevel" -> l.toString))
    if collapsed then attrs += ("collapsed" -> "1")
    if thickBot then attrs += ("thickBot" -> "1")
    if thickTop then attrs += ("thickTop" -> "1")
    dyDescent.foreach(d => attrs += ("x14ac:dyDescent" -> d.toString))

    SaxWriter.withAttributes(writer, attrs.result()*) {
      cells.sortBy(_.ref.col.index0).foreach(_.writeSax(writer))
    }

    writer.endElement() // row

  def toXml: Elem =
    // Excel expects attributes in specific order (not alphabetical!)
    // Order: r, spans, s, customFormat, ht, customHeight, hidden, outlineLevel, collapsed, thickBot, thickTop, x14ac:dyDescent
    val attrs = Seq.newBuilder[(String, String)]

    attrs += ("r" -> rowIndex.toString)
    spans.foreach(s => attrs += ("spans" -> s))
    style.foreach(s => attrs += ("s" -> s.toString))
    if customFormat then attrs += ("customFormat" -> "1")
    height.foreach(h => attrs += ("ht" -> h.toString))
    if customHeight then attrs += ("customHeight" -> "1")
    if hidden then attrs += ("hidden" -> "1")
    outlineLevel.foreach(l => attrs += ("outlineLevel" -> l.toString))
    if collapsed then attrs += ("collapsed" -> "1")
    if thickBot then attrs += ("thickBot" -> "1")
    if thickTop then attrs += ("thickTop" -> "1")
    dyDescent.foreach(d => attrs += ("x14ac:dyDescent" -> d.toString))

    elemOrdered("row", attrs.result()*)(
      cells.sortBy(_.ref.col.index0).map(_.toXml)*
    )
