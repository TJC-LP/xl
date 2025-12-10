package com.tjclp.xl.ooxml

import scala.xml.*
import XmlUtil.*
import com.tjclp.xl.addressing.* // For ARef, Column, Row types and extension methods
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.sheets.{ColumnProperties, RowProperties, Sheet}
import com.tjclp.xl.styles.color.Color
import SaxSupport.*

// Default namespaces for generated worksheets. Real files capture the original scope/attributes to
// avoid redundant declarations and preserve mc/x14/xr bindings from the source sheet.
private val defaultWorksheetScope =
  NamespaceBinding(null, nsSpreadsheetML, NamespaceBinding("r", nsRelationships, TopScope))

private def cleanNamespaces(elem: Elem): Elem =
  val cleanedChildren = elem.child.map {
    case e: Elem => cleanNamespaces(e)
    case other => other
  }
  elem.copy(scope = TopScope, child = cleanedChildren)

/**
 * Group consecutive columns with identical properties into spans.
 *
 * Excel's `<col>` element supports min/max attributes to apply the same properties to a range of
 * columns. This reduces file size by avoiding repeated `<col>` elements.
 *
 * @return
 *   Sequence of (minCol, maxCol, properties) tuples for span generation
 */
private def groupConsecutiveColumns(
  props: Map[Column, ColumnProperties]
): Seq[(Column, Column, ColumnProperties)] =
  if props.isEmpty then Seq.empty
  else
    props.toSeq
      .sortBy(_._1.index0)
      .foldLeft(Vector.empty[(Column, Column, ColumnProperties)]) { case (acc, (col, p)) =>
        acc.lastOption match
          case Some((minCol, maxCol, lastProps))
              if maxCol.index0 + 1 == col.index0 && lastProps == p =>
            // Extend current span
            acc.dropRight(1) :+ (minCol, col, p)
          case _ =>
            // Start new span
            acc :+ (col, col, p)
      }

/**
 * Build `<cols>` XML element from domain column properties.
 *
 * Generates OOXML-compliant column definitions: {{{<cols> <col min="1" max="3" width="15.5"
 * customWidth="1" hidden="1"/> </cols>}}}
 *
 * @return
 *   Some(cols element) if there are column properties, None otherwise
 */
private def buildColsElement(sheet: Sheet): Option[Elem] =
  val props = sheet.columnProperties
  if props.isEmpty then None
  else
    val spans = groupConsecutiveColumns(props)
    val colElems = spans.map { case (minCol, maxCol, p) =>
      // Build attribute sequence (only include non-default values)
      val attrs = Seq.newBuilder[(String, String)]
      attrs += ("min" -> (minCol.index0 + 1).toString) // OOXML is 1-based
      attrs += ("max" -> (maxCol.index0 + 1).toString)
      p.width.foreach { w =>
        attrs += ("width" -> w.toString)
        attrs += ("customWidth" -> "1")
      }
      if p.hidden then attrs += ("hidden" -> "1")
      p.outlineLevel.foreach(l => attrs += ("outlineLevel" -> l.toString))
      if p.collapsed then attrs += ("collapsed" -> "1")
      // Note: styleId would need remapping to workbook-level index (deferred)

      XmlUtil.elem("col", attrs.result()*)( /* no children */ )
    }
    Some(XmlUtil.elem("cols")(colElems*))

/**
 * Apply domain RowProperties to an OoxmlRow.
 *
 * Domain properties override existing row attributes (if any). This allows setting row height,
 * hidden state, and outline level from the domain model.
 */
private def applyDomainRowProps(row: OoxmlRow, props: RowProperties): OoxmlRow =
  row.copy(
    height = props.height.orElse(row.height),
    customHeight = props.height.isDefined || row.customHeight,
    hidden = props.hidden || row.hidden,
    outlineLevel = props.outlineLevel.orElse(row.outlineLevel),
    collapsed = props.collapsed || row.collapsed
    // Note: styleId would need remapping to workbook-level index (deferred)
  )

/** Cell data for worksheet - maps domain Cell to XML representation */
case class OoxmlCell(
  ref: ARef,
  value: CellValue,
  styleIndex: Option[Int] = None,
  cellType: String = "inlineStr" // "s" for SST, "inlineStr" for inline, "n" for number, etc.
):
  def toA1: String = ref.toA1

  def writeSax(writer: SaxWriter): Unit =
    writer.startElement("c")

    val attrs = Seq.newBuilder[(String, String)]
    attrs += ("r" -> toA1)
    styleIndex.foreach(s => attrs += ("s" -> s.toString))
    if cellType.nonEmpty then attrs += ("t" -> cellType)

    SaxWriter.withAttributes(writer, attrs.result()*) {
      value match
        case CellValue.Empty => ()

        case CellValue.Text(text) if cellType == "inlineStr" =>
          writer.startElement("is")
          writer.startElement("t")
          if needsXmlSpacePreserve(text) then writer.writeAttribute("xml:space", "preserve")
          writer.writeCharacters(text)
          writer.endElement() // t
          writer.endElement() // is

        case CellValue.Text(text) =>
          writer.startElement("v")
          writer.writeCharacters(text)
          writer.endElement() // v

        case CellValue.RichText(richText) =>
          writeRichTextSax(writer, richText)

        case CellValue.Number(num) =>
          writer.startElement("v")
          writer.writeCharacters(num.toString)
          writer.endElement() // v

        case CellValue.Bool(b) =>
          writer.startElement("v")
          writer.writeCharacters(if b then "1" else "0")
          writer.endElement() // v

        case CellValue.Formula(expr, cachedValue) =>
          writer.startElement("f")
          writer.writeCharacters(expr)
          writer.endElement() // f
          // Write cached value if present
          cachedValue.foreach {
            case CellValue.Number(num) =>
              writer.startElement("v")
              writer.writeCharacters(num.toString)
              writer.endElement()
            case CellValue.Text(s) =>
              writer.startElement("v")
              writer.writeCharacters(s)
              writer.endElement()
            case CellValue.Bool(b) =>
              writer.startElement("v")
              writer.writeCharacters(if b then "1" else "0")
              writer.endElement()
            case CellValue.Error(err) =>
              import com.tjclp.xl.cells.CellError.toExcel
              writer.startElement("v")
              writer.writeCharacters(err.toExcel)
              writer.endElement()
            case _ => () // Empty, RichText, Formula, DateTime - don't write
          }

        case CellValue.Error(err) =>
          import com.tjclp.xl.cells.CellError.toExcel
          writer.startElement("v")
          writer.writeCharacters(err.toExcel)
          writer.endElement() // v

        case CellValue.DateTime(dt) =>
          val serial = CellValue.dateTimeToExcelSerial(dt)
          writer.startElement("v")
          writer.writeCharacters(serial.toString)
          writer.endElement() // v
    }

    writer.endElement() // c

  private def writeRichTextSax(writer: SaxWriter, richText: com.tjclp.xl.richtext.RichText): Unit =
    writer.startElement("is")

    richText.runs.foreach { run =>
      writer.startElement("r")

      // Write rPr either from preserved raw XML or constructed from Font
      val preservedRpr = run.rawRPrXml.flatMap { xmlString =>
        XmlSecurity
          .parseSafe(xmlString, "worksheet richtext rPr")
          .toOption
          .map(XmlUtil.stripNamespaces)
      }

      preservedRpr match
        case Some(elem) =>
          writer.writeElem(elem)
        case None =>
          run.font.foreach(writeFontRPrSax(writer, _))

      writer.startElement("t")
      if needsXmlSpacePreserve(run.text) then writer.writeAttribute("xml:space", "preserve")
      writer.writeCharacters(run.text)
      writer.endElement() // t

      writer.endElement() // r
    }

    writer.endElement() // is

  private def writeFontRPrSax(writer: SaxWriter, font: com.tjclp.xl.styles.Font): Unit =
    writer.startElement("rPr")
    if font.bold then
      writer.startElement("b")
      writer.endElement()
    if font.italic then
      writer.startElement("i")
      writer.endElement()
    if font.underline then
      writer.startElement("u")
      writer.endElement()

    font.color.foreach {
      case Color.Rgb(argb) =>
        writer.startElement("color")
        writer.writeAttribute("rgb", f"$argb%08X")
        writer.endElement()
      case Color.Theme(slot, tint) =>
        writer.startElement("color")
        writer.writeAttribute("theme", slot.ordinal.toString)
        writer.writeAttribute("tint", tint.toString)
        writer.endElement()
    }

    writer.startElement("sz")
    writer.writeAttribute("val", font.sizePt.toString)
    writer.endElement()

    writer.startElement("name")
    writer.writeAttribute("val", font.name)
    writer.endElement()

    writer.endElement() // rPr

  def toXml: Elem =
    // Excel expects attributes in specific order: r, s, t
    val attrs = Seq.newBuilder[(String, String)]
    attrs += ("r" -> toA1)
    styleIndex.foreach(s => attrs += ("s" -> s.toString))
    if cellType.nonEmpty then attrs += ("t" -> cellType)

    val finalAttrs = attrs.result()

    val valueElem = value match
      case CellValue.Empty => Seq.empty
      case CellValue.Text(text) if cellType == "inlineStr" =>
        // Add xml:space="preserve" for text with leading/trailing/multiple spaces
        val needsPreserve = needsXmlSpacePreserve(text)
        val tElem =
          if needsPreserve then
            Elem(
              null,
              "t",
              PrefixedAttribute("xml", "space", "preserve", Null),
              TopScope,
              true,
              Text(text)
            )
          else elem("t")(Text(text))
        Seq(elem("is")(tElem))
      case CellValue.Text(text) => // SST index
        Seq(elem("v")(Text(text))) // text here would be the SST index as string
      case CellValue.RichText(richText) =>
        // Rich text: <is> with multiple <r> (text run) elements
        val runElems = richText.runs.map { run =>
          // Use preserved raw <rPr> if available (byte-perfect), otherwise build from Font
          val rPrElems = run.rawRPrXml.flatMap { xmlString =>
            // Parse preserved XML string back to Elem with XXE protection
            XmlSecurity.parseSafe(xmlString, "worksheet richtext rPr").toOption.map { elem =>
              // Strip redundant xmlns recursively from entire tree (namespace already on parent)
              XmlUtil.stripNamespaces(elem)
            }
          }.toList match
            case preserved if preserved.nonEmpty => preserved
            case _ =>
              // Build from Font model if no raw XML or parse failed
              run.font.map { f =>
                val fontProps = Seq.newBuilder[Elem]

                // Font style properties (order matters for OOXML)
                if f.bold then fontProps += elem("b")()
                if f.italic then fontProps += elem("i")()
                if f.underline then fontProps += elem("u")()

                // Font color
                f.color.foreach {
                  case Color.Rgb(argb) =>
                    fontProps += elem("color", "rgb" -> f"$argb%08X")()
                  case Color.Theme(slot, tint) =>
                    fontProps += elem(
                      "color",
                      "theme" -> slot.ordinal.toString,
                      "tint" -> tint.toString
                    )()
                }

                // Font size and name
                fontProps += elem("sz", "val" -> f.sizePt.toString)()
                fontProps += elem("name", "val" -> f.name)()

                elem("rPr")(fontProps.result()*)
              }.toList

          // Text run: <r> with optional <rPr> and <t>
          // Add xml:space="preserve" to preserve leading/trailing/multiple spaces
          val textElem =
            if needsXmlSpacePreserve(run.text) then
              Elem(
                null,
                "t",
                PrefixedAttribute("xml", "space", "preserve", Null),
                TopScope,
                true,
                Text(run.text)
              )
            else elem("t")(Text(run.text))

          elem("r")(
            rPrElems ++ Seq(textElem)*
          )
        }

        Seq(elem("is")(runElems*))
      case CellValue.Number(num) =>
        Seq(elem("v")(Text(num.toString)))
      case CellValue.Bool(b) =>
        Seq(elem("v")(Text(if b then "1" else "0")))
      case CellValue.Formula(expr, cachedValue) =>
        // Write formula element
        val formulaElem = elem("f")(Text(expr))
        // Write cached value if present
        val cachedElem = cachedValue.flatMap {
          case CellValue.Number(num) => Some(elem("v")(Text(num.toString)))
          case CellValue.Text(s) => Some(elem("v")(Text(s)))
          case CellValue.Bool(b) => Some(elem("v")(Text(if b then "1" else "0")))
          case CellValue.Error(err) =>
            import com.tjclp.xl.cells.CellError.toExcel
            Some(elem("v")(Text(err.toExcel)))
          case _ => None // Empty, RichText, Formula, DateTime - don't write
        }
        Seq(formulaElem) ++ cachedElem.toList
      case CellValue.Error(err) =>
        import com.tjclp.xl.cells.CellError.toExcel
        Seq(elem("v")(Text(err.toExcel)))
      case CellValue.DateTime(dt) =>
        // DateTime is serialized as number with Excel serial format
        val serial = CellValue.dateTimeToExcelSerial(dt)
        Seq(elem("v")(Text(serial.toString)))

    XmlUtil.elemOrdered("c", finalAttrs*)(valueElem*)

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

    XmlUtil.elemOrdered("row", attrs.result()*)(
      cells.sortBy(_.ref.col.index0).map(_.toXml)*
    )

/**
 * Worksheet for xl/worksheets/sheet#.xml
 *
 * Contains cell data in <sheetData> and preserves all worksheet metadata for surgical modification.
 * Unparsed elements are preserved as raw XML to maintain Excel compatibility.
 *
 * Elements are emitted in OOXML Part 1 schema order (§18.3.1.99).
 *
 * @param rows
 *   Row data with cells
 * @param mergedRanges
 *   Merged cell ranges
 * @param sheetPr
 *   Sheet properties (pageSetUpPr, outlinePr, tabColor, etc.)
 * @param dimension
 *   Used range reference (e.g., ref="B1:U104")
 * @param sheetViews
 *   View settings (zoom, gridLines, selection, tabSelected, pane, etc.)
 * @param sheetFormatPr
 *   Default row/column sizes and outline levels
 * @param cols
 *   Column definitions (CRITICAL - widths, styles, hidden, outlineLevel, etc.)
 * @param pageMargins
 *   Page margins (left, right, top, bottom, header, footer)
 * @param pageSetup
 *   Page setup (orientation, paperSize, scale, fitToWidth, fitToHeight, etc.)
 * @param headerFooter
 *   Header and footer content
 * @param drawing
 *   Drawing reference for charts (r:id link to drawing.xml)
 * @param legacyDrawing
 *   Legacy VML drawing reference for comments (r:id link to vmlDrawing.vml)
 * @param picture
 *   Background picture reference
 * @param oleObjects
 *   OLE objects
 * @param controls
 *   ActiveX controls
 * @param extLst
 *   Extensions
 * @param otherElements
 *   Any other elements not explicitly handled
 */
case class OoxmlWorksheet(
  rows: Seq[OoxmlRow],
  mergedRanges: Set[CellRange] = Set.empty,
  // Worksheet metadata (emitted in OOXML schema order)
  sheetPr: Option[Elem] = None,
  dimension: Option[Elem] = None,
  sheetViews: Option[Elem] = None,
  sheetFormatPr: Option[Elem] = None,
  cols: Option[Elem] = None,
  conditionalFormatting: Seq[Elem] = Seq.empty,
  printOptions: Option[Elem] = None,
  rowBreaks: Option[Elem] = None,
  colBreaks: Option[Elem] = None,
  customPropertiesWs: Option[Elem] = None,
  // Page layout
  pageMargins: Option[Elem] = None,
  pageSetup: Option[Elem] = None,
  headerFooter: Option[Elem] = None,
  // Drawings and objects
  drawing: Option[Elem] = None,
  legacyDrawing: Option[Elem] = None,
  picture: Option[Elem] = None,
  oleObjects: Option[Elem] = None,
  controls: Option[Elem] = None,
  // Tables
  tableParts: Option[Elem] = None,
  // Extensions
  extLst: Option[Elem] = None,
  otherElements: Seq[Elem] = Seq.empty,
  rootAttributes: MetaData = Null,
  rootScope: NamespaceBinding = defaultWorksheetScope
) extends XmlWritable,
      SaxSerializable:

  def writeSax(writer: SaxWriter): Unit =
    writer.startDocument()

    val scope = Option(rootScope).getOrElse(defaultWorksheetScope)
    val rootAttrs = Option(rootAttributes).getOrElse(Null)

    writer.startElement("worksheet")

    // Namespace declarations (deterministic order)
    SaxWriter.withAttributes(
      writer,
      writer.namespaceAttributes(scope) ++ writer.metaDataAttributes(rootAttrs)*
    ) {
      // Emit metadata in OOXML order
      sheetPr.foreach(writer.writeElem)
      dimension.foreach(writer.writeElem)
      sheetViews.foreach(writer.writeElem)
      sheetFormatPr.foreach(writer.writeElem)
      cols.foreach(writer.writeElem) // Column definitions BEFORE sheetData

      // sheetData
      writer.startElement("sheetData")
      rows.sortBy(_.rowIndex).foreach(_.writeSax(writer))
      writer.endElement() // sheetData

      // mergeCells
      if mergedRanges.nonEmpty then
        writer.startElement("mergeCells")
        writer.writeAttribute("count", mergedRanges.size.toString)
        mergedRanges.toSeq
          .sortBy(r => (r.start.row.index0, r.start.col.index0))
          .foreach { range =>
            writer.startElement("mergeCell")
            writer.writeAttribute("ref", range.toA1)
            writer.endElement()
          }
        writer.endElement() // mergeCells

      conditionalFormatting.foreach(writer.writeElem)

      printOptions.foreach(writer.writeElem)
      pageMargins.foreach(writer.writeElem)
      pageSetup.foreach(writer.writeElem)
      headerFooter.foreach(writer.writeElem)

      rowBreaks.foreach(writer.writeElem)
      colBreaks.foreach(writer.writeElem)
      customPropertiesWs.foreach(writer.writeElem)

      drawing.foreach(writer.writeElem)
      legacyDrawing.foreach(writer.writeElem)
      picture.foreach(writer.writeElem)
      oleObjects.foreach(writer.writeElem)
      controls.foreach(writer.writeElem)

      tableParts.foreach(writer.writeElem)

      extLst.foreach(writer.writeElem)

      otherElements.foreach(writer.writeElem)
    }

    writer.endElement() // worksheet
    writer.endDocument()
    writer.flush()

  def toXml: Elem =
    val children = Seq.newBuilder[Node]

    // Emit in OOXML Part 1 schema order (§18.3.1.99) - critical for Excel compatibility
    sheetPr.foreach(e => children += cleanNamespaces(e))
    dimension.foreach(e => children += cleanNamespaces(e))
    sheetViews.foreach(e => children += cleanNamespaces(e))
    sheetFormatPr.foreach(e => children += cleanNamespaces(e))
    cols.foreach(e => children += cleanNamespaces(e)) // Column definitions BEFORE sheetData

    // sheetData (always regenerated with current cell values)
    val sheetDataElem = elem("sheetData")(
      rows.sortBy(_.rowIndex).map(_.toXml)*
    )
    children += sheetDataElem

    // mergeCells (regenerated if present)
    if mergedRanges.nonEmpty then
      val mergeCellElems = mergedRanges.toSeq
        .sortBy(r => (r.start.row.index0, r.start.col.index0))
        .map(range => elem("mergeCell", "ref" -> range.toA1)())
      children += elem("mergeCells", "count" -> mergedRanges.size.toString)(mergeCellElems*)

    // Conditional formatting (multiple allowed)
    conditionalFormatting.foreach(e => children += cleanNamespaces(e))

    // Page layout (after sheetData/mergeCells)
    printOptions.foreach(e => children += cleanNamespaces(e))
    pageMargins.foreach(e => children += cleanNamespaces(e))
    pageSetup.foreach(e => children += cleanNamespaces(e))
    headerFooter.foreach(e => children += cleanNamespaces(e))

    rowBreaks.foreach(e => children += cleanNamespaces(e))
    colBreaks.foreach(e => children += cleanNamespaces(e))
    customPropertiesWs.foreach(e => children += cleanNamespaces(e))

    // Drawings and objects
    drawing.foreach(e => children += cleanNamespaces(e))
    legacyDrawing.foreach(e => children += cleanNamespaces(e))
    picture.foreach(e => children += cleanNamespaces(e))
    oleObjects.foreach(e => children += cleanNamespaces(e))
    controls.foreach(e => children += cleanNamespaces(e))

    // Tables
    tableParts.foreach(e => children += cleanNamespaces(e))

    // Extensions
    extLst.foreach(e => children += cleanNamespaces(e))

    // Any other elements
    otherElements.foreach(e => children += cleanNamespaces(e))

    val scope = Option(rootScope).getOrElse(defaultWorksheetScope)
    val attrs = Option(rootAttributes).getOrElse(Null)

    Elem(null, "worksheet", attrs, scope, minimizeEmpty = false, children.result()*)

object OoxmlWorksheet extends XmlReadable[OoxmlWorksheet]:
  /** Create minimal empty worksheet */
  def empty: OoxmlWorksheet = OoxmlWorksheet(Seq.empty)

  /** Create worksheet from domain Sheet (inline strings only) */
  def fromDomain(sheet: Sheet, styleRemapping: Map[Int, Int] = Map.empty): OoxmlWorksheet =
    fromDomainWithSST(sheet, None, styleRemapping, None)

  /**
   * Create worksheet from domain Sheet with optional SST and style remapping.
   *
   * @param sheet
   *   The domain Sheet to serialize
   * @param sst
   *   Optional SharedStrings table for string deduplication
   * @param styleRemapping
   *   Map from sheet-local styleId to workbook-level styleId
   * @param tableParts
   *   Optional tableParts XML element (generated from Sheet.tables)
   * @param escapeFormulas
   *   If true, escape text values starting with =, +, -, @ to prevent formula injection
   */
  def fromDomainWithSST(
    sheet: Sheet,
    sst: Option[SharedStrings],
    styleRemapping: Map[Int, Int] = Map.empty,
    tableParts: Option[Elem] = None,
    escapeFormulas: Boolean = false
  ): OoxmlWorksheet =
    fromDomainWithMetadata(sheet, sst, styleRemapping, None, tableParts, escapeFormulas)

  /**
   * Create worksheet from domain Sheet, preserving metadata from original worksheet XML.
   *
   * Used during surgical modification to regenerate sheetData while preserving all worksheet
   * metadata (column widths, view settings, page setup, etc.).
   *
   * @param sheet
   *   The domain Sheet with updated cell values
   * @param sst
   *   Optional SharedStrings table
   * @param styleRemapping
   *   Map from sheet-local styleId to workbook-level styleId
   * @param preservedMetadata
   *   Optional original worksheet to extract metadata from
   * @param tableParts
   *   Optional tableParts XML element (takes priority over preserved metadata)
   * @param escapeFormulas
   *   If true, escape text values starting with =, +, -, @ to prevent formula injection
   */
  def fromDomainWithMetadata(
    sheet: Sheet,
    sst: Option[SharedStrings],
    styleRemapping: Map[Int, Int] = Map.empty,
    preservedMetadata: Option[OoxmlWorksheet] = None,
    tableParts: Option[Elem] = None,
    escapeFormulas: Boolean = false
  ): OoxmlWorksheet =
    // Build a map of row indices to preserved row attributes
    val preservedRowAttrs = preservedMetadata
      .map { preserved =>
        preserved.rows.map(r => r.rowIndex -> r).toMap
      }
      .getOrElse(Map.empty)

    // Group cells by row
    // Optimization: Use TreeMap for auto-sorted grouping (avoids O(n log n) sort after groupBy)
    import scala.collection.immutable.TreeMap
    val cellsByRow = sheet.cells.values
      .groupBy(_.ref.row.index1) // 1-based row index
      .to(TreeMap) // Auto-sorted by key
      .toSeq

    // Create rows with cells (preserving attributes from original)
    val rowsWithCells = cellsByRow.map { case (rowIdx, cells) =>
      val ooxmlCells = cells.map { cell =>
        // Remap sheet-local styleId to workbook-level index
        val globalStyleIdx = cell.styleId.flatMap { localId =>
          // Look up in remapping table, fall back to 0 (default) if not found
          styleRemapping.get(localId.value).orElse(Some(0))
        }

        // Determine cell type and value based on CellValue type and SST availability
        val (cellType, value) = cell.value match
          case com.tjclp.xl.cells.CellValue.Text(s) =>
            // Apply formula injection escaping if enabled
            val safeText = if escapeFormulas then com.tjclp.xl.cells.CellValue.escape(s) else s
            sst.flatMap(_.indexOf(safeText)) match
              case Some(idx) => ("s", com.tjclp.xl.cells.CellValue.Text(idx.toString))
              case None => ("inlineStr", com.tjclp.xl.cells.CellValue.Text(safeText))
          case com.tjclp.xl.cells.CellValue.RichText(rt) =>
            // Check if RichText exists in SST (it can be shared!)
            sst.flatMap(_.indexOf(rt)) match
              case Some(idx) => ("s", com.tjclp.xl.cells.CellValue.Text(idx.toString))
              case None => ("inlineStr", cell.value)
          case com.tjclp.xl.cells.CellValue.Number(_) => ("n", cell.value)
          case com.tjclp.xl.cells.CellValue.Bool(_) => ("b", cell.value)
          case com.tjclp.xl.cells.CellValue.DateTime(dt) =>
            // Convert to Excel serial number
            val serial = com.tjclp.xl.cells.CellValue.dateTimeToExcelSerial(dt)
            ("n", com.tjclp.xl.cells.CellValue.Number(BigDecimal(serial)))
          case com.tjclp.xl.cells.CellValue.Formula(_, cachedValue) =>
            // Use cached value's type if available, otherwise "str"
            val cellType = cachedValue match
              case Some(com.tjclp.xl.cells.CellValue.Number(_)) => "n"
              case Some(com.tjclp.xl.cells.CellValue.Bool(_)) => "b"
              case Some(com.tjclp.xl.cells.CellValue.Error(_)) => "e"
              case _ => "str"
            (cellType, cell.value)
          case com.tjclp.xl.cells.CellValue.Error(_) => ("e", cell.value)
          case com.tjclp.xl.cells.CellValue.Empty => ("", cell.value)

        OoxmlCell(cell.ref, value, globalStyleIdx, cellType)
      }.toSeq

      // Preserve row attributes from original if available
      val baseRow = preservedRowAttrs.get(rowIdx) match
        case Some(original) =>
          // Merge: use original's attributes, replace cells with new data
          original.copy(cells = ooxmlCells)
        case None =>
          // New row - create with defaults
          OoxmlRow(rowIdx, ooxmlCells)

      // Apply domain row properties (height, hidden, outlineLevel, collapsed)
      val rowWithDomainProps = sheet.rowProperties.get(Row.from1(rowIdx)) match
        case Some(domainProps) => applyDomainRowProps(baseRow, domainProps)
        case None => baseRow

      rowWithDomainProps
    }

    // Preserve empty rows from original (critical for Row 1!)
    val emptyRowsFromOriginal = preservedMetadata.toList.flatMap { preserved =>
      preserved.rows.filter(_.cells.isEmpty)
    }

    // Generate empty rows for domain row properties not already represented
    val existingRowIndices =
      cellsByRow.map(_._1).toSet ++ emptyRowsFromOriginal.map(_.rowIndex).toSet
    val emptyRowsFromDomain = sheet.rowProperties
      .filterNot { case (row, _) => existingRowIndices.contains(row.index1) }
      .map { case (row, props) =>
        applyDomainRowProps(OoxmlRow(row.index1, Seq.empty), props)
      }
      .toSeq

    // Combine rows with cells + empty rows from original + empty rows from domain, sort by index
    val allRows = (rowsWithCells ++ emptyRowsFromOriginal ++ emptyRowsFromDomain).sortBy(_.rowIndex)

    // Generate legacyDrawing element if sheet has comments but no preserved legacyDrawing
    val legacyDrawingElem =
      if sheet.comments.nonEmpty then
        preservedMetadata.flatMap(_.legacyDrawing).orElse {
          // New comments - generate legacyDrawing reference to VML
          Some(
            Elem(
              prefix = null,
              label = "legacyDrawing",
              attributes = new PrefixedAttribute("r", "id", "rId2", Null),
              scope = NamespaceBinding("r", nsRelationships, TopScope),
              minimizeEmpty = true
            )
          )
        }
      else preservedMetadata.flatMap(_.legacyDrawing)

    // Generate cols from domain properties if not preserved
    val generatedCols = buildColsElement(sheet)

    // Calculate actual dimension from all rows (recalculate to reflect any new cells)
    val calculatedDimension: Option[Elem] =
      val allCells = allRows.flatMap(_.cells)
      for
        minCol <- allCells.map(_.ref.col.index0).minOption
        maxCol <- allCells.map(_.ref.col.index0).maxOption
        minRow <- allCells.map(_.ref.row.index0).minOption
        maxRow <- allCells.map(_.ref.row.index0).maxOption
      yield
        val startRef = ARef.from0(minCol, minRow)
        val endRef = ARef.from0(maxCol, maxRow)
        elem("dimension", "ref" -> s"${startRef.toA1}:${endRef.toA1}")()

    // If preservedMetadata is provided, use its metadata fields; otherwise use defaults (None)
    preservedMetadata match
      case Some(preserved) =>
        OoxmlWorksheet(
          allRows, // Use merged rows (with cells + empty rows)
          sheet.mergedRanges,
          // Preserve all metadata from original, but use recalculated dimension
          preserved.sheetPr,
          calculatedDimension.orElse(
            preserved.dimension
          ), // Use calculated dimension, fallback to preserved
          preserved.sheetViews,
          preserved.sheetFormatPr,
          preserved.cols.orElse(generatedCols), // Fallback to domain props if not preserved
          preserved.conditionalFormatting,
          preserved.printOptions,
          preserved.rowBreaks,
          preserved.colBreaks,
          preserved.customPropertiesWs,
          preserved.pageMargins,
          preserved.pageSetup,
          preserved.headerFooter,
          preserved.drawing,
          legacyDrawingElem, // Use computed element (preserved or generated)
          preserved.picture,
          preserved.oleObjects,
          preserved.controls,
          tableParts.orElse(preserved.tableParts), // Use parameter if provided, else preserve
          preserved.extLst,
          preserved.otherElements,
          preserved.rootAttributes,
          preserved.rootScope
        )
      case None =>
        // No preserved metadata - create minimal worksheet with cols from domain
        OoxmlWorksheet(
          rowsWithCells,
          sheet.mergedRanges,
          dimension = calculatedDimension,
          cols = generatedCols,
          legacyDrawing = legacyDrawingElem,
          tableParts = tableParts
        )

  /** Parse worksheet from XML (XmlReadable trait compatibility) */
  def fromXml(elem: Elem): Either[String, OoxmlWorksheet] =
    fromXmlWithSST(elem, None)

  /** Parse worksheet from XML with optional SharedStrings table */
  def fromXmlWithSST(elem: Elem, sst: Option[SharedStrings]): Either[String, OoxmlWorksheet] =
    for
      // Parse sheetData (required)
      sheetDataElem <- getChild(elem, "sheetData")
      rowElems = getChildren(sheetDataElem, "row")
      rows <- parseRows(rowElems, sst)

      // Parse mergeCells (optional)
      mergedRanges <- parseMergeCells(elem)

      // Extract ALL preserved metadata elements (all optional)
      sheetPr = (elem \ "sheetPr").headOption.collect { case e: Elem => cleanNamespaces(e) }
      dimension = (elem \ "dimension").headOption.collect { case e: Elem => cleanNamespaces(e) }
      sheetViews = (elem \ "sheetViews").headOption.collect { case e: Elem => cleanNamespaces(e) }
      sheetFormatPr = (elem \ "sheetFormatPr").headOption.collect { case e: Elem =>
        cleanNamespaces(e)
      }
      cols = (elem \ "cols").headOption.collect { case e: Elem => cleanNamespaces(e) }

      conditionalFormatting = elem.child.collect {
        case e: Elem if e.label == "conditionalFormatting" => cleanNamespaces(e)
      }
      printOptions = (elem \ "printOptions").headOption.collect { case e: Elem =>
        cleanNamespaces(e)
      }
      rowBreaks = (elem \ "rowBreaks").headOption.collect { case e: Elem => cleanNamespaces(e) }
      colBreaks = (elem \ "colBreaks").headOption.collect { case e: Elem => cleanNamespaces(e) }
      customPropertiesWs = (elem \ "customProperties").headOption.collect { case e: Elem =>
        cleanNamespaces(e)
      }

      pageMargins = (elem \ "pageMargins").headOption.collect { case e: Elem => cleanNamespaces(e) }
      pageSetup = (elem \ "pageSetup").headOption.collect { case e: Elem => cleanNamespaces(e) }
      headerFooter = (elem \ "headerFooter").headOption.collect { case e: Elem =>
        cleanNamespaces(e)
      }

      drawing = (elem \ "drawing").headOption.collect { case e: Elem => cleanNamespaces(e) }
      legacyDrawing = (elem \ "legacyDrawing").headOption.collect { case e: Elem =>
        cleanNamespaces(e)
      }
      picture = (elem \ "picture").headOption.collect { case e: Elem => cleanNamespaces(e) }
      oleObjects = (elem \ "oleObjects").headOption.collect { case e: Elem => cleanNamespaces(e) }
      controls = (elem \ "controls").headOption.collect { case e: Elem => cleanNamespaces(e) }

      tableParts = (elem \ "tableParts").headOption.collect { case e: Elem => cleanNamespaces(e) }

      extLst = (elem \ "extLst").headOption.collect { case e: Elem => cleanNamespaces(e) }

      // Collect any other elements we don't explicitly handle
      knownElements = Set(
        "sheetPr",
        "dimension",
        "sheetViews",
        "sheetFormatPr",
        "cols",
        "sheetData",
        "mergeCells",
        "pageMargins",
        "pageSetup",
        "headerFooter",
        "drawing",
        "legacyDrawing",
        "picture",
        "oleObjects",
        "controls",
        "extLst",
        // Additional elements from OOXML spec
        "sheetCalcPr",
        "sheetProtection",
        "protectedRanges",
        "scenarios",
        "autoFilter",
        "sortState",
        "dataConsolidate",
        "customSheetViews",
        "phoneticPr",
        "conditionalFormatting",
        "dataValidations",
        "hyperlinks",
        "printOptions",
        "rowBreaks",
        "colBreaks",
        "customProperties",
        "cellWatches",
        "ignoredErrors",
        "smartTags",
        "legacyDrawingHF",
        "webPublishItems",
        "tableParts"
      )
      otherElements = elem.child.collect {
        case e: Elem if !knownElements.contains(e.label) => e
      }
    yield OoxmlWorksheet(
      rows,
      mergedRanges,
      sheetPr,
      dimension,
      sheetViews,
      sheetFormatPr,
      cols,
      conditionalFormatting,
      printOptions,
      rowBreaks,
      colBreaks,
      customPropertiesWs,
      pageMargins,
      pageSetup,
      headerFooter,
      drawing,
      legacyDrawing,
      picture,
      oleObjects,
      controls,
      tableParts,
      extLst,
      otherElements.toSeq,
      rootAttributes = elem.attributes,
      rootScope = Option(elem.scope).getOrElse(defaultWorksheetScope)
    )

  private def parseMergeCells(worksheetElem: Elem): Either[String, Set[CellRange]] =
    // mergeCells is optional
    (worksheetElem \ "mergeCells").headOption match
      case None => Right(Set.empty)
      case Some(mergeCellsElem: Elem) =>
        val mergeCellElems = getChildren(mergeCellsElem, "mergeCell")
        val parsed = mergeCellElems.map { elem =>
          for
            refStr <- getAttr(elem, "ref")
            range <- CellRange.parse(refStr)
          yield range
        }
        val errors = parsed.collect { case Left(err) => err }
        if errors.nonEmpty then Left(s"MergeCell parse errors: ${errors.mkString(", ")}")
        else Right(parsed.collect { case Right(range) => range }.toSet)
      case _ => Right(Set.empty) // Non-Elem node, ignore

  private def parseRows(
    elems: Seq[Elem],
    sst: Option[SharedStrings]
  ): Either[String, Seq[OoxmlRow]] =
    val parsed = elems.map { e =>
      for
        rStr <- getAttr(e, "r")
        rowIdx <- rStr.toIntOption.toRight(s"Invalid row index: $rStr")

        // Optimization: Extract all attributes once into Map for O(1) lookups (was O(n) per attr = O(11n) total)
        // This avoids 11 DOM traversals per row (1-2% speedup for 10k rows)
        attrs = e.attributes.asAttrMap

        // Extract ALL row attributes for byte-perfect preservation
        spans = attrs.get("spans")
        style = attrs.get("s").flatMap(_.toIntOption)
        height = attrs.get("ht").flatMap(_.toDoubleOption)
        customHeight = attrs.get("customHeight").contains("1")
        customFormat = attrs.get("customFormat").contains("1")
        hidden = attrs.get("hidden").contains("1")
        outlineLevel = attrs.get("outlineLevel").flatMap(_.toIntOption)
        collapsed = attrs.get("collapsed").contains("1")
        thickBot = attrs.get("thickBot").contains("1")
        thickTop = attrs.get("thickTop").contains("1")
        dyDescent = attrs.get("x14ac:dyDescent").flatMap(_.toDoubleOption)

        cellElems = getChildren(e, "c")
        cells <- parseCells(cellElems, sst)
      yield OoxmlRow(
        rowIdx,
        cells,
        spans,
        style,
        height,
        customHeight,
        customFormat,
        hidden,
        outlineLevel,
        collapsed,
        thickBot,
        thickTop,
        dyDescent
      )
    }

    val errors = parsed.collect { case Left(err) => err }
    if errors.nonEmpty then Left(s"Row parse errors: ${errors.mkString(", ")}")
    else Right(parsed.collect { case Right(row) => row })

  private def parseCells(
    elems: Seq[Elem],
    sst: Option[SharedStrings]
  ): Either[String, Seq[OoxmlCell]] =
    val parsed = elems.map { e =>
      for
        refStr <- getAttr(e, "r")
        ref <- ARef.parse(refStr)
        cellType = getAttrOpt(e, "t").getOrElse("")
        styleIdx = getAttrOpt(e, "s").flatMap(_.toIntOption)
        value <- parseCellValue(e, cellType, sst)
      yield OoxmlCell(ref, value, styleIdx, cellType)
    }

    val errors = parsed.collect { case Left(err) => err }
    if errors.nonEmpty then Left(s"Cell parse errors: ${errors.mkString(", ")}")
    else Right(parsed.collect { case Right(cell) => cell })

  private def parseCellValue(
    elem: Elem,
    cellType: String,
    sst: Option[SharedStrings]
  ): Either[String, CellValue] =
    // Check for formula element first (before cellType dispatch)
    (elem \ "f").headOption.map(_.text.trim) match
      case Some(formulaExpr) if formulaExpr.nonEmpty =>
        // Formula cell - parse cached value from <v> if present
        val cachedValue: Option[CellValue] = (elem \ "v").headOption.map(_.text).flatMap { vText =>
          // Infer cached value type from cellType attribute
          cellType match
            case "n" | "" =>
              try Some(CellValue.Number(BigDecimal(vText)))
              catch case _: NumberFormatException => None
            case "b" =>
              vText match
                case "1" | "true" => Some(CellValue.Bool(true))
                case "0" | "false" => Some(CellValue.Bool(false))
                case _ => None
            case "e" =>
              import com.tjclp.xl.cells.CellError
              CellError.parse(vText).toOption.map(CellValue.Error.apply)
            case "str" | "inlineStr" =>
              Some(CellValue.Text(vText))
            case _ => None
        }
        Right(CellValue.Formula(formulaExpr, cachedValue))

      case _ =>
        // No formula - dispatch on cellType as before
        parseCellValueWithoutFormula(elem, cellType, sst)

  /** Parse cell value when no formula element is present */
  private def parseCellValueWithoutFormula(
    elem: Elem,
    cellType: String,
    sst: Option[SharedStrings]
  ): Either[String, CellValue] =
    cellType match
      case "inlineStr" | "str" =>
        // Both "inlineStr" and "str" cell types use <is> for inline strings
        // Check for rich text (<is><r>) vs simple text (<is><t>)
        (elem \ "is").headOption match
          case None =>
            // Fallback: "str" type may have text in <v> element (preserving whitespace)
            (elem \ "v").headOption
              .collect { case e: Elem => e }
              .map(getTextPreservingWhitespace) match
              case Some(text) => Right(CellValue.Text(text))
              case None => Left(s"$cellType cell missing <is> element and <v> element")
          case Some(isElem: Elem) =>
            val rElems = getChildren(isElem, "r")

            if rElems.nonEmpty then
              // Rich text: parse runs with formatting
              parseTextRuns(rElems).map(CellValue.RichText.apply)
            else
              // Simple text: extract from <t> (preserving whitespace)
              (isElem \ "t").headOption
                .collect { case e: Elem => e }
                .map(getTextPreservingWhitespace) match
                case Some(text) => Right(CellValue.Text(text))
                case None => Left(s"$cellType <is> missing <t> element and has no <r> runs")
          case _ => Left(s"$cellType <is> is not an Elem")

      case "s" =>
        // SST index - resolve using SharedStrings table
        (elem \ "v").headOption.map(_.text) match
          case Some(idxStr) =>
            idxStr.toIntOption match
              case Some(idx) =>
                (for {
                  sharedStrings <- sst
                  entry <- sharedStrings.apply(idx)
                } yield sharedStrings.toCellValue(entry)) match
                  case Some(cellValue) => Right(cellValue)
                  case None =>
                    // SST index out of bounds → CellError.Ref (not parse failure)
                    Right(CellValue.Error(com.tjclp.xl.cells.CellError.Ref))
              case None =>
                // Invalid SST index format → CellError.Value
                Right(CellValue.Error(com.tjclp.xl.cells.CellError.Value))
          case None => Left("SST cell missing <v>")

      case "n" | "" =>
        // Number
        (elem \ "v").headOption.map(_.text) match
          case Some(numStr) =>
            try Right(CellValue.Number(BigDecimal(numStr)))
            catch case _: NumberFormatException => Left(s"Invalid number: $numStr")
          case None => Right(CellValue.Empty) // Empty numeric cell

      case "b" =>
        // Boolean
        (elem \ "v").headOption.map(_.text) match
          case Some("1") | Some("true") => Right(CellValue.Bool(true))
          case Some("0") | Some("false") => Right(CellValue.Bool(false))
          case other => Left(s"Invalid boolean value: $other")

      case "e" =>
        // Error
        (elem \ "v").headOption.map(_.text) match
          case Some(errStr) =>
            import com.tjclp.xl.cells.CellError
            CellError.parse(errStr).map(CellValue.Error.apply)
          case None => Left("Error cell missing <v>")

      case other =>
        Left(s"Unsupported cell type: $other")
