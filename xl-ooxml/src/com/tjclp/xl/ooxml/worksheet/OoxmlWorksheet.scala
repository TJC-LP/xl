package com.tjclp.xl.ooxml.worksheet

import scala.xml.*

import com.tjclp.xl.addressing.{ARef, CellRange, Row}
import com.tjclp.xl.ooxml.XmlUtil.{elem, nsRelationships}
import com.tjclp.xl.ooxml.{SaxSerializable, SaxWriter, SharedStrings, XmlWritable}
import com.tjclp.xl.sheets.Sheet

/**
 * Worksheet for xl/worksheets/sheet#.xml
 *
 * Contains cell data in <sheetData> and preserves all worksheet metadata for surgical modification.
 * Unparsed elements are preserved as raw XML to maintain Excel compatibility.
 *
 * Elements are emitted in OOXML Part 1 schema order (section 18.3.1.99).
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
    import com.tjclp.xl.ooxml.SaxSupport.*
    SaxWriter.withAttributes(writer, writer.combinedAttributes(scope, rootAttrs)*) {
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

    // Emit in OOXML Part 1 schema order (section 18.3.1.99) - critical for Excel compatibility
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

object OoxmlWorksheet extends com.tjclp.xl.ooxml.XmlReadable[OoxmlWorksheet]:
  /** Create minimal empty worksheet */
  def empty: OoxmlWorksheet = OoxmlWorksheet(Seq.empty)

  /** Parse worksheet from XML (XmlReadable trait compatibility) */
  def fromXml(elem: scala.xml.Elem): Either[String, OoxmlWorksheet] =
    WorksheetReader.fromXml(elem)

  /** Parse worksheet from XML with optional SharedStrings table */
  def fromXmlWithSST(
    elem: scala.xml.Elem,
    sst: Option[SharedStrings]
  ): Either[String, OoxmlWorksheet] =
    WorksheetReader.fromXmlWithSST(elem, sst)

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
            // Cell type determined by cached value. When no cached value,
            // omit type attribute (empty string) to let Excel infer.
            // Using "str" for formulas without cached values causes
            // Excel to show a corruption warning.
            val cellType = cachedValue match
              case Some(com.tjclp.xl.cells.CellValue.Number(_)) => "n"
              case Some(com.tjclp.xl.cells.CellValue.Bool(_)) => "b"
              case Some(com.tjclp.xl.cells.CellValue.Error(_)) => "e"
              case Some(com.tjclp.xl.cells.CellValue.Text(_)) => "str"
              case Some(com.tjclp.xl.cells.CellValue.DateTime(_)) => "n"
              case _ => "" // No type attr
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
          generatedCols.orElse(preserved.cols), // Prefer domain props over preserved XML
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
