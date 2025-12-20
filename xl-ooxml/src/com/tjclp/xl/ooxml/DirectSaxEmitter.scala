package com.tjclp.xl.ooxml

import com.tjclp.xl.api.{Sheet, Cell}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.addressing.{ARef, Row}
import com.tjclp.xl.richtext.RichText
import com.tjclp.xl.styles.Color
import com.tjclp.xl.ooxml.SaxSupport.writeElem
import scala.xml.{NamespaceBinding, TopScope}
import scala.collection.immutable.TreeMap

/**
 * Direct SAX emission from domain model, bypassing intermediate OOXML types.
 *
 * This emitter writes worksheet data directly from domain Sheet to SAX events, avoiding the
 * allocation of OoxmlWorksheet, OoxmlRow, and OoxmlCell intermediate objects. Used by the SaxStax
 * backend for optimized write performance.
 *
 * Performance: Avoids O(n) object allocations where n = cells + rows, reducing GC pressure and
 * improving write throughput by 5-7x for large worksheets.
 */
object DirectSaxEmitter:

  private val nsSpreadsheetML = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  private val nsRelationships =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  private val nsX14ac = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
  private val nsMc = "http://schemas.openxmlformats.org/markup-compatibility/2006"

  /**
   * Emit worksheet XML directly from domain Sheet.
   *
   * @param writer
   *   SAX writer to emit to
   * @param sheet
   *   Domain sheet with cells
   * @param sst
   *   Optional shared strings table for string deduplication
   * @param styleRemapping
   *   Map from sheet-local styleId to workbook-level styleId
   * @param tablePartsXml
   *   Optional tableParts XML element
   * @param escapeFormulas
   *   If true, escape text values starting with =, +, -, @ to prevent formula injection
   */
  def emitWorksheet(
    writer: SaxWriter,
    sheet: Sheet,
    sst: Option[SharedStrings],
    styleRemapping: Map[Int, Int],
    tablePartsXml: Option[scala.xml.Elem] = None,
    escapeFormulas: Boolean = false
  ): Unit =
    writer.startDocument()
    writer.startElement("worksheet", nsSpreadsheetML)

    // Emit namespace declarations
    writer.writeAttribute("xmlns", nsSpreadsheetML)
    writer.writeAttribute("xmlns:r", nsRelationships)
    writer.writeAttribute("xmlns:mc", nsMc)
    writer.writeAttribute("xmlns:x14ac", nsX14ac)
    writer.writeAttribute("mc:Ignorable", "x14ac")

    // Calculate dimension from cells
    emitDimension(writer, sheet)

    // Emit sheetViews (minimal default)
    emitDefaultSheetViews(writer)

    // Emit sheetFormatPr (default row height)
    writer.startElement("sheetFormatPr")
    writer.writeAttribute("defaultRowHeight", "15")
    writer.writeAttribute("x14ac:dyDescent", "0.25")
    writer.endElement()

    // Emit column definitions if any
    emitCols(writer, sheet)

    // Emit sheetData directly from domain cells
    emitSheetData(writer, sheet, sst, styleRemapping, escapeFormulas)

    // Emit mergeCells if any
    emitMergeCells(writer, sheet)

    // Emit tableParts if provided
    tablePartsXml.foreach(writer.writeElem)

    writer.endElement() // worksheet
    writer.endDocument()
    writer.flush()

  /**
   * Emit dimension element based on cell extent.
   */
  private def emitDimension(writer: SaxWriter, sheet: Sheet): Unit =
    if sheet.cells.nonEmpty then
      val refs = sheet.cells.keys.toSeq
      val minCol = refs.map(_.col.index0).minOption.getOrElse(0)
      val maxCol = refs.map(_.col.index0).maxOption.getOrElse(0)
      val minRow = refs.map(_.row.index0).minOption.getOrElse(0)
      val maxRow = refs.map(_.row.index0).maxOption.getOrElse(0)

      val startRef = ARef.from0(minCol, minRow)
      val endRef = ARef.from0(maxCol, maxRow)

      writer.startElement("dimension")
      writer.writeAttribute("ref", s"${startRef.toA1}:${endRef.toA1}")
      writer.endElement()

  /**
   * Emit default sheetViews.
   */
  private def emitDefaultSheetViews(writer: SaxWriter): Unit =
    writer.startElement("sheetViews")
    writer.startElement("sheetView")
    writer.writeAttribute("tabSelected", "1")
    writer.writeAttribute("workbookViewId", "0")
    writer.endElement()
    writer.endElement()

  /**
   * Emit column definitions from sheet.
   */
  private def emitCols(writer: SaxWriter, sheet: Sheet): Unit =
    val colProps = sheet.columnProperties
    if colProps.nonEmpty then
      writer.startElement("cols")
      colProps.toSeq.sortBy(_._1.index0).foreach { case (col, props) =>
        val colIdx = col.index1 // 1-based for OOXML
        writer.startElement("col")
        writer.writeAttribute("min", colIdx.toString)
        writer.writeAttribute("max", colIdx.toString)
        props.width.foreach(w => writer.writeAttribute("width", w.toString))
        if props.hidden then writer.writeAttribute("hidden", "1")
        writer.writeAttribute("customWidth", "1")
        writer.endElement()
      }
      writer.endElement()

  /**
   * Emit sheetData directly from domain cells, grouped by row.
   */
  private def emitSheetData(
    writer: SaxWriter,
    sheet: Sheet,
    sst: Option[SharedStrings],
    styleRemapping: Map[Int, Int],
    escapeFormulas: Boolean
  ): Unit =
    writer.startElement("sheetData")

    // Group cells by row index, using TreeMap for auto-sorted iteration
    val cellsByRow = sheet.cells.groupBy(_._1.row.index1).to(TreeMap)

    // Also include empty rows with properties
    val allRowIndices =
      (cellsByRow.keys ++ sheet.rowProperties.keys.map(_.index1)).toSeq.distinct.sorted

    allRowIndices.foreach { rowIdx =>
      val cells = cellsByRow.getOrElse(rowIdx, Map.empty)
      val rowProps = sheet.rowProperties.get(Row.from1(rowIdx))

      emitRow(writer, rowIdx, cells, rowProps, sst, styleRemapping, escapeFormulas)
    }

    writer.endElement() // sheetData

  /**
   * Emit a single row element.
   */
  private def emitRow(
    writer: SaxWriter,
    rowIdx: Int,
    cells: Map[ARef, Cell],
    rowProps: Option[com.tjclp.xl.sheets.RowProperties],
    sst: Option[SharedStrings],
    styleRemapping: Map[Int, Int],
    escapeFormulas: Boolean
  ): Unit =
    writer.startElement("row")
    writer.writeAttribute("r", rowIdx.toString)

    // Calculate spans
    if cells.nonEmpty then
      val colIndices = cells.keys.map(_.col.index1).toSeq
      val minCol = colIndices.minOption.getOrElse(1)
      val maxCol = colIndices.maxOption.getOrElse(1)
      writer.writeAttribute("spans", s"$minCol:$maxCol")

    // Row properties
    rowProps.foreach { props =>
      props.height.foreach(h => writer.writeAttribute("ht", h.toString))
      if props.height.isDefined then writer.writeAttribute("customHeight", "1")
      if props.hidden then writer.writeAttribute("hidden", "1")
      props.outlineLevel.foreach(l => writer.writeAttribute("outlineLevel", l.toString))
      if props.collapsed then writer.writeAttribute("collapsed", "1")
    }

    // Emit cells sorted by column
    cells.toSeq.sortBy(_._1.col.index0).foreach { case (ref, cell) =>
      emitCell(writer, ref, cell, sst, styleRemapping, escapeFormulas)
    }

    writer.endElement() // row

  /**
   * Emit a single cell element directly from domain Cell.
   */
  private def emitCell(
    writer: SaxWriter,
    ref: ARef,
    cell: Cell,
    sst: Option[SharedStrings],
    styleRemapping: Map[Int, Int],
    escapeFormulas: Boolean
  ): Unit =
    // Determine cell type and prepare value
    val (cellType, preparedValue) = determineCellType(cell.value, sst, escapeFormulas)

    // Skip empty cells with no style
    if cellType.nonEmpty || cell.styleId.isDefined then
      writer.startElement("c")
      writer.writeAttribute("r", ref.toA1)

      // Style index (remapped)
      cell.styleId.foreach { localId =>
        val globalIdx = styleRemapping.getOrElse(localId.value, 0)
        if globalIdx > 0 then writer.writeAttribute("s", globalIdx.toString)
      }

      // Cell type
      if cellType.nonEmpty then writer.writeAttribute("t", cellType)

      // Cell value
      emitCellValue(writer, cellType, preparedValue, sst, escapeFormulas)

      writer.endElement() // c

  /**
   * Determine OOXML cell type from domain CellValue.
   */
  private def determineCellType(
    value: CellValue,
    sst: Option[SharedStrings],
    escapeFormulas: Boolean
  ): (String, CellValue) =
    value match
      case CellValue.Text(s) =>
        val safeText = if escapeFormulas then CellValue.escape(s) else s
        sst.flatMap(_.indexOf(safeText)) match
          case Some(idx) => ("s", CellValue.Text(idx.toString))
          case None => ("inlineStr", CellValue.Text(safeText))

      case CellValue.RichText(rt) =>
        sst.flatMap(_.indexOf(rt)) match
          case Some(idx) => ("s", CellValue.Text(idx.toString))
          case None => ("inlineStr", value)

      case CellValue.Number(_) => ("n", value)
      case CellValue.Bool(_) => ("b", value)

      case CellValue.DateTime(dt) =>
        val serial = CellValue.dateTimeToExcelSerial(dt)
        ("n", CellValue.Number(BigDecimal(serial)))

      case CellValue.Formula(_, cachedValue) =>
        val cType = cachedValue match
          case Some(CellValue.Number(_)) => "n"
          case Some(CellValue.Bool(_)) => "b"
          case Some(CellValue.Error(_)) => "e"
          case _ => "str"
        (cType, value)

      case CellValue.Error(_) => ("e", value)
      case CellValue.Empty => ("", value)

  /**
   * Emit cell value content.
   */
  private def emitCellValue(
    writer: SaxWriter,
    cellType: String,
    value: CellValue,
    sst: Option[SharedStrings],
    escapeFormulas: Boolean
  ): Unit =
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
        // SST reference
        writer.startElement("v")
        writer.writeCharacters(text)
        writer.endElement()

      case CellValue.RichText(rt) =>
        emitRichText(writer, rt)

      case CellValue.Number(num) =>
        writer.startElement("v")
        writer.writeCharacters(num.toString)
        writer.endElement()

      case CellValue.Bool(b) =>
        writer.startElement("v")
        writer.writeCharacters(if b then "1" else "0")
        writer.endElement()

      case CellValue.Formula(expr, cachedValue) =>
        writer.startElement("f")
        writer.writeCharacters(expr)
        writer.endElement()
        emitCachedValue(writer, cachedValue)

      case CellValue.Error(err) =>
        import com.tjclp.xl.cells.CellError.toExcel
        writer.startElement("v")
        writer.writeCharacters(err.toExcel)
        writer.endElement()

      case CellValue.DateTime(dt) =>
        val serial = CellValue.dateTimeToExcelSerial(dt)
        writer.startElement("v")
        writer.writeCharacters(serial.toString)
        writer.endElement()

  /**
   * Emit cached value for formula cells.
   */
  private def emitCachedValue(writer: SaxWriter, cachedValue: Option[CellValue]): Unit =
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
      case _ => ()
    }

  /**
   * Emit rich text inline string.
   */
  private def emitRichText(writer: SaxWriter, richText: RichText): Unit =
    writer.startElement("is")

    richText.runs.foreach { run =>
      writer.startElement("r")

      // Emit font properties if present
      run.font.foreach(emitFontRPr(writer, _))

      writer.startElement("t")
      if needsXmlSpacePreserve(run.text) then writer.writeAttribute("xml:space", "preserve")
      writer.writeCharacters(run.text)
      writer.endElement() // t

      writer.endElement() // r
    }

    writer.endElement() // is

  /**
   * Emit font run properties (rPr) for rich text.
   */
  private def emitFontRPr(writer: SaxWriter, font: com.tjclp.xl.styles.Font): Unit =
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

  /**
   * Emit merged cells.
   */
  private def emitMergeCells(writer: SaxWriter, sheet: Sheet): Unit =
    val mergedRanges = sheet.mergedRanges
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
      writer.endElement()

  /**
   * Check if text needs xml:space="preserve" attribute.
   */
  private def needsXmlSpacePreserve(text: String): Boolean =
    text.headOption.exists(_.isWhitespace) ||
      text.lastOption.exists(_.isWhitespace) ||
      text.contains('\n') ||
      text.contains('\r')
