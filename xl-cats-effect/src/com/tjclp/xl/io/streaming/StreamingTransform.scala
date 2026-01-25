package com.tjclp.xl.io.streaming

import java.io.{ByteArrayInputStream, ByteArrayOutputStream, InputStream, OutputStream}
import javax.xml.parsers.SAXParserFactory
import org.xml.sax.{Attributes, InputSource}
import org.xml.sax.helpers.DefaultHandler
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.ooxml.{SaxWriter, StaxSaxWriter, XmlUtil}
import scala.collection.mutable

/**
 * SAX→StAX streaming transformer for O(1) memory worksheet modifications.
 *
 * This transformer reads worksheet XML via SAX parser and emits to StAX writer, modifying only
 * cells in the target range. All other content passes through unchanged.
 *
 * Supports:
 *   - Style changes (s="N" attribute)
 *   - Value changes (text, number, boolean, formula)
 *   - Combined style + value changes
 *
 * Memory: O(1) - no cell data buffered, only current element state
 */
object StreamingTransform:

  /** Patch to apply to a cell during streaming */
  sealed trait CellPatch

  object CellPatch:
    /** Set cell style ID only */
    final case class SetStyle(styleId: Int) extends CellPatch

    /** Set cell value only (preserves existing style) */
    final case class SetValue(value: CellValue, preserveStyle: Boolean = true) extends CellPatch

    /** Set both style and value */
    final case class SetStyleAndValue(styleId: Int, value: CellValue) extends CellPatch

  /**
   * Analysis of patches to enable early-abort optimization.
   *
   * @param minRow
   *   Minimum 1-based row index in patches
   * @param maxRow
   *   Maximum 1-based row index in patches
   * @param patchCount
   *   Number of cells to patch
   */
  final case class PatchAnalysis(minRow: Int, maxRow: Int, patchCount: Int)

  /** Analyze patches to determine row bounds for early abort optimization. */
  @SuppressWarnings(Array("org.wartremover.warts.IterableOps"))
  def analyzePatches(patches: Map[ARef, CellPatch]): Option[PatchAnalysis] =
    if patches.isEmpty then None
    else
      val rows = patches.keys.map(_.row.index1).toSeq
      Some(PatchAnalysis(rows.min, rows.max, patches.size))

  /**
   * Exception signaling early abort during write transform.
   *
   * Carries the row index where we aborted so the caller can splice from that row.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Null"))
  private[streaming] final class EarlyAbortWrite(val abortedAtRow: Int)
      extends RuntimeException(null, null, false, false)

  /** Minimum rows in file to enable early abort (overhead not worth it for small files) */
  private val earlyAbortRowThreshold = 1000

  // ========== Shared Utility Functions ==========
  // These functions are extracted to reduce code duplication between
  // TransformHandler and EarlyAbortTransformHandler.

  /** Determine OOXML cell type for a CellValue */
  private[streaming] def cellTypeForValue(value: CellValue): String =
    value match
      case CellValue.Text(_) => "inlineStr"
      case CellValue.Number(_) => "" // Empty = number type
      case CellValue.Bool(_) => "b"
      case CellValue.Formula(_, cachedValue) =>
        cachedValue match
          case Some(CellValue.Number(_)) => ""
          case Some(CellValue.Bool(_)) => "b"
          case Some(CellValue.Error(_)) => "e"
          case _ => "str"
      case CellValue.Error(_) => "e"
      case CellValue.DateTime(_) => "" // Stored as number
      case CellValue.RichText(_) => "inlineStr"
      case CellValue.Empty => ""

  /** Write cell content for a CellValue to the given writer */
  private[streaming] def writeCellContent(writer: SaxWriter, value: CellValue): Unit =
    value match
      case CellValue.Empty => ()

      case CellValue.Text(s) =>
        // <is><t>text</t></is>
        writer.startElement("is")
        writer.startElement("t")
        if XmlUtil.needsXmlSpacePreserve(s) then writer.writeAttribute("xml:space", "preserve")
        writer.writeCharacters(s)
        writer.endElement() // t
        writer.endElement() // is

      case CellValue.Number(n) =>
        // <v>123</v>
        writer.startElement("v")
        writer.writeCharacters(n.toString)
        writer.endElement()

      case CellValue.Bool(b) =>
        // <v>1</v> or <v>0</v>
        writer.startElement("v")
        writer.writeCharacters(if b then "1" else "0")
        writer.endElement()

      case CellValue.Formula(expr, cachedValue) =>
        // <f>formula</f><v>cached</v>
        writer.startElement("f")
        writer.writeCharacters(expr)
        writer.endElement()
        cachedValue.foreach(cv => writeCachedValue(writer, cv))

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

      case CellValue.RichText(rt) =>
        // <is><r>...</r></is>
        writer.startElement("is")
        rt.runs.foreach { run =>
          writer.startElement("r")
          run.font.foreach { font =>
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
            font.color.foreach { color =>
              writer.startElement("color")
              color match
                case com.tjclp.xl.styles.Color.Rgb(argb) =>
                  writer.writeAttribute("rgb", f"$argb%08X")
                case com.tjclp.xl.styles.Color.Theme(slot, tint) =>
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
          }
          writer.startElement("t")
          if XmlUtil.needsXmlSpacePreserve(run.text) then
            writer.writeAttribute("xml:space", "preserve")
          writer.writeCharacters(run.text)
          writer.endElement() // t
          writer.endElement() // r
        }
        writer.endElement() // is

  /** Write cached formula value to the given writer */
  private[streaming] def writeCachedValue(writer: SaxWriter, value: CellValue): Unit =
    value match
      case CellValue.Number(n) =>
        writer.startElement("v")
        writer.writeCharacters(n.toString)
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

  /**
   * Transform worksheet XML, applying patches to target cells.
   *
   * Streams input→output with O(1) memory, modifying only cells in the patches map.
   *
   * @param input
   *   Source worksheet XML stream
   * @param output
   *   Target worksheet XML stream
   * @param patches
   *   Map of cell references to patches to apply
   */
  def transformWorksheet(
    input: InputStream,
    output: OutputStream,
    patches: Map[ARef, CellPatch]
  ): Unit =
    val writer = StaxSaxWriter.create(output)
    val handler = new TransformHandler(writer, patches)

    val factory = SAXParserFactory.newInstance()
    factory.setNamespaceAware(true)
    val parser = factory.newSAXParser()

    // Note: Caller is responsible for closing input stream.
    // When used with ZipInputStream, the entry is managed by the caller.
    parser.parse(InputSource(input), handler)
    writer.flush()

  /**
   * Result of early-abort transform operation.
   *
   * @param outputBytes
   *   Transformed XML bytes (partial if early abort, ends with unclosed sheetData)
   * @param aborted
   *   True if parsing was aborted early
   * @param abortedAtRow
   *   1-based row index where we aborted (0 if not aborted) - first row NOT included in output
   */
  final case class EarlyAbortResult(
    outputBytes: Array[Byte],
    aborted: Boolean,
    abortedAtRow: Int
  )

  /**
   * Transform worksheet XML with early-abort optimization.
   *
   * For patches targeting only early rows in a large file, this aborts SAX parsing after processing
   * all target rows and finding the end of sheetData section. The remaining bytes (mergeCells,
   * pageMargins, etc.) are spliced directly without parsing.
   *
   * @param entryBytes
   *   Source worksheet XML as byte array
   * @param patches
   *   Map of cell references to patches to apply
   * @return
   *   EarlyAbortResult with transformed bytes and abort status
   */
  def transformWorksheetWithEarlyAbort(
    entryBytes: Array[Byte],
    patches: Map[ARef, CellPatch]
  ): EarlyAbortResult =
    analyzePatches(patches) match
      case None =>
        // No patches - just return original
        EarlyAbortResult(entryBytes, aborted = false, abortedAtRow = 0)

      case Some(analysis) =>
        // Use early abort only if targeting early rows
        // (not worth it if patches are in the last 20% of file)
        val baos = new ByteArrayOutputStream()
        val writer = StaxSaxWriter.create(baos)
        val handler = new EarlyAbortTransformHandler(writer, patches, analysis.maxRow)

        val factory = SAXParserFactory.newInstance()
        factory.setNamespaceAware(true)
        val parser = factory.newSAXParser()

        try
          parser.parse(InputSource(new ByteArrayInputStream(entryBytes)), handler)
          writer.flush()
          // Full parse completed - no early abort
          EarlyAbortResult(baos.toByteArray, aborted = false, abortedAtRow = 0)
        catch
          case e: EarlyAbortWrite =>
            // Early abort triggered - need to splice remaining bytes from aborted row
            writer.flush()
            EarlyAbortResult(baos.toByteArray, aborted = true, abortedAtRow = e.abortedAtRow)

  /**
   * Pre-scan worksheet to collect existing style IDs for cells in range.
   *
   * Used for merge mode: need existing styles to merge with new style.
   *
   * @param input
   *   Source worksheet XML stream
   * @param range
   *   Cell references to scan for
   * @return
   *   Map of cell references to their existing style IDs (0 if none)
   */
  def scanExistingStyles(
    input: InputStream,
    range: Set[ARef]
  ): Map[ARef, Int] =
    val result = mutable.Map[ARef, Int]()
    val handler = new StyleScanHandler(range, result)

    val factory = SAXParserFactory.newInstance()
    factory.setNamespaceAware(true)
    val parser = factory.newSAXParser()

    try parser.parse(InputSource(input), handler)
    finally
      scala.util.Try(input.close()) // Log-worthy but not fatal; stream may already be closed

    result.toMap

  /**
   * SAX handler that transforms worksheet XML via StAX writer.
   *
   * For cells in patches map:
   *   - Style patches: modify s="N" attribute, pass through content
   *   - Value patches: replace cell content entirely
   *   - StyleAndValue patches: modify both
   */
  @SuppressWarnings(
    Array(
      "org.wartremover.warts.Var",
      "org.wartremover.warts.While",
      "org.wartremover.warts.Return"
    )
  )
  private class TransformHandler(
    writer: SaxWriter,
    patches: Map[ARef, CellPatch]
  ) extends DefaultHandler:

    private var inDocument = false
    private var currentCellRef: Option[String] = None
    private var currentPatch: Option[CellPatch] = None
    private var pendingCharacters: StringBuilder = new StringBuilder
    private var depth = 0
    private var cellDepth = 0 // Depth when we entered a patched cell
    private var skipContent = false // Skip content for value-patched cells
    private var contentWritten = false // Have we written replacement content?

    private val patchesByRow: Map[Int, Map[Int, CellPatch]] =
      patches.toSeq
        .groupBy { case (ref, _) => ref.row.index1 }
        .view
        .mapValues(entries => entries.map { case (ref, patch) => ref.col.index0 -> patch }.toMap)
        .toMap
    private val sortedPatchRows: Vector[Int] = patchesByRow.keys.toVector.sorted
    private var nextPatchRowPos = 0

    private var inSheetData = false
    private var inRow = false
    private var currentRowIndex = 0
    private var lastSourceRowIndex = 0
    private var currentRowPatches: Map[Int, CellPatch] = Map.empty
    private var currentRowPatchCols: Vector[Int] = Vector.empty
    private var nextPatchColPos = 0

    // Track namespaces for proper emission
    private val namespaceBindings = mutable.Map[String, String]()
    private val pendingNamespaces = mutable.ListBuffer[(String, String)]()

    override def startDocument(): Unit =
      writer.startDocument()
      inDocument = true

    override def endDocument(): Unit =
      writer.endDocument()
      writer.flush()
      inDocument = false

    override def startPrefixMapping(prefix: String, uri: String): Unit =
      namespaceBindings(prefix) = uri
      pendingNamespaces += ((prefix, uri))

    override def endPrefixMapping(prefix: String): Unit =
      namespaceBindings.remove(prefix)

    override def startElement(
      uri: String,
      localName: String,
      qName: String,
      attributes: Attributes
    ): Unit =
      flushCharacters()
      depth += 1

      // If we're skipping content inside a value-patched cell, don't emit child elements
      if skipContent && depth > cellDepth then return

      localName match
        case "sheetData" =>
          inSheetData = true
          if uri.nonEmpty then writer.startElement(qName, uri) else writer.startElement(qName)
          emitPendingNamespaces()
          for i <- 0 until attributes.getLength do
            writer.writeAttribute(attributes.getQName(i), attributes.getValue(i))

        case "row" =>
          val rowIndex = parseRowIndex(attributes)
          lastSourceRowIndex = rowIndex
          if inSheetData then
            emitMissingRowsBefore(rowIndex)
            currentRowIndex = rowIndex
            currentRowPatches = patchesByRow.getOrElse(rowIndex, Map.empty)
            currentRowPatchCols = currentRowPatches.keys.toVector.sorted
            nextPatchColPos = 0
            inRow = true
            if nextPatchRowPos < sortedPatchRows.size &&
              sortedPatchRows(nextPatchRowPos) == rowIndex
            then nextPatchRowPos += 1

          if uri.nonEmpty then writer.startElement(qName, uri) else writer.startElement(qName)
          emitPendingNamespaces()
          for i <- 0 until attributes.getLength do
            writer.writeAttribute(attributes.getQName(i), attributes.getValue(i))

        case "c" =>
          // Cell element - check if we need to patch it
          val cellRef = Option(attributes.getValue("r"))
          val parsedRef = cellRef.flatMap(ref => ARef.parse(ref).toOption)
          parsedRef.foreach { ref =>
            if inRow && currentRowPatchCols.nonEmpty then
              emitMissingCellsBefore(ref.col.index0)
              consumeCurrentCellIfPatched(ref.col.index0)
          }
          currentCellRef = cellRef

          // Parse cell reference and check for patch
          val patch = parsedRef.flatMap(patches.get)
          currentPatch = patch
          cellDepth = depth
          contentWritten = false

          // Determine if we need to skip original content (for value patches)
          skipContent = patch match
            case Some(_: CellPatch.SetValue) => true
            case Some(_: CellPatch.SetStyleAndValue) => true
            case _ => false

          // Start element with potentially modified attributes
          if uri.nonEmpty then writer.startElement(qName, uri) else writer.startElement(qName)

          // Emit pending namespace declarations
          emitPendingNamespaces()

          // Determine new cell type and style for value patches
          val (newType, newStyle) = patch match
            case Some(CellPatch.SetValue(value, preserveStyle)) =>
              val t = cellTypeForValue(value)
              val s =
                if preserveStyle then Option(attributes.getValue("s")).flatMap(_.toIntOption)
                else None
              (Some(t), s)
            case Some(CellPatch.SetStyleAndValue(styleId, value)) =>
              (Some(cellTypeForValue(value)), Some(styleId))
            case Some(CellPatch.SetStyle(styleId)) =>
              (None, Some(styleId))
            case None =>
              (None, None)

          // Write attributes
          var hasStyle = false
          var hasType = false
          for i <- 0 until attributes.getLength do
            val attrQName = attributes.getQName(i)
            val attrValue = attributes.getValue(i)

            if attrQName == "s" then
              hasStyle = true
              newStyle match
                case Some(sid) => writer.writeAttribute("s", sid.toString)
                case None => writer.writeAttribute(attrQName, attrValue)
            else if attrQName == "t" then
              hasType = true
              newType match
                case Some(t) if t.nonEmpty => writer.writeAttribute("t", t)
                case Some(_) => () // Empty type = number, don't write
                case None => writer.writeAttribute(attrQName, attrValue)
            else writer.writeAttribute(attrQName, attrValue)

          // Add missing attributes if needed
          newStyle.foreach { sid =>
            if !hasStyle && sid > 0 then writer.writeAttribute("s", sid.toString)
          }
          newType.foreach { t =>
            if !hasType && t.nonEmpty then writer.writeAttribute("t", t)
          }

        case _ =>
          // Other elements - pass through unchanged
          if uri.nonEmpty then writer.startElement(qName, uri) else writer.startElement(qName)
          emitPendingNamespaces()

          // Write all attributes unchanged
          for i <- 0 until attributes.getLength do
            writer.writeAttribute(attributes.getQName(i), attributes.getValue(i))

    override def endElement(uri: String, localName: String, qName: String): Unit =
      // If we're skipping content inside a value-patched cell
      if skipContent && depth > cellDepth then
        depth -= 1
        return

      flushCharacters()

      // For value-patched cells, write replacement content before closing
      if localName == "c" && skipContent && !contentWritten then
        currentPatch match
          case Some(CellPatch.SetValue(value, _)) =>
            writeCellContent(value)
            contentWritten = true
          case Some(CellPatch.SetStyleAndValue(_, value)) =>
            writeCellContent(value)
            contentWritten = true
          case _ => ()

      if localName == "row" && inSheetData then emitRemainingCells()
      if localName == "sheetData" then emitRemainingRows()

      writer.endElement()
      depth -= 1

      if localName == "c" then
        currentCellRef = None
        currentPatch = None
        skipContent = false
        cellDepth = 0
      if localName == "row" then
        inRow = false
        currentRowIndex = 0
        currentRowPatches = Map.empty
        currentRowPatchCols = Vector.empty
        nextPatchColPos = 0
      if localName == "sheetData" then inSheetData = false

    override def characters(ch: Array[Char], start: Int, length: Int): Unit =
      // Skip content for value-patched cells
      if skipContent && depth > cellDepth then return
      pendingCharacters.appendAll(ch, start, length)

    override def ignorableWhitespace(ch: Array[Char], start: Int, length: Int): Unit =
      if skipContent && depth > cellDepth then return
      pendingCharacters.appendAll(ch, start, length)

    private def emitPendingNamespaces(): Unit =
      pendingNamespaces.foreach { case (prefix, nsUri) =>
        if prefix.isEmpty then writer.writeAttribute("xmlns", nsUri)
        else writer.writeAttribute(s"xmlns:$prefix", nsUri)
      }
      pendingNamespaces.clear()

    private def flushCharacters(): Unit =
      if pendingCharacters.nonEmpty then
        // Skip flushing if we're inside a value-patched cell's content
        if !(skipContent && depth > cellDepth) then
          writer.writeCharacters(pendingCharacters.toString)
        pendingCharacters.clear()

    private def parseRowIndex(attributes: Attributes): Int =
      Option(attributes.getValue("r")).flatMap(_.toIntOption).getOrElse(lastSourceRowIndex + 1)

    private def emitMissingRowsBefore(targetRow: Int): Unit =
      while nextPatchRowPos < sortedPatchRows.size && sortedPatchRows(nextPatchRowPos) < targetRow
      do
        val rowIndex = sortedPatchRows(nextPatchRowPos)
        writePatchedRow(rowIndex)
        nextPatchRowPos += 1

    private def emitRemainingRows(): Unit =
      while nextPatchRowPos < sortedPatchRows.size do
        val rowIndex = sortedPatchRows(nextPatchRowPos)
        writePatchedRow(rowIndex)
        nextPatchRowPos += 1

    private def emitMissingCellsBefore(targetCol: Int): Unit =
      while nextPatchColPos < currentRowPatchCols.size &&
        currentRowPatchCols(nextPatchColPos) < targetCol
      do
        val colIndex = currentRowPatchCols(nextPatchColPos)
        val ref = ARef.from1(colIndex + 1, currentRowIndex)
        currentRowPatches.get(colIndex).foreach { patch =>
          writePatchedCell(ref, patch)
        }
        nextPatchColPos += 1

    private def emitRemainingCells(): Unit =
      while nextPatchColPos < currentRowPatchCols.size do
        val colIndex = currentRowPatchCols(nextPatchColPos)
        val ref = ARef.from1(colIndex + 1, currentRowIndex)
        currentRowPatches.get(colIndex).foreach { patch =>
          writePatchedCell(ref, patch)
        }
        nextPatchColPos += 1

    private def consumeCurrentCellIfPatched(colIndex: Int): Unit =
      if nextPatchColPos < currentRowPatchCols.size &&
        currentRowPatchCols(nextPatchColPos) == colIndex
      then nextPatchColPos += 1

    private def writePatchedRow(rowIndex: Int): Unit =
      val rowPatches = patchesByRow.getOrElse(rowIndex, Map.empty)
      if rowPatches.nonEmpty then
        writer.startElement("row")
        writer.writeAttribute("r", rowIndex.toString)
        val cols = rowPatches.keys.toVector.sorted
        cols.foreach { colIndex =>
          val ref = ARef.from1(colIndex + 1, rowIndex)
          writePatchedCell(ref, rowPatches(colIndex))
        }
        writer.endElement()

    private def writePatchedCell(ref: ARef, patch: CellPatch): Unit =
      writer.startElement("c")
      writer.writeAttribute("r", ref.toA1)
      val (newType, newStyle) = patch match
        case CellPatch.SetValue(value, _) =>
          (Some(cellTypeForValue(value)), None)
        case CellPatch.SetStyle(styleId) =>
          (None, Some(styleId))
        case CellPatch.SetStyleAndValue(styleId, value) =>
          (Some(cellTypeForValue(value)), Some(styleId))
      newStyle.foreach { sid =>
        if sid > 0 then writer.writeAttribute("s", sid.toString)
      }
      newType.foreach { t =>
        if t.nonEmpty then writer.writeAttribute("t", t)
      }
      patch match
        case CellPatch.SetValue(value, _) => writeCellContent(value)
        case CellPatch.SetStyleAndValue(_, value) => writeCellContent(value)
        case _ => ()
      writer.endElement()

    // Delegate to companion object shared utility functions
    private def cellTypeForValue(value: CellValue): String =
      StreamingTransform.cellTypeForValue(value)

    private def writeCellContent(value: CellValue): Unit =
      StreamingTransform.writeCellContent(writer, value)

  /**
   * SAX handler for scanning existing style IDs in a worksheet.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  private class StyleScanHandler(
    targetRefs: Set[ARef],
    result: mutable.Map[ARef, Int]
  ) extends DefaultHandler:

    override def startElement(
      uri: String,
      localName: String,
      qName: String,
      attributes: Attributes
    ): Unit =
      if localName == "c" then
        Option(attributes.getValue("r")).foreach { refStr =>
          ARef.parse(refStr).toOption.foreach { ref =>
            if targetRefs.contains(ref) then
              val styleId = Option(attributes.getValue("s")).flatMap(_.toIntOption).getOrElse(0)
              result(ref) = styleId
          }
        }

  /**
   * SAX handler that transforms worksheet XML with early-abort optimization.
   *
   * This handler extends the standard transform logic but tracks row indices. When we see a row
   * past the maximum target row, we close the sheetData element and throw EarlyAbortWrite to abort
   * parsing.
   *
   * The remaining bytes (mergeCells, pageMargins, etc.) are handled by the caller via byte
   * splicing.
   */
  @SuppressWarnings(
    Array(
      "org.wartremover.warts.Var",
      "org.wartremover.warts.While",
      "org.wartremover.warts.Return"
    )
  )
  private[streaming] class EarlyAbortTransformHandler(
    writer: SaxWriter,
    patches: Map[ARef, CellPatch],
    maxTargetRow: Int
  ) extends DefaultHandler:

    private var inDocument = false
    private var currentCellRef: Option[String] = None
    private var currentPatch: Option[CellPatch] = None
    private var pendingCharacters: StringBuilder = new StringBuilder
    private var depth = 0
    private var cellDepth = 0
    private var skipContent = false
    private var contentWritten = false

    private val patchesByRow: Map[Int, Map[Int, CellPatch]] =
      patches.toSeq
        .groupBy { case (ref, _) => ref.row.index1 }
        .view
        .mapValues(entries => entries.map { case (ref, patch) => ref.col.index0 -> patch }.toMap)
        .toMap
    private val sortedPatchRows: Vector[Int] = patchesByRow.keys.toVector.sorted
    private var nextPatchRowPos = 0

    // Row tracking for early abort
    private var currentRowIndex = 0
    private var lastSourceRowIndex = 0
    private var currentRowPatches: Map[Int, CellPatch] = Map.empty
    private var currentRowPatchCols: Vector[Int] = Vector.empty
    private var nextPatchColPos = 0
    private var inSheetData = false
    private var pastTargetRows = false
    private var sheetDataDepth = 0

    // Track namespaces for proper emission
    private val namespaceBindings = mutable.Map[String, String]()
    private val pendingNamespaces = mutable.ListBuffer[(String, String)]()

    override def startDocument(): Unit =
      writer.startDocument()
      inDocument = true

    override def endDocument(): Unit =
      writer.endDocument()
      writer.flush()
      inDocument = false

    override def startPrefixMapping(prefix: String, uri: String): Unit =
      namespaceBindings(prefix) = uri
      pendingNamespaces += ((prefix, uri))

    override def endPrefixMapping(prefix: String): Unit =
      namespaceBindings.remove(prefix)

    override def startElement(
      uri: String,
      localName: String,
      qName: String,
      attributes: Attributes
    ): Unit =
      flushCharacters()
      depth += 1

      // Track sheetData for early abort
      if localName == "sheetData" then
        inSheetData = true
        sheetDataDepth = depth

      // Check for early abort condition: first row past max target
      if localName == "row" && inSheetData then
        val rowIndex = parseRowIndex(attributes)
        currentRowIndex = rowIndex
        lastSourceRowIndex = rowIndex
        emitMissingRowsBefore(rowIndex)
        currentRowPatches = patchesByRow.getOrElse(rowIndex, Map.empty)
        currentRowPatchCols = currentRowPatches.keys.toVector.sorted
        nextPatchColPos = 0
        if nextPatchRowPos < sortedPatchRows.size &&
          sortedPatchRows(nextPatchRowPos) == rowIndex
        then nextPatchRowPos += 1
        if rowIndex > maxTargetRow && !pastTargetRows then
          // Abort WITHOUT closing sheetData - caller will splice remaining rows from original
          pastTargetRows = true
          writer.flush()
          throw new EarlyAbortWrite(rowIndex)

      // If past target rows, don't process (shouldn't reach here due to exception)
      if pastTargetRows then return

      // If we're skipping content inside a value-patched cell, don't emit child elements
      if skipContent && depth > cellDepth then return

      localName match
        case "sheetData" =>
          if uri.nonEmpty then writer.startElement(qName, uri) else writer.startElement(qName)
          emitPendingNamespaces()

          // Write all attributes unchanged
          for i <- 0 until attributes.getLength do
            writer.writeAttribute(attributes.getQName(i), attributes.getValue(i))

        case "row" =>
          if uri.nonEmpty then writer.startElement(qName, uri) else writer.startElement(qName)
          emitPendingNamespaces()

          // Write all attributes unchanged
          for i <- 0 until attributes.getLength do
            writer.writeAttribute(attributes.getQName(i), attributes.getValue(i))

        case "c" =>
          // Cell element - check if we need to patch it
          val cellRef = Option(attributes.getValue("r"))
          val parsedRef = cellRef.flatMap(ref => ARef.parse(ref).toOption)
          parsedRef.foreach { ref =>
            if inSheetData && currentRowPatchCols.nonEmpty then
              emitMissingCellsBefore(ref.col.index0)
              consumeCurrentCellIfPatched(ref.col.index0)
          }
          currentCellRef = cellRef

          // Parse cell reference and check for patch
          val patch = parsedRef.flatMap(patches.get)
          currentPatch = patch
          cellDepth = depth
          contentWritten = false

          // Determine if we need to skip original content (for value patches)
          skipContent = patch match
            case Some(_: CellPatch.SetValue) => true
            case Some(_: CellPatch.SetStyleAndValue) => true
            case _ => false

          // Start element with potentially modified attributes
          if uri.nonEmpty then writer.startElement(qName, uri) else writer.startElement(qName)

          // Emit pending namespace declarations
          emitPendingNamespaces()

          // Determine new cell type and style for value patches
          val (newType, newStyle) = patch match
            case Some(CellPatch.SetValue(value, preserveStyle)) =>
              val t = cellTypeForValue(value)
              val s =
                if preserveStyle then Option(attributes.getValue("s")).flatMap(_.toIntOption)
                else None
              (Some(t), s)
            case Some(CellPatch.SetStyleAndValue(styleId, value)) =>
              (Some(cellTypeForValue(value)), Some(styleId))
            case Some(CellPatch.SetStyle(styleId)) =>
              (None, Some(styleId))
            case None =>
              (None, None)

          // Write attributes
          var hasStyle = false
          var hasType = false
          for i <- 0 until attributes.getLength do
            val attrQName = attributes.getQName(i)
            val attrValue = attributes.getValue(i)

            if attrQName == "s" then
              hasStyle = true
              newStyle match
                case Some(sid) => writer.writeAttribute("s", sid.toString)
                case None => writer.writeAttribute(attrQName, attrValue)
            else if attrQName == "t" then
              hasType = true
              newType match
                case Some(t) if t.nonEmpty => writer.writeAttribute("t", t)
                case Some(_) => () // Empty type = number, don't write
                case None => writer.writeAttribute(attrQName, attrValue)
            else writer.writeAttribute(attrQName, attrValue)

          // Add missing attributes if needed
          newStyle.foreach { sid =>
            if !hasStyle && sid > 0 then writer.writeAttribute("s", sid.toString)
          }
          newType.foreach { t =>
            if !hasType && t.nonEmpty then writer.writeAttribute("t", t)
          }

        case _ =>
          // Other elements - pass through unchanged
          if uri.nonEmpty then writer.startElement(qName, uri) else writer.startElement(qName)
          emitPendingNamespaces()

          // Write all attributes unchanged
          for i <- 0 until attributes.getLength do
            writer.writeAttribute(attributes.getQName(i), attributes.getValue(i))

    override def endElement(uri: String, localName: String, qName: String): Unit =
      // If past target rows, don't process (shouldn't reach here due to exception)
      if pastTargetRows then
        depth -= 1
        return

      // If we're skipping content inside a value-patched cell
      if skipContent && depth > cellDepth then
        depth -= 1
        return

      flushCharacters()

      // For value-patched cells, write replacement content before closing
      if localName == "c" && skipContent && !contentWritten then
        currentPatch match
          case Some(CellPatch.SetValue(value, _)) =>
            writeCellContent(value)
            contentWritten = true
          case Some(CellPatch.SetStyleAndValue(_, value)) =>
            writeCellContent(value)
            contentWritten = true
          case _ => ()

      if localName == "row" && inSheetData then emitRemainingCells()
      if localName == "sheetData" then emitRemainingRows()

      writer.endElement()
      depth -= 1

      if localName == "c" then
        currentCellRef = None
        currentPatch = None
        skipContent = false
        cellDepth = 0

      if localName == "row" then
        currentRowIndex = 0
        currentRowPatches = Map.empty
        currentRowPatchCols = Vector.empty
        nextPatchColPos = 0

      if localName == "sheetData" then inSheetData = false

    override def characters(ch: Array[Char], start: Int, length: Int): Unit =
      if pastTargetRows then return
      // Skip content for value-patched cells
      if skipContent && depth > cellDepth then return
      pendingCharacters.appendAll(ch, start, length)

    override def ignorableWhitespace(ch: Array[Char], start: Int, length: Int): Unit =
      if pastTargetRows then return
      if skipContent && depth > cellDepth then return
      pendingCharacters.appendAll(ch, start, length)

    private def emitPendingNamespaces(): Unit =
      pendingNamespaces.foreach { case (prefix, nsUri) =>
        if prefix.isEmpty then writer.writeAttribute("xmlns", nsUri)
        else writer.writeAttribute(s"xmlns:$prefix", nsUri)
      }
      pendingNamespaces.clear()

    private def flushCharacters(): Unit =
      if pendingCharacters.nonEmpty then
        // Skip flushing if we're inside a value-patched cell's content
        if !(skipContent && depth > cellDepth) then
          writer.writeCharacters(pendingCharacters.toString)
        pendingCharacters.clear()

    private def parseRowIndex(attributes: Attributes): Int =
      Option(attributes.getValue("r")).flatMap(_.toIntOption).getOrElse(lastSourceRowIndex + 1)

    private def emitMissingRowsBefore(targetRow: Int): Unit =
      while nextPatchRowPos < sortedPatchRows.size && sortedPatchRows(nextPatchRowPos) < targetRow
      do
        val rowIndex = sortedPatchRows(nextPatchRowPos)
        writePatchedRow(rowIndex)
        nextPatchRowPos += 1

    private def emitRemainingRows(): Unit =
      while nextPatchRowPos < sortedPatchRows.size do
        val rowIndex = sortedPatchRows(nextPatchRowPos)
        writePatchedRow(rowIndex)
        nextPatchRowPos += 1

    private def emitMissingCellsBefore(targetCol: Int): Unit =
      while nextPatchColPos < currentRowPatchCols.size &&
        currentRowPatchCols(nextPatchColPos) < targetCol
      do
        val colIndex = currentRowPatchCols(nextPatchColPos)
        val ref = ARef.from1(colIndex + 1, currentRowIndex)
        currentRowPatches.get(colIndex).foreach { patch =>
          writePatchedCell(ref, patch)
        }
        nextPatchColPos += 1

    private def emitRemainingCells(): Unit =
      while nextPatchColPos < currentRowPatchCols.size do
        val colIndex = currentRowPatchCols(nextPatchColPos)
        val ref = ARef.from1(colIndex + 1, currentRowIndex)
        currentRowPatches.get(colIndex).foreach { patch =>
          writePatchedCell(ref, patch)
        }
        nextPatchColPos += 1

    private def consumeCurrentCellIfPatched(colIndex: Int): Unit =
      if nextPatchColPos < currentRowPatchCols.size &&
        currentRowPatchCols(nextPatchColPos) == colIndex
      then nextPatchColPos += 1

    private def writePatchedRow(rowIndex: Int): Unit =
      val rowPatches = patchesByRow.getOrElse(rowIndex, Map.empty)
      if rowPatches.nonEmpty then
        writer.startElement("row")
        writer.writeAttribute("r", rowIndex.toString)
        val cols = rowPatches.keys.toVector.sorted
        cols.foreach { colIndex =>
          val ref = ARef.from1(colIndex + 1, rowIndex)
          writePatchedCell(ref, rowPatches(colIndex))
        }
        writer.endElement()

    private def writePatchedCell(ref: ARef, patch: CellPatch): Unit =
      writer.startElement("c")
      writer.writeAttribute("r", ref.toA1)
      val (newType, newStyle) = patch match
        case CellPatch.SetValue(value, _) =>
          (Some(cellTypeForValue(value)), None)
        case CellPatch.SetStyle(styleId) =>
          (None, Some(styleId))
        case CellPatch.SetStyleAndValue(styleId, value) =>
          (Some(cellTypeForValue(value)), Some(styleId))
      newStyle.foreach { sid =>
        if sid > 0 then writer.writeAttribute("s", sid.toString)
      }
      newType.foreach { t =>
        if t.nonEmpty then writer.writeAttribute("t", t)
      }
      patch match
        case CellPatch.SetValue(value, _) => writeCellContent(value)
        case CellPatch.SetStyleAndValue(_, value) => writeCellContent(value)
        case _ => ()
      writer.endElement()

    // Delegate to companion object shared utility functions
    private def cellTypeForValue(value: CellValue): String =
      StreamingTransform.cellTypeForValue(value)

    private def writeCellContent(value: CellValue): Unit =
      StreamingTransform.writeCellContent(writer, value)
