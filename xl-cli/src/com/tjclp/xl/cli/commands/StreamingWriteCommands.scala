package com.tjclp.xl.cli.commands

import java.nio.charset.StandardCharsets
import java.nio.file.Path
import java.util.zip.ZipFile

import cats.effect.IO
import com.tjclp.xl.api.Workbook
import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.{FormulaParser, FormulaPrinter, FormulaShifter, ParseError}
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.io.streaming.{StreamingTransform, StylePatcher, ZipTransformer}
import com.tjclp.xl.ooxml.XmlSecurity
import com.tjclp.xl.ooxml.writer.WriterConfig
import com.tjclp.xl.sheets.{ColumnProperties, RowProperties}
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.cli.helpers.{BatchParser, StreamingCsvParser, StyleBuilder, ValueParser}
import scala.xml.Elem

/**
 * Streaming write command handlers.
 *
 * Provides two modes:
 *   1. True streaming (CSV import): End-to-end O(1) memory
 *   2. Hybrid streaming (workbook write): In-memory workbook → O(1) output memory
 */
object StreamingWriteCommands:

  /**
   * True streaming CSV import to new XLSX file.
   *
   * Streams CSV rows directly to XLSX with O(1) memory throughout. Creates a fresh file with a
   * single sheet - no styles to preserve.
   *
   * Uses writeStreamWithAutoDetect for automatic dimension detection (two-pass approach for
   * accurate bounds).
   *
   * @param csvPath
   *   Path to CSV file
   * @param outputPath
   *   Output XLSX file path
   * @param sheetName
   *   Name for the new sheet
   * @param options
   *   CSV parsing options
   * @return
   *   Success message with file info
   */
  def importCsvStream(
    csvPath: Path,
    outputPath: Path,
    sheetName: String,
    options: StreamingCsvParser.Options
  ): IO[String] =
    val excel = ExcelIO.instance[IO]

    StreamingCsvParser
      .streamCsv(csvPath, options)
      .through(excel.writeStreamWithAutoDetect(outputPath, sheetName))
      .compile
      .drain
      .map(_ => s"Streamed: ${csvPath.getFileName} → $outputPath (sheet: $sheetName)")

  /**
   * Hybrid streaming workbook write.
   *
   * Takes an in-memory workbook and writes it using the streaming writer for O(1) output memory.
   * Styles are fully preserved via StyleIndex.
   *
   * Best for: Large modified workbooks where output memory is the bottleneck.
   *
   * @param wb
   *   Workbook to write (already in memory)
   * @param outputPath
   *   Output file path
   * @param config
   *   Writer configuration
   * @return
   *   Unit on success
   */
  def writeWorkbookStreaming(
    wb: Workbook,
    outputPath: Path,
    config: WriterConfig = WriterConfig.default
  ): IO[Unit] =
    ExcelIO.instance[IO].writeWorkbookStream(wb, outputPath, config)

  /**
   * Hybrid streaming workbook write with success message.
   *
   * Same as writeWorkbookStreaming but returns a formatted success message.
   *
   * @param wb
   *   Workbook to write
   * @param outputPath
   *   Output file path
   * @param config
   *   Writer configuration
   * @param operation
   *   Description of the operation (e.g., "put A1 = Hello")
   * @return
   *   Formatted success message
   */
  def writeWorkbookStreamingWithMessage(
    wb: Workbook,
    outputPath: Path,
    config: WriterConfig,
    operation: String
  ): IO[String] =
    writeWorkbookStreaming(wb, outputPath, config).map { _ =>
      s"$operation\nSaved (streaming): $outputPath"
    }

  /**
   * Streaming put: write values to cells with O(1) memory.
   *
   * Uses SAX→StAX transform pipeline to modify only target cells.
   *
   * @param sourcePath
   *   Input XLSX file
   * @param outputPath
   *   Output XLSX file
   * @param sheetNameOpt
   *   Sheet name (required for multi-sheet files)
   * @param refStr
   *   Cell reference or range
   * @param values
   *   Values to write (single value for fill, multiple for batch)
   * @return
   *   Result message
   */
  def put(
    sourcePath: Path,
    outputPath: Path,
    sheetNameOpt: Option[String],
    refStr: String,
    values: List[String]
  ): IO[String] =
    import com.tjclp.xl.addressing.ARef
    import com.tjclp.xl.cells.CellValue
    import com.tjclp.xl.cli.helpers.ValueParser

    for
      // Resolve sheet path
      worksheetPath <- resolveSheetPath(sourcePath, sheetNameOpt)

      // Parse reference (single cell or range)
      refOrRange <- IO.fromEither(
        CellRange
          .parse(refStr)
          .map(Right(_))
          .left
          .flatMap { _ =>
            ARef.parse(refStr).map(Left(_))
          }
          .left
          .map(e => new Exception(s"Invalid reference: $refStr"))
      )

      // Build value map based on mode
      valueMap <- (refOrRange, values) match
        case (Left(ref), List(singleValue)) =>
          // Single cell
          IO.pure(Map(ref -> ValueParser.parseValue(singleValue)))

        case (Right(range), List(singleValue)) =>
          // Fill pattern: all cells get same value
          val value = ValueParser.parseValue(singleValue)
          IO.pure(range.cells.map(ref => ref -> value).toMap)

        case (Right(range), multipleValues) if multipleValues.length == range.cellCount.toInt =>
          // Batch values: exact count match
          val pairs = range.cellsRowMajor
            .zip(multipleValues.iterator)
            .map { (ref, v) =>
              ref -> ValueParser.parseValue(v)
            }
            .toMap
          IO.pure(pairs)

        case (Right(range), multipleValues) =>
          IO.raiseError(
            new Exception(
              s"Range ${range.toA1} has ${range.cellCount} cells but ${multipleValues.length} values provided"
            )
          )

        case (Left(ref), multipleValues) =>
          IO.raiseError(
            new Exception(
              s"Cannot put ${multipleValues.length} values to single cell ${ref.toA1}"
            )
          )

      // Execute streaming transform
      result <- ZipTransformer.transformValues[IO](sourcePath, outputPath, worksheetPath, valueMap)
    yield
      val desc = refOrRange match
        case Left(ref) => s"Put ${values.headOption.getOrElse("")} → ${ref.toA1}"
        case Right(range) if values.length == 1 =>
          s"Filled ${range.toA1} with ${values.headOption.getOrElse("")}"
        case Right(range) => s"Put ${values.length} values to ${range.toA1}"
      s"$desc (streaming)\nCells modified: ${result.cellCount}\nSaved (streaming): $outputPath"

  /**
   * Streaming putf: write formulas to cells with O(1) memory.
   *
   * Uses SAX→StAX transform pipeline to modify only target cells. Note: Formula dragging is NOT
   * supported in streaming mode.
   *
   * @param sourcePath
   *   Input XLSX file
   * @param outputPath
   *   Output XLSX file
   * @param sheetNameOpt
   *   Sheet name (required for multi-sheet files)
   * @param refStr
   *   Cell reference or range
   * @param formulas
   *   Formulas to write (single for fill, multiple for batch)
   * @return
   *   Result message
   */
  def putFormula(
    sourcePath: Path,
    outputPath: Path,
    sheetNameOpt: Option[String],
    refStr: String,
    formulas: List[String]
  ): IO[String] =
    import com.tjclp.xl.addressing.ARef
    import com.tjclp.xl.cells.CellValue

    for
      // Resolve sheet path
      worksheetPath <- resolveSheetPath(sourcePath, sheetNameOpt)

      // Parse reference
      refOrRange <- IO.fromEither(
        CellRange
          .parse(refStr)
          .map(Right(_))
          .left
          .flatMap { _ =>
            ARef.parse(refStr).map(Left(_))
          }
          .left
          .map(e => new Exception(s"Invalid reference: $refStr"))
      )

      // Build formula map
      valueMap <- (refOrRange, formulas) match
        case (Left(ref), List(singleFormula)) =>
          val formula =
            if singleFormula.startsWith("=") then singleFormula.drop(1) else singleFormula
          IO.pure(Map(ref -> CellValue.Formula(formula, None)))

        case (Right(range), List(singleFormula)) =>
          // Fill pattern: all cells get same formula (NO dragging in streaming mode)
          val formula =
            if singleFormula.startsWith("=") then singleFormula.drop(1) else singleFormula
          IO.pure(range.cells.map(ref => ref -> CellValue.Formula(formula, None)).toMap)

        case (Right(range), multipleFormulas) if multipleFormulas.length == range.cellCount.toInt =>
          // Batch formulas
          val pairs = range.cellsRowMajor
            .zip(multipleFormulas.iterator)
            .map { (ref, f) =>
              val formula = if f.startsWith("=") then f.drop(1) else f
              ref -> CellValue.Formula(formula, None)
            }
            .toMap
          IO.pure(pairs)

        case (Right(range), multipleFormulas) =>
          IO.raiseError(
            new Exception(
              s"Range ${range.toA1} has ${range.cellCount} cells but ${multipleFormulas.length} formulas provided"
            )
          )

        case (Left(ref), multipleFormulas) =>
          IO.raiseError(
            new Exception(
              s"Cannot put ${multipleFormulas.length} formulas to single cell ${ref.toA1}"
            )
          )

      // Execute streaming transform
      result <- ZipTransformer.transformValues[IO](sourcePath, outputPath, worksheetPath, valueMap)
    yield
      val desc = refOrRange match
        case Left(ref) => s"Put formula → ${ref.toA1}"
        case Right(range) if formulas.length == 1 => s"Filled ${range.toA1} with formula"
        case Right(range) => s"Put ${formulas.length} formulas to ${range.toA1}"
      s"$desc (streaming, no dragging)\nCells modified: ${result.cellCount}\nSaved (streaming): $outputPath"

  /**
   * Streaming style: apply styling with O(1) memory.
   *
   * Uses SAX→StAX transform pipeline to modify only target cells.
   */
  def style(
    sourcePath: Path,
    outputPath: Path,
    sheetNameOpt: Option[String],
    rangeStr: String,
    bold: Boolean,
    italic: Boolean,
    underline: Boolean,
    bg: Option[String],
    fg: Option[String],
    fontSize: Option[Double],
    fontName: Option[String],
    align: Option[String],
    valign: Option[String],
    wrap: Boolean,
    numFormat: Option[String],
    border: Option[String],
    borderTop: Option[String],
    borderRight: Option[String],
    borderBottom: Option[String],
    borderLeft: Option[String],
    borderColor: Option[String],
    replace: Boolean
  ): IO[String] =
    for
      // Resolve sheet path
      worksheetPath <- resolveSheetPath(sourcePath, sheetNameOpt)

      // Parse range
      range <- IO.fromEither(
        CellRange.parse(rangeStr).left.map(e => new Exception(s"Invalid range: $rangeStr"))
      )

      // Build style
      cellStyle <- StyleBuilder.buildCellStyle(
        bold,
        italic,
        underline,
        bg,
        fg,
        fontSize,
        fontName,
        align,
        valign,
        wrap,
        numFormat,
        border,
        borderTop,
        borderRight,
        borderBottom,
        borderLeft,
        borderColor
      )

      // Execute streaming transform
      result <- ZipTransformer.transformStyle[IO](
        sourcePath,
        outputPath,
        worksheetPath,
        range,
        cellStyle,
        replace
      )

      // Build result message
      modeLabel = if replace then "(replace, streaming)" else "(merge, streaming)"
      appliedList = StyleBuilder.buildStyleDescription(
        bold,
        italic,
        underline,
        bg,
        fg,
        fontSize,
        fontName,
        align,
        valign,
        wrap,
        numFormat,
        border
      )
    yield s"Styled: ${range.toA1} $modeLabel\n" +
      s"Applied: ${appliedList.mkString(", ")}\n" +
      s"Cells modified: ${result.cellCount}\n" +
      s"Styles added: ${result.stylesAdded}\n" +
      s"Saved (streaming): $outputPath"

  /**
   * Streaming batch: apply batch operations with O(1) worksheet memory.
   *
   * Uses SAX→StAX transform pipeline with worksheet metadata injection for:
   *   - Cell patches (put, putf, style)
   *   - Column widths (colwidth) → `<cols>` element
   *   - Merged cells (merge, unmerge) → `<mergeCells>` element
   *   - Row properties (rowheight) → row attributes
   *
   * @param sourcePath
   *   Input XLSX file
   * @param outputPath
   *   Output XLSX file
   * @param sheetNameOpt
   *   Sheet name (required for multi-sheet files)
   * @param batchSource
   *   JSON file path or "-" for stdin
   * @return
   *   Result message with operation count
   */
  def batch(
    sourcePath: Path,
    outputPath: Path,
    sheetNameOpt: Option[String],
    batchSource: String
  ): IO[String] =
    for
      // Resolve worksheet path first
      worksheetPath <- resolveSheetPath(sourcePath, sheetNameOpt)

      // Read and parse batch input
      input <- BatchParser.readBatchInput(batchSource)
      parseResult <- BatchParser.parseBatchOperations(input)
      _ <- IO(parseResult.warnings.foreach(System.err.println))
      ops = parseResult.ops

      // Separate operations into cell patches vs worksheet metadata
      (cellPatches, stylesXml, worksheetMetadata, summary) <-
        buildStreamingBatchPatches(sourcePath, worksheetPath, ops)

      // Execute streaming transform with metadata
      result <- ZipTransformer.transformWithMetadata[IO](
        sourcePath,
        outputPath,
        worksheetPath,
        cellPatches,
        worksheetMetadata,
        stylesXml
      )
    yield s"Applied ${ops.size} operations (streaming):\n$summary\nCells modified: ${result.cellCount}\nSaved (streaming): $outputPath"

  /**
   * Build streaming batch patches from batch operations.
   *
   * Separates operations into:
   *   - Cell patches (SetValue, SetStyle, SetStyleAndValue)
   *   - Worksheet metadata (columns, merges, row properties)
   *   - Updated styles.xml (if style operations present)
   *
   * @return
   *   (cellPatches, updatedStylesXml, worksheetMetadata, summary)
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  private def buildStreamingBatchPatches(
    sourcePath: Path,
    worksheetPath: String,
    ops: Vector[BatchParser.BatchOp]
  ): IO[
    (
      Map[ARef, StreamingTransform.CellPatch],
      Option[String],
      StreamingTransform.WorksheetMetadata,
      String
    )
  ] =
    IO.delay {
      import scala.collection.mutable

      // Read current styles.xml
      val zipFile = new ZipFile(sourcePath.toFile)
      val stylesXml =
        try
          val stylesEntry = zipFile.getEntry("xl/styles.xml")
          if stylesEntry != null then
            new String(zipFile.getInputStream(stylesEntry).readAllBytes(), StandardCharsets.UTF_8)
          else minimalStylesXml
        finally zipFile.close()

      // Accumulate patches and metadata
      val cellPatches = mutable.Map[ARef, StreamingTransform.CellPatch]()
      val columns = mutable.Map[Column, ColumnProperties]()
      val addMerges = mutable.Set[CellRange]()
      val removeMerges = mutable.Set[CellRange]()
      val rowProps = mutable.Map[Row, RowProperties]()
      val summaryLines = mutable.ListBuffer[String]()

      var currentStylesXml = stylesXml
      var stylesModified = false

      ops.foreach {
        case BatchParser.BatchOp.Put(refStr, cellValue, formatOpt) =>
          val ref = ARef.parse(refStr) match
            case Right(r) => r
            case Left(e) => throw new Exception(s"Invalid ref '$refStr': $e")
          // For streaming mode with format, we'd need to add style to styles.xml
          // For now, streaming mode just sets the value; format requires full writer
          formatOpt match
            case Some(numFmt) =>
              // Build style with numFmt and add to styles.xml
              val cellStyle = CellStyle.default.withNumFmt(numFmt)
              val (updatedStyles, styleId) =
                StylePatcher.addStyle(currentStylesXml, cellStyle) match
                  case Right(result) => result
                  case Left(e) => throw new Exception(s"Failed to add style: ${e.message}")
              currentStylesXml = updatedStyles
              stylesModified = true
              cellPatches(ref) = StreamingTransform.CellPatch.SetStyleAndValue(styleId, cellValue)
            case None =>
              cellPatches(ref) =
                StreamingTransform.CellPatch.SetValue(cellValue, preserveStyle = true)
          summaryLines += s"  PUT $refStr = $cellValue"

        case BatchParser.BatchOp.PutFormula(refStr, formula) =>
          val ref = ARef.parse(refStr) match
            case Right(r) => r
            case Left(e) => throw new Exception(s"Invalid ref '$refStr': $e")
          val formulaText = if formula.startsWith("=") then formula.drop(1) else formula
          cellPatches(ref) = StreamingTransform.CellPatch.SetValue(
            CellValue.Formula(formulaText, None),
            preserveStyle = true
          )
          summaryLines += s"  PUTF $refStr = $formula"

        case BatchParser.BatchOp.PutFormulaDragging(rangeStr, formula, fromRef) =>
          // Parse formula and apply with shifting (same as non-streaming batch mode)
          val fromARef = ARef.parse(fromRef) match
            case Right(r) => r
            case Left(e) => throw new Exception(s"Invalid 'from' ref '$fromRef': $e")
          val range = CellRange.parse(rangeStr) match
            case Right(r) => r
            case Left(e) => throw new Exception(s"Invalid range '$rangeStr': $e")
          val formulaText = if formula.startsWith("=") then formula.drop(1) else formula
          val fullFormula = s"=$formulaText"

          // Parse formula for shifting
          val parsedExpr = FormulaParser.parse(fullFormula) match
            case Right(expr) => expr
            case Left(e) =>
              throw new Exception(
                s"Invalid formula '$fullFormula': ${ParseError.formatWithContext(e, fullFormula)}"
              )

          // Apply formula with shifting
          val startCol = Column.index0(fromARef.col)
          val startRow = Row.index0(fromARef.row)

          range.cells.foreach { targetRef =>
            val colDelta = Column.index0(targetRef.col) - startCol
            val rowDelta = Row.index0(targetRef.row) - startRow
            val shiftedExpr = FormulaShifter.shift(parsedExpr, colDelta, rowDelta)
            val shiftedFormula = FormulaPrinter.print(shiftedExpr, includeEquals = false)
            cellPatches(targetRef) = StreamingTransform.CellPatch.SetValue(
              CellValue.Formula(shiftedFormula, None),
              preserveStyle = true
            )
          }
          summaryLines += s"  PUTF $rangeStr = $formula (from $fromRef, ${range.cells.size} formulas)"

        case BatchParser.BatchOp.PutFormulas(rangeStr, formulas) =>
          val range = CellRange.parse(rangeStr) match
            case Right(r) => r
            case Left(e) => throw new Exception(s"Invalid range '$rangeStr': $e")
          val cells = range.cellsRowMajor.toVector
          if cells.length != formulas.length then
            throw new Exception(
              s"Range $rangeStr has ${cells.length} cells but ${formulas.length} formulas provided"
            )
          cells.zip(formulas).foreach { case (ref, formula) =>
            val formulaText = if formula.startsWith("=") then formula.drop(1) else formula
            cellPatches(ref) = StreamingTransform.CellPatch.SetValue(
              CellValue.Formula(formulaText, None),
              preserveStyle = true
            )
          }
          summaryLines += s"  PUTF $rangeStr = [${formulas.length} formulas]"

        case BatchParser.BatchOp.Style(rangeStr, props) =>
          val range = CellRange.parse(rangeStr) match
            case Right(r) => r
            case Left(e) => throw new Exception(s"Invalid range '$rangeStr': $e")

          // Build CellStyle from props using same approach as StyleBuilder
          val cellStyle = buildCellStyleFromPropsSync(props)

          // Add style to styles.xml and get ID
          val (updatedStyles, styleId) = StylePatcher.addStyle(currentStylesXml, cellStyle) match
            case Right(result) => result
            case Left(e) => throw new Exception(s"Failed to add style: ${e.message}")
          currentStylesXml = updatedStyles
          stylesModified = true

          // Apply to all cells in range, merging with existing patches
          range.cells.foreach { ref =>
            cellPatches.get(ref) match
              case Some(StreamingTransform.CellPatch.SetValue(value, _)) =>
                // Cell already has a value patch - convert to SetStyleAndValue
                cellPatches(ref) = StreamingTransform.CellPatch.SetStyleAndValue(styleId, value)
              case Some(StreamingTransform.CellPatch.SetStyleAndValue(_, value)) =>
                // Cell already has SetStyleAndValue - update the style
                cellPatches(ref) = StreamingTransform.CellPatch.SetStyleAndValue(styleId, value)
              case Some(StreamingTransform.CellPatch.SetStyle(_)) =>
                // Cell already has style - replace it
                cellPatches(ref) = StreamingTransform.CellPatch.SetStyle(styleId)
              case None =>
                // No existing patch - just set style
                cellPatches(ref) = StreamingTransform.CellPatch.SetStyle(styleId)
          }
          summaryLines += s"  STYLE $rangeStr"

        case BatchParser.BatchOp.Merge(rangeStr) =>
          val range = CellRange.parse(rangeStr) match
            case Right(r) => r
            case Left(e) => throw new Exception(s"Invalid range '$rangeStr': $e")
          addMerges += range
          summaryLines += s"  MERGE $rangeStr"

        case BatchParser.BatchOp.Unmerge(rangeStr) =>
          val range = CellRange.parse(rangeStr) match
            case Right(r) => r
            case Left(e) => throw new Exception(s"Invalid range '$rangeStr': $e")
          removeMerges += range
          summaryLines += s"  UNMERGE $rangeStr"

        case BatchParser.BatchOp.ColWidth(colStr, width) =>
          val col = Column.fromLetter(colStr) match
            case Right(c) => c
            case Left(e) => throw new Exception(s"Invalid column '$colStr': $e")
          columns(col) = columns.getOrElse(col, ColumnProperties()).copy(width = Some(width))
          summaryLines += s"  COLWIDTH $colStr = $width"

        case BatchParser.BatchOp.RowHeight(rowNum, height) =>
          val row = Row.from1(rowNum)
          rowProps(row) = rowProps.getOrElse(row, RowProperties()).copy(height = Some(height))
          summaryLines += s"  ROWHEIGHT $rowNum = $height"
      }

      val worksheetMetadata = StreamingTransform.WorksheetMetadata(
        columns = columns.toMap,
        addMerges = addMerges.toSet,
        removeMerges = removeMerges.toSet,
        rowProps = rowProps.toMap
      )

      val updatedStylesXml = if stylesModified then Some(currentStylesXml) else None

      (cellPatches.toMap, updatedStylesXml, worksheetMetadata, summaryLines.mkString("\n"))
    }

  /**
   * Build CellStyle from batch StyleProps (synchronous version for use inside IO.delay).
   *
   * Uses the same approach as StyleBuilder but without IO wrapping.
   */
  private def buildCellStyleFromPropsSync(props: BatchParser.StyleProps): CellStyle =
    import scala.util.chaining.*
    import com.tjclp.xl.cli.ColorParser
    import com.tjclp.xl.styles.alignment.{Align, HAlign, VAlign}
    import com.tjclp.xl.styles.border.{Border, BorderSide, BorderStyle}
    import com.tjclp.xl.styles.fill.Fill
    import com.tjclp.xl.styles.font.Font
    import com.tjclp.xl.styles.numfmt.NumFmt

    // Parse colors
    val bgColor = props.bg.flatMap(s => ColorParser.parse(s).toOption)
    val fgColor = props.fg.flatMap(s => ColorParser.parse(s).toOption)
    val bdrColor = props.borderColor.flatMap(s => ColorParser.parse(s).toOption)

    // Parse alignments
    val hAlign = props.align.flatMap(parseHAlignSync)
    val vAlign = props.valign.flatMap(parseVAlignSync)

    // Parse border styles
    val bdrStyle = props.border.flatMap(parseBorderStyleSync)
    val bdrTopStyle = props.borderTop.flatMap(parseBorderStyleSync)
    val bdrRightStyle = props.borderRight.flatMap(parseBorderStyleSync)
    val bdrBottomStyle = props.borderBottom.flatMap(parseBorderStyleSync)
    val bdrLeftStyle = props.borderLeft.flatMap(parseBorderStyleSync)

    // Parse number format
    val nFmt = props.numFormat.flatMap(parseNumFmtSync)

    // Build font
    val font = Font.default
      .withBold(props.bold)
      .withItalic(props.italic)
      .withUnderline(props.underline)
      .pipe(f => fgColor.fold(f)(c => f.withColor(c)))
      .pipe(f => props.fontSize.fold(f)(s => f.withSize(s)))
      .pipe(f => props.fontName.fold(f)(n => f.withName(n)))

    // Build fill
    val fill = bgColor.map(Fill.Solid.apply).getOrElse(Fill.None)

    // Build border
    val cellBorder =
      buildBorderSync(bdrStyle, bdrTopStyle, bdrRightStyle, bdrBottomStyle, bdrLeftStyle, bdrColor)

    // Build alignment
    val alignment = Align.default
      .pipe(a => hAlign.fold(a)(h => a.withHAlign(h)))
      .pipe(a => vAlign.fold(a)(v => a.withVAlign(v)))
      .pipe(a => if props.wrap then a.withWrap() else a)

    CellStyle(
      font = font,
      fill = fill,
      border = cellBorder,
      numFmt = nFmt.getOrElse(NumFmt.General),
      align = alignment
    )

  private def buildBorderSync(
    allSides: Option[com.tjclp.xl.styles.border.BorderStyle],
    top: Option[com.tjclp.xl.styles.border.BorderStyle],
    right: Option[com.tjclp.xl.styles.border.BorderStyle],
    bottom: Option[com.tjclp.xl.styles.border.BorderStyle],
    left: Option[com.tjclp.xl.styles.border.BorderStyle],
    color: Option[com.tjclp.xl.styles.color.Color]
  ): com.tjclp.xl.styles.border.Border =
    import com.tjclp.xl.styles.border.{Border, BorderSide, BorderStyle}
    val base = allSides.getOrElse(BorderStyle.None)
    val topSide = BorderSide(top.getOrElse(base), color)
    val rightSide = BorderSide(right.getOrElse(base), color)
    val bottomSide = BorderSide(bottom.getOrElse(base), color)
    val leftSide = BorderSide(left.getOrElse(base), color)
    if topSide.style == BorderStyle.None && rightSide.style == BorderStyle.None &&
      bottomSide.style == BorderStyle.None && leftSide.style == BorderStyle.None
    then Border.none
    else Border(left = leftSide, right = rightSide, top = topSide, bottom = bottomSide)

  private def parseHAlignSync(s: String): Option[com.tjclp.xl.styles.alignment.HAlign] =
    import com.tjclp.xl.styles.alignment.HAlign
    s.toLowerCase match
      case "left" => Some(HAlign.Left)
      case "center" => Some(HAlign.Center)
      case "right" => Some(HAlign.Right)
      case "justify" => Some(HAlign.Justify)
      case "general" => Some(HAlign.General)
      case _ => None

  private def parseVAlignSync(s: String): Option[com.tjclp.xl.styles.alignment.VAlign] =
    import com.tjclp.xl.styles.alignment.VAlign
    s.toLowerCase match
      case "top" => Some(VAlign.Top)
      case "middle" | "center" => Some(VAlign.Middle)
      case "bottom" => Some(VAlign.Bottom)
      case _ => None

  private def parseBorderStyleSync(s: String): Option[com.tjclp.xl.styles.border.BorderStyle] =
    import com.tjclp.xl.styles.border.BorderStyle
    s.toLowerCase match
      case "none" => Some(BorderStyle.None)
      case "thin" => Some(BorderStyle.Thin)
      case "medium" => Some(BorderStyle.Medium)
      case "thick" => Some(BorderStyle.Thick)
      case "dashed" => Some(BorderStyle.Dashed)
      case "dotted" => Some(BorderStyle.Dotted)
      case "double" => Some(BorderStyle.Double)
      case _ => None

  private def parseNumFmtSync(s: String): Option[com.tjclp.xl.styles.numfmt.NumFmt] =
    import com.tjclp.xl.styles.numfmt.NumFmt
    s.toLowerCase match
      case "general" => Some(NumFmt.General)
      case "number" => Some(NumFmt.Decimal)
      case "currency" => Some(NumFmt.Currency)
      case "percent" => Some(NumFmt.Percent)
      case "date" => Some(NumFmt.Date)
      case "text" => Some(NumFmt.Text)
      case _ => None

  /** Minimal styles.xml for files that don't have one. */
  private val minimalStylesXml: String =
    """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      |<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      |<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
      |<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
      |<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
      |<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
      |<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
      |</styleSheet>""".stripMargin.replaceAll("\n", "")

  /**
   * Resolve worksheet path from name by parsing workbook.xml and workbook.xml.rels metadata.
   *
   * Lightweight operation - only reads workbook.xml/workbook.xml.rels, not worksheet data.
   */
  private val relsNamespace =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

  private def resolveSheetPath(
    sourcePath: Path,
    sheetNameOpt: Option[String]
  ): IO[String] =
    IO.delay {
      val zipFile = new ZipFile(sourcePath.toFile)
      try
        val wbEntry = zipFile.getEntry("xl/workbook.xml")
        if wbEntry == null then throw new Exception("Invalid XLSX: missing xl/workbook.xml")

        val content =
          new String(zipFile.getInputStream(wbEntry).readAllBytes(), StandardCharsets.UTF_8)
        XmlSecurity.parseSafe(content, "xl/workbook.xml") match
          case Left(err) =>
            throw new Exception(s"Failed to parse workbook.xml: ${err.message}")
          case Right(wbXml) =>
            val sheets = (wbXml \\ "sheet").collect { case elem: Elem => elem }.toSeq

            if sheets.isEmpty then throw new Exception("Workbook has no sheets")

            val targetSheet = sheetNameOpt match
              case None =>
                if sheets.size > 1 then
                  val names = sheets.map(s => (s \ "@name").text).mkString(", ")
                  throw new Exception(
                    s"Multiple sheets found: $names. Use --sheet to specify which sheet to modify."
                  )
                sheets.head
              case Some(targetName) =>
                val found = sheets.find { sheetElem =>
                  (sheetElem \ "@name").text == targetName
                }
                found.getOrElse {
                  val names = sheets.map(s => (s \ "@name").text).mkString(", ")
                  throw new Exception(
                    s"Sheet '$targetName' not found. Available sheets: $names"
                  )
                }

            val sheetIdStr = targetSheet \@ "sheetId"
            val sheetId = sheetIdStr.toIntOption.getOrElse {
              throw new Exception(s"Invalid sheetId: $sheetIdStr")
            }
            val rId = targetSheet.attribute(relsNamespace, "id").map(_.text).getOrElse("")

            val relsEntry = Option(zipFile.getEntry("xl/_rels/workbook.xml.rels"))
            val relsMap = relsEntry match
              case None => Map.empty[String, String]
              case Some(entry) =>
                val relsXml =
                  new String(zipFile.getInputStream(entry).readAllBytes(), StandardCharsets.UTF_8)
                XmlSecurity.parseSafe(relsXml, "xl/workbook.xml.rels") match
                  case Left(_) => Map.empty[String, String]
                  case Right(relsElem) =>
                    (relsElem \\ "Relationship").collect { case elem: Elem =>
                      val id = elem \@ "Id"
                      val target = elem \@ "Target"
                      id -> target
                    }.toMap

            val targetPath = relsMap.getOrElse(rId, s"xl/worksheets/sheet$sheetId.xml")
            val normalizedPath =
              val trimmed = if targetPath.startsWith("/") then targetPath.drop(1) else targetPath
              if trimmed.startsWith("xl/") then trimmed else s"xl/$trimmed"

            if zipFile.getEntry(normalizedPath) == null then
              throw new Exception(s"Worksheet not found: $normalizedPath")

            normalizedPath
      finally zipFile.close()
    }
