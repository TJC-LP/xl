package com.tjclp.xl.cli.commands

import java.nio.charset.StandardCharsets
import java.nio.file.Path
import java.util.zip.ZipFile

import cats.effect.IO
import com.tjclp.xl.api.Workbook
import com.tjclp.xl.addressing.{ARef, CellRange, Column, RefType, Row, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.error.XLError
import com.tjclp.xl.formula.{FormulaParser, FormulaPrinter, FormulaShifter, ParseError}
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.io.streaming.{StreamingTransform, StylePatcher, ZipTransformer}
import com.tjclp.xl.ooxml.XmlSecurity
import com.tjclp.xl.ooxml.metadata.WorkbookMetadataReader
import com.tjclp.xl.ooxml.writer.WriterConfig
import com.tjclp.xl.sheets.{ColumnProperties, RowProperties}
import com.tjclp.xl.styles.units.StyleId
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.numfmt.NumFmt
import com.tjclp.xl.cli.CliIO
import com.tjclp.xl.cli.batch.{FormatHint, OpRegistry, ScopedOp}
import com.tjclp.xl.cli.contract.{CliError, CliException, ErrorCode, Location}
import com.tjclp.xl.cli.helpers.{
  BatchParser,
  Resolve,
  StreamingCsvParser,
  StyleBuilder,
  ValueParser
}
import org.xml.sax.{Attributes, SAXException}
import org.xml.sax.helpers.DefaultHandler
import scala.collection.mutable
import scala.xml.Elem

/**
 * Streaming write command handlers.
 *
 * Provides two modes:
 *   1. True streaming (CSV import): End-to-end O(1) memory
 *   2. SAX/StAX workbook write: In-memory workbook → lower-allocation full OOXML writer
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
   * Takes an in-memory workbook and writes it using the SAX/StAX OOXML backend. Styles and workbook
   * metadata are preserved via the full writer pipeline.
   *
   * Best for: Modified workbooks that need the full OOXML feature set with lower writer allocation.
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
    values: List[String],
    detect: Boolean = true
  ): IO[String] =
    for
      // Parse the reference (single cell or range, optionally sheet-qualified)
      (qualified, refOrRange) <- parseTarget(refStr, "reference")

      // THE sheet rule over workbook.xml: qualifier > -s > the only sheet > SHEET_REQUIRED
      worksheetPath <- resolveSheetPath(sourcePath, sheetNameOpt, qualified, "put")

      // Build value map based on mode
      parsedValues <- (refOrRange, values) match
        case (Left(ref), List(singleValue)) =>
          // Single cell
          IO.pure(Vector(ref -> ValueParser.parsePutValue(singleValue, detect)))

        case (Right(range), List(singleValue)) =>
          // Fill pattern: all cells get same value
          val value = ValueParser.parsePutValue(singleValue, detect)
          IO.pure(range.cells.map(ref => ref -> value).toVector)

        case (Right(range), multipleValues) if multipleValues.length == range.cellCount.toInt =>
          // Batch values: exact count match
          val pairs = range.cellsRowMajor
            .zip(multipleValues.iterator)
            .map { (ref, v) =>
              ref -> ValueParser.parsePutValue(v, detect)
            }
            .toVector
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

      // Detected formats are Inferred hints: they apply onto General cells only (Sheet.put's rule)
      batchOps = parsedValues.zipWithIndex.map { case ((ref, formatted), i) =>
        val format = Option.when(formatted.numFmt != NumFmt.General)(formatted.numFmt)
        ScopedOp(BatchParser.BatchOp.Put(ref.toA1, formatted.value, format), None, i + 1)
      }
      patches <- buildStreamingBatchPatches(sourcePath, worksheetPath, batchOps)
      (cellPatches, updatedStylesXml, worksheetMetadata, _) = patches

      // Execute streaming transform, including any detected number formats.
      result <- ZipTransformer.transformWithMetadata[IO](
        sourcePath,
        outputPath,
        worksheetPath,
        cellPatches,
        worksheetMetadata,
        updatedStylesXml
      )
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
      // Parse the reference (single cell or range, optionally sheet-qualified)
      (qualified, refOrRange) <- parseTarget(refStr, "reference")

      // THE sheet rule over workbook.xml: qualifier > -s > the only sheet > SHEET_REQUIRED
      worksheetPath <- resolveSheetPath(sourcePath, sheetNameOpt, qualified, "putf")

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
      // Parse the range (optionally sheet-qualified; a single cell is a 1x1 range)
      (qualified, target) <- parseTarget(rangeStr, "range")
      range = target.fold(ref => CellRange(ref, ref), identity)

      // THE sheet rule over workbook.xml: qualifier > -s > the only sheet > SHEET_REQUIRED
      worksheetPath <- resolveSheetPath(sourcePath, sheetNameOpt, qualified, "style")

      // GH-475: same typo signal as the in-memory style command
      _ <- StyleBuilder.warnNumFmt(numFormat)

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
      modeLabel = if replace then "(replace, streaming)" else "(additive, streaming)"
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
   * @param stdin
   *   Where `batchSource == "-"` reads from (the CLI passes its `CliIO.stdin`)
   * @return
   *   Result message with operation count
   */
  def batch(
    sourcePath: Path,
    outputPath: Path,
    sheetNameOpt: Option[String],
    batchSource: String,
    stdin: IO[String] = CliIO.system.stdin
  ): IO[String] =
    for
      // Parse first: the refusals below come before the worksheet is even resolved
      input <- BatchParser.readBatchInput(batchSource, stdin)
      parseResult <- BatchParser.parseBatchOperations(input)
      _ <- IO(parseResult.warnings.foreach(System.err.println))
      scoped = parseResult.scoped

      // ADR-017 invariant 2: refuse by index what this writer cannot apply, before any byte is
      // read beyond workbook.xml or written
      _ <- refuseNonStreamable(scoped)

      // Resolve the worksheet (workbook.xml only: -s, else the only sheet, else SHEET_REQUIRED),
      // then refuse ops aimed at any other sheet
      (sheetName, worksheetPath) <- resolveSheetTarget(sourcePath, sheetNameOpt, None, "batch")
      _ <- refuseOtherSheets(scoped, sheetName)

      // A ref qualified with the streamed sheet is honoured as its bare form
      ops = scoped.map(s => s.copy(op = OpRegistry.unqualified(s.op)))

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
   * Whether [[buildStreamingBatchPatches]] has an arm for the op. Deliberately exhaustive with no
   * wildcard, like `WriteCommands.isCellMutating`: a new `BatchOp` must be classified here and in
   * `OpRegistry` — `OpRegistrySpec` asserts the two agree — or the match fails loudly. Dragging
   * `putf` IS streamable: the arm shifts exactly like the in-memory path (StreamingWriteSpec).
   */
  private[cli] def isStreamable(op: BatchParser.BatchOp): Boolean =
    op match
      case _: BatchParser.BatchOp.Put | _: BatchParser.BatchOp.PutFormula |
          _: BatchParser.BatchOp.PutFormulaDragging | _: BatchParser.BatchOp.PutFormulas |
          _: BatchParser.BatchOp.PutValues | _: BatchParser.BatchOp.Style |
          _: BatchParser.BatchOp.Merge | _: BatchParser.BatchOp.Unmerge |
          _: BatchParser.BatchOp.ColWidth | _: BatchParser.BatchOp.RowHeight |
          _: BatchParser.BatchOp.ColHide | _: BatchParser.BatchOp.ColShow |
          _: BatchParser.BatchOp.RowHide | _: BatchParser.BatchOp.RowShow =>
        true
      case _: BatchParser.BatchOp.AddComment | _: BatchParser.BatchOp.RemoveComment |
          _: BatchParser.BatchOp.Clear | _: BatchParser.BatchOp.AutoFit |
          _: BatchParser.BatchOp.AddSheet | _: BatchParser.BatchOp.RenameSheet |
          _: BatchParser.BatchOp.Freeze | BatchParser.BatchOp.Unfreeze |
          _: BatchParser.BatchOp.CopyRange | _: BatchParser.BatchOp.Hyperlink |
          _: BatchParser.BatchOp.AddChart | _: BatchParser.BatchOp.SetSheetView |
          _: BatchParser.BatchOp.SetTabColor | _: BatchParser.BatchOp.SetAutoFilter |
          _: BatchParser.BatchOp.GroupRows | _: BatchParser.BatchOp.GroupCols |
          _: BatchParser.BatchOp.UngroupRows | _: BatchParser.BatchOp.UngroupCols |
          _: BatchParser.BatchOp.SetPageSetup | _: BatchParser.BatchOp.SetHeaderFooter |
          _: BatchParser.BatchOp.AddConditionalFormat =>
        false

  /** `UNSUPPORTED_IN_STREAM` (exit 2) listing the offending ops as `[index op, …]`. */
  private def refuse(rejects: Vector[ScopedOp], reason: String): IO[Unit] =
    rejects.headOption match
      case None => IO.unit
      case Some(first) =>
        val listed = rejects.map(s => s"${s.index} ${OpRegistry.nameOf(s.op)}").mkString(", ")
        IO.raiseError(
          CliException(
            CliError(
              ErrorCode.UNSUPPORTED_IN_STREAM,
              s"ops [$listed] $reason; drop --stream to apply them in memory",
              hint = Some("drop --stream to apply them in memory"),
              location = Some(Location.none.copy(opIndex = Some(first.index)))
            )
          )
        )

  /** Ops this writer has no arm for. */
  private def refuseNonStreamable(scoped: Vector[ScopedOp]): IO[Unit] =
    refuse(scoped.filterNot(s => isStreamable(s.op)), "are not supported in streaming mode")

  /** Ops whose `sheet` key or qualified target ref names a sheet other than the streamed one. */
  private def refuseOtherSheets(scoped: Vector[ScopedOp], streamed: String): IO[Unit] =
    def namedSheets(s: ScopedOp): Vector[String] =
      val declared = s.sheet.toList.map(_.value)
      val qualified = OpRegistry
        .targetRefs(s.op)
        .flatMap(r => OpRegistry.qualifiedSheet(r).toList.map(_.value))
      (declared ++ qualified).toVector
    val rejects = scoped.filter(s => namedSheets(s).exists(_ != streamed))
    val others = rejects.flatMap(namedSheets).filter(_ != streamed).distinct
    refuse(
      rejects,
      s"target ${others.mkString(", ")} rather than the streamed worksheet '$streamed', " +
        "which is not supported in streaming mode"
    )

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
    ops: Vector[ScopedOp]
  ): IO[
    (
      Map[ARef, StreamingTransform.CellPatch],
      Option[String],
      StreamingTransform.WorksheetMetadata,
      String
    )
  ] =
    IO.delay {
      // Read current styles.xml
      val zipFile = new ZipFile(sourcePath.toFile)
      val stylesXml =
        try
          val stylesEntry = zipFile.getEntry("xl/styles.xml")
          if stylesEntry != null then
            new String(zipFile.getInputStream(stylesEntry).readAllBytes(), StandardCharsets.UTF_8)
          else minimalStylesXml
        finally zipFile.close()

      val plainOps = ops.map(_.op)
      val needsColumnMetadata = plainOps.exists {
        case BatchParser.BatchOp.ColWidth(_, _) | BatchParser.BatchOp.ColHide(_) |
            BatchParser.BatchOp.ColShow(_) =>
          true
        case _ => false
      }
      val targetRows = plainOps.collect {
        case BatchParser.BatchOp.RowHeight(rowNum, _) => Row.from1(rowNum)
        case BatchParser.BatchOp.RowHide(rowNum) => Row.from1(rowNum)
        case BatchParser.BatchOp.RowShow(rowNum) => Row.from1(rowNum)
      }.toSet
      val existingMetadata =
        readExistingWorksheetMetadata(sourcePath, worksheetPath, needsColumnMetadata, targetRows)
      // Every cell a formatted write lands on: its existing xf decides the outcome (GH-560), so
      // read those xfs once, without materializing the worksheet.
      val formattedRefs = plainOps.flatMap(formattedTargets).toSet
      val existingStyles =
        if formattedRefs.nonEmpty then scanExistingStyles(sourcePath, worksheetPath, formattedRefs)
        else Map.empty[ARef, Int]

      // Accumulate patches and metadata
      val cellPatches = mutable.Map[ARef, StreamingTransform.CellPatch]()
      val columns = mutable.Map[Column, ColumnProperties]() ++ existingMetadata.columns
      val addMerges = mutable.Set[CellRange]()
      val removeMerges = mutable.Set[CellRange]()
      val rowProps = mutable.Map[Row, RowProperties]() ++ existingMetadata.rowProps
      val summaryLines = mutable.ListBuffer[String]()
      val existingStyleCache = mutable.Map[Int, CellStyle]()
      val addedStyleIds = mutable.Map[String, Int]()

      var currentStylesXml = stylesXml
      var stylesModified = false

      def existingStyleOf(ref: ARef): (Option[Int], CellStyle) =
        val styleIdOpt = existingStyles.get(ref)
        val style = styleIdOpt
          .map { styleId =>
            existingStyleCache.getOrElseUpdate(
              styleId,
              StylePatcher.getStyle(stylesXml, styleId) match
                case Right(Some(style)) => style
                case Right(None) =>
                  throw new Exception(s"Existing style ID $styleId was not found")
                case Left(e) =>
                  throw new Exception(s"Failed to read existing style: ${e.message}")
            )
          }
          .getOrElse(CellStyle.default)
        (styleIdOpt, style)

      def registerStyle(cellStyle: CellStyle): Int =
        addedStyleIds.getOrElseUpdate(
          cellStyle.canonicalKey,
          StylePatcher.addStyle(currentStylesXml, cellStyle) match
            case Right((updatedStyles, addedStyleId)) =>
              currentStylesXml = updatedStyles
              stylesModified = true
              addedStyleId
            case Left(e) => throw new Exception(s"Failed to add style: ${e.message}")
        )

      /**
       * The xf a formatted write lands on — identical to the in-memory path (ADR-017 invariant 4,
       * GH-560): an Explicit hint replaces the numFmt on the cell's existing style, keeping its
       * font, fill and borders (the `applyNumFmt` path); an Inferred one applies only onto a
       * General cell and otherwise leaves the cell's xf as it is (`Sheet.put`'s merge rule).
       */
      def formattedStyleId(ref: ARef, numFmt: NumFmt, hint: FormatHint): Int =
        val (existingId, existingStyle) = existingStyleOf(ref)
        (hint, existingId) match
          case (FormatHint.Inferred, Some(id)) if existingStyle.numFmt != NumFmt.General => id
          case _ => registerStyle(existingStyle.withNumFmt(numFmt))

      ops.foreach { scoped =>
        val hint = scoped.hint
        scoped.op match
          case BatchParser.BatchOp.Put(refStr, cellValue, formatOpt) =>
            val ref = ARef.parse(refStr) match
              case Right(r) => r
              case Left(e) => throw new Exception(s"Invalid ref '$refStr': $e")
            formatOpt match
              case Some(numFmt) =>
                val styleId = formattedStyleId(ref, numFmt, hint)
                cellPatches(ref) = StreamingTransform.CellPatch.SetStyleAndValue(styleId, cellValue)
              case None =>
                cellPatches(ref) =
                  StreamingTransform.CellPatch.SetValue(cellValue, preserveStyle = true)
            summaryLines += s"  PUT $refStr = $cellValue"

          case BatchParser.BatchOp.PutFormula(refStr, formula, formatOpt) =>
            val ref = ARef.parse(refStr) match
              case Right(r) => r
              case Left(e) => throw new Exception(s"Invalid ref '$refStr': $e")
            val formulaText = if formula.startsWith("=") then formula.drop(1) else formula
            val formulaValue = CellValue.Formula(formulaText, None)
            formatOpt match
              case Some(numFmt) =>
                // GH-356: a putf format is always explicit — replace the numFmt, keep the font
                val styleId = formattedStyleId(ref, numFmt, FormatHint.Explicit)
                cellPatches(ref) =
                  StreamingTransform.CellPatch.SetStyleAndValue(styleId, formulaValue)
              case None =>
                cellPatches(ref) =
                  StreamingTransform.CellPatch.SetValue(formulaValue, preserveStyle = true)
            summaryLines += s"  PUTF $refStr = $formula"

          case BatchParser.BatchOp.PutFormulaDragging(rangeStr, formula, fromRef, formatOpt) =>
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

            // Apply formula with shifting; GH-356: the explicit format lands on each cell's own xf
            val startCol = Column.index0(fromARef.col)
            val startRow = Row.index0(fromARef.row)

            range.cells.foreach { targetRef =>
              val colDelta = Column.index0(targetRef.col) - startCol
              val rowDelta = Row.index0(targetRef.row) - startRow
              val shiftedExpr = FormulaShifter.shift(parsedExpr, colDelta, rowDelta)
              val shiftedFormula = FormulaPrinter.printFileForm(shiftedExpr)
              val formulaValue = CellValue.Formula(shiftedFormula, None)
              cellPatches(targetRef) = formatOpt match
                case Some(numFmt) =>
                  val styleId = formattedStyleId(targetRef, numFmt, FormatHint.Explicit)
                  StreamingTransform.CellPatch.SetStyleAndValue(styleId, formulaValue)
                case None =>
                  StreamingTransform.CellPatch.SetValue(formulaValue, preserveStyle = true)
            }
            summaryLines += s"  PUTF $rangeStr = $formula (from $fromRef, ${range.cells.size} formulas)"

          case BatchParser.BatchOp.PutFormulas(rangeStr, formulas, formatOpt) =>
            val range = CellRange.parse(rangeStr) match
              case Right(r) => r
              case Left(e) => throw new Exception(s"Invalid range '$rangeStr': $e")
            val cells = range.cellsRowMajor.toVector
            if cells.length != formulas.length then
              throw new Exception(
                s"Range $rangeStr has ${cells.length} cells but ${formulas.length} formulas provided"
              )
            // GH-356: the explicit format lands on each cell's own xf (font kept, numFmt replaced)
            cells.zip(formulas).foreach { case (ref, formula) =>
              val formulaText = if formula.startsWith("=") then formula.drop(1) else formula
              val formulaValue = CellValue.Formula(formulaText, None)
              cellPatches(ref) = formatOpt match
                case Some(numFmt) =>
                  val styleId = formattedStyleId(ref, numFmt, FormatHint.Explicit)
                  StreamingTransform.CellPatch.SetStyleAndValue(styleId, formulaValue)
                case None =>
                  StreamingTransform.CellPatch.SetValue(formulaValue, preserveStyle = true)
            }
            summaryLines += s"  PUTF $rangeStr = [${formulas.length} formulas]"

          case BatchParser.BatchOp.PutValues(rangeStr, values) =>
            val range = CellRange.parse(rangeStr) match
              case Right(r) => r
              case Left(e) => throw new Exception(s"Invalid range '$rangeStr': $e")
            val cells = range.cellsRowMajor.toVector
            if cells.length != values.length then
              throw new Exception(
                s"Range $rangeStr has ${cells.length} cells but ${values.length} values provided"
              )
            // The op-level `format` (GH-416) is the Explicit hint for every element; detected
            // formats inside the array are Inferred (GH-560)
            cells.zip(values).foreach { case (ref, pv) =>
              pv.format match
                case Some(numFmt) =>
                  val styleId = formattedStyleId(ref, numFmt, hint)
                  cellPatches(ref) =
                    StreamingTransform.CellPatch.SetStyleAndValue(styleId, pv.cellValue)
                case None =>
                  cellPatches(ref) =
                    StreamingTransform.CellPatch.SetValue(pv.cellValue, preserveStyle = true)
            }
            summaryLines += s"  PUT $rangeStr = [${values.length} values]"

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

          case BatchParser.BatchOp.ColHide(colStr) =>
            val col = Column.fromLetter(colStr) match
              case Right(c) => c
              case Left(e) => throw new Exception(s"Invalid column '$colStr': $e")
            val existing = columns.getOrElse(col, ColumnProperties())
            columns(col) = existing.copy(hidden = true)
            summaryLines += s"  COL-HIDE $colStr"

          case BatchParser.BatchOp.ColShow(colStr) =>
            val col = Column.fromLetter(colStr) match
              case Right(c) => c
              case Left(e) => throw new Exception(s"Invalid column '$colStr': $e")
            val existing = columns.getOrElse(col, ColumnProperties())
            columns(col) = existing.copy(hidden = false)
            summaryLines += s"  COL-SHOW $colStr"

          case BatchParser.BatchOp.RowHide(rowNum) =>
            val row = Row.from1(rowNum)
            val existing = rowProps.getOrElse(row, RowProperties())
            rowProps(row) = existing.copy(hidden = true)
            summaryLines += s"  ROW-HIDE $rowNum"

          case BatchParser.BatchOp.RowShow(rowNum) =>
            val row = Row.from1(rowNum)
            val existing = rowProps.getOrElse(row, RowProperties())
            rowProps(row) = existing.copy(hidden = false)
            summaryLines += s"  ROW-SHOW $rowNum"

          // Every op `isStreamable` rejects was refused by `batch` before any byte was written
          // (UNSUPPORTED_IN_STREAM, by index). Should a caller ever bypass that guard, the failure
          // stays typed — the same code and hint the guard raises, never a stack trace.
          case other =>
            throw CliException(
              CliError(
                ErrorCode.UNSUPPORTED_IN_STREAM,
                s"op '${OpRegistry.nameOf(other)}' is not supported in streaming mode; " +
                  "drop --stream to apply it in memory",
                hint = Some("drop --stream to apply it in memory")
              )
            )
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
   * The cells whose existing style a formatted write consults: a `put`/`putf` with a format, a
   * dragged or listed formula range with a format, a `values` array with any formatted element.
   * Unparseable refs are left for the arm to report.
   */
  private def formattedTargets(op: BatchParser.BatchOp): Vector[ARef] =
    def cellsOf(rangeStr: String): Vector[ARef] =
      CellRange.parse(rangeStr).toOption.toList.flatMap(_.cells).toVector
    def cellOf(refStr: String): Vector[ARef] = ARef.parse(refStr).toOption.toList.toVector
    op match
      case BatchParser.BatchOp.Put(refStr, _, Some(_)) => cellOf(refStr)
      case BatchParser.BatchOp.PutValues(rangeStr, values) if values.exists(_.format.isDefined) =>
        cellsOf(rangeStr)
      case BatchParser.BatchOp.PutFormula(refStr, _, Some(_)) => cellOf(refStr)
      case BatchParser.BatchOp.PutFormulaDragging(rangeStr, _, _, Some(_)) => cellsOf(rangeStr)
      case BatchParser.BatchOp.PutFormulas(rangeStr, _, Some(_)) => cellsOf(rangeStr)
      case _ => Vector.empty

  /** Read style IDs for selected cells without materializing the worksheet. */
  private def scanExistingStyles(
    sourcePath: Path,
    worksheetPath: String,
    refs: Set[ARef]
  ): Map[ARef, Int] =
    val normalizedPath =
      val trimmed = worksheetPath.stripPrefix("/")
      if trimmed.startsWith("xl/") then trimmed else s"xl/$trimmed"
    val zipFile = new ZipFile(sourcePath.toFile)
    try
      Option(zipFile.getEntry(normalizedPath)) match
        case Some(entry) =>
          StreamingTransform.scanExistingStyles(zipFile.getInputStream(entry), refs)
        case None => throw new Exception(s"Worksheet not found: $normalizedPath")
    finally zipFile.close()

  private final case class ExistingWorksheetMetadata(
    columns: Map[Column, ColumnProperties],
    rowProps: Map[Row, RowProperties]
  )

  private final class StopWorksheetScan extends SAXException("stop worksheet scan")

  @SuppressWarnings(Array("org.wartremover.warts.Null"))
  private def readExistingWorksheetMetadata(
    sourcePath: Path,
    worksheetPath: String,
    needsColumns: Boolean,
    targetRows: Set[Row]
  ): ExistingWorksheetMetadata =
    if !needsColumns && targetRows.isEmpty then ExistingWorksheetMetadata(Map.empty, Map.empty)
    else
      val parsedColumns = mutable.Map.empty[Column, ColumnProperties]
      val parsedRows = mutable.Map.empty[Row, RowProperties]
      val pendingRows = mutable.Set.empty[Int] ++ targetRows.map(_.index1)

      val zipFile = new ZipFile(sourcePath.toFile)
      try
        val worksheetEntry = zipFile.getEntry(worksheetPath)
        if worksheetEntry == null then
          throw new Exception(s"Worksheet not found while loading metadata: $worksheetPath")

        // GH-350/GH-457: build from the shared hardened factory (doctype-strip + lifted JAXP
        // entity-size limits) instead of re-inlining the XXE posture per-site.
        val parser = XmlSecurity.secureSaxParserFactory().newSAXParser()

        val inputStream =
          XmlSecurity.stripLeadingDoctypeStream(zipFile.getInputStream(worksheetEntry))
        try
          val handler = new DefaultHandler:
            override def startElement(
              uri: String,
              localName: String,
              qName: String,
              attributes: Attributes
            ): Unit =
              val name = if localName.nonEmpty then localName else qName

              if needsColumns && name == "col" then
                parseColumnPropertiesFromAttributes(attributes).foreach {
                  case (minCol, maxCol, props) =>
                    (minCol to maxCol).foreach { index1 =>
                      parsedColumns(Column.from1(index1)) = props
                    }
                }

              if pendingRows.nonEmpty && name == "row" then
                Option(attributes.getValue("r")).flatMap(_.toIntOption).filter(_ > 0).foreach {
                  rowIndex =>
                    if pendingRows.contains(rowIndex) then
                      parsedRows(Row.from1(rowIndex)) = parseRowPropertiesFromAttributes(attributes)
                      pendingRows -= rowIndex
                      if pendingRows.isEmpty then throw new StopWorksheetScan
                    else
                      val maxPendingRow =
                        pendingRows.foldLeft(Int.MinValue)((currentMax, nextRow) =>
                          math.max(currentMax, nextRow)
                        )
                      if rowIndex > maxPendingRow then
                        // Row indices are ascending in worksheet XML, so remaining targets won't appear.
                        throw new StopWorksheetScan
                }

              if name == "sheetData" && needsColumns && pendingRows.isEmpty then
                // Column definitions are fully parsed before sheetData.
                throw new StopWorksheetScan

          try parser.parse(inputStream, handler)
          catch case _: StopWorksheetScan => ()
        finally inputStream.close()

        ExistingWorksheetMetadata(parsedColumns.toMap, parsedRows.toMap)
      finally zipFile.close()

  private def parseColumnPropertiesFromAttributes(
    attributes: Attributes
  ): Option[(Int, Int, ColumnProperties)] =
    for
      minCol <- Option(attributes.getValue("min")).flatMap(_.toIntOption).filter(_ > 0)
      maxCol <- Option(attributes.getValue("max")).flatMap(_.toIntOption).filter(_ >= minCol)
    yield
      val width = Option(attributes.getValue("width")).flatMap(_.toDoubleOption)
      val styleId = Option(attributes.getValue("style")).flatMap(_.toIntOption).map(StyleId.apply)
      val outlineLevel =
        Option(attributes.getValue("outlineLevel")).flatMap(_.toIntOption).filter(_ > 0)
      val props = ColumnProperties(
        width = width,
        hidden = isXmlTrue(attributes.getValue("hidden")),
        styleId = styleId,
        outlineLevel = outlineLevel,
        collapsed = isXmlTrue(attributes.getValue("collapsed"))
      )
      (minCol, maxCol, props)

  private def parseRowPropertiesFromAttributes(attributes: Attributes): RowProperties =
    RowProperties(
      height = Option(attributes.getValue("ht")).flatMap(_.toDoubleOption),
      hidden = isXmlTrue(attributes.getValue("hidden")),
      styleId = Option(attributes.getValue("s")).flatMap(_.toIntOption).map(StyleId.apply),
      outlineLevel =
        Option(attributes.getValue("outlineLevel")).flatMap(_.toIntOption).filter(_ > 0),
      collapsed = isXmlTrue(attributes.getValue("collapsed"))
    )

  private def isXmlTrue(value: String): Boolean =
    Option(value).exists(v => v == "1" || v.equalsIgnoreCase("true"))

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
    import com.tjclp.xl.styles.font.{Font, Underline}
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
      .withUnderline(if props.underline then Underline.Single else Underline.None)
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

  /**
   * GH-475: streaming numFmt resolution is the SAME function as the in-memory path.
   *
   * This used to be a hand-copied subset ending in `case _ => None`, so a style op carrying a
   * custom code (or any name outside the six it knew) reported success and left the cell General.
   */
  private def parseNumFmtSync(s: String): Option[com.tjclp.xl.styles.numfmt.NumFmt] =
    StyleBuilder.parseNumFmt(s).toOption

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
   * A target ref as its optional sheet qualifier and its cell or range; `label` names it in the
   * error (`Invalid reference: …`, `Invalid range: …`).
   */
  private def parseTarget(
    refStr: String,
    label: String
  ): IO[(Option[SheetName], Either[ARef, CellRange])] =
    IO.fromEither(
      RefType
        .parse(refStr)
        .left
        .map(_ => new Exception(s"Invalid $label: $refStr"))
        .map {
          case RefType.Cell(ref) => (None, Left(ref))
          case RefType.Range(range) => (None, Right(range))
          case RefType.QualifiedCell(sheet, ref) => (Some(sheet), Left(ref))
          case RefType.QualifiedRange(sheet, range) => (Some(sheet), Right(range))
        }
    )

  private val relsNamespace =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

  private def resolveSheetPath(
    sourcePath: Path,
    sheetNameOpt: Option[String],
    qualified: Option[SheetName],
    verb: String
  ): IO[String] =
    resolveSheetTarget(sourcePath, sheetNameOpt, qualified, verb).map(_._2)

  /**
   * The streamed worksheet as `(sheet name, worksheet part path)`: the name by THE sheet rule over
   * `workbook.xml` ([[Resolve.sheetName]], ADR-017 §2.5 — the ref's qualifier, else `-s`, else the
   * only sheet of a single-sheet book, else `SHEET_REQUIRED`), the part path from
   * `workbook.xml.rels` ([[worksheetPath]]). Lightweight: reads those two parts and each
   * worksheet's opening `<dimension>` only, never cell data.
   */
  private def resolveSheetTarget(
    sourcePath: Path,
    sheetNameOpt: Option[String],
    qualified: Option[SheetName],
    verb: String
  ): IO[(String, String)] =
    for
      meta <- IO.fromEither(
        WorkbookMetadataReader.read(sourcePath).left.map(readFailure(sourcePath, _))
      )
      name <- IO.fromEither(
        Resolve.sheetName(meta, sheetNameOpt, qualified, verb).left.map(CliException(_))
      )
      path <- worksheetPath(sourcePath, name.value)
    yield (name.value, path)

  /** An unreadable input is `IO_READ` (exit 3), as on every other verb. */
  private def readFailure(sourcePath: Path, err: XLError): CliException =
    CliException(
      CliError(
        ErrorCode.IO_READ,
        err.message,
        location = Some(Location.file(sourcePath.toString)),
        cause = Some(err)
      )
    )

  /**
   * The worksheet part path of the named sheet: its `<sheet r:id>` in workbook.xml resolved through
   * workbook.xml.rels, with the `sheet<sheetId>.xml` convention as the fallback.
   */
  private def worksheetPath(sourcePath: Path, sheetName: String): IO[String] =
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
            val targetSheet = sheets
              .find(sheetElem => (sheetElem \ "@name").text == sheetName)
              .getOrElse(throw new Exception(s"Worksheet not found in workbook.xml: $sheetName"))

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
