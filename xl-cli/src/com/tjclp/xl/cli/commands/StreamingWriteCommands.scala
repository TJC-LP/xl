package com.tjclp.xl.cli.commands

import java.nio.charset.StandardCharsets
import java.nio.file.Path
import java.util.zip.ZipFile

import cats.effect.IO
import com.tjclp.xl.api.Workbook
import com.tjclp.xl.addressing.CellRange
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.io.streaming.ZipTransformer
import com.tjclp.xl.ooxml.XmlSecurity
import com.tjclp.xl.ooxml.writer.WriterConfig
import com.tjclp.xl.cli.helpers.{StreamingCsvParser, StyleBuilder}
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
