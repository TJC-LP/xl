package com.tjclp.xl.cli.commands

import java.nio.file.{Path, Paths}

import cats.effect.{IO, Resource}
import com.tjclp.xl.{Workbook, Sheet, style, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.helpers.{CsvParser, MarkdownTableParser, StreamingCsvParser}
import com.tjclp.xl.cli.helpers.MarkdownTableParser.ColumnAlignment
import com.tjclp.xl.formatted.{Formatted, FormattedParsers}
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.ooxml.writer.WriterConfig
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.alignment.{Align, HAlign}
import com.tjclp.xl.styles.numfmt.NumFmt

/**
 * Command handlers for CSV and markdown table import operations.
 *
 * Supports:
 *   - Import to existing sheet at position
 *   - Import to new sheet
 *   - Create new workbook from CSV
 *   - True streaming import for new sheets (O(1) memory)
 *   - GFM markdown table import with smart type detection (GH-159)
 */
object ImportCommands:

  /** Write workbook using the standard or SAX/StAX backend based on mode */
  private def writeWorkbook(
    wb: Workbook,
    outputPath: Path,
    config: WriterConfig,
    stream: Boolean
  ): IO[Unit] =
    val excel = ExcelIO.instance[IO]
    if stream then excel.writeWorkbookStream(wb, outputPath, config)
    else excel.writeWith(wb, outputPath, config)

  /** Build save message suffix based on write mode */
  private def saveSuffix(outputPath: Path, stream: Boolean): String =
    if stream then s"Saved (streaming): $outputPath"
    else s"Saved: $outputPath"

  /**
   * Import CSV data into a workbook.
   *
   * @param wb
   *   Existing workbook (may be empty for new workbook creation)
   * @param sheetOpt
   *   Target sheet (required unless --new-sheet is used)
   * @param csvPath
   *   Path to CSV file
   * @param startRefStr
   *   Top-left cell position (e.g., "A1", "B5")
   * @param delimiter
   *   CSV field delimiter
   * @param skipHeader
   *   Whether to skip first row (default: false, imports all rows including headers)
   * @param encoding
   *   File encoding
   * @param newSheetName
   *   If Some, create new sheet with this name
   * @param outputPath
   *   Output Excel file path
   * @param config
   *   Writer configuration
   * @param stream
   *   If true and --new-sheet, uses true streaming (O(1) memory for entire operation)
   * @return
   *   Success message with row/column count
   */
  def importCsv(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    csvPath: String,
    startRefStr: Option[String],
    delimiter: Char,
    skipHeader: Boolean,
    encoding: String,
    newSheetName: Option[String],
    noTypeInference: Boolean,
    outputPath: Path,
    config: WriterConfig,
    stream: Boolean = false
  ): IO[String] =
    val csvPathResolved = Paths.get(csvPath)
    val options = CsvParser.ImportOptions(
      delimiter = delimiter,
      skipHeader = skipHeader,
      encoding = encoding,
      sampleRows = 10,
      inferTypes = !noTypeInference
    )

    newSheetName match
      case Some(name) if stream && wb.sheets.isEmpty =>
        // True streaming: CSV → XLSX with O(1) memory (only when creating new workbook)
        importToNewSheetStreaming(csvPathResolved, name, options, outputPath)

      case Some(name) =>
        // Hybrid or regular import to new sheet (workbook has existing sheets)
        importToNewSheet(wb, csvPathResolved, name, options, outputPath, config, stream)

      case None =>
        // Import to existing sheet at position
        sheetOpt match
          case None =>
            IO.raiseError(
              new Exception(
                "Import to existing sheet requires --sheet flag. Use --new-sheet to create a new sheet."
              )
            )
          case Some(sheet) =>
            val startRef = startRefStr match
              case Some(refStr) =>
                ARef.parse(refStr) match
                  case Right(ref) => IO.pure(ref)
                  case Left(err) =>
                    IO.raiseError(new Exception(s"Invalid start reference '$refStr': $err"))
              case None =>
                IO.pure(ARef.from0(0, 0)) // Default to A1

            startRef.flatMap { ref =>
              importToPosition(wb, sheet, csvPathResolved, ref, options, outputPath, config, stream)
            }

  /**
   * Import CSV to an existing sheet at a specified position.
   *
   * Overwrites any existing data at that position.
   *
   * @param stream
   *   If true, uses the SAX/StAX workbook writer
   */
  private def importToPosition(
    wb: Workbook,
    sheet: Sheet,
    csvPath: Path,
    startRef: ARef,
    options: CsvParser.ImportOptions,
    outputPath: Path,
    config: WriterConfig,
    stream: Boolean
  ): IO[String] =
    for
      // Parse CSV into (ARef, CellValue) tuples
      updates <- CsvParser.parseCsv(csvPath, startRef, options)

      // Apply batch put to sheet (O(N) with style deduplication)
      updatedSheet = sheet.put(updates*)

      // Replace sheet in workbook
      updatedWb = wb.put(updatedSheet)

      // Write to output file
      _ <- writeWorkbook(updatedWb, outputPath, config, stream)

      // Calculate import stats
      rowCount = updates.map(_._1.row).distinct.size
      colCount = updates.map(_._1.col).distinct.size
      cellCount = updates.size
    yield s"""Imported: ${csvPath.getFileName} → ${sheet.name} (${rowCount} rows, ${colCount} cols, ${cellCount} cells)
${saveSuffix(outputPath, stream)}"""

  /**
   * Import CSV to a new sheet (created as part of the import).
   *
   * Sheet name must not already exist in the workbook.
   *
   * @param stream
   *   If true, uses the SAX/StAX workbook writer
   */
  private def importToNewSheet(
    wb: Workbook,
    csvPath: Path,
    sheetName: String,
    options: CsvParser.ImportOptions,
    outputPath: Path,
    config: WriterConfig,
    stream: Boolean
  ): IO[String] =
    // Validate sheet name doesn't already exist
    if wb.sheets.exists(_.name.value == sheetName) then
      IO.raiseError(
        new Exception(
          s"Sheet '$sheetName' already exists. Use a different name or import to existing sheet."
        )
      )
    else
      for
        // Parse CSV starting at A1
        updates <- CsvParser.parseCsv(csvPath, ARef.from0(0, 0), options)

        // Validate sheet name using safe constructor
        sheetNameValidated <- SheetName(sheetName) match
          case Right(name) => IO.pure(name)
          case Left(err) =>
            IO.raiseError(new Exception(s"Invalid sheet name '$sheetName': $err"))

        // Create new sheet with CSV data
        newSheet = Sheet(sheetNameValidated).put(updates*)

        // Add sheet to workbook
        updatedWb = wb.put(newSheet)

        // Write to output file
        _ <- writeWorkbook(updatedWb, outputPath, config, stream)

        // Calculate import stats
        rowCount = updates.map(_._1.row).distinct.size
        colCount = updates.map(_._1.col).distinct.size
        cellCount = updates.size
      yield s"""Imported: ${csvPath.getFileName} → new sheet '$sheetName' (${rowCount} rows, ${colCount} cols, ${cellCount} cells)
${saveSuffix(outputPath, stream)}"""

  /**
   * True streaming CSV import - O(1) memory for entire operation.
   *
   * Only available when creating a new workbook (no existing sheets to preserve). Uses
   * StreamingCsvParser + writeStreamWithAutoDetect for end-to-end streaming.
   */
  private def importToNewSheetStreaming(
    csvPath: Path,
    sheetName: String,
    options: CsvParser.ImportOptions,
    outputPath: Path
  ): IO[String] =
    val streamingOpts = StreamingCsvParser.Options(
      delimiter = options.delimiter,
      skipHeader = options.skipHeader,
      encoding = options.encoding,
      inferTypes = options.inferTypes
    )

    StreamingCsvParser
      .streamCsv(csvPath, streamingOpts)
      .through(ExcelIO.instance[IO].writeStreamWithAutoDetect(outputPath, sheetName))
      .compile
      .drain
      .map(_ =>
        s"Streamed: ${csvPath.getFileName} → new sheet '$sheetName'\nSaved (streaming): $outputPath"
      )

  // ==========================================================================
  // Markdown table import (GH-159)
  // ==========================================================================

  /**
   * Import a GFM markdown table from a file (or stdin when `mdPath` is "-").
   *
   * Mirrors the CSV import workflow: read → parse → typed put → write. Values go through the shared
   * smart detection ([[FormattedParsers.detect]]: currency, percent, ISO dates, numbers, booleans)
   * unless `noTypeInference` is set. GFM alignment markers (`:---`, `:---:`, `---:`) map to cell
   * horizontal alignment.
   */
  def importMarkdown(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    mdPath: String,
    startRefStr: Option[String],
    skipHeader: Boolean,
    newSheetName: Option[String],
    noTypeInference: Boolean,
    outputPath: Path,
    config: WriterConfig,
    stream: Boolean
  ): IO[String] =
    val contentIO =
      if mdPath == "-" then
        IO.blocking(scala.io.Source.fromInputStream(System.in)(using scala.io.Codec.UTF8).mkString)
      else
        Resource
          .fromAutoCloseable(
            IO.blocking(scala.io.Source.fromFile(mdPath)(using scala.io.Codec.UTF8))
          )
          .use(src => IO.blocking(src.mkString))
          .handleErrorWith(e =>
            IO.raiseError(new Exception(s"Failed to read markdown file '$mdPath': ${e.getMessage}"))
          )
    val sourceName = if mdPath == "-" then "(stdin)" else Paths.get(mdPath).getFileName.toString
    contentIO.flatMap { content =>
      importMarkdownContent(
        wb,
        sheetOpt,
        content,
        sourceName,
        startRefStr,
        skipHeader,
        newSheetName,
        noTypeInference,
        outputPath,
        config,
        stream
      )
    }

  /**
   * Import a GFM markdown table from already-read content (seam for tests and stdin).
   *
   * @param sourceName
   *   Display name for messages (file name or "(stdin)")
   */
  def importMarkdownContent(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    content: String,
    sourceName: String,
    startRefStr: Option[String],
    skipHeader: Boolean,
    newSheetName: Option[String],
    noTypeInference: Boolean,
    outputPath: Path,
    config: WriterConfig,
    stream: Boolean
  ): IO[String] =
    for
      table <- MarkdownTableParser.parse(content) match
        case Right(t) => IO.pure(t)
        case Left(err) =>
          IO.raiseError(new Exception(s"Failed to parse markdown table from $sourceName: $err"))
      rows = if skipHeader then table.rows.drop(1) else table.rows
      _ <- IO.raiseWhen(rows.isEmpty)(
        new Exception(s"Markdown table in $sourceName has no rows to import")
      )
      startRef <- startRefStr match
        case Some(refStr) =>
          ARef.parse(refStr) match
            case Right(ref) => IO.pure(ref)
            case Left(err) =>
              IO.raiseError(new Exception(s"Invalid start reference '$refStr': $err"))
        case None => IO.pure(ARef.from0(0, 0))
      result <- newSheetName match
        case Some(name) =>
          if wb.sheets.exists(_.name.value == name) then
            IO.raiseError(
              new Exception(
                s"Sheet '$name' already exists. Use a different name or import to existing sheet."
              )
            )
          else
            for
              sheetName <- SheetName(name) match
                case Right(n) => IO.pure(n)
                case Left(err) => IO.raiseError(new Exception(s"Invalid sheet name '$name': $err"))
              filled <- applyMarkdownTable(
                Sheet(sheetName),
                rows,
                table.alignments,
                startRef,
                !noTypeInference
              )
              _ <- writeWorkbook(wb.put(filled), outputPath, config, stream)
            yield importMessage(sourceName, s"new sheet '$name'", rows, table.columnCount) +
              "\n" + saveSuffix(outputPath, stream)
        case None =>
          sheetOpt match
            case None =>
              IO.raiseError(
                new Exception(
                  "Import to existing sheet requires --sheet flag. Use --new-sheet to create a new sheet."
                )
              )
            case Some(sheet) =>
              for
                filled <- applyMarkdownTable(
                  sheet,
                  rows,
                  table.alignments,
                  startRef,
                  !noTypeInference
                )
                _ <- writeWorkbook(wb.put(filled), outputPath, config, stream)
              yield importMessage(sourceName, sheet.name.value, rows, table.columnCount) +
                "\n" + saveSuffix(outputPath, stream)
    yield result

  private def importMessage(
    sourceName: String,
    target: String,
    rows: Vector[Vector[String]],
    colCount: Int
  ): String =
    val cellCount = rows.map(_.length).sum
    s"Imported: $sourceName → $target (${rows.length} rows, $colCount cols, $cellCount cells)"

  /**
   * Put typed values into the sheet and apply detected number formats plus GFM column alignment.
   */
  private def applyMarkdownTable(
    sheet: Sheet,
    rows: Vector[Vector[String]],
    alignments: Vector[ColumnAlignment],
    startRef: ARef,
    inferTypes: Boolean
  ): IO[Sheet] =
    val rowCount = rows.length
    val colCount = alignments.length
    val endCol = startRef.col.index0 + colCount - 1
    val endRow = startRef.row.index0 + rowCount - 1

    if endCol > 16383 then
      IO.raiseError(
        new Exception(
          s"Markdown table exceeds Excel column limit: start=${startRef.col.toLetter}, " +
            s"columns=$colCount (max: XFD)"
        )
      )
    else if endRow > 1048575 then
      IO.raiseError(
        new Exception(
          s"Markdown table exceeds Excel row limit: start=${startRef.row.index1}, " +
            s"rows=$rowCount (max: 1048576)"
        )
      )
    else
      IO.pure {
        val typed: Vector[(ARef, CellValue, NumFmt)] =
          for
            (row, rowIdx) <- rows.zipWithIndex
            (raw, colIdx) <- row.zipWithIndex
          yield
            val ref = ARef.from0(startRef.col.index0 + colIdx, startRef.row.index0 + rowIdx)
            val Formatted(value, numFmt) = detectValue(raw, inferTypes)
            (ref, value, numFmt)

        val filled = sheet.put(typed.map((ref, value, _) => (ref, value))*)

        // Apply detected number formats and GFM column alignment as cell styles
        typed.foldLeft(filled) { case (acc, (ref, value, numFmt)) =>
          val halign = alignments.lift(ref.col.index0 - startRef.col.index0) match
            case Some(ColumnAlignment.Left) => Some(HAlign.Left)
            case Some(ColumnAlignment.Center) => Some(HAlign.Center)
            case Some(ColumnAlignment.Right) => Some(HAlign.Right)
            case _ => None
          val wantsStyle = numFmt != NumFmt.General || halign.isDefined
          if value == CellValue.Empty || !wantsStyle then acc
          else
            val base = CellStyle.default.withNumFmt(numFmt)
            val styled = halign.fold(base)(h => base.withAlign(Align(horizontal = h)))
            acc.style(ref, styled)
        }
      }

  /** Type a raw markdown cell: Empty for blanks, smart detection unless opted out. */
  private def detectValue(raw: String, inferTypes: Boolean): Formatted =
    if raw.trim.isEmpty then Formatted(CellValue.Empty, NumFmt.General)
    else if inferTypes then FormattedParsers.detect(raw)
    else Formatted(CellValue.Text(raw), NumFmt.General)
