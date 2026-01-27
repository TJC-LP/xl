package com.tjclp.xl.cli.commands

import java.nio.file.{Path, Paths}

import cats.effect.IO
import com.tjclp.xl.{Workbook, Sheet, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cli.helpers.{CsvParser, StreamingCsvParser}
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.ooxml.writer.WriterConfig

/**
 * Command handlers for CSV import operations.
 *
 * Supports:
 *   - Import to existing sheet at position
 *   - Import to new sheet
 *   - Create new workbook from CSV
 *   - True streaming import for new sheets (O(1) memory)
 */
object ImportCommands:

  /** Write workbook using standard or streaming writer based on mode */
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
   *   If true, uses streaming writer for O(1) output memory (hybrid streaming)
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
   *   If true, uses streaming writer for O(1) output memory (hybrid streaming)
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
