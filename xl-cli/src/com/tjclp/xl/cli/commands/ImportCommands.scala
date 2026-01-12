package com.tjclp.xl.cli.commands

import java.nio.file.{Path, Paths}

import cats.effect.IO
import com.tjclp.xl.{Workbook, Sheet, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cli.helpers.CsvParser
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.ooxml.writer.WriterConfig

/**
 * Command handlers for CSV import operations.
 *
 * Supports:
 *   - Import to existing sheet at position
 *   - Import to new sheet
 *   - Create new workbook from CSV
 */
object ImportCommands:

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
   * @param hasHeader
   *   Whether first row contains headers (not imported as data)
   * @param encoding
   *   File encoding
   * @param newSheetName
   *   If Some, create new sheet with this name
   * @param outputPath
   *   Output Excel file path
   * @param config
   *   Writer configuration
   * @return
   *   Success message with row/column count
   */
  def importCsv(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    csvPath: String,
    startRefStr: Option[String],
    delimiter: Char,
    hasHeader: Boolean,
    encoding: String,
    newSheetName: Option[String],
    noTypeInference: Boolean,
    outputPath: Path,
    config: WriterConfig
  ): IO[String] =
    val csvPathResolved = Paths.get(csvPath)
    val options = CsvParser.ImportOptions(
      delimiter = delimiter,
      hasHeader = hasHeader,
      encoding = encoding,
      sampleRows = 10,
      inferTypes = !noTypeInference
    )

    newSheetName match
      case Some(name) =>
        // Import to new sheet
        importToNewSheet(wb, csvPathResolved, name, options, outputPath, config)

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
              importToPosition(wb, sheet, csvPathResolved, ref, options, outputPath, config)
            }

  /**
   * Import CSV to an existing sheet at a specified position.
   *
   * Overwrites any existing data at that position.
   */
  private def importToPosition(
    wb: Workbook,
    sheet: Sheet,
    csvPath: Path,
    startRef: ARef,
    options: CsvParser.ImportOptions,
    outputPath: Path,
    config: WriterConfig
  ): IO[String] =
    for
      // Parse CSV into (ARef, CellValue) tuples
      updates <- CsvParser.parseCsv(csvPath, startRef, options)

      // Apply batch put to sheet (O(N) with style deduplication)
      updatedSheet = sheet.put(updates*)

      // Replace sheet in workbook
      updatedWb = wb.put(updatedSheet)

      // Write to output file
      _ <- ExcelIO.instance[IO].writeWith(updatedWb, outputPath, config)

      // Calculate import stats
      rowCount = updates.map(_._1.row).distinct.size
      colCount = updates.map(_._1.col).distinct.size
      cellCount = updates.size
    yield s"""Imported: ${csvPath.getFileName} → ${sheet.name} (${rowCount} rows, ${colCount} cols, ${cellCount} cells)
Saved: ${outputPath}"""

  /**
   * Import CSV to a new sheet (created as part of the import).
   *
   * Sheet name must not already exist in the workbook.
   */
  private def importToNewSheet(
    wb: Workbook,
    csvPath: Path,
    sheetName: String,
    options: CsvParser.ImportOptions,
    outputPath: Path,
    config: WriterConfig
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
        _ <- ExcelIO.instance[IO].writeWith(updatedWb, outputPath, config)

        // Calculate import stats
        rowCount = updates.map(_._1.row).distinct.size
        colCount = updates.map(_._1.col).distinct.size
        cellCount = updates.size
      yield s"""Imported: ${csvPath.getFileName} → new sheet '$sheetName' (${rowCount} rows, ${colCount} cols, ${cellCount} cells)
Saved: ${outputPath}"""
