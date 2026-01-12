package com.tjclp.xl.cli.helpers

import java.nio.file.Path
import java.time.{LocalDate, LocalDateTime}
import java.time.format.DateTimeFormatter
import scala.util.Try

import cats.effect.IO
import com.github.tototoshi.csv.{CSVReader, DefaultCSVFormat}
import com.tjclp.xl.addressing.{ARef, Column, Row}
import com.tjclp.xl.cells.CellValue

/**
 * CSV parser for importing CSV files into Excel workbooks.
 *
 * Features:
 *   - RFC 4180 compliant via scala-csv
 *   - Column-based type inference (Number, Boolean, Date, Text)
 *   - Configurable delimiter, encoding, header detection
 *   - Helpful error messages with row/column indices
 */
object CsvParser:

  /** Configuration options for CSV import */
  final case class ImportOptions(
    delimiter: Char = ',',
    hasHeader: Boolean = true,
    encoding: String = "UTF-8",
    sampleRows: Int = 10,
    inferTypes: Boolean = true
  )

  /** Column type detected via sampling */
  enum ColumnType:
    case Number, Boolean, Date, Text

  /**
   * Parse CSV file and return (ARef, CellValue) tuples for batch put.
   *
   * Algorithm:
   *   1. Read all CSV rows
   *   2. Infer column types from first N rows (if enabled)
   *   3. Map each (row, col) to (ARef + offset, CellValue)
   *   4. Return flat vector suitable for Sheet.put(updates*)
   *
   * @param csvPath
   *   Path to CSV file
   * @param startRef
   *   Top-left cell position (default: A1)
   * @param options
   *   Import configuration
   * @return
   *   Vector of (ARef, CellValue) tuples
   */
  def parseCsv(
    csvPath: Path,
    startRef: ARef,
    options: ImportOptions
  ): IO[Vector[(ARef, CellValue)]] =
    IO {
      // Custom CSV format with configured delimiter
      implicit val csvFormat: DefaultCSVFormat = new DefaultCSVFormat {
        override val delimiter: Char = options.delimiter
      }

      // Open CSV file with encoding
      val reader = CSVReader.open(csvPath.toFile, options.encoding)
      try
        val allRows = reader.all() // Read all rows into memory

        if allRows.isEmpty then Vector.empty
        else
          // Skip header row if configured
          val dataRows = if options.hasHeader then allRows.drop(1) else allRows

          if dataRows.isEmpty then Vector.empty
          else
            // Infer column types from sample
            val columnTypes =
              if options.inferTypes then inferColumnTypes(dataRows, options.sampleRows)
              else Vector.fill(dataRows.headOption.map(_.length).getOrElse(0))(ColumnType.Text)

            // Pre-check CSV dimensions fit within Excel bounds
            val csvRows = dataRows.length
            val csvCols = dataRows.headOption.map(_.length).getOrElse(0)
            val endCol = startRef.col.index0 + csvCols - 1
            val endRow = startRef.row.index0 + csvRows - 1

            if endCol > 16383 then
              throw new Exception(
                s"CSV exceeds Excel column limit: start=${startRef.col.toLetter}, " +
                  s"columns=$csvCols, end=${Column.from0(endCol).toLetter} (max: XFD)"
              )

            if endRow > 1048575 then
              throw new Exception(
                s"CSV exceeds Excel row limit: start=${startRef.row.index1}, " +
                  s"rows=$csvRows, end=${endRow + 1} (max: 1048576)"
              )

            // Map each CSV cell to (ARef, CellValue)
            val updates = scala.collection.mutable.ArrayBuffer[(ARef, CellValue)]()

            dataRows.zipWithIndex.foreach { (dataRow, rowIdx) =>
              dataRow.zipWithIndex.foreach { (value, colIdx) =>
                val ref = ARef.from0(
                  startRef.col.index0 + colIdx,
                  startRef.row.index0 + rowIdx
                )
                val cellValue =
                  parseValue(value, columnTypes.lift(colIdx).getOrElse(ColumnType.Text))
                updates += ((ref, cellValue))
              }
            }

            updates.toVector
      finally reader.close()
    }.handleErrorWith { e =>
      IO.raiseError(
        new Exception(s"Failed to parse CSV file '${csvPath.getFileName}': ${e.getMessage}", e)
      )
    }

  /**
   * Infer column types from sample rows.
   *
   * Strategy: For each column, check if ALL sampled values match a type pattern. If unanimous →
   * assign type, else fallback to Text.
   *
   * Priority: Number → Boolean → Date → Text
   */
  private def inferColumnTypes(rows: Seq[Seq[String]], sampleSize: Int): Vector[ColumnType] =
    if rows.isEmpty then Vector.empty
    else
      val maxCols = rows.map(_.length).maxOption.getOrElse(0)
      val sampleRows = rows.take(sampleSize)

      (0 until maxCols).map { colIdx =>
        // Get all values for this column from sample rows
        val columnValues = sampleRows.flatMap(_.lift(colIdx))
        val nonEmptyValues = columnValues.filter(_.trim.nonEmpty)

        // Only infer type if we have non-empty values (prevents all-empty columns from being typed as Number)
        if nonEmptyValues.isEmpty then ColumnType.Text
        else if nonEmptyValues.forall(isNumber) then ColumnType.Number
        else if nonEmptyValues.forall(isBoolean) then ColumnType.Boolean
        else if nonEmptyValues.forall(isDate) then ColumnType.Date
        else ColumnType.Text
      }.toVector

  /**
   * Parse a string value based on inferred column type.
   *
   * @param value
   *   Raw CSV value
   * @param colType
   *   Inferred column type
   * @return
   *   Typed CellValue (Empty for blank strings)
   */
  private def parseValue(value: String, colType: ColumnType): CellValue =
    // Empty strings become Empty cells
    if value.trim.isEmpty then CellValue.Empty
    else
      colType match
        case ColumnType.Number =>
          // Try parsing as BigDecimal
          Try(BigDecimal(value.trim))
            .map(CellValue.Number.apply)
            .getOrElse(CellValue.Text(value)) // Fallback to text if parse fails

        case ColumnType.Boolean =>
          // Case-insensitive true/false
          value.trim.toLowerCase match
            case "true" => CellValue.Bool(true)
            case "false" => CellValue.Bool(false)
            case _ => CellValue.Text(value) // Fallback

        case ColumnType.Date =>
          // Try parsing as LocalDate (ISO 8601: YYYY-MM-DD), convert to LocalDateTime at midnight
          Try(LocalDate.parse(value.trim, DateTimeFormatter.ISO_LOCAL_DATE))
            .map(_.atStartOfDay()) // Convert LocalDate to LocalDateTime
            .map(CellValue.DateTime.apply)
            .getOrElse(CellValue.Text(value)) // Fallback

        case ColumnType.Text =>
          CellValue.Text(value)

  // ========== Type Detection Helpers ==========

  /**
   * Check if a string value can be parsed as a number.
   *
   * Empty strings are allowed (they don't violate the Number type).
   */
  private def isNumber(s: String): Boolean =
    s.trim.isEmpty || Try(BigDecimal(s.trim)).isSuccess

  /**
   * Check if a string value is a boolean.
   *
   * Accepts: true, false (case-insensitive), or empty
   */
  private def isBoolean(s: String): Boolean =
    val normalized = s.trim.toLowerCase
    normalized.isEmpty || normalized == "true" || normalized == "false"

  /**
   * Check if a string value can be parsed as a date.
   *
   * Supports ISO 8601 format: YYYY-MM-DD Empty strings are allowed.
   */
  private def isDate(s: String): Boolean =
    s.trim.isEmpty || Try(LocalDate.parse(s.trim, DateTimeFormatter.ISO_LOCAL_DATE)).isSuccess
