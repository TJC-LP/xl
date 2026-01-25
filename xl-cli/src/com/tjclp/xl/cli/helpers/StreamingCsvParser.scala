package com.tjclp.xl.cli.helpers

import java.nio.charset.Charset
import java.nio.file.Path
import java.time.{LocalDate, LocalDateTime}
import java.time.format.DateTimeFormatter
import scala.util.Try

import cats.effect.IO
import fs2.{Stream, Pipe, text}
import fs2.io.file.{Files, Path as Fs2Path}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.io.RowData

/**
 * Streaming CSV parser for O(1) memory CSV import.
 *
 * Uses fs2 for true streaming - rows are parsed and converted to RowData as they're read, never
 * materializing the entire file in memory.
 *
 * Features:
 *   - O(1) memory regardless of file size
 *   - Configurable delimiter, encoding, header handling
 *   - Per-cell type inference (Number, Boolean, Date, Text)
 *   - RFC 4180-ish parsing (quoted fields, escaped quotes)
 */
object StreamingCsvParser:

  /** Configuration options for streaming CSV import */
  final case class Options(
    delimiter: Char = ',',
    skipHeader: Boolean = false,
    encoding: String = "UTF-8",
    inferTypes: Boolean = true
  )

  /**
   * Stream CSV file as RowData for direct streaming write.
   *
   * Each row is emitted as a RowData with 1-based rowIndex and 0-based column keys. Suitable for
   * passing directly to ExcelIO.writeStream or writeStreamWithAutoDetect.
   *
   * @param csvPath
   *   Path to CSV file
   * @param options
   *   Parsing options
   * @return
   *   Stream of RowData, one per CSV row (excluding header if skipHeader=true)
   */
  def streamCsv(csvPath: Path, options: Options): Stream[IO, RowData] =
    val fs2Path = Fs2Path.fromNioPath(csvPath)

    Files[IO]
      .readAll(fs2Path)
      .through(text.decodeWithCharset(Charset.forName(options.encoding)))
      .through(text.lines)
      .zipWithIndex
      .through(skipHeaderIfNeeded(options.skipHeader))
      .map { case (line, originalIdx) =>
        // Row index is 1-based in Excel (A1 is row 1)
        // After skipHeaderIfNeeded, indices are re-indexed from 0, so always add 1
        val rowIndex = originalIdx.toInt + 1
        val fields = parseCsvLine(line, options.delimiter)
        val cells = fields.zipWithIndex.map { case (value, colIdx) =>
          val cellValue = if options.inferTypes then inferAndParse(value) else CellValue.Text(value)
          colIdx -> cellValue
        }.toMap
        RowData(rowIndex, cells)
      }
      .filter(_.cells.nonEmpty) // Skip empty rows

  /**
   * Stream CSV with explicit start row offset for appending to existing data.
   *
   * @param csvPath
   *   Path to CSV file
   * @param startRow
   *   1-based row index to start numbering from (e.g., 1 for A1, 100 for row 100)
   * @param options
   *   Parsing options
   * @return
   *   Stream of RowData with row indices starting from startRow
   */
  def streamCsvWithOffset(
    csvPath: Path,
    startRow: Int,
    options: Options
  ): Stream[IO, RowData] =
    val fs2Path = Fs2Path.fromNioPath(csvPath)

    Files[IO]
      .readAll(fs2Path)
      .through(text.decodeWithCharset(Charset.forName(options.encoding)))
      .through(text.lines)
      .zipWithIndex
      .through(skipHeaderIfNeeded(options.skipHeader))
      .map { case (line, idx) =>
        // Adjust row index to start from specified row
        val rowIndex = startRow + idx.toInt
        val fields = parseCsvLine(line, options.delimiter)
        val cells = fields.zipWithIndex.map { case (value, colIdx) =>
          val cellValue = if options.inferTypes then inferAndParse(value) else CellValue.Text(value)
          colIdx -> cellValue
        }.toMap
        RowData(rowIndex, cells)
      }
      .filter(_.cells.nonEmpty)

  // ========== Private Helpers ==========

  /** Skip first element (header) if needed, re-index remaining elements */
  private def skipHeaderIfNeeded(
    skip: Boolean
  ): Pipe[IO, (String, Long), (String, Long)] =
    if skip then _.drop(1).zipWithIndex.map { case ((line, _), newIdx) => (line, newIdx) }
    else identity

  /**
   * Parse a single CSV line into fields.
   *
   * Handles:
   *   - Quoted fields (fields containing delimiter, newline, or quotes)
   *   - Escaped quotes ("" inside quoted field)
   *   - Mixed quoted/unquoted fields
   *
   * Note: This is a simplified parser that handles most common cases. For complex CSVs with
   * embedded newlines, consider a specialized library.
   */
  private def parseCsvLine(line: String, delimiter: Char): Vector[String] =
    val result = Vector.newBuilder[String]
    val field = new StringBuilder
    var inQuotes = false
    var i = 0

    while i < line.length do
      val c = line.charAt(i)

      if inQuotes then
        if c == '"' then
          // Check for escaped quote
          if i + 1 < line.length && line.charAt(i + 1) == '"' then
            field.append('"')
            i += 1
          else inQuotes = false
        else field.append(c)
      else if c == '"' then inQuotes = true
      else if c == delimiter then
        result += field.result()
        field.clear()
      else field.append(c)

      i += 1

    // Add final field
    result += field.result()
    result.result()

  /**
   * Infer type and parse a single value.
   *
   * Priority: Empty → Number → Boolean → Date → Text
   *
   * Each cell is independently typed based on its content. This differs from batch CsvParser which
   * uses column-based sampling.
   */
  private def inferAndParse(value: String): CellValue =
    val trimmed = value.trim

    // Empty check
    if trimmed.isEmpty then return CellValue.Empty

    // Number check (integers and decimals)
    Try(BigDecimal(trimmed)).toOption match
      case Some(n) => return CellValue.Number(n)
      case None => ()

    // Boolean check (case-insensitive)
    trimmed.toLowerCase match
      case "true" => return CellValue.Bool(true)
      case "false" => return CellValue.Bool(false)
      case _ => ()

    // Date check (ISO 8601: YYYY-MM-DD)
    Try(LocalDate.parse(trimmed, DateTimeFormatter.ISO_LOCAL_DATE)).toOption match
      case Some(date) => return CellValue.DateTime(date.atStartOfDay())
      case None => ()

    // Default to text
    CellValue.Text(value) // Use original value, not trimmed
