package com.tjclp.xl.cli.helpers

import java.nio.charset.Charset
import java.nio.file.Path
import java.time.{LocalDate, LocalDateTime}
import java.time.format.DateTimeFormatter
import scala.util.Try
import scala.util.boundary
import scala.util.boundary.break

import cats.effect.IO
import fs2.{Stream, Pipe, text}
import fs2.io.file.{Files, Path as Fs2Path}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formatted.Formatted
import com.tjclp.xl.io.StyledRowData
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.numfmt.NumFmt

/**
 * Streaming CSV parser for O(1) memory CSV import.
 *
 * Uses fs2 for true streaming - rows are parsed and converted to [[StyledRowData]] as they're read,
 * never materializing the entire file in memory. A detected cell's format is chosen where it is
 * detected (the `Formatted` vocabulary of the in-memory `CsvParser`): a date carries `NumFmt.Date`,
 * and its row names the matching entry of [[Styles]], the table the caller hands the styled
 * streaming writer (`writeStreamStyledWithAutoDetect`), so the date displays as a date (GH-675).
 *
 * Features:
 *   - O(1) memory regardless of file size
 *   - Configurable delimiter, encoding, header handling
 *   - Per-cell type inference (Number, Boolean, Date, Text)
 *   - RFC 4180-ish parsing (quoted fields, escaped quotes)
 */
object StreamingCsvParser:

  /**
   * The style table every streamed row's `cellStyles` indexes: one entry per non-General format the
   * detector emits. `General` has no entry, so an undetected or plain value stays unstyled. A
   * future detected format (datetime, say) is one more entry here.
   */
  val Styles: Vector[CellStyle] = Vector(CellStyle.default.withNumFmt(NumFmt.Date))

  private val styleIndexOf: Map[NumFmt, Int] =
    Styles.zipWithIndex.map((style, i) => style.numFmt -> i).toMap

  /** Configuration options for streaming CSV import */
  final case class Options(
    delimiter: Char = ',',
    skipHeader: Boolean = false,
    encoding: String = "UTF-8",
    inferTypes: Boolean = true
  )

  /**
   * Stream a CSV file as rows for the styled streaming writer.
   *
   * Each row has a 1-based rowIndex and 0-based column keys; `cellStyles` indexes [[Styles]] for
   * every cell whose detected format has an entry there (dates), and is empty when `inferTypes` is
   * off. Suitable for passing directly to
   * `ExcelIO.writeStreamStyledWithAutoDetect(path, sheet, StreamingCsvParser.Styles)`.
   *
   * @param csvPath
   *   Path to CSV file
   * @param options
   *   Parsing options
   * @return
   *   Stream of rows, one per non-empty CSV row (excluding the header if skipHeader=true)
   */
  def streamCsv(csvPath: Path, options: Options): Stream[IO, StyledRowData] =
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
        val fields = parseCsvLine(line, options.delimiter).zipWithIndex
        if options.inferTypes then
          val typed = fields.map((value, colIdx) => colIdx -> inferAndParse(value))
          StyledRowData(
            rowIndex,
            typed.map((col, f) => col -> f.value).toMap,
            typed.flatMap((col, f) => styleIndexOf.get(f.numFmt).map(col -> _)).toMap
          )
        else StyledRowData(rowIndex, fields.map((v, col) => col -> CellValue.Text(v)).toMap)
      }
      .filter(_.cells.nonEmpty) // Skip empty rows

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
  // Imperative parsing for performance; early-return pattern for clarity
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
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
   * Infer type and parse a single value, with the format it displays in.
   *
   * Priority: Empty → Number → Boolean → Date → Text. A date is `NumFmt.Date` (the format `put` and
   * the in-memory `CsvParser` give the same text, GH-667/GH-675); every other value is `General`.
   *
   * Each cell is independently typed based on its content. This differs from batch CsvParser which
   * uses column-based sampling.
   */
  private def inferAndParse(value: String): Formatted = boundary:
    def general(v: CellValue): Formatted = Formatted(v, NumFmt.General)
    val trimmed = value.trim

    // Empty check
    if trimmed.isEmpty then break(general(CellValue.Empty))

    // Number check (integers and decimals)
    Try(BigDecimal(trimmed)).toOption match
      case Some(n) => break(general(CellValue.Number(n)))
      case None => ()

    // Boolean check (case-insensitive)
    trimmed.toLowerCase match
      case "true" => break(general(CellValue.Bool(true)))
      case "false" => break(general(CellValue.Bool(false)))
      case _ => ()

    // Date check (ISO 8601: YYYY-MM-DD)
    Try(LocalDate.parse(trimmed, DateTimeFormatter.ISO_LOCAL_DATE)).toOption match
      case Some(date) => break(Formatted(CellValue.DateTime(date.atStartOfDay()), NumFmt.Date))
      case None => ()

    // Default to text
    general(CellValue.Text(value)) // Use original value, not trimmed
