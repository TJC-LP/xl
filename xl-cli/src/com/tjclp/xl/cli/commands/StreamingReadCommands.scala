package com.tjclp.xl.cli.commands

import java.nio.file.Path

import cats.effect.IO
import fs2.Stream
import com.tjclp.xl.addressing.{ARef, CellRange, Column, RefType, Row}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.ViewFormat
import com.tjclp.xl.cli.helpers.ValueParser
import com.tjclp.xl.cli.output.{CsvRenderer, JsonRenderer, Markdown}
import com.tjclp.xl.io.{ExcelIO, RowData}

/**
 * Streaming implementations of read-only CLI commands.
 *
 * Uses O(1) memory by streaming rows instead of loading entire workbook. Suitable for large files
 * (100k+ rows).
 *
 * Supported commands: search, stats, bounds, view (markdown/csv/json only)
 */
object StreamingReadCommands:

  private val excel = ExcelIO.instance[IO]

  /**
   * Search for cells matching pattern using streaming.
   *
   * Emits matches as found during scan. Limit stops scan early for efficiency.
   */
  def search(
    filePath: Path,
    sheetNameOpt: Option[String],
    pattern: String,
    limit: Int
  ): IO[String] =
    IO.fromEither(
      scala.util
        .Try(pattern.r)
        .toEither
        .left
        .map(e => new Exception(s"Invalid regex pattern: ${e.getMessage}"))
    ).flatMap { regex =>
      val rowStream = sheetNameOpt match
        case Some(name) => excel.readSheetStream(filePath, name)
        case None => excel.readStream(filePath)

      rowStream
        .flatMap { row =>
          Stream.emits(
            row.cells.toSeq.collect { case (colIdx, value) =>
              val text = ValueParser.formatCellValue(value)
              if regex.findFirstIn(text).isDefined then
                Some((ARef.from0(colIdx, row.rowIndex - 1), text))
              else None
            }.flatten
          )
        }
        .take(limit)
        .compile
        .toVector
        .map { results =>
          val sheetDesc = sheetNameOpt.getOrElse("first sheet")
          s"Found ${results.size} matches in $sheetDesc (streaming):\n\n${formatSearchResults(results)}"
        }
    }

  /**
   * Calculate statistics for range using streaming aggregation.
   *
   * Memory: O(1) - only accumulates count/sum/min/max/sum-of-squares.
   */
  def stats(
    filePath: Path,
    sheetNameOpt: Option[String],
    refStr: String
  ): IO[String] =
    parseRangeFromRef(refStr).flatMap { range =>
      val rowStream = sheetNameOpt match
        case Some(name) => excel.readSheetStream(filePath, name)
        case None => excel.readStream(filePath)

      rowStream
        .filter(row => rowInRange(row, range))
        .flatMap(row => Stream.emits(cellsInRange(row, range)))
        .collect { case CellValue.Number(n) => n }
        .compile
        .fold(StatsAccumulator.empty)(_.add(_))
        .flatMap { acc =>
          if acc.count == 0 then
            IO.raiseError(new Exception(s"No numeric values in range ${range.toA1}"))
          else IO.pure(acc.format)
        }
    }

  /**
   * Compute used range bounds using streaming.
   *
   * Tracks min/max row/col during scan. Memory: O(1).
   */
  def bounds(
    filePath: Path,
    sheetNameOpt: Option[String]
  ): IO[String] =
    val rowStream = sheetNameOpt match
      case Some(name) => excel.readSheetStream(filePath, name)
      case None => excel.readStream(filePath)

    rowStream.compile
      .fold(BoundsAccumulator.empty)(_.update(_))
      .map { acc =>
        val sheetName = sheetNameOpt.getOrElse("Sheet1")
        acc.format(sheetName)
      }

  /**
   * View range using streaming (markdown/csv/json only).
   *
   * HTML/SVG/PDF require styles which aren't available in streaming mode.
   */
  def view(
    filePath: Path,
    sheetNameOpt: Option[String],
    rangeStr: String,
    showFormulas: Boolean,
    limit: Int,
    format: ViewFormat,
    showLabels: Boolean,
    skipEmpty: Boolean,
    headerRow: Option[Int]
  ): IO[String] =
    // Reject non-streamable formats
    format match
      case ViewFormat.Html | ViewFormat.Svg | ViewFormat.Png | ViewFormat.Jpeg | ViewFormat.WebP |
          ViewFormat.Pdf =>
        IO.raiseError(
          new Exception(
            s"--stream not supported for ${format.toString.toLowerCase} (needs styles). " +
              "Remove --stream flag or use markdown/csv/json format."
          )
        )
      case _ =>
        parseRangeFromRef(rangeStr).flatMap { range =>
          val rowStream = sheetNameOpt match
            case Some(name) => excel.readSheetStream(filePath, name)
            case None => excel.readStream(filePath)

          // Limit rows and filter to range
          val limitedRange = limitRange(range, limit)

          rowStream
            .filter(row => rowInRange(row, limitedRange))
            .compile
            .toVector
            .map { rows =>
              format match
                case ViewFormat.Markdown =>
                  formatMarkdown(rows, limitedRange, showFormulas, skipEmpty, showLabels)
                case ViewFormat.Csv =>
                  formatCsv(rows, limitedRange, showFormulas, skipEmpty, showLabels)
                case ViewFormat.Json =>
                  formatJson(rows, limitedRange, showFormulas, skipEmpty, headerRow)
                case _ => "" // unreachable due to earlier check
            }
        }

  // ==========================================================================
  // Helper types and functions
  // ==========================================================================

  /** Accumulator for streaming statistics computation. */
  private case class StatsAccumulator(
    count: Long,
    sum: BigDecimal,
    min: BigDecimal,
    max: BigDecimal
  ):
    def add(n: BigDecimal): StatsAccumulator =
      StatsAccumulator(
        count = count + 1,
        sum = sum + n,
        min = if count == 0 then n else min.min(n),
        max = if count == 0 then n else max.max(n)
      )

    def format: String =
      val mean = if count > 0 then sum / count else BigDecimal(0)
      f"count: $count, sum: $sum%.2f, min: $min%.2f, max: $max%.2f, mean: $mean%.2f"

  private object StatsAccumulator:
    def empty: StatsAccumulator = StatsAccumulator(0, BigDecimal(0), BigDecimal(0), BigDecimal(0))

  /** Accumulator for streaming bounds computation. */
  private case class BoundsAccumulator(
    minRow: Option[Int],
    maxRow: Option[Int],
    minCol: Option[Int],
    maxCol: Option[Int],
    cellCount: Long
  ):
    def update(row: RowData): BoundsAccumulator =
      if row.cells.isEmpty then this
      else
        val cols = row.cells.keys
        BoundsAccumulator(
          minRow = Some(minRow.fold(row.rowIndex)(_ min row.rowIndex)),
          maxRow = Some(maxRow.fold(row.rowIndex)(_ max row.rowIndex)),
          minCol = Some(minCol.fold(cols.min)(_ min cols.min)),
          maxCol = Some(maxCol.fold(cols.max)(_ max cols.max)),
          cellCount = cellCount + row.cells.size
        )

    def format(sheetName: String): String =
      (minRow, maxRow, minCol, maxCol) match
        case (Some(r1), Some(r2), Some(c1), Some(c2)) =>
          val startRef = ARef.from0(c1, r1 - 1) // rowIndex is 1-based
          val endRef = ARef.from0(c2, r2 - 1)
          val rowCount = r2 - r1 + 1
          val colCount = c2 - c1 + 1
          s"""Sheet: $sheetName (streaming)
             |Used range: ${startRef.toA1}:${endRef.toA1}
             |Rows: $r1-$r2 ($rowCount total)
             |Columns: ${Column.from0(c1).toLetter}-${Column.from0(c2).toLetter} ($colCount total)
             |Non-empty: $cellCount cells""".stripMargin
        case _ =>
          s"""Sheet: $sheetName (streaming)
             |Used range: (empty)
             |Non-empty: 0 cells""".stripMargin

  private object BoundsAccumulator:
    def empty: BoundsAccumulator = BoundsAccumulator(None, None, None, None, 0)

  /** Parse range from ref string (handles both single cell and range). */
  private def parseRangeFromRef(refStr: String): IO[CellRange] =
    IO.fromEither(
      RefType.parse(refStr).left.map(e => new Exception(e))
    ).flatMap {
      case RefType.Cell(ref) => IO.pure(CellRange(ref, ref))
      case RefType.Range(range) => IO.pure(range)
      case RefType.QualifiedCell(_, ref) => IO.pure(CellRange(ref, ref))
      case RefType.QualifiedRange(_, range) => IO.pure(range)
    }

  /** Check if row falls within range (by row index). */
  private def rowInRange(row: RowData, range: CellRange): Boolean =
    val rowIdx = row.rowIndex - 1 // Convert to 0-based
    rowIdx >= range.start.row.index0 && rowIdx <= range.end.row.index0

  /** Extract cells from row that fall within range columns. */
  private def cellsInRange(row: RowData, range: CellRange): Seq[CellValue] =
    val startCol = range.start.col.index0
    val endCol = range.end.col.index0
    row.cells.toSeq
      .filter { case (col, _) => col >= startCol && col <= endCol }
      .map(_._2)

  /** Limit range to max rows. */
  private def limitRange(range: CellRange, maxRows: Int): CellRange =
    val rowCount = range.end.row.index0 - range.start.row.index0 + 1
    if rowCount <= maxRows then range
    else
      val newEndRow = range.start.row.index0 + maxRows - 1
      CellRange(range.start, ARef.from0(range.end.col.index0, newEndRow))

  /** Format search results as markdown table. */
  private def formatSearchResults(results: Vector[(ARef, String)]): String =
    if results.isEmpty then "No matches found."
    else
      val sb = new StringBuilder
      sb.append("| Ref | Value |\n")
      sb.append("|-----|-------|\n")
      results.foreach { case (ref, value) =>
        val escaped = value.replace("|", "\\|").take(50)
        sb.append(s"| ${ref.toA1} | $escaped |\n")
      }
      sb.toString

  /** Format rows as markdown table. */
  private def formatMarkdown(
    rows: Vector[RowData],
    range: CellRange,
    showFormulas: Boolean,
    skipEmpty: Boolean,
    showLabels: Boolean
  ): String =
    if rows.isEmpty then return "(empty range)"

    val startCol = range.start.col.index0
    val endCol = range.end.col.index0
    val colCount = endCol - startCol + 1

    val sb = new StringBuilder

    // Header row (column letters if showLabels, else generic)
    if showLabels then
      sb.append("|   |")
      (startCol to endCol).foreach(c => sb.append(s" ${Column.from0(c).toLetter} |"))
      sb.append("\n|---|")
      (startCol to endCol).foreach(_ => sb.append("---|"))
      sb.append("\n")
    else
      sb.append("|")
      (0 until colCount).foreach(_ => sb.append(" |"))
      sb.append("\n|")
      (0 until colCount).foreach(_ => sb.append("---|"))
      sb.append("\n")

    // Data rows
    rows.foreach { row =>
      if showLabels then sb.append(s"| ${row.rowIndex} |")
      else sb.append("|")

      (startCol to endCol).foreach { colIdx =>
        val value = row.cells.get(colIdx) match
          case Some(v) => formatCellValue(v, showFormulas)
          case None => ""
        val escaped = value.replace("|", "\\|")
        sb.append(s" $escaped |")
      }
      sb.append("\n")
    }

    sb.toString

  /** Format rows as CSV. */
  private def formatCsv(
    rows: Vector[RowData],
    range: CellRange,
    showFormulas: Boolean,
    skipEmpty: Boolean,
    showLabels: Boolean
  ): String =
    val startCol = range.start.col.index0
    val endCol = range.end.col.index0

    val sb = new StringBuilder

    // Header row if showLabels
    if showLabels then
      sb.append(",")
      sb.append((startCol to endCol).map(c => Column.from0(c).toLetter).mkString(","))
      sb.append("\n")

    // Data rows
    rows.foreach { row =>
      if showLabels then sb.append(s"${row.rowIndex},")

      val values = (startCol to endCol).map { colIdx =>
        row.cells.get(colIdx) match
          case Some(v) =>
            val formatted = formatCellValue(v, showFormulas)
            if formatted.contains(",") || formatted.contains("\"") || formatted.contains("\n") then
              "\"" + formatted.replace("\"", "\"\"") + "\""
            else formatted
          case None => ""
      }
      sb.append(values.mkString(","))
      sb.append("\n")
    }

    sb.toString

  /** Format rows as JSON. */
  private def formatJson(
    rows: Vector[RowData],
    range: CellRange,
    showFormulas: Boolean,
    skipEmpty: Boolean,
    headerRow: Option[Int]
  ): String =
    val startCol = range.start.col.index0
    val endCol = range.end.col.index0

    // Get header keys if headerRow specified
    val headers: Option[Map[Int, String]] = headerRow.flatMap { hr =>
      rows.find(_.rowIndex == hr).map { row =>
        (startCol to endCol).flatMap { colIdx =>
          row.cells.get(colIdx).map(v => colIdx -> formatCellValue(v, false))
        }.toMap
      }
    }

    val dataRows = headerRow match
      case Some(hr) => rows.filterNot(_.rowIndex == hr)
      case None => rows

    val jsonRows = dataRows.map { row =>
      val cells = (startCol to endCol).flatMap { colIdx =>
        val valueOpt = row.cells.get(colIdx)
        if skipEmpty && valueOpt.isEmpty then None
        else
          val key = headers.flatMap(_.get(colIdx)).getOrElse(Column.from0(colIdx).toLetter)
          val value = valueOpt.map(v => formatCellValue(v, showFormulas)).getOrElse("")
          Some(s""""$key":"${escapeJson(value)}"""")
      }
      "{" + cells.mkString(",") + "}"
    }

    "[" + jsonRows.mkString(",\n") + "]"

  /** Escape string for JSON. */
  private def escapeJson(s: String): String =
    s.replace("\\", "\\\\")
      .replace("\"", "\\\"")
      .replace("\n", "\\n")
      .replace("\r", "\\r")
      .replace("\t", "\\t")

  /** Format cell value for display. */
  private def formatCellValue(value: CellValue, showFormulas: Boolean): String =
    value match
      case CellValue.Formula(formula, cached) if showFormulas => formula
      case CellValue.Formula(_, Some(cached)) => formatCellValue(cached, false)
      case CellValue.Formula(formula, None) => formula
      case other => ValueParser.formatCellValue(other)
