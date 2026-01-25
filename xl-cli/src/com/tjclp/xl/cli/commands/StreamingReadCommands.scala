package com.tjclp.xl.cli.commands

import java.nio.file.Path

import cats.effect.IO
import cats.implicits.*
import fs2.Stream
import com.tjclp.xl.addressing.{ARef, CellRange, Column, RefType, SheetName}
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
    limit: Int,
    sheetsFilter: Option[String]
  ): IO[String] =
    IO.fromEither(
      scala.util
        .Try(pattern.r)
        .toEither
        .left
        .map(e => new Exception(s"Invalid regex pattern: ${e.getMessage}"))
    ).flatMap { regex =>
      resolveSearchSheets(filePath, sheetNameOpt, sheetsFilter).flatMap { targetSheets =>
        val rowStream = Stream
          .emits(targetSheets)
          .flatMap { sheetName =>
            excel.readSheetStream(filePath, sheetName).map(row => (sheetName, row))
          }

        rowStream
          .flatMap { case (sheetName, row) =>
            Stream.emits(
              row.cells.toSeq.collect { case (colIdx, value) =>
                val text = ValueParser.formatCellValue(value)
                if regex.findFirstIn(text).isDefined then
                  Some((sheetName, ARef.from0(colIdx, row.rowIndex - 1), text))
                else None
              }.flatten
            )
          }
          .take(limit)
          .compile
          .toVector
          .map { results =>
            val sheetDesc =
              if targetSheets.size == 1 then targetSheets.head
              else s"${targetSheets.size} sheets"
            val formatted = Markdown.renderSearchResultsWithRef(
              results.map { case (sheetName, ref, value) =>
                (s"$sheetName!${ref.toA1}", value)
              }
            )
            s"Found ${results.size} matches in $sheetDesc (streaming):\n\n$formatted"
          }
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
    parseRangeFromRef(refStr).flatMap { case (refSheetOpt, range) =>
      resolveSheetName(sheetNameOpt, refSheetOpt, "stats", refStr).flatMap { resolvedSheetOpt =>
        val rowStream = resolvedSheetOpt match
          case Some(name) => excel.readSheetStreamRange(filePath, name, range)
          case None => excel.readStreamRange(filePath, range)

        rowStream
          .flatMap(row => Stream.emits(cellsInRange(row, range)))
          .collect {
            case CellValue.Number(n) => n
            case CellValue.Formula(_, Some(CellValue.Number(n))) => n
          }
          .compile
          .fold(StatsAccumulator.empty)(_.add(_))
          .flatMap { acc =>
            if acc.count == 0 then
              IO.raiseError(new Exception(s"No numeric values in range ${range.toA1}"))
            else IO.pure(acc.format)
          }
      }
    }

  /**
   * Compute used range bounds using dimension element (instant for any file size).
   *
   * Uses <dimension ref="..."> from worksheet metadata. Falls back to streaming scan if dimension
   * is missing.
   */
  def bounds(
    filePath: Path,
    sheetNameOpt: Option[String]
  ): IO[String] =
    boundsDimension(filePath, sheetNameOpt)

  /**
   * Compute used range bounds using dimension element (instant).
   *
   * Reads only the <dimension ref="..."> element from worksheet XML. Typically completes in <100ms.
   */
  def boundsDimension(
    filePath: Path,
    sheetNameOpt: Option[String]
  ): IO[String] =
    excel.readMetadata(filePath).flatMap { meta =>
      // Find sheet index
      val sheetIndexEither: Either[String, Int] = sheetNameOpt match
        case Some(name) =>
          meta.sheets.indexWhere(_.name.value == name) match
            case -1 => Left(s"Sheet not found: $name")
            case idx => Right(idx + 1) // 1-based
        case None => Right(1) // First sheet

      sheetIndexEither match
        case Left(err) => IO.raiseError(new Exception(err))
        case Right(sheetIndex) =>
          val sheetInfo = meta.sheets.lift(sheetIndex - 1)
          val sheetName = sheetInfo.map(_.name.value).getOrElse(s"Sheet$sheetIndex")
          val dimension = sheetInfo.flatMap(_.dimension)

          dimension match
            case Some(range) =>
              val rowCount = range.end.row.index1 - range.start.row.index1 + 1
              val colCount = range.end.col.index0 - range.start.col.index0 + 1
              IO.pure(
                s"""Sheet: $sheetName
                   |Used range: ${range.toA1} (from dimension element)
                   |Rows: ${range.start.row.index1}-${range.end.row.index1} ($rowCount total)
                   |Columns: ${range.start.col.toLetter}-${range.end.col.toLetter} ($colCount total)""".stripMargin
              )
            case None =>
              // Fallback to streaming scan
              boundsScan(filePath, sheetNameOpt)
    }

  /**
   * Compute used range bounds using streaming scan (accurate but slower).
   *
   * Tracks min/max row/col during full scan. Memory: O(1). Time: O(rows).
   */
  def boundsScan(
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
        acc.format(sheetName, fromScan = true)
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
        parseRangeFromRef(rangeStr).flatMap { case (refSheetOpt, range) =>
          resolveSheetName(sheetNameOpt, refSheetOpt, "view", rangeStr).flatMap {
            resolvedSheetOpt =>
              // Limit rows and filter to range
              val limitedRange = limitRange(range, limit)
              val rowStream = resolvedSheetOpt match
                case Some(name) => excel.readSheetStreamRange(filePath, name, limitedRange)
                case None => excel.readStreamRange(filePath, limitedRange)

              rowStream.compile.toVector
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

    def format(sheetName: String, fromScan: Boolean = false): String =
      val source = if fromScan then "(from scan)" else "(streaming)"
      (minRow, maxRow, minCol, maxCol) match
        case (Some(r1), Some(r2), Some(c1), Some(c2)) =>
          val startRef = ARef.from0(c1, r1 - 1) // rowIndex is 1-based
          val endRef = ARef.from0(c2, r2 - 1)
          val rowCount = r2 - r1 + 1
          val colCount = c2 - c1 + 1
          s"""Sheet: $sheetName
             |Used range: ${startRef.toA1}:${endRef.toA1} $source
             |Rows: $r1-$r2 ($rowCount total)
             |Columns: ${Column.from0(c1).toLetter}-${Column.from0(c2).toLetter} ($colCount total)
             |Non-empty: $cellCount cells""".stripMargin
        case _ =>
          s"""Sheet: $sheetName
             |Used range: (empty) $source
             |Non-empty: 0 cells""".stripMargin

  private object BoundsAccumulator:
    def empty: BoundsAccumulator = BoundsAccumulator(None, None, None, None, 0)

  /** Parse range and optional sheet name from ref string (handles both single cell and range). */
  private def parseRangeFromRef(refStr: String): IO[(Option[String], CellRange)] =
    IO.fromEither(
      RefType.parse(refStr).left.map(e => new Exception(e))
    ).flatMap {
      case RefType.Cell(ref) => IO.pure((None, CellRange(ref, ref)))
      case RefType.Range(range) => IO.pure((None, range))
      case RefType.QualifiedCell(sheet, ref) => IO.pure((Some(sheet.value), CellRange(ref, ref)))
      case RefType.QualifiedRange(sheet, range) => IO.pure((Some(sheet.value), range))
    }

  /** Resolve sheet selection for streaming range commands, honoring qualified refs. */
  private def resolveSheetName(
    sheetNameOpt: Option[String],
    refSheetOpt: Option[String],
    context: String,
    refStr: String
  ): IO[Option[String]] =
    (refSheetOpt, sheetNameOpt) match
      case (Some(refSheet), Some(cliSheet)) if refSheet != cliSheet =>
        IO.raiseError(
          new Exception(
            s"$context ref '$refStr' targets sheet '$refSheet' but --sheet is '$cliSheet'"
          )
        )
      case (Some(refSheet), _) => IO.pure(Some(refSheet))
      case (None, Some(cliSheet)) => IO.pure(Some(cliSheet))
      case (None, None) => IO.pure(None)

  /** Resolve target sheets for streaming search, matching --sheets/--sheet semantics. */
  private def resolveSearchSheets(
    filePath: Path,
    sheetNameOpt: Option[String],
    sheetsFilter: Option[String]
  ): IO[Vector[String]] =
    excel.readMetadata(filePath).flatMap { meta =>
      val available = meta.sheets.map(_.name.value)
      val availableList = available.mkString(", ")

      (sheetsFilter, sheetNameOpt) match
        case (Some(filterStr), _) =>
          val names = filterStr.split(",").map(_.trim).filter(_.nonEmpty).toVector
          if names.isEmpty then
            IO.raiseError(new Exception("--sheets requires at least one sheet name"))
          else
            names.traverse { name =>
              IO.fromEither(SheetName(name).left.map(e => new Exception(e))).flatMap { sn =>
                if available.contains(sn.value) then IO.pure(sn.value)
                else
                  IO.raiseError(
                    new Exception(s"Sheet not found: $name. Available: $availableList")
                  )
              }
            }
        case (None, Some(sheetName)) =>
          IO.fromEither(SheetName(sheetName).left.map(e => new Exception(e))).flatMap { sn =>
            if available.contains(sn.value) then IO.pure(Vector(sn.value))
            else
              IO.raiseError(
                new Exception(s"Sheet not found: $sheetName. Available: $availableList")
              )
          }
        case (None, None) =>
          IO.pure(available)
    }

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

  /** Check if a streamed cell value is effectively empty. */
  private def isCellEmptyValue(value: CellValue): Boolean =
    value match
      case CellValue.Empty => true
      case CellValue.Text(s) if s.trim.isEmpty => true
      case CellValue.Formula(_, Some(CellValue.Empty)) => true
      case CellValue.Formula(_, Some(CellValue.Text(s))) if s.trim.isEmpty => true
      case _ => false

  /** Select rows and columns to render, honoring skipEmpty semantics. */
  private def selectRowsAndCols(
    rows: Vector[RowData],
    range: CellRange,
    skipEmpty: Boolean
  ): (Vector[RowData], Vector[Int]) =
    val startCol = range.start.col.index0
    val endCol = range.end.col.index0
    val cols = (startCol to endCol).toVector
    if !skipEmpty then (rows, cols)
    else
      val nonEmptyCols = cols.filter { colIdx =>
        rows.exists { row =>
          row.cells.get(colIdx).exists(value => !isCellEmptyValue(value))
        }
      }
      val nonEmptyRows = rows.filter { row =>
        nonEmptyCols.exists { colIdx =>
          row.cells.get(colIdx).exists(value => !isCellEmptyValue(value))
        }
      }
      (nonEmptyRows, nonEmptyCols)

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

    val (selectedRows, selectedCols) = selectRowsAndCols(rows, range, skipEmpty)
    val colCount = selectedCols.size

    val sb = new StringBuilder

    // Header row (column letters if showLabels, else generic)
    if showLabels then
      sb.append("|   |")
      selectedCols.foreach(c => sb.append(s" ${Column.from0(c).toLetter} |"))
      sb.append("\n|---|")
      selectedCols.foreach(_ => sb.append("---|"))
      sb.append("\n")
    else
      sb.append("|")
      (0 until colCount).foreach(_ => sb.append(" |"))
      sb.append("\n|")
      (0 until colCount).foreach(_ => sb.append("---|"))
      sb.append("\n")

    // Data rows
    selectedRows.foreach { row =>
      if showLabels then sb.append(s"| ${row.rowIndex} |")
      else sb.append("|")

      selectedCols.foreach { colIdx =>
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
    val (selectedRows, selectedCols) = selectRowsAndCols(rows, range, skipEmpty)

    val sb = new StringBuilder

    // Header row if showLabels
    if showLabels then
      sb.append(",")
      sb.append(selectedCols.map(c => Column.from0(c).toLetter).mkString(","))
      sb.append("\n")

    // Data rows
    selectedRows.foreach { row =>
      if showLabels then sb.append(s"${row.rowIndex},")

      val values = selectedCols.map { colIdx =>
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
