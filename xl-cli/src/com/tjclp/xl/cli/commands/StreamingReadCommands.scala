package com.tjclp.xl.cli.commands

import java.nio.file.Path

import scala.util.boundary
import scala.util.boundary.break

import cats.effect.IO
import cats.implicits.*
import fs2.Stream
import com.tjclp.xl.addressing.{ARef, CellRange, Column, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.{CliIO, ViewFormat}
import com.tjclp.xl.cli.contract.{
  CliError,
  CliException,
  Diagnostics,
  ErrorCode,
  Warning,
  WarningCode
}
import com.tjclp.xl.error.XLError
import com.tjclp.xl.cli.helpers.{Resolve, ValueParser}
import com.tjclp.xl.cli.output.{CsvRenderer, Format, JsonRenderer, Markdown, RendererCommon}
import com.tjclp.xl.display.NumFmtFormatter
import com.tjclp.xl.io.{ExcelIO, RowData}
import com.tjclp.xl.ooxml.metadata.LightMetadata
import com.tjclp.xl.ooxml.style.WorkbookStyles
import com.tjclp.xl.styles.numfmt.NumFmt

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
        .map(e => CliException(CliError.usage(s"Invalid regex pattern: ${e.getMessage}", None)))
    ).flatMap { regex =>
      resolveSearchSheets(filePath, sheetNameOpt, sheetsFilter).flatMap { targetSheets =>
        val rowStream = Stream
          .emits(targetSheets)
          .flatMap { sheetName =>
            excel.readSheetStream(filePath, sheetName).map(row => (sheetName, row))
          }

        val matchStream = rowStream
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

        // GH-351: --limit 0 means "no limit"; otherwise the scan aborts early at the limit.
        val limited = if limit > 0 then matchStream.take(limit) else matchStream

        limited.compile.toVector
          .map { results =>
            val sheetDesc = targetSheets match
              case Vector(single) => single
              case multiple => s"${multiple.size} sheets"
            val formatted = Markdown.renderSearchResultsWithRef(
              results.map { case (sheetName, ref, value) =>
                (s"$sheetName!${ref.toA1}", value)
              }
            )
            val base = s"Found ${results.size} matches in $sheetDesc (streaming):\n\n$formatted"
            // The streaming scan stops at --limit, so the true total is unknown; flag
            // the possible clip when the limit was reached (GH-351).
            if limit > 0 && results.size == limit then
              s"$base\n… showing first $limit matches (streaming scan stopped at --limit; use --limit to raise; --limit 0 = no limit)"
            else base
          }
      }
    }

  /**
   * Get cell details using streaming with O(1) worksheet memory.
   *
   * Pre-loads styles/sharedStrings/comments once, then streams worksheet with early-abort. Memory:
   * ~3MB max regardless of worksheet size.
   */
  def cell(
    filePath: Path,
    sheetNameOpt: Option[String],
    refStr: String,
    noStyle: Boolean
  ): IO[String] =
    for
      // Parse the ref (may include sheet name like "Sheet1!A1"): INVALID_REFERENCE either way —
      // the parser's own text, or the same shape refusal the in-memory `cell` and `deps` give
      parsed <- IO.fromEither(Resolve.ref(refStr).left.map(CliException(_)))
      (refSheetOpt, ref) <- parsed match
        case (qualifier, Resolve.Target.Cell(r)) => IO.pure((qualifier, r))
        case (_, Resolve.Target.Range(range)) =>
          IO.raiseError(
            CliException(
              Resolve.invalidReference(
                s"cell requires a single cell, not the range ${range.toA1}"
              )
            )
          )

      // THE sheet rule over workbook.xml: qualifier > -s > the only sheet > SHEET_REQUIRED
      targetSheet <- resolveSheetName(
        filePath,
        sheetNameOpt,
        refSheetOpt,
        s"cell with unqualified ref '$refStr'"
      )

      // Get streaming cell details
      details <- excel.streamCellDetails(filePath, targetSheet, ref)

      // Format output similar to non-streaming version
      style = if noStyle then None else details.style
      numFmt = style.map(_.numFmt).getOrElse(NumFmt.General)
      valueToFormat = details.value match
        case CellValue.Formula(_, Some(cached), _) => cached
        case other => other
      formatted = NumFmtFormatter.formatValue(valueToFormat, numFmt)

      // Format dependencies as strings (cell refs)
      deps = details.dependencies

      // Dependents not available in streaming mode
      dependentsStr = "(not available in streaming mode)"
    yield formatStreamingCellInfo(
      details.ref,
      details.value,
      formatted,
      style,
      details.comment,
      None, // hyperlinks not available in streaming mode
      deps,
      dependentsStr
    )

  /**
   * Format streaming cell info output (similar to Format.cellInfo but with streaming-specific
   * notes).
   */
  private def formatStreamingCellInfo(
    ref: ARef,
    value: CellValue,
    formatted: String,
    style: Option[com.tjclp.xl.styles.CellStyle],
    comment: Option[com.tjclp.xl.cells.Comment],
    hyperlink: Option[String],
    dependencies: Vector[String],
    dependentsStr: String
  ): String =
    val sb = new StringBuilder
    sb.append(s"Cell: ${ref.toA1}\n")
    sb.append(s"Type: ${valueType(value)}\n")

    // For formulas, show expression and cached value separately
    value match
      case CellValue.Formula(expr, cached, _) =>
        val displayExpr = if expr.startsWith("=") then expr else s"=$expr"
        sb.append(s"Formula: $displayExpr\n")
        cached.foreach { v =>
          sb.append(s"Cached: ${formatValue(v)}\n")
          val cachedRaw = formatValue(v)
          if formatted != cachedRaw && formatted != cachedRaw.stripPrefix("\"").stripSuffix("\"")
          then sb.append(s"Formatted: $formatted\n")
        }
      case CellValue.Empty =>
        sb.append("Value: (empty)\n")
      case _ =>
        sb.append(s"Raw: ${formatValue(value)}\n")
        val rawStr = formatValue(value)
        if formatted != rawStr && formatted != rawStr.stripPrefix("\"").stripSuffix("\"") then
          sb.append(s"Formatted: $formatted\n")

    // Style (non-default properties only)
    formatStyleForStreaming(style).foreach(s => sb.append(s).append("\n"))

    // Comment
    comment.foreach { c =>
      val authorStr = c.author.map(a => s" (Author: $a)").getOrElse("")
      sb.append(s"""Comment: "${c.text.toPlainText}"$authorStr\n""")
    }

    // Hyperlink
    hyperlink.foreach(h => sb.append(s"Hyperlink: $h\n"))

    // Dependencies and dependents
    val depsStr = if dependencies.isEmpty then "(none)" else dependencies.mkString(", ")
    sb.append(s"Dependencies: $depsStr\n")
    sb.append(s"Dependents: $dependentsStr")

    sb.toString

  private def valueType(value: CellValue): String =
    value match
      case CellValue.Text(_) => "text"
      case CellValue.Number(_) => "number"
      case CellValue.Bool(_) => "boolean"
      case CellValue.DateTime(_) => "datetime"
      case CellValue.Error(_) => "error"
      case CellValue.RichText(_) => "richtext"
      case CellValue.Empty => "empty"
      case CellValue.Formula(_, _, _) => "formula"

  private def formatValue(value: CellValue): String =
    value match
      case CellValue.Text(s) => s"\"$s\""
      case CellValue.Number(n) =>
        if n.isWhole then n.toBigInt.toString
        else n.underlying.stripTrailingZeros.toPlainString
      case CellValue.Bool(b) => if b then "TRUE" else "FALSE"
      case CellValue.DateTime(dt) => dt.toString
      case CellValue.Error(err) => err.toExcel
      case CellValue.RichText(rt) => s"\"${rt.toPlainText}\""
      case CellValue.Empty => "(empty)"
      case CellValue.Formula(expr, cached, _) =>
        val displayExpr = if expr.startsWith("=") then expr else s"=$expr"
        cached.map(formatValue).getOrElse(displayExpr)

  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  private def formatStyleForStreaming(
    style: Option[com.tjclp.xl.styles.CellStyle]
  ): Option[String] =
    import com.tjclp.xl.styles.font.{Font, Underline}
    import com.tjclp.xl.styles.fill.Fill
    import com.tjclp.xl.styles.border.{Border, BorderStyle}
    import com.tjclp.xl.styles.alignment.Align
    import com.tjclp.xl.styles.color.Color

    style.flatMap { s =>
      val parts = Vector.newBuilder[String]

      // Font (if non-default)
      if s.font != Font.default then
        var fontDesc = Vector(s.font.name, s"${s.font.sizePt}pt")
        if s.font.bold then fontDesc = fontDesc :+ "bold"
        if s.font.italic then fontDesc = fontDesc :+ "italic"
        s.font.underline match
          case Underline.None => ()
          case Underline.Single => fontDesc = fontDesc :+ "underline"
          case other => fontDesc = fontDesc :+ s"underline=${Underline.token(other)}"
        s.font.color.foreach {
          case Color.Rgb(argb) =>
            fontDesc = fontDesc :+ f"${argb & 0xffffff}%06X"
          case Color.Theme(slot, tint) =>
            val tintStr = if tint == 0.0 then "" else f" tint=$tint%.2f"
            fontDesc = fontDesc :+ s"$slot$tintStr"
        }
        parts += s"Font: ${fontDesc.mkString(" ")}"

      // Fill (if non-default)
      s.fill match
        case Fill.Solid(color) =>
          val colorStr = color match
            case Color.Rgb(argb) => f"${argb & 0xffffff}%06X"
            case Color.Theme(slot, tint) =>
              val tintStr = if tint == 0.0 then "" else f" tint=$tint%.2f"
              s"$slot$tintStr"
          parts += s"Fill: $colorStr (solid)"
        case Fill.Pattern(_, _, _) => parts += "Fill: (pattern)"
        case Fill.None => ()

      // NumFmt (if non-General)
      if s.numFmt != NumFmt.General then
        val idStr = s.numFmtId.map(id => s" (id: $id)").getOrElse("")
        val codeStr = s.numFmt match
          case NumFmt.Custom(code) => code
          case other => other.toString
        parts += s"NumFmt: $codeStr$idStr"

      // Alignment (if non-default)
      if s.align != Align.default then
        val wrapStr = if s.align.wrapText then ", wrap" else ""
        parts += s"Align: ${s.align.horizontal}, ${s.align.vertical}$wrapStr"

      // Border (if non-default)
      if s.border != Border.none then
        val sides = Vector(
          if s.border.top.style != BorderStyle.None then Some(s"top: ${s.border.top.style}")
          else None,
          if s.border.bottom.style != BorderStyle.None then
            Some(s"bottom: ${s.border.bottom.style}")
          else None,
          if s.border.left.style != BorderStyle.None then Some(s"left: ${s.border.left.style}")
          else None,
          if s.border.right.style != BorderStyle.None then Some(s"right: ${s.border.right.style}")
          else None
        ).flatten
        if sides.nonEmpty then
          val allSame = s.border.top == s.border.bottom &&
            s.border.bottom == s.border.left &&
            s.border.left == s.border.right &&
            s.border.top.style != BorderStyle.None
          if allSame then parts += s"Border: ${s.border.top.style} (all sides)"
          else parts += s"Border: ${sides.mkString(", ")}"

      val result = parts.result()
      if result.isEmpty then None
      else Some(result.map("  " + _).mkString("Style:\n", "\n", ""))
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
    parseRangeFromRef(refStr).flatMap { case (refSheetOpt, range, unqualified) =>
      resolveSheetName(filePath, sheetNameOpt, refSheetOpt, s"stats $unqualified").flatMap {
        sheetName =>
          val rowStream = excel.readSheetStreamRange(filePath, sheetName, range)

          rowStream
            .flatMap(row => Stream.emits(cellsInRange(row, range)))
            .collect {
              case CellValue.Number(n) => n
              case CellValue.Formula(_, Some(CellValue.Number(n)), _) => n
            }
            .compile
            .fold(StatsAccumulator.empty)(_.add(_))
            .flatMap { acc =>
              if acc.count == 0 then
                IO.raiseError(
                  CliException(
                    CliError.fromXLError(
                      XLError.Other(s"No numeric values in range ${range.toA1}"),
                      None
                    )
                  )
                )
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
    excel.readMetadata(filePath).flatMap(meta => boundsFromMetadata(meta, filePath, sheetNameOpt))

  /**
   * [[boundsDimension]] over metadata the caller has already read — the runner reads it under its
   * `IO_READ` classification, so a missing or unreadable file gets the same code as on every verb
   * while a sheet the metadata does not list keeps its own failure.
   */
  def boundsFromMetadata(
    meta: LightMetadata,
    filePath: Path,
    sheetNameOpt: Option[String]
  ): IO[String] =
    boundsTarget(meta, sheetNameOpt) match
      case Left(err) => IO.raiseError(CliException(err))
      case Right((sheetName, Some(range))) =>
        val rowCount = range.end.row.index1 - range.start.row.index1 + 1
        val colCount = range.end.col.index0 - range.start.col.index0 + 1
        IO.pure(
          s"""Sheet: $sheetName
             |Used range: ${range.toA1} (from dimension element)
             |Rows: ${range.start.row.index1}-${range.end.row.index1} ($rowCount total)
             |Columns: ${range.start.col.toLetter}-${range.end.col.toLetter} ($colCount total)""".stripMargin
        )
      case Right((_, None)) =>
        // Fallback to streaming scan
        boundsScan(filePath, sheetNameOpt)

  /**
   * `bounds` as data (`bounds --json`): `{sheet, range, dimension}` — `range` the used range in A1
   * form or `null` for an empty sheet, `dimension` true when it came from the worksheet's
   * `<dimension>` element and false when from a streaming scan (`--scan`, or a sheet without one).
   * Takes the metadata already read (see [[boundsFromMetadata]]).
   */
  def boundsData(
    meta: LightMetadata,
    filePath: Path,
    sheetNameOpt: Option[String],
    scan: Boolean
  ): IO[ujson.Value] =
    def scanned: IO[ujson.Value] =
      scanBounds(filePath, sheetNameOpt).map { (sheetName, acc) =>
        boundsJson(sheetName, acc.range, fromDimension = false)
      }
    if scan then scanned
    else
      boundsTarget(meta, sheetNameOpt) match
        case Left(err) => IO.raiseError(CliException(err))
        case Right((sheetName, Some(range))) =>
          IO.pure(boundsJson(sheetName, Some(range), fromDimension = true))
        case Right((_, None)) => scanned

  private def boundsJson(
    sheetName: String,
    range: Option[CellRange],
    fromDimension: Boolean
  ): ujson.Value =
    ujson.Obj(
      "sheet" -> ujson.Str(sheetName),
      "range" -> range.fold[ujson.Value](ujson.Null)(r => ujson.Str(r.toA1)),
      "dimension" -> ujson.Bool(fromDimension)
    )

  /**
   * The sheet `bounds` reads, by THE sheet rule over the metadata ([[Resolve.sheetName]]: `-s`,
   * else the only sheet, else `SHEET_REQUIRED`) — as its display name and the `<dimension>` range
   * the metadata carries for it (None when the worksheet has none).
   */
  private def boundsTarget(
    meta: LightMetadata,
    sheetNameOpt: Option[String]
  ): Either[CliError, (String, Option[CellRange])] =
    Resolve.sheetName(meta, sheetNameOpt, None, "bounds").map { name =>
      (name.value, meta.sheets.find(_.name == name).flatMap(_.dimension))
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
    scanBounds(filePath, sheetNameOpt).map { (sheetName, acc) =>
      acc.format(sheetName, fromScan = true)
    }

  /**
   * The scan itself: the sheet THE rule selects (its name is what the result is reported under) and
   * the accumulated bounds.
   */
  private def scanBounds(
    filePath: Path,
    sheetNameOpt: Option[String]
  ): IO[(String, BoundsAccumulator)] =
    resolveSheetName(filePath, sheetNameOpt, None, "bounds").flatMap { sheetName =>
      excel
        .readSheetStream(filePath, sheetName)
        .compile
        .fold(BoundsAccumulator.empty)(_.update(_))
        .map(acc => (sheetName, acc))
    }

  /**
   * View range using streaming (markdown/csv/json only).
   *
   * HTML/SVG/PDF require styles which aren't available in streaming mode.
   *
   * @param skipHidden
   *   GH-474: cannot be honored here (the streaming reader never parses row/column properties).
   *   Passing it emits [[RendererCommon.streamingSkipHiddenNotice]] as a `FLAG_IGNORED` warning
   *   rather than being silently ignored.
   * @param warn
   *   where the out-of-band notices go (that one, and the csv/json truncation notice); the runner
   *   collects them for the run's stderr or the `--json` envelope, the default prints
   *   `Warning[CODE]: …` straight to the process stderr — the same channel the in-memory view uses
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
    headerRow: Option[Int],
    skipHidden: Boolean = false,
    warn: Warning => IO[Unit] = Diagnostics.warn(_, CliIO.system)
  ): IO[String] =
    // Reject non-streamable formats
    format match
      case ViewFormat.Html | ViewFormat.Svg | ViewFormat.Png | ViewFormat.Jpeg | ViewFormat.WebP |
          ViewFormat.Pdf =>
        IO.raiseError(
          CliException(
            CliError(
              ErrorCode.UNSUPPORTED_IN_STREAM,
              s"--stream not supported for ${format.toString.toLowerCase} (needs styles). " +
                "Remove --stream flag or use markdown/csv/json format.",
              hint = Some("omit --stream, or use --format markdown, csv or json")
            )
          )
        )
      case _ =>
        // GH-474: the flag is unsupported here — announce it instead of no-oping in silence.
        warn(Warning(WarningCode.FLAG_IGNORED, RendererCommon.streamingSkipHiddenNotice))
          .whenA(skipHidden) *>
          parseRangeFromRef(rangeStr).flatMap { case (refSheetOpt, range, unqualified) =>
            resolveSheetName(filePath, sheetNameOpt, refSheetOpt, s"view $unqualified").flatMap {
              sheetName =>
                // Limit rows and filter to range
                val limitedRange = limitRange(range, limit)
                // GH-351: report truncation when --limit clips the requested range
                val totalRows = range.end.row.index0 - range.start.row.index0 + 1
                val shownRows = limitedRange.end.row.index0 - limitedRange.start.row.index0 + 1
                val isTruncated = shownRows < totalRows
                val notice = RendererCommon.truncationNotice(shownRows, totalRows)
                val truncated = Warning(WarningCode.TRUNCATED, notice)
                val rowStream = excel.readSheetStreamRange(filePath, sheetName, limitedRange)

                excel.loadStyles(filePath).flatMap { styles =>
                  rowStream.compile.toVector
                    .flatMap { rows =>
                      val s = Some(styles)
                      format match
                        case ViewFormat.Markdown =>
                          val table =
                            formatMarkdown(
                              rows,
                              limitedRange,
                              showFormulas,
                              skipEmpty,
                              showLabels,
                              s
                            )
                          IO.pure(if isTruncated then s"$table\n$notice" else table)
                        case ViewFormat.Csv =>
                          // stdout must stay machine-parseable: the notice is a warning only
                          warn(truncated)
                            .whenA(isTruncated)
                            .as(
                              formatCsv(rows, limitedRange, showFormulas, skipEmpty, showLabels, s)
                            )
                        case ViewFormat.Json =>
                          // Streaming JSON is a bare array (no top-level object to extend):
                          // the notice is a warning only
                          warn(truncated)
                            .whenA(isTruncated)
                            .as(
                              formatJson(rows, limitedRange, showFormulas, skipEmpty, headerRow, s)
                            )
                        case _ => IO.pure("") // unreachable due to earlier check
                    }
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
    // IterableOps: .min/.max safe because row.cells.isEmpty checked first
    @SuppressWarnings(Array("org.wartremover.warts.IterableOps"))
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

    /** The bounding range of every non-empty cell seen, None when the sheet is empty. */
    def range: Option[CellRange] = (minRow, maxRow, minCol, maxCol) match
      case (Some(r1), Some(r2), Some(c1), Some(c2)) =>
        Some(CellRange(ARef.from0(c1, r1 - 1), ARef.from0(c2, r2 - 1))) // rowIndex is 1-based
      case _ => None

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

  /**
   * Parse a ref string into its optional sheet qualifier, its range (a single cell is a 1x1 range)
   * and the words the `SHEET_REQUIRED` message uses for it (`with unqualified ref 'A1'`, exactly as
   * the in-memory twin says).
   */
  private def parseRangeFromRef(refStr: String): IO[(Option[SheetName], CellRange, String)] =
    IO.fromEither(Resolve.ref(refStr).left.map(CliException(_))).map {
      case (sheet, Resolve.Target.Cell(ref)) =>
        (sheet, CellRange(ref, ref), s"with unqualified ref '$refStr'")
      case (sheet, Resolve.Target.Range(range)) =>
        (sheet, range, s"with unqualified range '$refStr'")
    }

  /**
   * THE sheet rule over `workbook.xml` ([[Resolve.sheetName]], ADR-017 §2.5): the ref's qualifier,
   * else `-s`, else the only sheet of a single-sheet book, else `SHEET_REQUIRED` naming `context`.
   * Reads the metadata part only.
   */
  private def resolveSheetName(
    filePath: Path,
    sheetNameOpt: Option[String],
    qualified: Option[SheetName],
    context: String
  ): IO[String] =
    excel.readMetadata(filePath).flatMap { meta =>
      IO.fromEither(
        Resolve
          .sheetName(meta, sheetNameOpt, qualified, context)
          .map(_.value)
          .left
          .map(CliException(_))
      )
    }

  /**
   * Resolve target sheets for streaming search, matching --sheets/--sheet semantics: every name
   * goes through [[Resolve.knownName]], so a miss is `SHEET_NOT_FOUND` with the nearest names and a
   * name the validator refuses is `INVALID_SHEET_NAME` — the codes the in-memory twin gives.
   */
  private def resolveSearchSheets(
    filePath: Path,
    sheetNameOpt: Option[String],
    sheetsFilter: Option[String]
  ): IO[Vector[String]] =
    excel.readMetadata(filePath).flatMap { meta =>
      def known(name: String): IO[String] =
        IO.fromEither(Resolve.knownName(meta, name).map(_.value).left.map(CliException(_)))
      (sheetsFilter, sheetNameOpt) match
        case (Some(filterStr), _) =>
          val names = filterStr.split(",").map(_.trim).filter(_.nonEmpty).toVector
          if names.isEmpty then
            IO.raiseError(
              CliException(CliError.usage("--sheets requires at least one sheet name", None))
            )
          else names.traverse(known)
        case (None, Some(sheetName)) => known(sheetName).map(Vector(_))
        case (None, None) => IO.pure(meta.sheets.map(_.name.value))
    }

  /** Extract cells from row that fall within range columns. */
  private def cellsInRange(row: RowData, range: CellRange): Seq[CellValue] =
    val startCol = range.start.col.index0
    val endCol = range.end.col.index0
    row.cells.toSeq
      .filter { case (col, _) => col >= startCol && col <= endCol }
      .map(_._2)

  /** Limit range to max rows. maxRows <= 0 means "no limit" (GH-351). */
  private def limitRange(range: CellRange, maxRows: Int): CellRange =
    val rowCount = range.end.row.index0 - range.start.row.index0 + 1
    if maxRows <= 0 || rowCount <= maxRows then range
    else
      val newEndRow = range.start.row.index0 + maxRows - 1
      CellRange(range.start, ARef.from0(range.end.col.index0, newEndRow))

  /** Check if a streamed cell value is effectively empty. */
  private def isCellEmptyValue(value: CellValue): Boolean =
    value match
      case CellValue.Empty => true
      case CellValue.Text(s) if s.trim.isEmpty => true
      case CellValue.Formula(_, Some(CellValue.Empty), _) => true
      case CellValue.Formula(_, Some(CellValue.Text(s)), _) if s.trim.isEmpty => true
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
    showLabels: Boolean,
    styles: Option[WorkbookStyles] = None
  ): String = boundary:
    if rows.isEmpty then break("(empty range)")

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
          case Some(v) => formatCellValue(v, showFormulas, resolveNumFmt(colIdx, row, styles))
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
    showLabels: Boolean,
    styles: Option[WorkbookStyles] = None
  ): String =
    val (selectedRows, selectedCols) = selectRowsAndCols(rows, range, skipEmpty)

    val sb = new StringBuilder
    val lastRowIndex = selectedRows.lastOption.map(_.rowIndex)

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
            val formatted = formatCellValue(v, showFormulas, resolveNumFmt(colIdx, row, styles))
            if formatted.contains(",") || formatted.contains("\"") || formatted.contains("\n") then
              "\"" + formatted.replace("\"", "\"\"") + "\""
            else formatted
          case None => ""
      }
      sb.append(values.mkString(","))
      if !lastRowIndex.contains(row.rowIndex) then sb.append("\n")
    }

    sb.toString

  /** Format rows as JSON. */
  private def formatJson(
    rows: Vector[RowData],
    range: CellRange,
    showFormulas: Boolean,
    skipEmpty: Boolean,
    headerRow: Option[Int],
    styles: Option[WorkbookStyles] = None
  ): String =
    val startCol = range.start.col.index0
    val endCol = range.end.col.index0

    // Get header keys if headerRow specified
    val headers: Option[Map[Int, String]] = headerRow.flatMap { hr =>
      rows.find(_.rowIndex == hr).map { row =>
        (startCol to endCol).flatMap { colIdx =>
          row.cells
            .get(colIdx)
            .map(v => colIdx -> formatCellValue(v, false, resolveNumFmt(colIdx, row, styles)))
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
          val value = valueOpt
            .map(v => formatCellValue(v, showFormulas, resolveNumFmt(colIdx, row, styles)))
            .getOrElse("")
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

  /** Resolve NumFmt for a cell from streaming style info. */
  private def resolveNumFmt(
    colIdx: Int,
    row: RowData,
    styles: Option[WorkbookStyles]
  ): NumFmt =
    (for
      s <- styles
      sid <- row.cellStyles.get(colIdx)
      cs <- s.styleAt(sid)
    yield cs.numFmt).getOrElse(NumFmt.General)

  /** Format cell value for display, applying number format when available. */
  private def formatCellValue(
    value: CellValue,
    showFormulas: Boolean,
    numFmt: NumFmt = NumFmt.General
  ): String =
    value match
      case CellValue.Formula(formula, _, _) if showFormulas => formula
      case CellValue.Formula(_, Some(cached), _) => formatCellValue(cached, false, numFmt)
      case CellValue.Formula(formula, None, _) => formula
      case other => NumFmtFormatter.formatValue(other, numFmt)
