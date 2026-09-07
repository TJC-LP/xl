package com.tjclp.xl.cli.read

import scala.util.Try

import cats.effect.{IO, Ref}
import cats.syntax.all.*
import fs2.Stream

import com.tjclp.xl.addressing.{ARef, CellRange, Column, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.{FilterFormat, ViewFormat}
import com.tjclp.xl.cli.contract.{
  CliError,
  CliException,
  CliSignal,
  Outcome,
  OutputMode,
  Payload,
  Warning,
  WarningCode
}
import com.tjclp.xl.cli.helpers.{FilterPredicate, Resolve}
import com.tjclp.xl.cli.output.{CsvRenderer, Escape, Format, JsonRenderer, Markdown, RendererCommon}
import com.tjclp.xl.error.XLError

/**
 * The read verbs as one function (W2.4): `ReadQuery => SheetSource => IO[Payload]`, with
 * [[outcome]] the `IO[Outcome]` form the parity law runs. `view`, `cell`, `search`, `stats` and
 * `filter` resolve their sheet through THE sheet rule over the source's names ([[Resolve]]), read
 * [[CellRecord]]s from the source and render them; the source's strategy — loaded or streaming —
 * never shows in the rendering, only in what it can answer ([[Capability]]).
 *
 * @param sheetFlag
 *   the run's `-s`, applied as step 2 of the rule
 * @param mode
 *   selects the typed JSON payloads of `search`, `stats` and `cell` under `--json`; the
 *   pass-through formats of `view` and `filter` are already resolved in the query
 * @param warn
 *   the run's warning sink (truncation, hidden lines, an ignored flag, an advisory `--eval`
 *   failure)
 */
object Reads:

  def run(
    query: ReadQuery,
    source: SheetSource,
    sheetFlag: Option[String],
    mode: OutputMode,
    warn: Warning => IO[Unit]
  ): IO[Payload] =
    refuse(query, source) *> (query match
      case q: ReadQuery.View => view(q, source, sheetFlag, mode, warn)
      case q: ReadQuery.Cell => cell(q, source, sheetFlag, mode)
      case q: ReadQuery.Search => search(q, source, sheetFlag, mode)
      case q: ReadQuery.Stats => stats(q, source, sheetFlag, mode)
      case q: ReadQuery.Filter => filter(q, source, sheetFlag, mode))

  /**
   * [[run]] as a complete [[Outcome]]: the payload with the warnings the run raised, or the failure
   * classified like every other verb's ([[CliError.fromThrowable]]). What the parity law compares.
   */
  def outcome(
    query: ReadQuery,
    source: SheetSource,
    sheetFlag: Option[String],
    mode: OutputMode
  ): IO[Outcome] =
    Ref.of[IO, Vector[Warning]](Vector.empty).flatMap { warnings =>
      run(query, source, sheetFlag, mode, w => warnings.update(_ :+ w)).attempt.flatMap { attempt =>
        warnings.get.map { collected =>
          attempt match
            case Right(payload) => Outcome.ok(query.verb, payload, collected)
            case Left(signal: CliSignal) =>
              Outcome.signal(query.verb, signal.payload, signal.error, collected)
            case Left(err) => Outcome.failed(query.verb, CliError.fromThrowable(err), collected)
        }
      }
    }

  // --- Refusals ---------------------------------------------------------------------------------

  /**
   * A query needing a capability the source lacks is refused before anything is read:
   * `UNSUPPORTED_IN_STREAM` with the in-memory alternative as the hint (the only source lacking
   * capabilities is the streaming one).
   */
  private def refuse(query: ReadQuery, source: SheetSource): IO[Unit] =
    val missing = query.needs -- source.capabilities
    query match
      case _ if missing.isEmpty => IO.unit
      case _: ReadQuery.View if missing.contains(Capability.Eval) =>
        IO.raiseError(
          SheetSource.unsupported(
            "--eval is not supported with --stream (streaming view uses cached values only)",
            "omit --stream to evaluate formulas; use --max-size <MB> for a large file"
          )
        )
      case v: ReadQuery.View if missing.contains(Capability.Render) =>
        IO.raiseError(
          SheetSource.unsupported(
            s"--stream not supported for ${v.format.toString.toLowerCase} (needs styles). " +
              "Remove --stream flag or use markdown/csv/json format.",
            "omit --stream, or use --format markdown, csv or json"
          )
        )
      case _ =>
        IO.raiseError(
          SheetSource.unsupported(
            s"${query.verb} needs ${missing.toVector.map(_.name).sorted.mkString(", ")}, which " +
              "--stream cannot provide",
            "omit --stream; use --max-size <MB> to load a large file in memory"
          )
        )

  private def lift[A](result: Either[CliError, A]): IO[A] =
    IO.fromEither(result.left.map(CliException(_)))

  /** A wrong flag or argument combination: `USAGE` (exit 2). */
  private def usage(message: String): CliException = CliException(CliError.usage(message, None))

  /** A ref of the wrong shape (a range where one cell is needed): `INVALID_REFERENCE`. */
  private def invalidRef(reason: String): CliException =
    CliException(Resolve.invalidReference(reason))

  /** THE sheet rule over the source's names: qualifier, `-s`, the only sheet, `SHEET_REQUIRED`. */
  private def resolveSheet(
    source: SheetSource,
    sheetFlag: Option[String],
    qualified: Option[SheetName],
    context: String
  ): IO[SheetName] =
    source.sheets.flatMap(names =>
      lift(Resolve.sheetNameAmong(names, sheetFlag, qualified, context))
    )

  /**
   * JSON text as the run's payload: spliced under `--json` ([[Payload.Raw]]), printed otherwise.
   */
  private def jsonPayload(mode: OutputMode, json: String): Payload = mode match
    case OutputMode.Json => Payload.Raw(json)
    case OutputMode.Text => Payload.text(json)

  // --- view -------------------------------------------------------------------------------------

  private def view(
    q: ReadQuery.View,
    source: SheetSource,
    sheetFlag: Option[String],
    mode: OutputMode,
    warn: Warning => IO[Unit]
  ): IO[Payload] =
    for
      parsed <- q.range.traverse(r => lift(Resolve.ref(r)))
      context = (q.range, parsed) match
        case (Some(refStr), Some((_, target))) => Resolve.unqualified("view", refStr, target)
        case _ => "view"
      sheet <- resolveSheet(source, sheetFlag, parsed.flatMap(_._1), context)
      base <- parsed match
        case Some((_, Resolve.Target.Cell(ref))) => IO.pure(Some(CellRange(ref, ref)))
        case Some((_, Resolve.Target.Range(range))) => IO.pure(Some(range))
        case None => source.usedRange(sheet) // no range given: the used range
      payload <- base match
        case Some(range) => viewWindow(q, source, sheet, range, mode, warn)
        case None => viewEmptySheet(q, source, sheet, mode, warn)
    yield payload

  /** `view` with no range on an empty sheet: nothing to address. */
  private def viewEmptySheet(
    q: ReadQuery.View,
    source: SheetSource,
    sheet: SheetName,
    mode: OutputMode,
    warn: Warning => IO[Unit]
  ): IO[Payload] =
    q.format match
      case ViewFormat.Markdown => IO.pure(Payload.text("(empty sheet)"))
      case ViewFormat.Csv => IO.pure(Payload.text(""))
      case ViewFormat.Json =>
        IO.pure(jsonPayload(mode, JsonRenderer.renderEmptySheet(sheet, q.headerRow.isDefined)))
      case _ =>
        // A picture of an empty sheet: its first cell
        val origin = ARef.from0(0, 0)
        source.render(sheet, CellRange(origin, origin), renderSpec(q), warn).map(Payload.text)

  private def renderSpec(q: ReadQuery.View): RenderSpec =
    RenderSpec(
      q.format,
      q.evalFormulas,
      q.strict,
      q.printScale,
      q.showGridlines,
      q.showLabels,
      q.dpi,
      q.quality,
      q.rasterOutput,
      q.rasterizer
    )

  private def viewWindow(
    q: ReadQuery.View,
    source: SheetSource,
    sheet: SheetName,
    range: CellRange,
    mode: OutputMode,
    warn: Warning => IO[Unit]
  ): IO[Payload] =
    val totalRows = range.height
    val totalCols = range.width
    for
      _ <- IO.raiseError(usage("--offset must be 0 or more")).whenA(q.offset < 0)
      _ <- IO.raiseError(usage("--max-cols must be 0 or more")).whenA(q.maxCols < 0)
      _ <- IO
        .raiseError(
          usage(s"--offset ${q.offset} skips every row of ${range.toA1} ($totalRows rows)")
        )
        .whenA(q.offset >= totalRows)
      afterOffset = CellRange(
        ARef.from0(range.start.col.index0, range.start.row.index0 + q.offset),
        range.end
      )
      window = limitCols(limitRows(afterOffset, q.limit), q.maxCols)
      shownRows = window.height
      shownCols = window.width
      rowsClipped = shownRows < totalRows
      colsClipped = shownCols < totalCols
      // GH-351: report truncation when --limit (or --offset/--max-cols) clips the requested range
      rowNotice =
        if q.offset == 0 then RendererCommon.truncationNotice(shownRows, totalRows)
        else RendererCommon.pagingNotice(q.offset + 1, q.offset + shownRows, totalRows)
      notices = Vector(
        Option.when(rowsClipped)(Warning(WarningCode.TRUNCATED, rowNotice)),
        Option.when(colsClipped)(
          Warning(WarningCode.TRUNCATED, RendererCommon.columnNotice(shownCols, totalCols))
        )
      ).flatten
      // GH-474: --skip-hidden cannot be honoured without the hidden capability — say so
      _ <- warn(Warning(WarningCode.FLAG_IGNORED, RendererCommon.streamingSkipHiddenNotice))
        .whenA(q.skipHidden && !source.capabilities.contains(Capability.Hidden))
      payload <- q.format match
        case ViewFormat.Markdown =>
          gridFor(q, source, sheet, window, warn).map { grid =>
            val table = Markdown.render(grid, q.showFormulas, q.skipEmpty, q.skipHidden)
            val hiddenNote = RendererCommon.hiddenNotice(grid, q.skipHidden)
            Payload.text((table +: notices.map(_.message)).appendedAll(hiddenNote).mkString("\n"))
          }
        case ViewFormat.Json =>
          gridFor(q, source, sheet, window, warn).flatMap { grid =>
            headerRecords(q.headerRow, grid, source, sheet).map { header =>
              jsonPayload(
                mode,
                JsonRenderer.render(
                  grid,
                  header,
                  q.skipEmpty,
                  Option.when(rowsClipped)(totalRows),
                  Option.when(colsClipped)(totalCols),
                  q.skipHidden
                )
              )
            }
          }
        case ViewFormat.Csv =>
          gridFor(q, source, sheet, window, warn).flatMap { grid =>
            val csv =
              CsvRenderer.render(grid, q.showFormulas, q.showLabels, q.skipEmpty, q.skipHidden)
            val hiddenWarning =
              RendererCommon
                .hiddenNotice(grid, q.skipHidden)
                .map(Warning(WarningCode.HIDDEN_OMITTED, _))
            // CSV stdout must stay machine-parseable: notices are warnings only
            (notices ++ hiddenWarning).traverse_(warn).as(Payload.text(csv))
          }
        case ViewFormat.Html =>
          source.render(sheet, window, renderSpec(q), warn).flatMap { html =>
            // In-band marker as trailing HTML comments (comments after the root element are valid
            // HTML), plus the TRUNCATED warnings out of band
            val marked = (html +: notices.map(n => s"<!-- ${n.message} -->")).mkString("\n")
            notices.traverse_(warn).as(Payload.text(marked))
          }
        case ViewFormat.Svg =>
          // SVG stdout must stay a clean XML document: the notices are warnings only
          source.render(sheet, window, renderSpec(q), warn).flatMap { svg =>
            notices.traverse_(warn).as(Payload.text(svg))
          }
        case ViewFormat.Png | ViewFormat.Jpeg | ViewFormat.WebP | ViewFormat.Pdf =>
          // Binary goes to --raster-output, so the notices can ride the status line
          source.render(sheet, window, renderSpec(q), warn).map { exported =>
            Payload.text((exported +: notices.map(_.message)).mkString("\n"))
          }
    yield payload

  /** The window's records, evaluated first under `--eval`. */
  private def gridFor(
    q: ReadQuery.View,
    source: SheetSource,
    sheet: SheetName,
    window: CellRange,
    warn: Warning => IO[Unit]
  ): IO[RecordGrid] =
    if q.evalFormulas then source.evaluated(sheet, window, q.strict, warn)
    else source.grid(sheet, window)

  /**
   * The `--header-row` records for the JSON `records` mode: the grid's own row when the header lies
   * inside the window, else the header row read over the window's columns.
   */
  private def headerRecords(
    headerRow: Option[Int],
    grid: RecordGrid,
    source: SheetSource,
    sheet: SheetName
  ): IO[Option[(Int, Vector[CellRecord])]] =
    headerRow.traverse { rowNum =>
      val idx = rowNum - 1
      grid.rows.lift(idx - grid.range.start.row.index0) match
        case Some(records) => IO.pure((idx, records))
        case None =>
          val headerRange = CellRange(
            ARef.from0(grid.range.start.col.index0, idx),
            ARef.from0(grid.range.end.col.index0, idx)
          )
          source
            .rows(sheet, headerRange)
            .compile
            .last
            .map(records => (idx, records.getOrElse(Vector.empty)))
    }

  /** Cap the rows of a range; `maxRows <= 0` means no limit (GH-351). */
  private def limitRows(range: CellRange, maxRows: Int): CellRange =
    if maxRows <= 0 || range.height <= maxRows then range
    else
      val newEndRow = range.start.row.index0 + maxRows - 1
      CellRange(range.start, ARef.from0(range.end.col.index0, newEndRow))

  /** Cap the columns of a range; `maxCols <= 0` means no limit. */
  private def limitCols(range: CellRange, maxCols: Int): CellRange =
    if maxCols <= 0 || range.width <= maxCols then range
    else
      val newEndCol = range.start.col.index0 + maxCols - 1
      CellRange(range.start, ARef.from0(newEndCol, range.end.row.index0))

  // --- cell -------------------------------------------------------------------------------------

  private def cell(
    q: ReadQuery.Cell,
    source: SheetSource,
    sheetFlag: Option[String],
    mode: OutputMode
  ): IO[Payload] =
    for
      parsed <- lift(Resolve.ref(q.ref))
      (qualifier, target) = parsed
      sheet <- resolveSheet(
        source,
        sheetFlag,
        qualifier,
        Resolve.unqualified("cell", q.ref, target)
      )
      ref <- target match
        case Resolve.Target.Cell(r) => IO.pure(r)
        case Resolve.Target.Range(range) =>
          IO.raiseError(invalidRef(s"cell requires a single cell, not the range ${range.toA1}"))
      detail <- source.detail(sheet, ref, withStyle = !q.noStyle)
    yield mode match
      case OutputMode.Text => Payload.text(Format.cellInfo(detail))
      case OutputMode.Json => Payload.Raw(cellJson(detail))

  /**
   * `cell --json`: the typed record plus `style`, `comment`, `hyperlink`, `dependencies` and
   * `dependents` (`null` when the source cannot compute them).
   */
  private def cellJson(detail: CellDetail): String =
    val record = detail.record
    val style = record.style.fold("null")(CellRecord.styleJson)
    val comment = detail.comment.fold("null") { c =>
      s"""{"text": ${Escape.json(c.text.toPlainText)}, "author": ${c.author.fold("null")(
          Escape.json
        )}}"""
    }
    val hyperlink = detail.hyperlink.fold("null")(Escape.json)
    val dependencies = detail.dependencies.map(Escape.json).mkString("[", ", ", "]")
    val dependents = detail.dependents.fold("null")(_.map(Escape.json).mkString("[", ", ", "]"))
    s"""{${record.typedFields}, "style": $style, "comment": $comment, "hyperlink": $hyperlink, """ +
      s""""dependencies": $dependencies, "dependents": $dependents}"""

  // --- search -----------------------------------------------------------------------------------

  private def search(
    q: ReadQuery.Search,
    source: SheetSource,
    sheetFlag: Option[String],
    mode: OutputMode
  ): IO[Payload] =
    for
      regex <- IO.fromEither(
        Try(q.pattern.r).toEither.left.map(e => usage(s"Invalid regex pattern: ${e.getMessage}"))
      )
      names <- source.sheets
      // Which sheets to search: --sheets, else -s, else every sheet
      targets <- (q.sheetsFilter, sheetFlag) match
        case (Some(filterStr), _) =>
          val listed = filterStr.split(",").map(_.trim).filter(_.nonEmpty).toVector
          if listed.isEmpty then IO.raiseError(usage("--sheets requires at least one sheet name"))
          else listed.traverse(name => lift(Resolve.knownAmong(names, name)))
        case (None, Some(flag)) => lift(Resolve.knownAmong(names, flag)).map(Vector(_))
        case (None, None) => IO.pure(names)
      // GH-351: the true total is always reported; only the first --limit matches are kept
      scanned <- Stream
        .emits(targets)
        .flatMap(source.occupied)
        .filter(record => regex.findFirstIn(record.searchText).isDefined)
        .compile
        .fold((Vector.empty[CellRecord], 0)) { case ((kept, total), record) =>
          (if q.limit <= 0 || kept.size < q.limit then kept :+ record else kept, total + 1)
        }
      (shown, total) = scanned
    yield
      val body = mode match
        case OutputMode.Text =>
          val sheetDesc = targets match
            case Vector(single) => single.value
            case many => s"${many.size} sheets"
          val table = Markdown.renderSearchResultsWithRef(
            shown.map(r => (s"${r.sheet.value}!${r.ref.toA1}", r.searchText))
          )
          val base = s"Found $total matches in $sheetDesc:\n\n$table"
          if shown.size < total then
            s"$base\n${RendererCommon.truncationNotice(shown.size, total, "matches")}"
          else base
        case OutputMode.Json =>
          val sheets = targets.map(n => Escape.json(n.value)).mkString("[", ", ", "]")
          val matches = shown.map(_.toJson(legacyKeys = false)).mkString("[", ", ", "]")
          s"""{"pattern": ${Escape.json(
              q.pattern
            )}, "sheets": $sheets, "count": ${shown.size}, """ +
            s""""total": $total, "matches": $matches}"""
      mode match
        case OutputMode.Json => Payload.Raw(body)
        case OutputMode.Text => Payload.text(body)

  // --- stats ------------------------------------------------------------------------------------

  /** Running statistics over the numeric records of a range: O(1) whatever the range. */
  private final case class StatsAcc(
    count: Long,
    sum: BigDecimal,
    min: Option[BigDecimal],
    max: Option[BigDecimal]
  ):
    def add(n: BigDecimal): StatsAcc =
      StatsAcc(count + 1, sum + n, Some(min.fold(n)(_ min n)), Some(max.fold(n)(_ max n)))

  private def stats(
    q: ReadQuery.Stats,
    source: SheetSource,
    sheetFlag: Option[String],
    mode: OutputMode
  ): IO[Payload] =
    for
      parsed <- lift(Resolve.ref(q.ref))
      (qualifier, target) = parsed
      sheet <- resolveSheet(
        source,
        sheetFlag,
        qualifier,
        Resolve.unqualified("stats", q.ref, target)
      )
      range = target.toEither.fold(ref => CellRange(ref, ref), identity)
      acc <- source
        .rows(sheet, range)
        .flatMap(Stream.emits)
        .collect { record =>
          record.value match
            case CellValue.Number(n) => n
        }
        .compile
        .fold(StatsAcc(0, BigDecimal(0), None, None))(_.add(_))
      _ <- IO
        .raiseError(
          CliException(
            CliError.fromXLError(XLError.Other(s"No numeric values in range ${range.toA1}"), None)
          )
        )
        .whenA(acc.count == 0)
    yield
      val count = acc.count
      val sum = acc.sum
      val min = acc.min.getOrElse(BigDecimal(0))
      val max = acc.max.getOrElse(BigDecimal(0))
      val mean = sum / BigDecimal(count)
      mode match
        case OutputMode.Text =>
          Payload.text(
            f"count: $count, sum: $sum%.2f, min: $min%.2f, max: $max%.2f, mean: $mean%.2f"
          )
        case OutputMode.Json =>
          Payload.Raw(
            s"""{"sheet": ${Escape.json(
                sheet.value
              )}, "range": "${range.toA1}", "count": $count, """ +
              s""""sum": ${CellRecord.numberLexeme(sum)}, "min": ${CellRecord.numberLexeme(
                  min
                )}, """ +
              s""""max": ${CellRecord.numberLexeme(max)}, "mean": ${CellRecord.numberLexeme(
                  mean
                )}}"""
          )

  // --- filter -----------------------------------------------------------------------------------

  private def filter(
    q: ReadQuery.Filter,
    source: SheetSource,
    sheetFlag: Option[String],
    mode: OutputMode
  ): IO[Payload] =
    for
      sheet <- resolveSheet(source, sheetFlag, None, "filter")
      pred <- lift(
        FilterPredicate
          .parse(q.where)
          .left
          .map(err => CliError.usage(s"Invalid filter predicate: $err", None))
      )
      used <- source.usedRange(sheet)
      text <- used match
        case None => IO.pure(filterNoData(q.format, Vector.empty))
        case Some(range) => runFilter(q, source, sheet, range, pred)
    yield if q.format == FilterFormat.Json then jsonPayload(mode, text) else Payload.text(text)

  private def runFilter(
    q: ReadQuery.Filter,
    source: SheetSource,
    sheet: SheetName,
    range: CellRange,
    pred: FilterPredicate.Pred
  ): IO[String] =
    val firstCol = range.start.col.index0
    val usedCols = (firstCol to range.end.col.index0).toVector
    val headerRowIdx = range.start.row.index0
    val headerRange =
      CellRange(ARef.from0(firstCol, headerRowIdx), ARef.from0(range.end.col.index0, headerRowIdx))
    for
      header <-
        if q.header then source.rows(sheet, headerRange).compile.last.map(_.getOrElse(Vector.empty))
        else IO.pure(Vector.empty[CellRecord])
      // Header names from the first used row (original casing; matched case-insensitively)
      headerNames = header.flatMap { record =>
        FilterPredicate
          .textOf(record.cellValue)
          .map(_.trim)
          .filter(_.nonEmpty)
          .map(_ -> record.ref.col.index0)
      }
      headers = headerNames.map((name, col) => name.toLowerCase -> col).toMap
      resolve = (name: String) =>
        headers
          .get(name.toLowerCase)
          .orElse(
            Column.fromLetter(name.toUpperCase).toOption.map(_.index0).filter(usedCols.contains)
          )
      // Validate every referenced column upfront for a proper error message
      unresolved = FilterPredicate.columnRefs(pred).filter(resolve(_).isEmpty)
      _ <- IO
        .raiseError(
          invalidRef {
            val available =
              if q.header then
                s"available headers: ${headerNames.map(_._1).mkString(", ")}; " +
                  s"used columns: ${range.start.col.toLetter}:${range.end.col.toLetter}"
              else s"used columns: ${range.start.col.toLetter}:${range.end.col.toLetter}"
            s"Unknown column(s) in predicate: ${unresolved.toVector.sorted.mkString(", ")} ($available)"
          }
        )
        .whenA(unresolved.nonEmpty)
      selectedCols <- lift(parseColumns(q.columns, usedCols).left.map(CliError.usage(_, None)))
      // One pass over the used range: keep the first --limit matching rows, count them all
      scanned <- source
        .rows(sheet, range)
        .filter { row =>
          val rowIdx = row.headOption.fold(-1)(_.ref.row.index0)
          !(q.header && rowIdx == headerRowIdx) &&
          FilterPredicate.evaluate(pred, resolve, col => row.lift(col - firstCol).map(_.cellValue))
        }
        .compile
        .fold((Vector.empty[Vector[CellRecord]], 0)) { case ((kept, total), row) =>
          (if kept.size < math.max(0, q.limit) then kept :+ row else kept, total + 1)
        }
      (shown, total) = scanned
    yield
      val labels = selectedCols.map { col =>
        val letter = Column.from0(col).toLetter
        if q.header then
          header
            .lift(col - firstCol)
            .flatMap(record => FilterPredicate.textOf(record.cellValue))
            .map(_.trim)
            .filter(_.nonEmpty)
            .getOrElse(letter)
        else letter
      }
      def cellsOf(row: Vector[CellRecord]): Vector[Option[CellRecord]] =
        selectedCols.map(col => row.lift(col - firstCol))
      q.format match
        case FilterFormat.Markdown => filterMarkdown(shown.map(cellsOf), total, labels)
        case FilterFormat.Csv => filterCsv(shown.map(cellsOf), labels)
        case FilterFormat.Json => filterJson(shown.map(cellsOf), labels)

  /** Parse "A,C:E" into 0-based column indices; None = all used columns. */
  private def parseColumns(
    spec: Option[String],
    usedCols: Vector[Int]
  ): Either[String, Vector[Int]] =
    spec match
      case None => Right(usedCols)
      case Some(s) =>
        s.split(",").toVector.map(_.trim).filter(_.nonEmpty).flatTraverse { token =>
          token.split(":", -1) match
            case Array(single) =>
              Column.fromLetter(single.trim.toUpperCase).map(c => Vector(c.index0))
            case Array(lo, hi) =>
              for
                l <- Column.fromLetter(lo.trim.toUpperCase)
                h <- Column.fromLetter(hi.trim.toUpperCase)
              yield (math.min(l.index0, h.index0) to math.max(l.index0, h.index0)).toVector
            case _ => Left(s"Invalid --columns token '$token' (use A or A:C)")
        }

  /** The row number of a rendered filter row: every record of the row shares it. */
  private def rowNumber(cells: Vector[Option[CellRecord]]): String =
    cells.flatten.headOption.fold("")(r => (r.ref.row.index0 + 1).toString)

  private def filterNoData(format: FilterFormat, labels: Vector[String]): String =
    format match
      case FilterFormat.Markdown => "No rows matched."
      case FilterFormat.Csv => ("row" +: labels).mkString(",")
      case FilterFormat.Json => "[]"

  private def filterMarkdown(
    shown: Vector[Vector[Option[CellRecord]]],
    totalMatched: Int,
    labels: Vector[String]
  ): String =
    if totalMatched == 0 then "No rows matched."
    else
      val headers = "Row" +: labels
      val rows = shown.map { cells =>
        rowNumber(cells) +: cells.map(_.fold("")(_.text(showFormulas = false)))
      }
      val widths = headers.indices.map { i =>
        val dataMax = rows.map(r => Escape.markdown(r(i)).length).maxOption.getOrElse(0)
        math.max(3, math.max(headers(i).length, dataMax))
      }
      val sb = new StringBuilder
      sb.append(headers.zip(widths).map((h, w) => s" ${h.padTo(w, ' ')} ").mkString("|", "|", "|"))
      sb.append("\n")
      sb.append(widths.map(w => "-" * (w + 2)).mkString("|", "|", "|"))
      sb.append("\n")
      rows.foreach { r =>
        sb.append(
          r.zip(widths)
            .map((c, w) => s" ${Escape.markdown(c).padTo(w, ' ')} ")
            .mkString("|", "|", "|")
        )
        sb.append("\n")
      }
      if totalMatched > shown.length then
        sb.append(s"\nMatched $totalMatched row(s); showing first ${shown.length} (--limit).\n")
      else sb.append(s"\n$totalMatched row(s) matched.\n")
      sb.toString

  private def filterCsv(shown: Vector[Vector[Option[CellRecord]]], labels: Vector[String]): String =
    val header = ("row" +: labels).map(Escape.csv).mkString(",")
    val lines = shown.map { cells =>
      (rowNumber(cells) +: cells.map(_.fold("")(_.text(showFormulas = false))))
        .map(Escape.csv)
        .mkString(",")
    }
    (header +: lines).mkString("\n")

  /**
   * `[{"row": n, "cells": {label: value}}]`, every number lexeme exact; an uncached formula shows
   * its expression as a string, as it always has.
   */
  private def filterJson(
    shown: Vector[Vector[Option[CellRecord]]],
    labels: Vector[String]
  ): String =
    def valueJson(record: CellRecord): String = record.formula match
      case Some(f) if !f.cached => Escape.json(f.text)
      case _ => record.rawJson
    val rows = shown.map { cells =>
      val fields = cells.zip(labels).map { (cell, label) =>
        s"${Escape.json(label)}: ${cell.fold("null")(valueJson)}"
      }
      s"""{"row": ${rowNumber(cells)}, "cells": {${fields.mkString(", ")}}}"""
    }
    val text = rows.mkString("[", ", ", "]")
    Try(ujson.reformat(text, indent = 2)).getOrElse(text)
