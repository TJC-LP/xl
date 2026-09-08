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
  ErrorCode,
  Outcome,
  OutputMode,
  Payload,
  StreamedBody,
  Warning,
  WarningCode
}
import com.tjclp.xl.cli.helpers.{FilterPredicate, Resolve}
import com.tjclp.xl.cli.output.{CsvRenderer, Escape, Format, JsonRenderer, Markdown, RendererCommon}

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
      case q: ReadQuery.Stats => stats(q, source, sheetFlag, mode, warn)
      case q: ReadQuery.Filter => filter(q, source, sheetFlag, mode, warn))

  /**
   * [[run]] as a complete [[Outcome]]: the payload — a streamed table gathered into the text it
   * composes ([[Payload.materialise]]) — with the warnings the run raised, or the failure
   * classified like every other verb's ([[CliError.fromThrowable]]). What the parity law compares.
   */
  def outcome(
    query: ReadQuery,
    source: SheetSource,
    sheetFlag: Option[String],
    mode: OutputMode
  ): IO[Outcome] =
    Ref.of[IO, Vector[Warning]](Vector.empty).flatMap { warnings =>
      run(query, source, sheetFlag, mode, w => warnings.update(_ :+ w))
        .flatMap(Payload.materialise(_))
        .attempt
        .flatMap { attempt =>
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
      // A 1-based row: 0 or less would address a row above the sheet
      _ <- IO
        .raiseError(usage(s"--header-row must be 1 or more (got ${q.headerRow.getOrElse(0)})"))
        .whenA(q.headerRow.exists(_ < 1))
      // The counts: 0 means no limit, below 0 means nothing (GH-635 — -1 used to read as 0)
      _ <- IO.raiseError(usage(s"--limit must be 0 or more (got ${q.limit})")).whenA(q.limit < 0)
      _ <- IO.raiseError(usage("--offset must be 0 or more")).whenA(q.offset < 0)
      _ <- IO.raiseError(usage("--max-cols must be 0 or more")).whenA(q.maxCols < 0)
      parsed <- q.range.traverse(r => lift(Resolve.ref(r)))
      context = (q.range, parsed) match
        case (Some(refStr), Some((_, target))) => Resolve.unqualified("view", refStr, target)
        case _ => "view"
      sheet <- resolveSheet(source, sheetFlag, parsed.flatMap(_._1), context)
      base <- parsed match
        case Some((_, Resolve.Target.Cell(ref))) => IO.pure(Some(CellRange(ref, ref)))
        case Some((_, Resolve.Target.Range(range))) => clampSpan(source, sheet, range).map(Some(_))
        case None => source.usedRange(sheet) // no range given: the used range
      payload <- base match
        case Some(range) => viewWindow(q, source, sheet, range, mode, warn)
        case None => viewEmptySheet(q, source, sheet, mode, warn)
    yield payload

  /**
   * A whole-column (`B:B`) or whole-row (`3:3`) span clamped to the sheet's used range on the axis
   * it left open (GH-641): the rows of the used range for a column span, its columns for a row span
   * — so `view B:B` renders the column's used rows (and `totalRows` counts them), not the
   * 1,048,576-row axis; the columns (rows) named stay exactly as asked, so `view Z:Z` on a sheet
   * used through C shows an empty column Z over the used rows. An empty sheet clamps to row 1 (or
   * column A). Any other range is returned unchanged.
   */
  private def clampSpan(source: SheetSource, sheet: SheetName, range: CellRange): IO[CellRange] =
    if !range.isFullColumn && !range.isFullRow then IO.pure(range)
    else
      source.usedRange(sheet).map { used =>
        if range.isFullColumn then
          val (top, bottom) = used.fold((0, 0))(u => (u.start.row.index0, u.end.row.index0))
          CellRange(
            ARef.from0(range.start.col.index0, top),
            ARef.from0(range.end.col.index0, bottom)
          )
        else
          val (left, right) = used.fold((0, 0))(u => (u.start.col.index0, u.end.col.index0))
          CellRange(
            ARef.from0(left, range.start.row.index0),
            ARef.from0(right, range.end.row.index0)
          )
      }

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
        case ViewFormat.Markdown | ViewFormat.Json | ViewFormat.Csv =>
          table(
            q,
            source,
            sheet,
            window,
            notices,
            Option.when(rowsClipped)(totalRows),
            Option.when(colsClipped)(totalCols),
            warn
          )
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

  /**
   * The three table formats, written row by row (GH-635): the window's rows stream from the source
   * — the loaded sheet's from memory, the streaming reader's from the file — through the renderer
   * into a [[Payload.Streamed]], so no window is ever held. Markdown, and csv under `--skip-empty`,
   * first fold the rows once for what every row must know about the others (the column widths, the
   * empty columns — [[ColumnFacts]], O(columns) memory), then stream them a second time; under
   * `--stream` that is two reads of the window. `--eval` evaluates the window first (the loaded
   * workbook only) and streams the evaluated grid. Every notice — truncation, hidden lines, an
   * ignored flag — is decided here, before the first row: markdown appends them after its table,
   * csv raises them as warnings so stdout stays machine-parseable, json carries them as fields.
   */
  private def table(
    q: ReadQuery.View,
    source: SheetSource,
    sheet: SheetName,
    window: CellRange,
    notices: Vector[Warning],
    truncatedTotalRows: Option[Int],
    truncatedTotalCols: Option[Int],
    warn: Warning => IO[Unit]
  ): IO[Payload] =
    for
      evaluated <-
        if q.evalFormulas then source.evaluated(sheet, window, q.strict, warn).map(Some(_))
        else IO.pure(None)
      hidden <- evaluated.fold(source.hiddenLines(sheet, window))(g =>
        IO.pure((g.hiddenRows, g.hiddenCols))
      )
      shape = RecordWindow(sheet, window, hidden._1, hidden._2)
      rows = evaluated.fold(source.rows(sheet, window))(g => Stream.emits(g.rows).covary[IO])
      hiddenNote = RendererCommon.hiddenNotice(shape, q.skipHidden)
      payload <- q.format match
        case ViewFormat.Markdown =>
          ColumnFacts
            .of(shape, rows, Markdown.cellText(q.showFormulas), q.skipEmpty, q.skipHidden)
            .compile
            .lastOrError
            .map { facts =>
              val cols = Markdown.columns(shape, facts, q.skipEmpty, q.skipHidden)
              val widths = Markdown.columnWidths(cols, facts)
              val label = facts.labelWidth
              val body = Markdown
                .lines(shape, rows, cols, widths, label, q.showFormulas, q.skipEmpty, q.skipHidden)
              // The table's own trailing newline, then the notices, as the text always read
              val after = ("" +: notices.map(_.message)).appendedAll(hiddenNote)
              Payload.Streamed(
                StreamedBody.Lines(Markdown.header(cols, widths, label), body, after)
              )
            }
        case ViewFormat.Csv =>
          val facts =
            if CsvRenderer.needsFacts(q.skipEmpty) then
              ColumnFacts
                .of(shape, rows, _.text(q.showFormulas), q.skipEmpty, q.skipHidden)
                .compile
                .lastOrError
            else IO.pure(ColumnFacts.empty)
          facts.flatMap { facts =>
            val cols = CsvRenderer.columns(shape, facts, q.skipEmpty, q.skipHidden)
            val body = CsvRenderer
              .lines(shape, rows, cols, q.showFormulas, q.showLabels, q.skipEmpty, q.skipHidden)
            val hiddenWarning = hiddenNote.map(Warning(WarningCode.HIDDEN_OMITTED, _))
            // CSV stdout must stay machine-parseable: notices are warnings only
            (notices ++ hiddenWarning)
              .traverse_(warn)
              .as(
                Payload.Streamed(
                  StreamedBody.Lines(CsvRenderer.header(cols, q.showLabels), body, Vector.empty)
                )
              )
          }
        case _ =>
          headerRecords(q.headerRow, evaluated, source, sheet, window).map { header =>
            Payload.Streamed(
              StreamedBody.JsonArray(
                JsonRenderer.head(shape, header.isDefined, truncatedTotalRows, truncatedTotalCols),
                JsonRenderer.elements(shape, rows, header, q.skipEmpty, q.skipHidden),
                JsonRenderer.tail
              )
            )
          }
    yield payload

  /**
   * The `--header-row` records for the JSON `records` mode: the evaluated grid's own row when
   * `--eval` computed it and the header lies inside the window (so a formula header shows its
   * value), else the header row read over the window's columns — one row, from either source.
   */
  private def headerRecords(
    headerRow: Option[Int],
    evaluated: Option[RecordGrid],
    source: SheetSource,
    sheet: SheetName,
    window: CellRange
  ): IO[Option[(Int, Vector[CellRecord])]] =
    headerRow.traverse { rowNum =>
      val idx = rowNum - 1
      evaluated.flatMap(_.rows.lift(idx - window.start.row.index0)) match
        case Some(records) => IO.pure((idx, records))
        case None =>
          val headerRange = CellRange(
            ARef.from0(window.start.col.index0, idx),
            ARef.from0(window.end.col.index0, idx)
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
   * `dependents` — the two graph lists `null` when the source cannot compute them.
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
    def refList(refs: Option[Vector[String]]): String =
      refs.fold("null")(_.map(Escape.json).mkString("[", ", ", "]"))
    val dependencies = refList(detail.dependencies)
    val dependents = refList(detail.dependents)
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
      // GH-351: only the first --limit matches are listed, the total is reported. GH-637: by
      // default the scan stops ONE match past the limit — enough to know more exist without reading
      // the rest of a million-row sheet — and the total is a lower bound; --total (or --limit 0)
      // reads every cell for the exact total. The sources are lazy: a streamed sheet's parser is
      // interrupted where the pull ends, and a sheet after the one that filled the limit is never
      // opened
      matched = Stream
        .emits(targets)
        .flatMap(source.occupied)
        .filter(record => regex.findFirstIn(record.searchText).isDefined)
      bounded = q.limit > 0 && !q.exactTotal
      scanned <- (if bounded then matched.take(q.limit.toLong + 1) else matched).compile
        .fold((Vector.empty[CellRecord], 0)) { case ((kept, seen), record) =>
          (if q.limit <= 0 || kept.size < q.limit then kept :+ record else kept, seen + 1)
        }
      (shown, total) = scanned
      // The bounded scan that ran past the limit saw `limit + 1` matches: at least that many exist
      totalExact = !bounded || total <= q.limit
    yield
      val body = mode match
        case OutputMode.Text =>
          val sheetDesc = targets match
            case Vector(single) => single.value
            case many => s"${many.size} sheets"
          // the qualifier as a formula spells it (`'On-Premise'!A1`), the printer `cell`, `deps`
          // and FormulaPrinter share (`SheetName.quoteForFormula`, GH-609)
          val table = Markdown.renderSearchResultsWithRef(
            shown.map(r =>
              (s"${SheetName.quoteForFormula(r.sheet.value)}!${r.ref.toA1}", r.searchText)
            )
          )
          val found =
            if totalExact then s"Found $total matches" else s"Found at least $total matches"
          val base = s"$found in $sheetDesc:\n\n$table"
          if shown.size < total then
            val notice =
              if totalExact then RendererCommon.truncationNotice(shown.size, total, "matches")
              else RendererCommon.searchStoppedNotice(shown.size)
            s"$base\n$notice"
          else base
        case OutputMode.Json =>
          val sheets = targets.map(n => Escape.json(n.value)).mkString("[", ", ", "]")
          val matches = shown.map(_.toJson(legacyKeys = false)).mkString("[", ", ", "]")
          s"""{"pattern": ${Escape.json(
              q.pattern
            )}, "sheets": $sheets, "count": ${shown.size}, """ +
            s""""total": $total, "totalExact": $totalExact, "matches": $matches}"""
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

  /**
   * `stats <range>`: count, sum, min, max and mean of the numeric records. A range holding no
   * numbers is a legitimate result (GH-641) — zero-count statistics with `min`/`max`/`mean` absent
   * (`n/a` in text, `null` in JSON), exit 0 — flagged `NO_NUMERIC_VALUES` out of band. Whole-column
   * spans (`AM:AM`) are accepted like `view`'s ([[Resolve.ref]]), labelled as spelled and folded
   * over the used rows only ([[clampSpan]]).
   */
  private def stats(
    q: ReadQuery.Stats,
    source: SheetSource,
    sheetFlag: Option[String],
    mode: OutputMode,
    warn: Warning => IO[Unit]
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
      spelled = target.toEither.fold(ref => CellRange(ref, ref), identity)
      // the label is the span as asked (`B:B`); the scan covers its used rows only (GH-641)
      label = Resolve.rangeLabel(spelled)
      range <- clampSpan(source, sheet, spelled)
      acc <- source
        .rows(sheet, range)
        .flatMap(Stream.emits)
        .collect { record =>
          record.value match
            case CellValue.Number(n) => n
        }
        .compile
        .fold(StatsAcc(0, BigDecimal(0), None, None))(_.add(_))
      _ <- warn(Warning(WarningCode.NO_NUMERIC_VALUES, s"No numeric values in range $label"))
        .whenA(acc.count == 0)
    yield
      val count = acc.count
      val sum = acc.sum
      val mean = Option.when(count > 0)(sum / BigDecimal(count))
      mode match
        case OutputMode.Text =>
          def cell(v: Option[BigDecimal]): String = v.fold("n/a")(n => f"$n%.2f")
          Payload.text(
            f"count: $count, sum: $sum%.2f, min: ${cell(acc.min)}, max: ${cell(acc.max)}, " +
              s"mean: ${cell(mean)}"
          )
        case OutputMode.Json =>
          def lexeme(v: Option[BigDecimal]): String = v.fold("null")(CellRecord.numberLexeme)
          Payload.Raw(
            s"""{"sheet": ${Escape.json(sheet.value)}, "range": "$label", "count": $count, """ +
              s""""sum": ${CellRecord.numberLexeme(sum)}, "min": ${lexeme(acc.min)}, """ +
              s""""max": ${lexeme(acc.max)}, "mean": ${lexeme(mean)}}"""
          )

  // --- filter -----------------------------------------------------------------------------------

  /**
   * `filter --where`: the matching rows of the used range. `--limit` caps the rows shown, `0` means
   * no limit (GH-639, as for `view` and `search`); the match total always travels with the result —
   * in the markdown footer, as the JSON document's `matched`/`shown`/`truncated`/`limit` fields,
   * and for CSV (whose stdout must stay parseable) as a `TRUNCATED` warning when rows were clipped.
   */
  private def filter(
    q: ReadQuery.Filter,
    source: SheetSource,
    sheetFlag: Option[String],
    mode: OutputMode,
    warn: Warning => IO[Unit]
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
        case None => IO.pure(filterNoData(q, Vector.empty))
        case Some(range) => runFilter(q, source, sheet, range, pred, warn)
    yield if q.format == FilterFormat.Json then jsonPayload(mode, text) else Payload.text(text)

  /** The row cap `--limit` denotes: `None` for `0` (no limit). */
  private def filterCap(limit: Int): Option[Int] = Option.when(limit > 0)(limit)

  private def runFilter(
    q: ReadQuery.Filter,
    source: SheetSource,
    sheet: SheetName,
    range: CellRange,
    pred: FilterPredicate.Pred,
    warn: Warning => IO[Unit]
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
      // One pass over the used range: keep the first --limit matching rows, count them all. A row's
      // number is its position in the dense window, never recovered from the cells kept for it: a
      // --columns token outside the used range selects no cell, and the row is still a match.
      // `--limit 0` keeps every match (GH-639: no limit, as `view` and `search` read it)
      cap = filterCap(q.limit)
      scanned <- source
        .rows(sheet, range)
        .zipWithIndex
        .map { (row, i) => (range.start.row.index0 + i.toInt, row) }
        .filter { (rowIdx, row) =>
          !(q.header && rowIdx == headerRowIdx) &&
          FilterPredicate.evaluate(pred, resolve, col => row.lift(col - firstCol).map(_.cellValue))
        }
        .compile
        .fold((Vector.empty[FilterRow], 0)) { case ((kept, total), (rowIdx, row)) =>
          val keep = cap.forall(kept.size < _)
          val cells = selectedCols.map(col => row.lift(col - firstCol))
          (if keep then kept :+ FilterRow(rowIdx + 1, cells) else kept, total + 1)
        }
      (shown, total) = scanned
      labels = selectedCols.map { col =>
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
      truncated = total > shown.size
      // CSV stdout must stay machine-parseable: the clip is a warning, as `view --format csv` does;
      // markdown carries it in its footer and JSON in its own fields
      _ <- warn(
        Warning(
          WarningCode.TRUNCATED,
          RendererCommon.truncationNotice(shown.size, total, "matching rows")
        )
      ).whenA(truncated && q.format == FilterFormat.Csv)
      text <- q.format match
        case FilterFormat.Markdown => IO.pure(filterMarkdown(shown, total, labels))
        case FilterFormat.Csv => IO.pure(filterCsv(shown, labels))
        case FilterFormat.Json => filterJson(shown, total, cap, labels)
    yield text

  /**
   * One matching row as rendered: its 1-based number and the selected cells, `None` where the
   * column lies outside the used range (a `--columns` token past it, or left of it).
   */
  private final case class FilterRow(number: Int, cells: Vector[Option[CellRecord]])

  /**
   * Parse "A,C:E" into 0-based column indices; None = all used columns. A column may be named once:
   * a repeat would be a second column of the same label (the same JSON key), so it is refused.
   */
  private def parseColumns(
    spec: Option[String],
    usedCols: Vector[Int]
  ): Either[String, Vector[Int]] =
    spec match
      case None => Right(usedCols)
      case Some(s) =>
        s.split(",")
          .toVector
          .map(_.trim)
          .filter(_.nonEmpty)
          .flatTraverse { token =>
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
          .flatMap { cols =>
            val repeated = cols.diff(cols.distinct).distinct.map(Column.from0(_).toLetter)
            if repeated.isEmpty then Right(cols)
            else Left(s"Duplicate column(s) in --columns: ${repeated.mkString(", ")}")
          }

  private def filterNoData(q: ReadQuery.Filter, labels: Vector[String]): String =
    q.format match
      case FilterFormat.Markdown => "No rows matched."
      case FilterFormat.Csv => ("row" +: labels).mkString(",")
      case FilterFormat.Json => filterJsonDocument(Vector.empty, 0, filterCap(q.limit))

  private def filterMarkdown(
    shown: Vector[FilterRow],
    totalMatched: Int,
    labels: Vector[String]
  ): String =
    if totalMatched == 0 then "No rows matched."
    else
      val headers = "Row" +: labels
      val rows = shown.map { row =>
        row.number.toString +: row.cells.map(_.fold("")(_.text(showFormulas = false)))
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

  private def filterCsv(shown: Vector[FilterRow], labels: Vector[String]): String =
    val header = ("row" +: labels).map(Escape.csv).mkString(",")
    val lines = shown.map { row =>
      (row.number.toString +: row.cells.map(_.fold("")(_.text(showFormulas = false))))
        .map(Escape.csv)
        .mkString(",")
    }
    (header +: lines).mkString("\n")

  /**
   * `{"matched": N, "shown": M, "truncated": bool, "limit": n|null, "rows": [{"row": n, "cells":
   * {label: value}}]}` (GH-639: the match total and the clip travel with the rows, as `view`'s
   * `totalRows`/`truncated` and `search`'s `count`/`total` do), every number lexeme exact; an
   * uncached formula shows its expression as a string, as it always has; a column outside the used
   * range is `null`. Two selected columns under one header name share a key, which keeps the
   * first's position and the last's value — `ujson.Obj` semantics, as before. The text is
   * re-emitted through `ujson`: a lexeme this renderer got wrong is its defect (`INTERNAL`), never
   * malformed stdout.
   */
  private def filterJson(
    shown: Vector[FilterRow],
    matched: Int,
    cap: Option[Int],
    labels: Vector[String]
  ): IO[String] =
    def valueJson(record: CellRecord): String = record.formula match
      case Some(f) if !f.cached => Escape.json(f.text)
      case _ => record.rawJson
    val rows = shown.map { row =>
      val fields = uniqueKeys(row.cells.zip(labels).map { (cell, label) =>
        Escape.json(label) -> cell.fold("null")(valueJson)
      })
      s"""{"row": ${row.number}, "cells": {${fields.map((k, v) => s"$k: $v").mkString(", ")}}}"""
    }
    IO(ujson.reformat(filterJsonDocument(rows, matched, cap), indent = 2)).adaptError { case e =>
      CliException(
        CliError(ErrorCode.INTERNAL, s"filter produced malformed JSON: ${CliError.messageOf(e)}")
      )
    }

  /** The filter document around already-rendered row objects. */
  private def filterJsonDocument(rows: Vector[String], matched: Int, cap: Option[Int]): String =
    s"""{"matched": $matched, "shown": ${rows.size}, "truncated": ${matched > rows.size}, """ +
      s""""limit": ${cap.fold("null")(_.toString)}, "rows": ${rows.mkString("[", ", ", "]")}}"""

  /** A repeated key keeps its first position and takes its last value (`ujson.Obj` on update). */
  private def uniqueKeys(fields: Vector[(String, String)]): Vector[(String, String)] =
    fields.foldLeft(Vector.empty[(String, String)]) { (acc, field) =>
      acc.indexWhere(_._1 == field._1) match
        case -1 => acc :+ field
        case i => acc.updated(i, field)
    }
