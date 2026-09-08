package com.tjclp.xl.cli.read

import java.nio.file.Path

import cats.effect.IO
import cats.syntax.all.*
import fs2.{Chunk, Pull, Stream}

import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cli.contract.{CliError, CliException, ErrorCode, Location, Warning}
import com.tjclp.xl.io.{ExcelIO, RowData}
import com.tjclp.xl.ooxml.metadata.LightMetadata
import com.tjclp.xl.ooxml.style.WorkbookStyles

/**
 * The `--stream` strategy: the SAX row reader over one worksheet at a time, `styles.xml` loaded
 * once for number formats, `workbook.xml` for the sheet list and dimensions. O(1) memory in the
 * worksheet; [[Capability.streaming]] only — the reader never parses row/column properties, merges,
 * hyperlinks or other sheets' formulas, so `eval` and `render` are refused in band and the fields
 * of `hidden`, `merges`, `hyperlinks` and `graph` are reported unknown (`null`), never guessed.
 */
final class StreamingSource(path: Path, excel: ExcelIO[IO]) extends SheetSource:

  val capabilities: Set[Capability] = Capability.streaming

  def sheets: IO[Vector[SheetName]] = metadata.map(_.sheets.map(_.name))

  /** The worksheet's `<dimension>` when it has one, else a scan of its occupied cells. */
  def usedRange(sheet: SheetName): IO[Option[CellRange]] =
    metadata.flatMap { meta =>
      meta.sheets.find(_.name == sheet).flatMap(_.dimension) match
        case Some(range) => IO.pure(Some(range))
        case None => scanBounds(sheet)
    }

  def rows(sheet: SheetName, window: CellRange): Stream[IO, Vector[CellRecord]] =
    Stream.eval(excel.loadStyles(path)).flatMap { styles =>
      StreamingSource.dense(
        excel.readSheetStreamRange(path, sheet.value, window),
        sheet,
        window,
        styles
      )
    }

  def hiddenLines(sheet: SheetName, window: CellRange): IO[(Set[Int], Set[Int])] =
    IO.pure((Set.empty, Set.empty))

  def evaluated(
    sheet: SheetName,
    window: CellRange,
    strict: Boolean,
    warn: Warning => IO[Unit]
  ): IO[RecordGrid] =
    IO.raiseError(
      SheetSource.unsupported(
        "--eval is not supported with --stream (streaming view uses cached values only)",
        "omit --stream to evaluate formulas; use --max-size <MB> for a large file"
      )
    )

  def occupied(sheet: SheetName): Stream[IO, CellRecord] =
    Stream.eval(excel.loadStyles(path)).flatMap { styles =>
      excel.readSheetStream(path, sheet.value).flatMap { row =>
        Stream.emits(StreamingSource.records(row, sheet, styles))
      }
    }

  def detail(sheet: SheetName, ref: ARef, withStyle: Boolean): IO[CellDetail] =
    excel.streamCellDetails(path, sheet.value, ref).map { details =>
      val style = if withStyle then details.style else None
      CellDetail(
        CellRecord.of(sheet, ref, details.value, style, hidden = None, mergedInto = None),
        details.comment,
        hyperlink = None,
        // No graph without the workbook: the reader's token list is not the precedent set (a range
        // token plus its endpoints is neither the cells it covers nor exact), so say so instead
        dependencies = None,
        dependents = None
      )
    }

  def render(
    sheet: SheetName,
    window: CellRange,
    spec: RenderSpec,
    warn: Warning => IO[Unit]
  ): IO[String] =
    IO.raiseError(StreamingSource.renderUnsupported(spec))

  /**
   * `workbook.xml`: the input read of every streaming verb, so a missing or unreadable file is
   * `IO_READ` here as on every other verb (the metadata reader itself raises a plain exception).
   */
  private def metadata: IO[LightMetadata] =
    excel.readMetadata(path).adaptError {
      case cli: CliException => cli
      case other =>
        CliException(
          CliError(
            ErrorCode.IO_READ,
            CliError.messageOf(other),
            location = Some(Location.file(path.toString))
          )
        )
    }

  /** The bounding box of every non-empty streamed cell; `None` when the sheet has none. */
  private def scanBounds(sheet: SheetName): IO[Option[CellRange]] =
    excel
      .readSheetStream(path, sheet.value)
      .compile
      .fold(Option.empty[(Int, Int, Int, Int)]) { (acc, row) =>
        if row.cells.isEmpty then acc
        else
          val r = row.rowIndex - 1 // rowIndex is 1-based
          val cMin = row.cells.keys.foldLeft(Int.MaxValue)(_ min _)
          val cMax = row.cells.keys.foldLeft(Int.MinValue)(_ max _)
          Some(acc.fold((r, r, cMin, cMax)) { case (r1, r2, c1, c2) =>
            (r1 min r, r2 max r, c1 min cMin, c2 max cMax)
          })
      }
      .map(_.map { case (r1, r2, c1, c2) => CellRange(ARef.from0(c1, r1), ARef.from0(c2, r2)) })

object StreamingSource:

  /** The refusal every styled format gets under `--stream`: today's exact text. */
  def renderUnsupported(spec: RenderSpec): CliException =
    SheetSource.unsupported(
      s"--stream not supported for ${spec.format.toString.toLowerCase} (needs styles). " +
        "Remove --stream flag or use markdown/csv/json format.",
      "omit --stream, or use --format markdown, csv or json"
    )

  /** One streamed row's occupied cells as records, left to right. */
  private def records(row: RowData, sheet: SheetName, styles: WorkbookStyles): Vector[CellRecord] =
    row.cells.toVector.sortBy(_._1).map { (col, value) =>
      CellRecord.of(
        sheet,
        ARef.from0(col, row.rowIndex - 1),
        value,
        row.cellStyles.get(col).flatMap(styles.styleAt),
        hidden = None,
        mergedInto = None
      )
    }

  /**
   * The streamed rows of a window as dense records: every row of the window in order, the rows the
   * reader never emitted (no cells) filled with empty records, every column of the window present.
   */
  private def dense(
    streamed: Stream[IO, RowData],
    sheet: SheetName,
    window: CellRange,
    styles: WorkbookStyles
  ): Stream[IO, Vector[CellRecord]] =
    val firstRow = window.start.row.index0
    val lastRow = window.end.row.index0
    val cols = (window.start.col.index0 to window.end.col.index0).toVector
    def emptyRow(row: Int): Vector[CellRecord] =
      cols.map(col => CellRecord.empty(sheet, ARef.from0(col, row), hidden = None, None))
    def denseRow(row: RowData, rowIdx: Int): Vector[CellRecord] =
      cols.map { col =>
        val ref = ARef.from0(col, rowIdx)
        row.cells.get(col) match
          case Some(value) =>
            CellRecord.of(
              sheet,
              ref,
              value,
              row.cellStyles.get(col).flatMap(styles.styleAt),
              hidden = None,
              mergedInto = None
            )
          case None => CellRecord.empty(sheet, ref, hidden = None, None)
      }
    def gap(from: Int, until: Int): Pull[IO, Vector[CellRecord], Unit] =
      if from >= until then Pull.done
      else Pull.output(Chunk.from((from until until).map(emptyRow)))
    def go(next: Int, s: Stream[IO, RowData]): Pull[IO, Vector[CellRecord], Unit] =
      s.pull.uncons1.flatMap {
        case Some((row, tail)) =>
          val idx = row.rowIndex - 1
          // Outside the window, or not after the last emitted row: the reader is row-ordered, so
          // such a row is a duplicate and is skipped
          if idx < next || idx > lastRow then go(next, tail)
          else gap(next, idx) >> Pull.output1(denseRow(row, idx)) >> go(idx + 1, tail)
        case None => gap(next, lastRow + 1)
      }
    go(firstRow, streamed).stream
