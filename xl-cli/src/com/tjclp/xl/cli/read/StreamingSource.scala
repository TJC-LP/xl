package com.tjclp.xl.cli.read

import java.nio.file.Path

import cats.effect.IO
import cats.syntax.all.*
import fs2.{Chunk, Pull, Stream}
import org.xml.sax.SAXParseException

import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cli.contract.{CliError, CliException, ErrorCode, Location, Warning}
import com.tjclp.xl.error.{XLError, XLException}
import com.tjclp.xl.io.{ExcelIO, RowData}
import com.tjclp.xl.ooxml.SharedStrings
import com.tjclp.xl.ooxml.XlsxReader.ReaderConfig
import com.tjclp.xl.ooxml.metadata.LightMetadata
import com.tjclp.xl.ooxml.style.WorkbookStyles

/**
 * The `--stream` strategy: the SAX row reader over one worksheet at a time, `styles.xml` loaded
 * once for number formats, `workbook.xml` for the sheet list and dimensions, and the shared-string
 * table parsed ONCE for every read of the run (GH-640: the table is the one part a streaming read
 * holds in memory; `cell` used to parse it again through a DOM, three times the size). The three
 * are memoised effects: the first read that needs one pays for it, the next reads — a table's
 * second pass, the header row of `--header-row`, `cell` after `view` — share it. O(1) memory in the
 * worksheet; [[Capability.streaming]] only — the reader never parses row/column properties, merges,
 * hyperlinks or other sheets' formulas, so `eval` and `render` are refused in band and the fields
 * of `hidden`, `merges`, `hyperlinks` and `graph` are reported unknown (`null`), never guessed.
 *
 * The table is held to the reader's ZIP-bomb limits (`--max-size`, the same `ReaderConfig` the
 * in-memory load honours): a table past them is `SECURITY_ERROR` (exit 3) with the flag to raise as
 * the hint — under `--stream` the flag bounds this one part.
 */
final class StreamingSource private (
  path: Path,
  excel: ExcelIO[IO],
  metadata: IO[LightMetadata],
  sharedStrings: IO[Option[SharedStrings]],
  styles: IO[WorkbookStyles]
) extends SheetSource:

  val capabilities: Set[Capability] = Capability.streaming

  def sheets: IO[Vector[SheetName]] = metadata.map(_.sheets.map(_.name))

  /**
   * The worksheet's `<dimension>` as written, else the bounding box of its non-empty cells. A
   * single-cell `<dimension>` is not trusted: Excel writes `<dimension ref="A1"/>` for an EMPTY
   * sheet, which the loaded workbook (no stored cell) reports as no used range, so a one-cell
   * extent is re-derived by the scan — a pass over at most one cell when the declaration is honest,
   * and the right answer when it is not.
   */
  def usedRange(sheet: SheetName): IO[Option[CellRange]] =
    metadata.flatMap { meta =>
      meta.sheets.find(_.name == sheet).flatMap(_.dimension) match
        case Some(range) if range.width > 1 || range.height > 1 => IO.pure(Some(range))
        case _ => scanBounds(sheet)
    }

  def rows(sheet: SheetName, window: CellRange): Stream[IO, Vector[CellRecord]] =
    Stream
      .eval((sharedStrings, styles).tupled)
      .flatMap { (table, st) =>
        StreamingSource.dense(
          excel.readSheetStreamRange(path, sheet.value, window, table),
          sheet,
          window,
          st
        )
      }
      .handleErrorWith(e => Stream.raiseError[IO](located(sheet)(e)))

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
    Stream
      .eval((sharedStrings, styles).tupled)
      .flatMap { (table, st) =>
        excel.readSheetStream(path, sheet.value, table).flatMap { row =>
          Stream.emits(StreamingSource.records(row, sheet, st))
        }
      }
      .handleErrorWith(e => Stream.raiseError[IO](located(sheet)(e)))

  def detail(sheet: SheetName, ref: ARef, withStyle: Boolean): IO[CellDetail] =
    (sharedStrings, styles).tupled.flatMap { (table, st) =>
      excel.streamCellDetails(path, sheet.value, ref, table, st).adaptError(located(sheet)).map {
        details =>
          val style = if withStyle then details.style else None
          CellDetail(
            CellRecord.of(sheet, ref, details.value, style, hidden = None, mergedInto = None),
            details.comment,
            hyperlink = None,
            // No graph without the workbook: the reader's token list is not the precedent set (a
            // range token plus its endpoints is neither the cells it covers nor exact), so say so
            // instead
            dependencies = None,
            dependents = None
          )
      }
    }

  def render(
    sheet: SheetName,
    window: CellRange,
    spec: RenderSpec,
    warn: Warning => IO[Unit]
  ): IO[String] =
    IO.raiseError(StreamingSource.renderUnsupported(spec))

  /**
   * A worksheet the parser rejects, from any read of `sheet`: `IO_READ` with the parser's message
   * and position, the file and the sheet in `location` — what the loaded reader reports for the
   * same part (GH-635). Every other failure passes as raised.
   */
  private def located(sheet: SheetName): PartialFunction[Throwable, Throwable] =
    case sax: SAXParseException =>
      CliException(
        CliError(
          ErrorCode.IO_READ,
          s"Parse error at worksheet '${sheet.value}': ${CliError.messageOf(sax)} " +
            s"(line ${sax.getLineNumber}, column ${sax.getColumnNumber})",
          location = Some(Location(Some(path.toString), Some(sheet.value), None, None))
        )
      )
    case other => other

  /** The bounding box of every non-empty streamed cell; `None` when the sheet has none. */
  private def scanBounds(sheet: SheetName): IO[Option[CellRange]] =
    sharedStrings.flatMap { table =>
      excel
        .readSheetStream(path, sheet.value, table)
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
        .adaptError(located(sheet))
    }

object StreamingSource:

  /**
   * The source over `path`, its three shared parts memoised: `workbook.xml` (the input read of
   * every streaming verb, so a missing or unreadable file is `IO_READ` here as on every other
   * verb), the shared-string table under `config`'s limits, and `styles.xml`. Each is read on first
   * use, at most once per source.
   */
  def apply(path: Path, excel: ExcelIO[IO], config: ReaderConfig): IO[StreamingSource] =
    (
      metadata(path, excel).memoize,
      sharedStrings(path, excel, config).memoize,
      excel.loadStyles(path).memoize
    ).mapN(new StreamingSource(path, excel, _, _, _))

  /** The refusal every styled format gets under `--stream`: today's exact text. */
  def renderUnsupported(spec: RenderSpec): CliException =
    SheetSource.unsupported(
      s"--stream not supported for ${spec.format.toString.toLowerCase} (needs styles). " +
        "Remove --stream flag or use markdown/csv/json format.",
      "omit --stream, or use --format markdown, csv or json"
    )

  /** The hint of a shared-string table past `--max-size` under `--stream`. */
  val tableLimitHint: String =
    "raise --max-size <MB> (0 = unlimited): under --stream it bounds the shared-string table, " +
      "the one part a streaming read holds in memory"

  /** `workbook.xml` under the `IO_READ` classification (the metadata reader raises plain text). */
  private def metadata(path: Path, excel: ExcelIO[IO]): IO[LightMetadata] =
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

  /**
   * The shared-string table under `config`'s ZIP-bomb limits (GH-640): a breach is the reader's own
   * `SecurityError`, coded `SECURITY_ERROR`, with [[tableLimitHint]] — the library's hint names
   * `--stream` as the way out, which this already is.
   */
  private def sharedStrings(
    path: Path,
    excel: ExcelIO[IO],
    config: ReaderConfig
  ): IO[Option[SharedStrings]] =
    excel.loadSharedStrings(path, config).adaptError { case x: XLException =>
      x.error match
        case security: XLError.SecurityError =>
          CliException(
            CliError(
              security.code,
              security.message,
              hint = Some(tableLimitHint),
              location = Some(Location.file(path.toString)),
              cause = Some(security)
            )
          )
        case _ => x
    }

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
  private[read] def dense(
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
          // The reader is row-ordered: a row past the window ends the pull (its tail is never read);
          // one not after the last emitted row is a duplicate and is skipped
          if idx > lastRow then gap(next, lastRow + 1)
          else if idx < next then go(next, tail)
          else gap(next, idx) >> Pull.output1(denseRow(row, idx)) >> go(idx + 1, tail)
        case None => gap(next, lastRow + 1)
      }
    go(firstRow, streamed).stream
