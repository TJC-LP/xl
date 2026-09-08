package com.tjclp.xl.cli.read

import java.nio.file.Path

import cats.effect.IO
import cats.syntax.all.*
import fs2.Stream

import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cli.ViewFormat
import com.tjclp.xl.cli.contract.{CliError, CliException, ErrorCode, Warning}
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.ooxml.XlsxReader.ReaderConfig
import com.tjclp.xl.workbooks.Workbook

/**
 * How `view`'s styled formats are drawn: the html/svg text, or a raster export to `rasterOutput`.
 * Evaluation happens inside the rendering so cross-sheet formulas see the workbook; `strict` is
 * `view --eval --strict`'s gate for html and svg (raster formats have never gated).
 */
final case class RenderSpec(
  format: ViewFormat,
  evalFormulas: Boolean,
  strict: Boolean,
  printScale: Boolean,
  showGridlines: Boolean,
  showLabels: Boolean,
  dpi: Int,
  quality: Int,
  rasterOutput: Option[Path],
  rasterizer: Option[String]
) derives CanEqual

/**
 * Where a read verb's cells come from (W2.4): a loaded workbook ([[SheetSource.inMemory]]) or the
 * O(1)-memory streaming reader over a file ([[SheetSource.streaming]]). Every read verb is a
 * function of a [[ReadQuery]] and a source ([[Reads.run]]); the two strategies differ only in
 * [[capabilities]], and a query that needs more than the source has is refused in band as
 * `UNSUPPORTED_IN_STREAM` before anything is read.
 *
 * Positions are dense: [[rows]] yields one record per column of the window for every row of the
 * window, an empty record where the sheet has no cell, so the renderers never look at a `Sheet`.
 */
trait SheetSource:

  def capabilities: Set[Capability]

  /** The sheet names in workbook order. */
  def sheets: IO[Vector[SheetName]]

  /** The bounding box of the sheet's occupied cells; `None` for an empty sheet. */
  def usedRange(sheet: SheetName): IO[Option[CellRange]]

  /** The window's records, one dense row per row of `window`, top to bottom. */
  def rows(sheet: SheetName, window: CellRange): Stream[IO, Vector[CellRecord]]

  /** The 0-based hidden row and column indices inside `window` (empty without `hidden`). */
  def hiddenLines(sheet: SheetName, window: CellRange): IO[(Set[Int], Set[Int])]

  /** [[rows]] materialised with the hidden lines: what the table renderers consume. */
  def grid(sheet: SheetName, window: CellRange): IO[RecordGrid] =
    (rows(sheet, window).compile.toVector, hiddenLines(sheet, window)).mapN {
      case (rs, (hiddenRows, hiddenCols)) => RecordGrid(sheet, window, rs, hiddenRows, hiddenCols)
    }

  /**
   * [[grid]] after evaluating the window's formulas (`view --eval`; capability `eval`): an
   * evaluation failure is the `--strict` gate (`RECALC_GATE`) or an `EVAL_FAILED` warning through
   * `warn` with the cached values rendered instead.
   */
  def evaluated(
    sheet: SheetName,
    window: CellRange,
    strict: Boolean,
    warn: Warning => IO[Unit]
  ): IO[RecordGrid]

  /** Every occupied cell of the sheet, row-major (what `search` scans). */
  def occupied(sheet: SheetName): Stream[IO, CellRecord]

  /**
   * One cell with everything the sheet attaches to it. `withStyle = false` projects the record
   * without its style, so its display text is the General-format one (`cell --no-style`).
   */
  def detail(sheet: SheetName, ref: ARef, withStyle: Boolean): IO[CellDetail]

  /** A styled rendering of the window (capability `render`): html/svg text, or the export line. */
  def render(
    sheet: SheetName,
    window: CellRange,
    spec: RenderSpec,
    warn: Warning => IO[Unit]
  ): IO[String]

object SheetSource:

  /** Every capability: the loaded workbook answers everything. */
  def inMemory(wb: Workbook): SheetSource = InMemorySource(wb)

  /**
   * The streaming reader over `path`: [[Capability.streaming]], O(1) memory in the worksheet. Its
   * shared parts — `workbook.xml`, the shared-string table under `config`'s limits, `styles.xml` —
   * are read at most once per source (GH-640), which is why the source is built in `IO`.
   */
  def streaming(
    path: Path,
    excel: ExcelIO[IO],
    config: ReaderConfig = ReaderConfig.default
  ): IO[SheetSource] =
    StreamingSource(path, excel, config).widen

  /** A `--stream` refusal (exit 2), naming the in-memory alternative as the hint. */
  def unsupported(message: String, alternative: String): CliException =
    CliException(CliError(ErrorCode.UNSUPPORTED_IN_STREAM, message, hint = Some(alternative)))
