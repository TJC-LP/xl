package com.tjclp.xl.cli.read

import fs2.Stream

import com.tjclp.xl.addressing.{CellRange, Column, SheetName}
import com.tjclp.xl.cells.Comment

/**
 * The window a table renderer draws, without its rows (GH-635): the range and the 0-based indices
 * of the hidden lines inside it (empty from a source without the `hidden` capability). Everything a
 * renderer decides before the first row arrives — which columns are drawn, which rows, the
 * hidden-line notice — is a function of this alone; the rows then stream past it one at a time.
 */
final case class RecordWindow(
  sheet: SheetName,
  range: CellRange,
  hiddenRows: Set[Int],
  hiddenCols: Set[Int]
) derives CanEqual:

  /** 0-based row indices of the window, top to bottom. */
  def rowIndices: Vector[Int] = (range.start.row.index0 to range.end.row.index0).toVector

  /** 0-based column indices of the window, left to right. */
  def colIndices: Vector[Int] = (range.start.col.index0 to range.end.col.index0).toVector

  /** The rows a renderer shows: every row of the window, or the visible ones under `skipHidden`. */
  def renderedRows(skipHidden: Boolean): Vector[Int] =
    if skipHidden then rowIndices.filterNot(hiddenRows.contains) else rowIndices

  /** Whether the 0-based `row` is one a renderer shows (see [[renderedRows]]). */
  def isRendered(row: Int, skipHidden: Boolean): Boolean =
    !(skipHidden && hiddenRows.contains(row))

  /** The columns a renderer shows (see [[renderedRows]]). */
  def renderedCols(skipHidden: Boolean): Vector[Int] =
    if skipHidden then colIndices.filterNot(hiddenCols.contains) else colIndices

  /** 1-based numbers of the hidden rows, ascending — what the JSON `hiddenRows` field lists. */
  def hiddenRowNumbers: Vector[Int] = rowIndices.filter(hiddenRows.contains).map(_ + 1)

  /** Letters of the hidden columns, left to right — what the JSON `hiddenCols` field lists. */
  def hiddenColLetters: Vector[String] =
    colIndices.filter(hiddenCols.contains).map(c => Column.from0(c).toLetter)

  /** The record of a dense row at absolute 0-based `col`, when inside the window. */
  def cell(row: Vector[CellRecord], col: Int): Option[CellRecord] =
    row.lift(col - range.start.col.index0)

  /**
   * Whether a dense row is drawn: shown under `skipHidden`, and — under `skipEmpty` — holding a
   * non-empty record in one of `cols`. (A row with such a record makes its column non-empty, so
   * this is exactly the row set the whole-grid renderers kept: the rows with a record in a
   * non-empty column.)
   */
  def isDrawn(
    row: Vector[CellRecord],
    rowIdx: Int,
    cols: Vector[Int],
    skipEmpty: Boolean,
    skipHidden: Boolean
  ): Boolean =
    isRendered(rowIdx, skipHidden) &&
      (!skipEmpty || cols.exists(col => cell(row, col).exists(r => !r.isEmpty)))

/**
 * What a table renderer must know about every row before it writes the first (GH-635): per column
 * of the window, the widest display text among the drawn rows (markdown's column widths) and
 * whether any drawn row holds a non-empty record there (`--skip-empty`'s column pruning). Folded
 * over the rows in O(columns) memory, so the pass reads the window without holding it; a renderer
 * that needs neither skips the pass.
 */
final case class ColumnFacts(widths: Map[Int, Int], nonEmpty: Set[Int]) derives CanEqual:

  /** The widest drawn text in `col` (0 when nothing was drawn there). */
  def width(col: Int): Int = widths.getOrElse(col, 0)

  /** `cols` minus the ones no drawn row fills. */
  def nonEmptyAmong(cols: Vector[Int]): Vector[Int] = cols.filter(nonEmpty.contains)

object ColumnFacts:

  val empty: ColumnFacts = ColumnFacts(Map.empty, Set.empty)

  /**
   * The facts of a window's dense rows, as a one-element stream (a fold, so it runs on the same
   * effect as the rows). `text` is the display text a renderer measures — the markdown cell text —
   * and only the drawn rows ([[RecordWindow.isDrawn]] over the shown columns) count.
   */
  def of[F[_]](
    window: RecordWindow,
    rows: Stream[F, Vector[CellRecord]],
    text: CellRecord => String,
    skipEmpty: Boolean,
    skipHidden: Boolean
  ): Stream[F, ColumnFacts] =
    val cols = window.renderedCols(skipHidden)
    val firstRow = window.range.start.row.index0
    rows.zipWithIndex.fold(empty) { case (facts, (row, i)) =>
      val rowIdx = firstRow + i.toInt
      if !window.isDrawn(row, rowIdx, cols, skipEmpty, skipHidden) then facts
      else
        cols.foldLeft(facts) { (acc, col) =>
          window.cell(row, col) match
            case None => acc
            case Some(record) =>
              val width = math.max(acc.width(col), text(record).length)
              val nonEmpty = if record.isEmpty then acc.nonEmpty else acc.nonEmpty + col
              ColumnFacts(acc.widths.updated(col, width), nonEmpty)
        }
    }

/**
 * A rectangular window of records: one dense row of [[CellRecord]]s per row of `range`, one record
 * per column of `range` in each — what the whole-grid renderer adapters consume (W2.4).
 * `hiddenRows` and `hiddenCols` are the 0-based indices of the hidden lines inside the window
 * (empty from a source without the `hidden` capability). The renderers themselves stream
 * ([[RecordWindow]] plus a row stream, GH-635); a grid is that stream held in a `Vector`.
 */
final case class RecordGrid(
  sheet: SheetName,
  range: CellRange,
  rows: Vector[Vector[CellRecord]],
  hiddenRows: Set[Int],
  hiddenCols: Set[Int]
) derives CanEqual:

  /** The window without its rows. */
  def window: RecordWindow = RecordWindow(sheet, range, hiddenRows, hiddenCols)

  /** 0-based row indices of the window, top to bottom. */
  def rowIndices: Vector[Int] = window.rowIndices

  /** 0-based column indices of the window, left to right. */
  def colIndices: Vector[Int] = window.colIndices

  /** The record at absolute 0-based `(row, col)`, when inside the window. */
  def at(row: Int, col: Int): Option[CellRecord] =
    rows
      .lift(row - range.start.row.index0)
      .flatMap(_.lift(col - range.start.col.index0))

  /** The rows a renderer shows: every row of the window, or the visible ones under `skipHidden`. */
  def renderedRows(skipHidden: Boolean): Vector[Int] = window.renderedRows(skipHidden)

  /** The columns a renderer shows (see [[renderedRows]]). */
  def renderedCols(skipHidden: Boolean): Vector[Int] = window.renderedCols(skipHidden)

  /** Whether the position holds an empty record ([[CellRecord.isEmpty]]) or lies off the grid. */
  def isEmptyAt(row: Int, col: Int): Boolean = at(row, col).forall(_.isEmpty)

  /** The given columns that hold at least one non-empty record among `rows`. */
  def nonEmptyCols(cols: Vector[Int], rows: Vector[Int]): Vector[Int] =
    cols.filter(col => rows.exists(row => !isEmptyAt(row, col)))

  /** The given rows that hold at least one non-empty record among `cols`. */
  def nonEmptyRows(rows: Vector[Int], cols: Vector[Int]): Vector[Int] =
    rows.filter(row => cols.exists(col => !isEmptyAt(row, col)))

  /** 1-based numbers of the hidden rows, ascending — what the JSON `hiddenRows` field lists. */
  def hiddenRowNumbers: Vector[Int] = window.hiddenRowNumbers

  /** Letters of the hidden columns, left to right — what the JSON `hiddenCols` field lists. */
  def hiddenColLetters: Vector[String] = window.hiddenColLetters

/**
 * Everything `cell` reports about one cell: the record plus what the sheet attaches to it.
 *
 * @param hyperlink
 *   the cell's hyperlink — `None` when it has none, and also when the source lacks `hyperlinks`
 * @param dependencies
 *   the cells the formula reads, as `A1` / `Sheet!A1` text (ranges as their occupied cells) —
 *   `None` when the source cannot know (no `graph` capability)
 * @param dependents
 *   the formulas that read the cell — `None` when the source cannot know (no `graph` capability)
 */
final case class CellDetail(
  record: CellRecord,
  comment: Option[Comment],
  hyperlink: Option[String],
  dependencies: Option[Vector[String]],
  dependents: Option[Vector[String]]
) derives CanEqual
