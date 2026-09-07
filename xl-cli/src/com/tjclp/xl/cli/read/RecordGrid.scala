package com.tjclp.xl.cli.read

import com.tjclp.xl.addressing.{CellRange, Column, SheetName}
import com.tjclp.xl.cells.Comment

/**
 * A rectangular window of records: one dense row of [[CellRecord]]s per row of `range`, one record
 * per column of `range` in each — what every table renderer consumes (W2.4). `hiddenRows` and
 * `hiddenCols` are the 0-based indices of the hidden lines inside the window (empty from a source
 * without the `hidden` capability).
 */
final case class RecordGrid(
  sheet: SheetName,
  range: CellRange,
  rows: Vector[Vector[CellRecord]],
  hiddenRows: Set[Int],
  hiddenCols: Set[Int]
) derives CanEqual:

  /** 0-based row indices of the window, top to bottom. */
  def rowIndices: Vector[Int] = (range.start.row.index0 to range.end.row.index0).toVector

  /** 0-based column indices of the window, left to right. */
  def colIndices: Vector[Int] = (range.start.col.index0 to range.end.col.index0).toVector

  /** The record at absolute 0-based `(row, col)`, when inside the window. */
  def at(row: Int, col: Int): Option[CellRecord] =
    rows
      .lift(row - range.start.row.index0)
      .flatMap(_.lift(col - range.start.col.index0))

  /** The rows a renderer shows: every row of the window, or the visible ones under `skipHidden`. */
  def renderedRows(skipHidden: Boolean): Vector[Int] =
    if skipHidden then rowIndices.filterNot(hiddenRows.contains) else rowIndices

  /** The columns a renderer shows (see [[renderedRows]]). */
  def renderedCols(skipHidden: Boolean): Vector[Int] =
    if skipHidden then colIndices.filterNot(hiddenCols.contains) else colIndices

  /** Whether the position holds an empty record ([[CellRecord.isEmpty]]) or lies off the grid. */
  def isEmptyAt(row: Int, col: Int): Boolean = at(row, col).forall(_.isEmpty)

  /** The given columns that hold at least one non-empty record among `rows`. */
  def nonEmptyCols(cols: Vector[Int], rows: Vector[Int]): Vector[Int] =
    cols.filter(col => rows.exists(row => !isEmptyAt(row, col)))

  /** The given rows that hold at least one non-empty record among `cols`. */
  def nonEmptyRows(rows: Vector[Int], cols: Vector[Int]): Vector[Int] =
    rows.filter(row => cols.exists(col => !isEmptyAt(row, col)))

  /** 1-based numbers of the hidden rows, ascending — what the JSON `hiddenRows` field lists. */
  def hiddenRowNumbers: Vector[Int] = rowIndices.filter(hiddenRows.contains).map(_ + 1)

  /** Letters of the hidden columns, left to right — what the JSON `hiddenCols` field lists. */
  def hiddenColLetters: Vector[String] =
    colIndices.filter(hiddenCols.contains).map(c => Column.from0(c).toLetter)

/**
 * Everything `cell` reports about one cell: the record plus what the sheet attaches to it.
 *
 * @param dependencies
 *   the cells the formula reads, as `A1` / `Sheet!A1` text
 * @param dependents
 *   the formulas that read the cell — `None` when the source cannot know (no `graph` capability)
 */
final case class CellDetail(
  record: CellRecord,
  comment: Option[Comment],
  hyperlink: Option[String],
  dependencies: Vector[String],
  dependents: Option[Vector[String]]
) derives CanEqual
