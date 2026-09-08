package com.tjclp.xl.cli.output

import fs2.Stream

import com.tjclp.xl.addressing.{CellRange, Column}
import com.tjclp.xl.cli.read.{CellRecord, ColumnFacts, InMemorySource, RecordGrid, RecordWindow}
import com.tjclp.xl.sheets.Sheet

/**
 * CSV renderer for xl CLI output.
 *
 * Produces RFC 4180-compliant CSV output with optional row/column labels. The rows stream past a
 * [[RecordWindow]] one at a time (GH-635): [[header]] is the line before them, [[lines]] one line
 * per drawn row, and nothing follows — so a streamed `view` never holds its window. Only
 * `--skip-empty` needs to see every row before the first is written (the empty columns are pruned
 * across the whole window); that is one [[ColumnFacts]] pass. [[render]] is the same lines over a
 * grid held in memory, joined.
 */
object CsvRenderer:

  /**
   * Render a range as CSV.
   *
   * @param sheet
   *   Sheet to render from
   * @param range
   *   Cell range to render
   * @param showFormulas
   *   If true, show formulas instead of computed values
   * @param showLabels
   *   If true, include column letters as header row and row numbers as first column
   * @param skipEmpty
   *   If true, skip entirely empty rows and columns from output
   * @param skipHidden
   *   GH-474: if true, omit hidden rows/columns; default false renders them
   * @return
   *   CSV string
   */
  def renderRange(
    sheet: Sheet,
    range: CellRange,
    showFormulas: Boolean = false,
    showLabels: Boolean = false,
    skipEmpty: Boolean = false,
    evalFormulas: Boolean = false,
    skipHidden: Boolean = false
  ): String =
    val grid =
      if evalFormulas then InMemorySource.evaluatedGrid(sheet, range)
      else InMemorySource.grid(sheet, range)
    render(grid, showFormulas, showLabels, skipEmpty, skipHidden)

  /** Render a grid as CSV (see [[renderRange]] for the flags). No trailing newline. */
  def render(
    grid: RecordGrid,
    showFormulas: Boolean,
    showLabels: Boolean,
    skipEmpty: Boolean,
    skipHidden: Boolean
  ): String =
    val window = grid.window
    val rows = Stream.emits(grid.rows)
    val facts =
      if needsFacts(skipEmpty) then
        ColumnFacts
          .of(window, rows, _.text(showFormulas), skipEmpty, skipHidden)
          .toList
          .headOption
          .getOrElse(ColumnFacts.empty)
      else ColumnFacts.empty
    val cols = columns(window, facts, skipEmpty, skipHidden)
    val body = lines(window, rows, cols, showFormulas, showLabels, skipEmpty, skipHidden).toList
    (header(cols, showLabels) ++ body).mkString("\n")

  /**
   * Whether the renderer must see every row before it writes the first: only under `--skip-empty`,
   * whose column pruning is a fact about the whole window.
   */
  def needsFacts(skipEmpty: Boolean): Boolean = skipEmpty

  /**
   * The columns drawn: the shown ones ([[RecordWindow.renderedCols]]), minus — under `skipEmpty` —
   * those no drawn row fills.
   */
  def columns(
    window: RecordWindow,
    facts: ColumnFacts,
    skipEmpty: Boolean,
    skipHidden: Boolean
  ): Vector[Int] =
    val shown = window.renderedCols(skipHidden)
    if skipEmpty then facts.nonEmptyAmong(shown) else shown

  /** The lines before the rows: the column-letter header row under `showLabels`, else none. */
  def header(cols: Vector[Int], showLabels: Boolean): Vector[String] =
    if showLabels then Vector("," + cols.map(colIdx => Column.from0(colIdx).toLetter).mkString(","))
    else Vector.empty

  /**
   * One CSV line per drawn row of the window ([[RecordWindow.isDrawn]] over `cols`), in order: the
   * row number first under `showLabels`, then the escaped display text of every drawn column.
   */
  def lines[F[_]](
    window: RecordWindow,
    rows: Stream[F, Vector[CellRecord]],
    cols: Vector[Int],
    showFormulas: Boolean,
    showLabels: Boolean,
    skipEmpty: Boolean,
    skipHidden: Boolean
  ): Stream[F, String] =
    val firstRow = window.range.start.row.index0
    rows.zipWithIndex.map { (row, i) =>
      val rowIdx = firstRow + i.toInt
      Option.when(window.isDrawn(row, rowIdx, cols, skipEmpty, skipHidden)) {
        val label = if showLabels then s"${rowIdx + 1}," else ""
        val cells = cols.map { colIdx =>
          window.cell(row, colIdx).fold("")(record => Escape.csv(record.text(showFormulas)))
        }
        label + cells.mkString(",")
      }
    }.unNone
