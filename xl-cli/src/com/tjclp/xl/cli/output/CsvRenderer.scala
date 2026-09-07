package com.tjclp.xl.cli.output

import com.tjclp.xl.addressing.{CellRange, Column}
import com.tjclp.xl.cli.read.{InMemorySource, RecordGrid}
import com.tjclp.xl.sheets.Sheet

/**
 * CSV renderer for xl CLI output.
 *
 * Produces RFC 4180-compliant CSV output with optional row/column labels, from a [[RecordGrid]].
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
    // GH-474: hidden rows/columns render unless --skip-hidden asked for the visible-only view
    val visibleCols = grid.renderedCols(skipHidden)
    val visibleRows = grid.renderedRows(skipHidden)

    // Filter empty columns/rows if skipEmpty is true
    val nonEmptyCols =
      if skipEmpty then grid.nonEmptyCols(visibleCols, visibleRows) else visibleCols
    val nonEmptyRows =
      if skipEmpty then grid.nonEmptyRows(visibleRows, nonEmptyCols) else visibleRows

    val sb = new StringBuilder

    // Header row with column letters (if showLabels)
    if showLabels then
      sb.append(",") // Empty cell for row number column
      sb.append(nonEmptyCols.map(colIdx => Column.from0(colIdx).toLetter).mkString(","))
      sb.append("\n")

    // Data rows
    val lastRowIdx = nonEmptyRows.lastOption
    nonEmptyRows.foreach { rowIdx =>
      if showLabels then
        sb.append((rowIdx + 1).toString)
        sb.append(",")

      val cellValues = nonEmptyCols.map { colIdx =>
        grid.at(rowIdx, colIdx).fold("")(record => Escape.csv(record.text(showFormulas)))
      }
      sb.append(cellValues.mkString(","))

      // Add newline (except after last row)
      if !lastRowIdx.contains(rowIdx) then sb.append("\n")
    }

    sb.toString
