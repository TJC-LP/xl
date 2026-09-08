package com.tjclp.xl.cli.output

import com.tjclp.xl.addressing.{CellRange, Column}
import com.tjclp.xl.cli.read.{InMemorySource, RecordGrid}
import com.tjclp.xl.ooxml.metadata.SheetInfo
import com.tjclp.xl.sheets.Sheet

/**
 * Markdown table rendering for xl CLI output.
 *
 * All output includes row numbers and column letters for LLM consumption. Range tables render from
 * a [[RecordGrid]], the same rows whether they came from a loaded sheet or the streaming reader.
 */
object Markdown:

  /**
   * Render a range as a markdown table with row/column headers.
   *
   * Example output:
   * {{{
   * |   | A       | B       | C          |
   * |---|---------|---------|------------|
   * | 1 | Revenue |         | $1,000,000 |
   * | 2 | COGS    |         | $400,000   |
   * }}}
   *
   * @param skipEmpty
   *   If true, skip entirely empty rows and columns from output (reduces table size for sparse
   *   ranges)
   * @param skipHidden
   *   GH-474: if true, omit hidden rows/columns; default false renders them (an explicitly
   *   requested range never silently loses cells)
   */
  def renderRange(
    sheet: Sheet,
    range: CellRange,
    showFormulas: Boolean = false,
    skipEmpty: Boolean = false,
    evalFormulas: Boolean = false,
    skipHidden: Boolean = false
  ): String =
    val grid =
      if evalFormulas then InMemorySource.evaluatedGrid(sheet, range)
      else InMemorySource.grid(sheet, range)
    render(grid, showFormulas, skipEmpty, skipHidden)

  /** Render a grid as a markdown table (see [[renderRange]] for the flags). */
  def render(
    grid: RecordGrid,
    showFormulas: Boolean,
    skipEmpty: Boolean,
    skipHidden: Boolean
  ): String =
    val sb = new StringBuilder

    // GH-474: hidden rows/columns render unless --skip-hidden asked for the visible-only view
    val visibleCols = grid.renderedCols(skipHidden)
    val visibleRows = grid.renderedRows(skipHidden)

    // Filter empty columns/rows if skipEmpty is true
    val nonEmptyCols =
      if skipEmpty then grid.nonEmptyCols(visibleCols, visibleRows) else visibleCols
    val nonEmptyRows =
      if skipEmpty then grid.nonEmptyRows(visibleRows, nonEmptyCols) else visibleRows

    def text(rowIdx: Int, colIdx: Int): String =
      grid.at(rowIdx, colIdx).fold("")(_.text(showFormulas))

    // Calculate column widths for better formatting (only visible rows/cols)
    val colWidths = nonEmptyCols.map { col =>
      val header = Column.from0(col).toLetter
      val maxContent = nonEmptyRows.map(row => text(row, col).length).maxOption.getOrElse(0)
      math.max(header.length, math.max(maxContent, 3)) // Minimum width 3
    }

    // GH-641: the row-label column is as wide as the widest row number shown, so `| 10 |` lines
    // up under `| 9  |` instead of drifting once the window passes row 9
    val labelWidth = nonEmptyRows.map(row => (row + 1).toString.length).maxOption.getOrElse(1)

    // Header row with column letters
    sb.append(s"| ${" " * labelWidth} |")
    colWidths.zip(nonEmptyCols).foreach { case (width, col) =>
      val header = Column.from0(col).toLetter
      sb.append(s" ${header.padTo(width, ' ')} |")
    }
    sb.append("\n")

    // Separator row
    sb.append(s"|${"-" * (labelWidth + 2)}|")
    colWidths.foreach { (width: Int) =>
      sb.append("-" * (width + 2))
      sb.append("|")
    }
    sb.append("\n")

    // Data rows with row numbers (only visible rows)
    for row <- nonEmptyRows do
      val rowNum = (row + 1).toString
      sb.append(s"| ${rowNum.padTo(labelWidth, ' ')} |")
      colWidths.zip(nonEmptyCols).foreach { case (width, col) =>
        val escaped = Escape.markdown(text(row, col))
        sb.append(s" ${escaped.padTo(width, ' ')} |")
      }
      sb.append("\n")

    sb.toString

  /**
   * Render a list of sheets as a markdown table with proper column alignment.
   */
  def renderSheetList(sheets: Vector[(String, Option[CellRange], Int, Int)]): String =
    val headers = Vector("#", "Name", "Range", "Cells", "Formulas")
    val rows = sheets.zipWithIndex.map { case ((name, range, cells, formulas), idx) =>
      Vector(
        (idx + 1).toString,
        name,
        range.map(_.toA1).getOrElse("(empty)"),
        cells.toString,
        formulas.toString
      )
    }

    renderTable(headers, rows)

  /**
   * Render a list of sheets from lightweight metadata (quick mode).
   *
   * Shows name, dimension (from worksheet metadata), and visibility state. No cell counting.
   */
  def renderSheetListQuick(sheets: Vector[SheetInfo]): String =
    val headers = Vector("#", "Name", "Dimension", "State")
    val rows = sheets.zipWithIndex.map { case (info, idx) =>
      val stateStr = info.state match
        case Some("hidden") => "(hidden)"
        case Some("veryHidden") => "(very hidden)"
        case _ => ""
      Vector(
        (idx + 1).toString,
        info.name.value,
        info.dimension.map(_.toA1).getOrElse("(unknown)"),
        stateStr
      )
    }

    renderTable(headers, rows)

  /**
   * Render search results with qualified refs (`Sheet!A1`) as a two-column table.
   */
  def renderSearchResultsWithRef(results: Vector[(String, String)]): String =
    val headers = Vector("Ref", "Value")
    val rows = results.map { case (ref, value) =>
      Vector(ref, value)
    }
    renderTable(headers, rows)

  /**
   * Generic table renderer with proper column alignment.
   */
  def renderTable(headers: Vector[String], rows: Vector[Vector[String]]): String =
    val sb = new StringBuilder

    // Calculate column widths
    val colWidths = headers.indices.map { col =>
      val headerWidth = headers(col).length
      val maxDataWidth =
        rows.map(row => row.lift(col).map(_.length).getOrElse(0)).maxOption.getOrElse(0)
      math.max(headerWidth, math.max(maxDataWidth, 3))
    }

    // Header row
    sb.append("|")
    headers.zip(colWidths).foreach { case (header, width) =>
      sb.append(s" ${header.padTo(width, ' ')} |")
    }
    sb.append("\n")

    // Separator row
    sb.append("|")
    colWidths.foreach { width =>
      sb.append("-" * (width + 2))
      sb.append("|")
    }
    sb.append("\n")

    // Data rows
    rows.foreach { row =>
      sb.append("|")
      row.zip(colWidths).foreach { case (cell, width) =>
        sb.append(s" ${Escape.markdown(cell).padTo(width, ' ')} |")
      }
      sb.append("\n")
    }

    sb.toString
