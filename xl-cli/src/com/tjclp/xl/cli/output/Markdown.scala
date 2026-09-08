package com.tjclp.xl.cli.output

import fs2.Stream

import com.tjclp.xl.addressing.{CellRange, Column}
import com.tjclp.xl.cli.read.{CellRecord, ColumnFacts, InMemorySource, RecordGrid, RecordWindow}
import com.tjclp.xl.ooxml.metadata.SheetInfo
import com.tjclp.xl.sheets.Sheet

/**
 * Markdown table rendering for xl CLI output.
 *
 * All output includes row numbers and column letters for LLM consumption. Range tables render the
 * same rows whether they came from a loaded sheet or the streaming reader, and they stream
 * (GH-635): a table's column widths are a fact about every row, so one [[ColumnFacts]] pass folds
 * the window first — O(columns) memory — then [[header]] and [[lines]] write it row by row.
 * [[render]] is the same lines over a grid held in memory, joined.
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

  /**
   * Render a grid as a markdown table (see [[renderRange]] for the flags): every line — header,
   * separator, one per drawn row — terminated by a newline.
   */
  def render(
    grid: RecordGrid,
    showFormulas: Boolean,
    skipEmpty: Boolean,
    skipHidden: Boolean
  ): String =
    val window = grid.window
    val rows = Stream.emits(grid.rows)
    val facts = ColumnFacts
      .of(window, rows, _.text(showFormulas), skipEmpty, skipHidden)
      .toList
      .headOption
      .getOrElse(ColumnFacts.empty)
    val cols = columns(window, facts, skipEmpty, skipHidden)
    val widths = columnWidths(cols, facts)
    val body = lines(window, rows, cols, widths, showFormulas, skipEmpty, skipHidden).toList
    (header(cols, widths) ++ body :+ "").mkString("\n")

  /** The text a table cell shows and is measured by: the record's display text. */
  def cellText(showFormulas: Boolean): CellRecord => String = _.text(showFormulas)

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

  /** Each drawn column's width: its letter, its widest drawn text, never under 3. */
  def columnWidths(cols: Vector[Int], facts: ColumnFacts): Vector[Int] =
    cols.map { col =>
      val letter = Column.from0(col).toLetter
      math.max(letter.length, math.max(facts.width(col), 3))
    }

  /** The two lines before the rows: the column-letter header and the separator. */
  def header(cols: Vector[Int], widths: Vector[Int]): Vector[String] =
    val letters = cols.zip(widths).map { (col, width) =>
      s" ${Column.from0(col).toLetter.padTo(width, ' ')} |"
    }
    val rule = widths.map(width => "-" * (width + 2) + "|")
    Vector("|   |" + letters.mkString, "|---|" + rule.mkString)

  /**
   * One table line per drawn row of the window ([[RecordWindow.isDrawn]] over `cols`): the row
   * number, then every drawn column's escaped text padded to its width.
   */
  def lines[F[_]](
    window: RecordWindow,
    rows: Stream[F, Vector[CellRecord]],
    cols: Vector[Int],
    widths: Vector[Int],
    showFormulas: Boolean,
    skipEmpty: Boolean,
    skipHidden: Boolean
  ): Stream[F, String] =
    val firstRow = window.range.start.row.index0
    val text = cellText(showFormulas)
    rows.zipWithIndex.map { (row, i) =>
      val rowIdx = firstRow + i.toInt
      Option.when(window.isDrawn(row, rowIdx, cols, skipEmpty, skipHidden)) {
        val cells = cols.zip(widths).map { (col, width) =>
          val escaped = Escape.markdown(window.cell(row, col).fold("")(text))
          s" ${escaped.padTo(width, ' ')} |"
        }
        s"| ${(rowIdx + 1).toString.padTo(2, ' ')}|" + cells.mkString
      }
    }.unNone

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
