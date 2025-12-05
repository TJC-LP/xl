package com.tjclp.xl.cli.output

import com.tjclp.xl.addressing.{ARef, CellRange, Column}
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.display.NumFmtFormatter
import com.tjclp.xl.formula.SheetEvaluator
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.numfmt.NumFmt

/**
 * Markdown table rendering for xl CLI output.
 *
 * All output includes row numbers and column letters for LLM consumption.
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
   */
  def renderRange(
    sheet: Sheet,
    range: CellRange,
    showFormulas: Boolean = false,
    skipEmpty: Boolean = false,
    evalFormulas: Boolean = false
  ): String =
    val sb = new StringBuilder
    val startCol = range.start.col.index0
    val endCol = range.end.col.index0
    val startRow = range.start.row.index0
    val endRow = range.end.row.index0

    // Filter out hidden columns and rows
    val visibleCols = RendererCommon.visibleColumns(sheet, startCol, endCol)
    val visibleRows = RendererCommon.visibleRows(sheet, startRow, endRow)

    // Filter empty columns/rows if skipEmpty is true
    val nonEmptyCols =
      if skipEmpty then RendererCommon.nonEmptyColumns(sheet, visibleCols, visibleRows)
      else visibleCols

    val nonEmptyRows =
      if skipEmpty then RendererCommon.nonEmptyRows(sheet, visibleRows, nonEmptyCols)
      else visibleRows

    // Calculate column widths for better formatting (only visible rows/cols)
    val colWidths = nonEmptyCols.map { col =>
      val header = Column.from0(col).toLetter
      val maxContent = nonEmptyRows
        .map { row =>
          val ref = ARef.from0(col, row)
          sheet.cells
            .get(ref)
            .map(c => formatCell(c, sheet, showFormulas, evalFormulas).length)
            .getOrElse(0)
        }
        .maxOption
        .getOrElse(0)
      math.max(header.length, math.max(maxContent, 3)) // Minimum width 3
    }

    // Header row with column letters
    sb.append("|   |")
    colWidths.zip(nonEmptyCols).foreach { case (width, col) =>
      val header = Column.from0(col).toLetter
      sb.append(s" ${header.padTo(width, ' ')} |")
    }
    sb.append("\n")

    // Separator row
    sb.append("|---|")
    colWidths.foreach { (width: Int) =>
      sb.append("-" * (width + 2))
      sb.append("|")
    }
    sb.append("\n")

    // Data rows with row numbers (only visible rows)
    for row <- nonEmptyRows do
      val rowNum = (row + 1).toString
      sb.append(s"| ${rowNum.padTo(2, ' ')}|")
      colWidths.zip(nonEmptyCols).foreach { case (width, col) =>
        val ref = ARef.from0(col, row)
        val value =
          sheet.cells
            .get(ref)
            .map(c => formatCell(c, sheet, showFormulas, evalFormulas))
            .getOrElse("")
        val escaped = escapeMarkdown(value)
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
   * Render label-value pairs as a markdown table.
   */
  def renderLabels(labels: Vector[(String, String, ARef, ARef, String)]): String =
    val headers = Vector("Label", "Value", "Label Ref", "Value Ref", "Position")
    val rows = labels.map { case (label, value, labelRef, valueRef, position) =>
      Vector(label, value, labelRef.toA1, valueRef.toA1, position)
    }
    s"Found ${labels.size} label-value pairs:\n\n${renderTable(headers, rows)}"

  /**
   * Render search results as a markdown table.
   */
  def renderSearchResults(results: Vector[(ARef, String, String)]): String =
    val headers = Vector("Ref", "Value", "Context")
    val rows = results.map { case (ref, value, context) =>
      Vector(ref.toA1, value, context)
    }
    s"Found ${results.size} matches:\n\n${renderTable(headers, rows)}"

  /**
   * Render search results with qualified refs (for --all-sheets mode).
   */
  def renderSearchResultsWithRef(results: Vector[(String, String)]): String =
    val headers = Vector("Ref", "Value")
    val rows = results.map { case (ref, value) =>
      Vector(ref, value)
    }
    renderTable(headers, rows)

  /**
   * Format a cell value for display with Excel-style number formatting.
   *
   * @param evalFormulas
   *   If true, evaluate formulas live (compute values). Takes precedence over cached values.
   */
  private def formatCell(
    cell: Cell,
    sheet: Sheet,
    showFormulas: Boolean,
    evalFormulas: Boolean
  ): String =
    // Extract NumFmt from cell's style
    val numFmt = cell.styleId
      .flatMap(sheet.styleRegistry.get)
      .map(_.numFmt)
      .getOrElse(NumFmt.General)

    cell.value match
      case CellValue.Formula(expr, cached) =>
        val displayExpr = if expr.startsWith("=") then expr else s"=$expr"
        if showFormulas then displayExpr
        else if evalFormulas then
          // Evaluate formula live
          SheetEvaluator.evaluateFormula(sheet)(displayExpr) match
            case Right(result) => NumFmtFormatter.formatValue(result, numFmt)
            case Left(err) => RendererCommon.formatEvalError(err.message)
        else cached.map(cv => NumFmtFormatter.formatValue(cv, numFmt)).getOrElse(displayExpr)
      case CellValue.RichText(rt) =>
        // Rich text has its own formatting, don't apply NumFmt
        rt.toPlainText
      case CellValue.Empty => ""
      case other => NumFmtFormatter.formatValue(other, numFmt)

  private def escapeMarkdown(s: String): String =
    s.replace("|", "\\|").replace("\n", " ").replace("\r", "")

  /**
   * Generic table renderer with proper column alignment.
   */
  private def renderTable(headers: Vector[String], rows: Vector[Vector[String]]): String =
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
        sb.append(s" ${escapeMarkdown(cell).padTo(width, ' ')} |")
      }
      sb.append("\n")
    }

    sb.toString
