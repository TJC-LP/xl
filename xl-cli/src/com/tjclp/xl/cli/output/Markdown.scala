package com.tjclp.xl.cli.output

import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row}
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.sheets.Sheet

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
   */
  def renderRange(sheet: Sheet, range: CellRange, showFormulas: Boolean = false): String =
    val sb = new StringBuilder
    val startCol = range.start.col.index0
    val endCol = range.end.col.index0
    val startRow = range.start.row.index0
    val endRow = range.end.row.index0

    // Calculate column widths for better formatting
    val colWidths = (startCol to endCol).map { col =>
      val header = Column.from0(col).toLetter
      val maxContent = (startRow to endRow)
        .map { row =>
          val ref = ARef.from0(col, row)
          sheet.cells.get(ref).map(c => formatCell(c, showFormulas).length).getOrElse(0)
        }
        .maxOption
        .getOrElse(0)
      math.max(header.length, math.max(maxContent, 3)) // Minimum width 3
    }

    // Header row with column letters
    sb.append("|   |")
    colWidths.zip(startCol to endCol).foreach { case (width, col) =>
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

    // Data rows with row numbers
    for row <- startRow to endRow do
      val rowNum = (row + 1).toString
      sb.append(s"| ${rowNum.padTo(2, ' ')}|")
      colWidths.zip(startCol to endCol).foreach { case (width, col) =>
        val ref = ARef.from0(col, row)
        val value = sheet.cells.get(ref).map(c => formatCell(c, showFormulas)).getOrElse("")
        val escaped = escapeMarkdown(value)
        sb.append(s" ${escaped.padTo(width, ' ')} |")
      }
      sb.append("\n")

    sb.toString

  /**
   * Render a list of sheets as a markdown table.
   */
  def renderSheetList(sheets: Vector[(String, Option[CellRange], Int, Int)]): String =
    val sb = new StringBuilder
    sb.append("| # | Name | Range | Cells | Formulas |\n")
    sb.append("|---|------|-------|-------|----------|\n")
    sheets.zipWithIndex.foreach { case ((name, range, cells, formulas), idx) =>
      val rangeStr = range.map(_.toA1).getOrElse("(empty)")
      sb.append(s"| ${idx + 1} | $name | $rangeStr | $cells | $formulas |\n")
    }
    sb.toString

  /**
   * Render label-value pairs as a markdown table.
   */
  def renderLabels(labels: Vector[(String, String, ARef, ARef, String)]): String =
    val sb = new StringBuilder
    sb.append(s"Found ${labels.size} label-value pairs:\n\n")
    sb.append("| Label | Value | Label Ref | Value Ref | Position |\n")
    sb.append("|-------|-------|-----------|-----------|----------|\n")
    labels.foreach { case (label, value, labelRef, valueRef, position) =>
      sb.append(s"| $label | $value | ${labelRef.toA1} | ${valueRef.toA1} | $position |\n")
    }
    sb.toString

  /**
   * Render search results as a markdown table.
   */
  def renderSearchResults(results: Vector[(ARef, String, String)]): String =
    val sb = new StringBuilder
    sb.append(s"Found ${results.size} matches:\n\n")
    sb.append("| Ref | Value | Context |\n")
    sb.append("|-----|-------|--------|\n")
    results.foreach { case (ref, value, context) =>
      sb.append(s"| ${ref.toA1} | ${escapeMarkdown(value)} | ${escapeMarkdown(context)} |\n")
    }
    sb.toString

  /**
   * Format a cell value for display.
   */
  private def formatCell(cell: Cell, showFormulas: Boolean): String =
    cell.value match
      case CellValue.Formula(expr, cached) =>
        if showFormulas then s"=$expr"
        else cached.map(formatValue).getOrElse(s"=$expr")
      case other => formatValue(other)

  private def formatValue(value: CellValue): String =
    value match
      case CellValue.Text(s) => s
      case CellValue.Number(n) => formatNumber(n)
      case CellValue.Bool(b) => if b then "TRUE" else "FALSE"
      case CellValue.DateTime(dt) => dt.toLocalDate.toString
      case CellValue.Error(err) => err.toExcel
      case CellValue.RichText(rt) => rt.toPlainText
      case CellValue.Empty => ""
      case CellValue.Formula(_, Some(cached)) => formatValue(cached)
      case CellValue.Formula(expr, None) => s"=$expr"

  private def formatNumber(n: BigDecimal): String =
    // Simple formatting - could be enhanced with style-based formatting
    if n.isWhole then n.toBigInt.toString
    else n.underlying.stripTrailingZeros.toPlainString

  private def escapeMarkdown(s: String): String =
    s.replace("|", "\\|").replace("\n", " ").replace("\r", "")
