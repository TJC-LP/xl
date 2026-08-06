package com.tjclp.xl.cli.output

import com.tjclp.xl.addressing.{ARef, CellRange, Column}
import com.tjclp.xl.cells.{Cell, CellValue, FormulaKind}
import com.tjclp.xl.display.NumFmtFormatter
import com.tjclp.xl.formula.SheetEvaluator
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.numfmt.NumFmt

/**
 * CSV renderer for xl CLI output.
 *
 * Produces RFC 4180-compliant CSV output with optional row/column labels.
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
    val startCol = range.start.col.index0
    val endCol = range.end.col.index0
    val startRow = range.start.row.index0
    val endRow = range.end.row.index0

    // GH-474: hidden rows/columns render unless --skip-hidden asked for the visible-only view
    val visibleCols = RendererCommon.renderedColumns(sheet, startCol, endCol, skipHidden)
    val visibleRows = RendererCommon.renderedRows(sheet, startRow, endRow, skipHidden)

    // Filter empty columns/rows if skipEmpty is true
    val nonEmptyCols =
      if skipEmpty then RendererCommon.nonEmptyColumns(sheet, visibleCols, visibleRows)
      else visibleCols

    val nonEmptyRows =
      if skipEmpty then RendererCommon.nonEmptyRows(sheet, visibleRows, nonEmptyCols)
      else visibleRows

    val sb = new StringBuilder

    // Header row with column letters (if showLabels)
    if showLabels then
      val headerCells = nonEmptyCols.map { colIdx =>
        Column.from0(colIdx).toLetter
      }
      sb.append(",") // Empty cell for row number column
      sb.append(headerCells.mkString(","))
      sb.append("\n")

    // Data rows
    val lastRowIdx = nonEmptyRows.lastOption
    nonEmptyRows.foreach { rowIdx =>
      val rowNum = rowIdx + 1

      if showLabels then
        sb.append(rowNum.toString)
        sb.append(",")

      val cellValues = nonEmptyCols.map { colIdx =>
        val ref = ARef.from0(colIdx, rowIdx)
        sheet.cells.get(ref) match
          case Some(cell) => formatCell(cell, sheet, showFormulas, evalFormulas)
          case None => ""
      }
      sb.append(cellValues.mkString(","))

      // Add newline (except after last row)
      if !lastRowIdx.contains(rowIdx) then sb.append("\n")
    }

    sb.toString

  private def formatCell(
    cell: Cell,
    sheet: Sheet,
    showFormulas: Boolean,
    evalFormulas: Boolean
  ): String =
    val numFmt = cell.styleId
      .flatMap(sheet.styleRegistry.get)
      .map(_.numFmt)
      .getOrElse(NumFmt.General)

    val raw = cell.value match
      case CellValue.Text(s) => s

      case CellValue.Number(n) =>
        NumFmtFormatter.formatValue(cell.value, numFmt)

      case CellValue.Bool(b) =>
        if b then "TRUE" else "FALSE"

      case CellValue.DateTime(dt) =>
        NumFmtFormatter.formatValue(cell.value, numFmt)

      case CellValue.Error(err) =>
        err.toExcel

      case CellValue.RichText(rt) =>
        rt.toPlainText

      case CellValue.Empty =>
        ""

      case CellValue.Formula(expr, cached, kind) =>
        val evalExpr = if expr.startsWith("=") then expr else s"=$expr"
        val displayExpr = RendererCommon.formulaDisplay(expr, kind)
        if showFormulas then displayExpr
        else
          kind match
            case _: FormulaKind.DataTable =>
              // GH-430: TABLE(...) is a record, never evaluable — the cache IS the value
              cached.map(cv => NumFmtFormatter.formatValue(cv, numFmt)).getOrElse(displayExpr)
            case _ if evalFormulas =>
              SheetEvaluator.evaluateFormula(sheet)(evalExpr) match
                case Right(result) => NumFmtFormatter.formatValue(result, numFmt)
                case Left(err) => RendererCommon.formatEvalError(err.message)
            case _ =>
              cached.map(cv => NumFmtFormatter.formatValue(cv, numFmt)).getOrElse(displayExpr)

    escapeCsv(raw)

  /**
   * Escape a value for CSV output per RFC 4180.
   *
   * Rules:
   *   - If value contains comma, quote, or newline: wrap in quotes
   *   - Escape quotes by doubling them
   */
  private def escapeCsv(s: String): String =
    val needsQuoting = s.contains(',') || s.contains('"') || s.contains('\n') || s.contains('\r')
    if needsQuoting then
      val escaped = s.replace("\"", "\"\"")
      s"\"$escaped\""
    else s
