package com.tjclp.xl.cli.output

import com.tjclp.xl.addressing.{ARef, CellRange, Column}
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.display.NumFmtFormatter
import com.tjclp.xl.formula.SheetEvaluator
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.numfmt.NumFmt

/**
 * JSON renderer for xl CLI output.
 *
 * Produces structured JSON suitable for LLM consumption with cell references, types, raw values,
 * and formatted values.
 *
 * No external JSON library dependency - uses manual string building for simplicity and to match the
 * project's minimal-dependency philosophy.
 */
object JsonRenderer:

  /**
   * Render a range as JSON.
   *
   * Output format (default):
   * {{{
   * {
   *   "sheet": "Sheet1",
   *   "range": "A1:D5",
   *   "rows": [
   *     {
   *       "row": 1,
   *       "cells": [
   *         {"ref": "A1", "type": "text", "value": "Revenue", "formatted": "Revenue"},
   *         {"ref": "B1", "type": "number", "value": 1000000, "formatted": "$1,000,000"}
   *       ]
   *     }
   *   ]
   * }
   * }}}
   *
   * Output format (with headerRow):
   * {{{
   * {
   *   "sheet": "Sheet1",
   *   "range": "A1:D5",
   *   "records": [
   *     {"Name": "Widget", "Price": 19.99, "Quantity": 100},
   *     {"Name": "Gadget", "Price": 29.99, "Quantity": 50}
   *   ]
   * }
   * }}}
   *
   * @param skipEmpty
   *   If true, omit cells where type is "empty" from output (reduces token usage for sparse ranges)
   * @param headerRow
   *   If provided, use values from this row (1-based) as object keys in JSON output
   */
  def renderRange(
    sheet: Sheet,
    range: CellRange,
    showFormulas: Boolean = false,
    skipEmpty: Boolean = false,
    headerRow: Option[Int] = None,
    evalFormulas: Boolean = false
  ): String =
    headerRow match
      case Some(headerRowNum) =>
        renderAsRecords(sheet, range, showFormulas, skipEmpty, headerRowNum, evalFormulas)
      case None =>
        renderAsRows(sheet, range, showFormulas, skipEmpty, evalFormulas)

  /**
   * Render range as array of records with header row values as keys.
   */
  private def renderAsRecords(
    sheet: Sheet,
    range: CellRange,
    showFormulas: Boolean,
    skipEmpty: Boolean,
    headerRowNum: Int,
    evalFormulas: Boolean
  ): String =
    val startCol = range.start.col.index0
    val endCol = range.end.col.index0
    val startRow = range.start.row.index0
    val endRow = range.end.row.index0
    val headerRowIdx = headerRowNum - 1 // Convert to 0-based

    // Filter hidden columns
    val visibleCols = RendererCommon.visibleColumns(sheet, startCol, endCol)

    // Get header values
    val headers: Map[Int, String] = visibleCols.flatMap { colIdx =>
      val ref = ARef.from0(colIdx, headerRowIdx)
      sheet.cells.get(ref).map { cell =>
        val headerName = getCellTextValue(cell, sheet)
        // Use column letter as fallback if header is empty
        val name = if headerName.trim.isEmpty then Column.from0(colIdx).toLetter else headerName
        colIdx -> name
      }
    }.toMap

    // Filter out hidden rows and header row itself
    val dataRows = RendererCommon.visibleRows(sheet, startRow, endRow).filterNot(_ == headerRowIdx)

    val sb = new StringBuilder
    sb.append("{\n")
    sb.append(s"""  "sheet": ${escapeJsonString(sheet.name.value)},\n""")
    sb.append(s"""  "range": "${range.toA1}",\n""")
    sb.append("""  "records": [""")

    val recordJsons = dataRows.flatMap { rowIdx =>
      val fields = visibleCols.flatMap { colIdx =>
        val ref = ARef.from0(colIdx, rowIdx)
        val headerName = headers.getOrElse(colIdx, Column.from0(colIdx).toLetter)

        sheet.cells.get(ref) match
          case Some(cell) =>
            val isEmpty = RendererCommon.isCellEmpty(cell)
            if skipEmpty && isEmpty then None
            else
              Some(
                s"${escapeJsonString(headerName)}: ${renderCellValue(cell, sheet, showFormulas, evalFormulas)}"
              )
          case None =>
            if skipEmpty then None
            else Some(s"${escapeJsonString(headerName)}: null")
      }
      // Skip entire record if all fields are empty
      if skipEmpty && fields.isEmpty then None
      else Some(s"    {${fields.mkString(", ")}}")
    }

    if recordJsons.nonEmpty then
      sb.append("\n")
      sb.append(recordJsons.mkString(",\n"))
      sb.append("\n  ")

    sb.append("]\n")
    sb.append("}")
    sb.toString

  /**
   * Original row-based rendering.
   */
  private def renderAsRows(
    sheet: Sheet,
    range: CellRange,
    showFormulas: Boolean,
    skipEmpty: Boolean,
    evalFormulas: Boolean
  ): String =
    val startCol = range.start.col.index0
    val endCol = range.end.col.index0
    val startRow = range.start.row.index0
    val endRow = range.end.row.index0

    // Filter hidden rows/cols (same as Markdown renderer)
    val visibleCols = RendererCommon.visibleColumns(sheet, startCol, endCol)
    val visibleRows = RendererCommon.visibleRows(sheet, startRow, endRow)

    val sb = new StringBuilder
    sb.append("{\n")
    sb.append(s"""  "sheet": ${escapeJsonString(sheet.name.value)},\n""")
    sb.append(s"""  "range": "${range.toA1}",\n""")
    sb.append("""  "rows": [""")

    val rowJsons = visibleRows.flatMap { rowIdx =>
      val rowNum = rowIdx + 1
      val cellJsons = visibleCols.flatMap { colIdx =>
        val ref = ARef.from0(colIdx, rowIdx)
        sheet.cells.get(ref) match
          case Some(cell) =>
            // Check if cell is effectively empty (including formulas returning empty)
            val isEmpty = RendererCommon.isCellEmpty(cell)
            if skipEmpty && isEmpty then None
            else Some(renderCell(ref, cell, sheet, showFormulas, evalFormulas))
          case None =>
            if skipEmpty then None
            else Some(renderEmptyCell(ref))
      }
      // Skip entire row if all cells are empty (when skipEmpty is true)
      if skipEmpty && cellJsons.isEmpty then None
      else Some(s"""    {"row": $rowNum, "cells": [${cellJsons.mkString(", ")}]}""")
    }

    if rowJsons.nonEmpty then
      sb.append("\n")
      sb.append(rowJsons.mkString(",\n"))
      sb.append("\n  ")

    sb.append("]\n")
    sb.append("}")
    sb.toString

  /**
   * Render search results as JSON.
   */
  def renderSearchResults(results: Vector[(ARef, String, String)]): String =
    val sb = new StringBuilder
    sb.append("{\n")
    sb.append(s"""  "count": ${results.size},\n""")
    sb.append("""  "matches": [""")

    val matchJsons = results.map { case (ref, value, context) =>
      s"""{"ref": "${ref.toA1}", "value": ${escapeJsonString(value)}, "context": ${escapeJsonString(
          context
        )}}"""
    }

    if matchJsons.nonEmpty then
      sb.append("\n    ")
      sb.append(matchJsons.mkString(",\n    "))
      sb.append("\n  ")

    sb.append("]\n")
    sb.append("}")
    sb.toString

  private def renderCell(
    ref: ARef,
    cell: Cell,
    sheet: Sheet,
    showFormulas: Boolean,
    evalFormulas: Boolean
  ): String =
    val numFmt = cell.styleId
      .flatMap(sheet.styleRegistry.get)
      .map(_.numFmt)
      .getOrElse(NumFmt.General)

    val (typeStr, rawValue, formatted) = cell.value match
      case CellValue.Text(s) =>
        ("text", escapeJsonString(s), escapeJsonString(s))

      case CellValue.Number(n) =>
        val raw =
          if n.isWhole then n.toBigInt.toString
          else n.underlying.stripTrailingZeros.toPlainString
        ("number", raw, escapeJsonString(NumFmtFormatter.formatValue(cell.value, numFmt)))

      case CellValue.Bool(b) =>
        val boolStr = if b then "true" else "false"
        ("boolean", boolStr, escapeJsonString(if b then "TRUE" else "FALSE"))

      case CellValue.DateTime(dt) =>
        (
          "datetime",
          escapeJsonString(dt.toString),
          escapeJsonString(NumFmtFormatter.formatValue(cell.value, numFmt))
        )

      case CellValue.Error(err) =>
        ("error", escapeJsonString(err.toExcel), escapeJsonString(err.toExcel))

      case CellValue.RichText(rt) =>
        val plain = rt.toPlainText
        ("richtext", escapeJsonString(plain), escapeJsonString(plain))

      case CellValue.Empty =>
        ("empty", "null", "\"\"")

      case CellValue.Formula(expr, cached) =>
        val displayExpr = if expr.startsWith("=") then expr else s"=$expr"
        if showFormulas then
          ("formula", escapeJsonString(displayExpr), escapeJsonString(displayExpr))
        else if evalFormulas then
          SheetEvaluator.evaluateFormula(sheet)(displayExpr) match
            case Right(result) =>
              val formatted = NumFmtFormatter.formatValue(result, numFmt)
              ("formula", escapeJsonString(displayExpr), escapeJsonString(formatted))
            case Left(err) =>
              val errStr = RendererCommon.formatEvalError(err.message)
              ("formula", escapeJsonString(displayExpr), escapeJsonString(errStr))
        else
          val cachedValue =
            cached.map(cv => NumFmtFormatter.formatValue(cv, numFmt)).getOrElse(displayExpr)
          ("formula", escapeJsonString(displayExpr), escapeJsonString(cachedValue))

    s"""{"ref": "${ref.toA1}", "type": "$typeStr", "value": $rawValue, "formatted": $formatted}"""

  private def renderEmptyCell(ref: ARef): String =
    s"""{"ref": "${ref.toA1}", "type": "empty", "value": null, "formatted": ""}"""

  /**
   * Escape a string for JSON output.
   *
   * Handles special characters per JSON spec (RFC 8259).
   */
  private def escapeJsonString(s: String): String =
    val sb = new StringBuilder
    sb.append('"')
    s.foreach {
      case '"' => sb.append("\\\"")
      case '\\' => sb.append("\\\\")
      case '\n' => sb.append("\\n")
      case '\r' => sb.append("\\r")
      case '\t' => sb.append("\\t")
      case '\b' => sb.append("\\b")
      case '\f' => sb.append("\\f")
      case c if c < 32 => sb.append(f"\\u${c.toInt}%04x")
      case c => sb.append(c)
    }
    sb.append('"')
    sb.toString

  /** Get text value from cell for use as header */
  private def getCellTextValue(cell: Cell, sheet: Sheet): String =
    val numFmt = cell.styleId
      .flatMap(sheet.styleRegistry.get)
      .map(_.numFmt)
      .getOrElse(NumFmt.General)

    cell.value match
      case CellValue.Text(s) => s
      case CellValue.Number(n) => NumFmtFormatter.formatValue(cell.value, numFmt)
      case CellValue.Bool(b) => if b then "TRUE" else "FALSE"
      case CellValue.DateTime(dt) => NumFmtFormatter.formatValue(cell.value, numFmt)
      case CellValue.RichText(rt) => rt.toPlainText
      case CellValue.Formula(_, Some(cached)) => getCellTextValueFromCellValue(cached, numFmt)
      case CellValue.Formula(expr, None) => expr
      case CellValue.Error(err) => err.toExcel
      case CellValue.Empty => ""

  private def getCellTextValueFromCellValue(value: CellValue, numFmt: NumFmt): String =
    value match
      case CellValue.Text(s) => s
      case CellValue.Number(n) => NumFmtFormatter.formatValue(value, numFmt)
      case CellValue.Bool(b) => if b then "TRUE" else "FALSE"
      case CellValue.DateTime(dt) => NumFmtFormatter.formatValue(value, numFmt)
      case CellValue.RichText(rt) => rt.toPlainText
      case CellValue.Error(err) => err.toExcel
      case CellValue.Empty => ""
      case CellValue.Formula(_, _) => "" // Shouldn't happen

  /** Render cell value as JSON value (unquoted for numbers/booleans) */
  private def renderCellValue(
    cell: Cell,
    sheet: Sheet,
    showFormulas: Boolean,
    evalFormulas: Boolean
  ): String =
    val numFmt = cell.styleId
      .flatMap(sheet.styleRegistry.get)
      .map(_.numFmt)
      .getOrElse(NumFmt.General)

    cell.value match
      case CellValue.Text(s) => escapeJsonString(s)
      case CellValue.Number(n) =>
        if n.isWhole then n.toBigInt.toString
        else n.underlying.stripTrailingZeros.toPlainString
      case CellValue.Bool(b) => if b then "true" else "false"
      case CellValue.DateTime(dt) => escapeJsonString(dt.toString)
      case CellValue.RichText(rt) => escapeJsonString(rt.toPlainText)
      case CellValue.Error(err) => escapeJsonString(err.toExcel)
      case CellValue.Empty => "null"
      case CellValue.Formula(expr, cached) =>
        val displayExpr = if expr.startsWith("=") then expr else s"=$expr"
        if showFormulas then escapeJsonString(displayExpr)
        else if evalFormulas then
          SheetEvaluator.evaluateFormula(sheet)(displayExpr) match
            case Right(result) => renderCellValueFromCellValue(result, numFmt)
            case Left(err) => escapeJsonString(RendererCommon.formatEvalError(err.message))
        else
          cached match
            case Some(cv) => renderCellValueFromCellValue(cv, numFmt)
            case None => escapeJsonString(displayExpr)

  private def renderCellValueFromCellValue(value: CellValue, numFmt: NumFmt): String =
    value match
      case CellValue.Text(s) => escapeJsonString(s)
      case CellValue.Number(n) =>
        if n.isWhole then n.toBigInt.toString
        else n.underlying.stripTrailingZeros.toPlainString
      case CellValue.Bool(b) => if b then "true" else "false"
      case CellValue.DateTime(dt) => escapeJsonString(dt.toString)
      case CellValue.RichText(rt) => escapeJsonString(rt.toPlainText)
      case CellValue.Error(err) => escapeJsonString(err.toExcel)
      case CellValue.Empty => "null"
      case CellValue.Formula(_, _) => "null" // Shouldn't happen
