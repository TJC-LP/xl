package com.tjclp.xl.cli.output

import com.tjclp.xl.addressing.{ARef, CellRange, Column}
import com.tjclp.xl.cells.{Cell, CellValue, FormulaKind}
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
   *         {"ref": "B1", "type": "number", "value": 1000000, "formatted": "$1,000,000"},
   *         {"ref": "C1", "type": "formula", "formula": "=B1*2", "value": 2000000, "formatted": "$2,000,000"}
   *       ]
   *     }
   *   ]
   * }
   * }}}
   *
   * Formula cells always carry the expression in a dedicated `formula` field (GH-357); `value` and
   * `formatted` hold the computed (`evalFormulas`) or cached value — `null`/`""` when uncached.
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
   * @param showFormulas
   *   Ignored for JSON output (GH-357): the formula expression is always present in the `formula`
   *   field, so there is nothing to swap. Retained for signature compatibility with the other
   *   renderers — `--formulas` only controls non-JSON display formats.
   * @param skipEmpty
   *   If true, omit cells where type is "empty" from output (reduces token usage for sparse ranges)
   * @param headerRow
   *   If provided, use values from this row (1-based) as object keys in JSON output
   * @param truncatedTotalRows
   *   When --limit clipped the requested range (GH-351), the total row count that would have been
   *   rendered without the limit. Emits top-level `"truncated": true` and `"totalRows": N` fields.
   *   Fields are omitted entirely when output is not clipped, keeping the payload byte-identical to
   *   previous releases for unclipped output.
   */
  def renderRange(
    sheet: Sheet,
    range: CellRange,
    showFormulas: Boolean = false,
    skipEmpty: Boolean = false,
    headerRow: Option[Int] = None,
    evalFormulas: Boolean = false,
    truncatedTotalRows: Option[Int] = None
  ): String =
    headerRow match
      case Some(headerRowNum) =>
        renderAsRecords(
          sheet,
          range,
          skipEmpty,
          headerRowNum,
          evalFormulas,
          truncatedTotalRows
        )
      case None =>
        renderAsRows(sheet, range, skipEmpty, evalFormulas, truncatedTotalRows)

  /**
   * Render range as array of records with header row values as keys.
   */
  private def renderAsRecords(
    sheet: Sheet,
    range: CellRange,
    skipEmpty: Boolean,
    headerRowNum: Int,
    evalFormulas: Boolean,
    truncatedTotalRows: Option[Int]
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
    appendTruncationFields(sb, truncatedTotalRows)
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
                s"${escapeJsonString(headerName)}: ${renderCellValue(cell, sheet, evalFormulas)}"
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
    skipEmpty: Boolean,
    evalFormulas: Boolean,
    truncatedTotalRows: Option[Int]
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
    appendTruncationFields(sb, truncatedTotalRows)
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
            else Some(renderCell(ref, cell, sheet, evalFormulas))
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

  /** Append `"truncated": true` / `"totalRows": N` fields when --limit clipped output (GH-351). */
  private def appendTruncationFields(sb: StringBuilder, truncatedTotalRows: Option[Int]): Unit =
    truncatedTotalRows.foreach { total =>
      sb.append("  \"truncated\": true,\n")
      sb.append(s"""  "totalRows": $total,\n""")
    }

  private def renderCell(
    ref: ARef,
    cell: Cell,
    sheet: Sheet,
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

      case CellValue.Formula(expr, cached, kind) =>
        // GH-430: a dataTable record is never evaluated — its cache IS the value
        val evaluable = kind match
          case _: FormulaKind.DataTable => false
          case _ => true
        val (raw, fmt) =
          formulaValueJson(
            sheet,
            displayExpression(expr),
            cached,
            numFmt,
            evalFormulas && evaluable
          )
        ("formula", raw, fmt)

    // GH-357: formula cells always carry the expression in a dedicated field; value/formatted
    // hold the computed or cached value. --formulas only affects non-JSON display formats.
    // GH-430: non-Normal record kinds surface additively as "formulaKind".
    val formulaField = cell.value match
      case CellValue.Formula(expr, _, kind) =>
        val kindField = kind match
          case FormulaKind.Normal => ""
          case _: FormulaKind.ArrayFormula => """, "formulaKind": "array""""
          case _: FormulaKind.DataTable => """, "formulaKind": "dataTable""""
        s""", "formula": ${escapeJsonString(displayExpression(expr))}$kindField"""
      case _ => ""

    s"""{"ref": "${ref.toA1}", "type": "$typeStr"$formulaField, "value": $rawValue, "formatted": $formatted}"""

  /** Formula expression as displayed: always with a leading `=`. */
  private def displayExpression(expr: String): String =
    if expr.startsWith("=") then expr else s"=$expr"

  /**
   * Raw JSON value + formatted string for a formula cell (GH-357): evaluated when `evalFormulas`
   * (error token on failure), else the cached value; `null`/`""` when uncached.
   */
  private def formulaValueJson(
    sheet: Sheet,
    displayExpr: String,
    cached: Option[CellValue],
    numFmt: NumFmt,
    evalFormulas: Boolean
  ): (String, String) =
    if evalFormulas then
      SheetEvaluator.evaluateFormula(sheet)(displayExpr) match
        case Right(result) =>
          (
            renderCellValueFromCellValue(result, numFmt),
            escapeJsonString(NumFmtFormatter.formatValue(result, numFmt))
          )
        case Left(err) =>
          val errStr = escapeJsonString(RendererCommon.formatEvalError(err.message))
          (errStr, errStr)
    else
      cached match
        case Some(cv) =>
          (
            renderCellValueFromCellValue(cv, numFmt),
            escapeJsonString(NumFmtFormatter.formatValue(cv, numFmt))
          )
        case None => ("null", "\"\"")

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
      case CellValue.Formula(_, Some(cached), _) => getCellTextValueFromCellValue(cached, numFmt)
      case CellValue.Formula(expr, None, _) => expr
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
      case CellValue.Formula(_, _, _) => "" // Shouldn't happen

  /** Render cell value as JSON value (unquoted for numbers/booleans) */
  private def renderCellValue(
    cell: Cell,
    sheet: Sheet,
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
      case CellValue.Formula(expr, cached, kind) =>
        // GH-357: records mode is a scalar projection — always the computed/cached value,
        // never the expression; null when uncached (matches empty cells).
        // GH-430: a dataTable record is never evaluated — its cache IS the value.
        val evaluable = kind match
          case _: FormulaKind.DataTable => false
          case _ => true
        formulaValueJson(
          sheet,
          displayExpression(expr),
          cached,
          numFmt,
          evalFormulas && evaluable
        )._1

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
      case CellValue.Formula(_, _, _) => "null" // Shouldn't happen
