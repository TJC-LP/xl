package com.tjclp.xl.cli.output

import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row}
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.display.NumFmtFormatter
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
   * Output format:
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
   */
  def renderRange(sheet: Sheet, range: CellRange, showFormulas: Boolean = false): String =
    val startCol = range.start.col.index0
    val endCol = range.end.col.index0
    val startRow = range.start.row.index0
    val endRow = range.end.row.index0

    // Filter hidden rows/cols (same as Markdown renderer)
    val visibleCols = (startCol to endCol).filterNot { col =>
      sheet.getColumnProperties(Column.from0(col)).hidden
    }
    val visibleRows = (startRow to endRow).filterNot { row =>
      sheet.getRowProperties(Row.from0(row)).hidden
    }

    val sb = new StringBuilder
    sb.append("{\n")
    sb.append(s"""  "sheet": ${escapeJsonString(sheet.name.value)},\n""")
    sb.append(s"""  "range": "${range.toA1}",\n""")
    sb.append("""  "rows": [""")

    val rowJsons = visibleRows.map { rowIdx =>
      val rowNum = rowIdx + 1
      val cellJsons = visibleCols.map { colIdx =>
        val ref = ARef.from0(colIdx, rowIdx)
        sheet.cells.get(ref) match
          case Some(cell) => renderCell(ref, cell, sheet, showFormulas)
          case None => renderEmptyCell(ref)
      }
      s"""    {"row": $rowNum, "cells": [${cellJsons.mkString(", ")}]}"""
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

  private def renderCell(ref: ARef, cell: Cell, sheet: Sheet, showFormulas: Boolean): String =
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
