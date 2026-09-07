package com.tjclp.xl.cli.output

import com.tjclp.xl.addressing.{ARef, CellRange, Column, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.read.{CellRecord, InMemorySource, RecordGrid}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.numfmt.NumFmt

/**
 * JSON renderer for xl CLI output.
 *
 * Produces structured JSON suitable for LLM consumption with cell references, types, raw values,
 * and formatted values. Consumes a [[RecordGrid]] — the same rows whether they came from a loaded
 * sheet or the streaming reader — and writes text rather than a JSON tree so every number lexeme is
 * exactly the record's ([[CellRecord.rawJson]]).
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
   * @param skipHidden
   *   GH-474: if true, omit hidden rows/columns from the payload; default false renders them.
   *   Either way the hidden lines inside the range are reported in the top-level
   *   `hiddenRows`/`hiddenCols` fields (omitted when the range holds none), so JSON never drops an
   *   addressed cell in silence.
   */
  def renderRange(
    sheet: Sheet,
    range: CellRange,
    showFormulas: Boolean = false,
    skipEmpty: Boolean = false,
    headerRow: Option[Int] = None,
    evalFormulas: Boolean = false,
    truncatedTotalRows: Option[Int] = None,
    skipHidden: Boolean = false
  ): String =
    val grid =
      if evalFormulas then InMemorySource.evaluatedGrid(sheet, range)
      else InMemorySource.grid(sheet, range)
    val header = headerRow.map { rowNum =>
      val idx = rowNum - 1
      val records = grid.rows
        .lift(idx - range.start.row.index0)
        .getOrElse {
          val headerRange =
            CellRange(
              ARef.from0(range.start.col.index0, idx),
              ARef.from0(range.end.col.index0, idx)
            )
          val projected =
            if evalFormulas then InMemorySource.evaluatedGrid(sheet, headerRange)
            else InMemorySource.grid(sheet, headerRange)
          projected.rows.headOption.getOrElse(Vector.empty)
        }
      (idx, records)
    }
    render(grid, header, skipEmpty, truncatedTotalRows, None, skipHidden)

  /**
   * Render a grid: `rows` mode, or `records` mode when `header` gives the header row's 0-based
   * index and its records (one per column of the grid; the header row itself is never a record).
   */
  def render(
    grid: RecordGrid,
    header: Option[(Int, Vector[CellRecord])],
    skipEmpty: Boolean,
    truncatedTotalRows: Option[Int],
    truncatedTotalCols: Option[Int],
    skipHidden: Boolean
  ): String =
    header match
      case Some((headerRowIdx, records)) =>
        renderAsRecords(
          grid,
          headerRowIdx,
          records,
          skipEmpty,
          truncatedTotalRows,
          truncatedTotalCols,
          skipHidden
        )
      case None => renderAsRows(grid, skipEmpty, truncatedTotalRows, truncatedTotalCols, skipHidden)

  /**
   * An empty sheet viewed without a range: nothing to address, so `range` is `null` and the rows
   * (or records) are empty.
   */
  def renderEmptySheet(sheet: SheetName, records: Boolean): String =
    val body = if records then "records" else "rows"
    s"""{
       |  "sheet": ${Escape.json(sheet.value)},
       |  "range": null,
       |  "$body": []
       |}""".stripMargin

  /** Render range as array of records with header row values as keys. */
  private def renderAsRecords(
    grid: RecordGrid,
    headerRowIdx: Int,
    headerRecords: Vector[CellRecord],
    skipEmpty: Boolean,
    truncatedTotalRows: Option[Int],
    truncatedTotalCols: Option[Int],
    skipHidden: Boolean
  ): String =
    // GH-474: hidden columns render unless --skip-hidden asked for the visible-only view
    val visibleCols = grid.renderedCols(skipHidden)
    val firstCol = grid.range.start.col.index0

    // Header names: the header cell's display text, the column letter when it is blank
    val headers: Map[Int, String] = visibleCols.flatMap { colIdx =>
      headerRecords.lift(colIdx - firstCol).map { record =>
        val headerName = headerText(record)
        val name = if headerName.trim.isEmpty then Column.from0(colIdx).toLetter else headerName
        colIdx -> name
      }
    }.toMap

    // GH-474: hidden rows render unless --skip-hidden; the header row itself is never a record
    val dataRows = grid.renderedRows(skipHidden).filterNot(_ == headerRowIdx)

    val sb = new StringBuilder
    sb.append("{\n")
    sb.append(s"""  "sheet": ${Escape.json(grid.sheet.value)},\n""")
    sb.append(s"""  "range": "${grid.range.toA1}",\n""")
    appendTruncationFields(sb, truncatedTotalRows, truncatedTotalCols)
    appendHiddenFields(sb, grid)
    sb.append("""  "records": [""")

    val recordJsons = dataRows.flatMap { rowIdx =>
      val fields = visibleCols.flatMap { colIdx =>
        val headerName = headers.getOrElse(colIdx, Column.from0(colIdx).toLetter)
        grid.at(rowIdx, colIdx) match
          case Some(record) if skipEmpty && record.isEmpty => None
          case Some(record) => Some(s"${Escape.json(headerName)}: ${record.rawJson}")
          case None => if skipEmpty then None else Some(s"${Escape.json(headerName)}: null")
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

  /** Original row-based rendering. */
  private def renderAsRows(
    grid: RecordGrid,
    skipEmpty: Boolean,
    truncatedTotalRows: Option[Int],
    truncatedTotalCols: Option[Int],
    skipHidden: Boolean
  ): String =
    // GH-474: hidden rows/cols render unless --skip-hidden (same as Markdown renderer)
    val visibleCols = grid.renderedCols(skipHidden)
    val visibleRows = grid.renderedRows(skipHidden)

    val sb = new StringBuilder
    sb.append("{\n")
    sb.append(s"""  "sheet": ${Escape.json(grid.sheet.value)},\n""")
    sb.append(s"""  "range": "${grid.range.toA1}",\n""")
    appendTruncationFields(sb, truncatedTotalRows, truncatedTotalCols)
    appendHiddenFields(sb, grid)
    sb.append("""  "rows": [""")

    val rowJsons = visibleRows.flatMap { rowIdx =>
      val rowNum = rowIdx + 1
      val cellJsons = visibleCols.flatMap { colIdx =>
        grid.at(rowIdx, colIdx) match
          case Some(record) if skipEmpty && record.isEmpty => None
          case Some(record) => Some(record.toJson(legacyKeys = true))
          case None =>
            if skipEmpty then None
            else
              Some(
                CellRecord
                  .empty(grid.sheet, ARef.from0(colIdx, rowIdx), hidden = false, None)
                  .toJson(legacyKeys = true)
              )
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
   * Append `"truncated": true` / `"totalRows": N` fields when --limit clipped output (GH-351), and
   * `"totalCols": M` when --max-cols clipped the columns.
   */
  private def appendTruncationFields(
    sb: StringBuilder,
    truncatedTotalRows: Option[Int],
    truncatedTotalCols: Option[Int]
  ): Unit =
    if truncatedTotalRows.isDefined || truncatedTotalCols.isDefined then
      sb.append("  \"truncated\": true,\n")
    truncatedTotalRows.foreach(total => sb.append(s"""  "totalRows": $total,\n"""))
    truncatedTotalCols.foreach(total => sb.append(s"""  "totalCols": $total,\n"""))

  /**
   * GH-474: report the hidden rows/columns inside the requested range so a consumer can never
   * mistake an elided (or merely hidden) line for missing data. Emitted whether or not the lines
   * were rendered; omitted entirely when the range holds no hidden lines, keeping the payload
   * byte-identical to previous releases for ordinary ranges.
   */
  private def appendHiddenFields(sb: StringBuilder, grid: RecordGrid): Unit =
    val rows = grid.hiddenRowNumbers
    val cols = grid.hiddenColLetters
    if rows.nonEmpty then sb.append(s"""  "hiddenRows": [${rows.mkString(", ")}],\n""")
    if cols.nonEmpty then
      sb.append(s"""  "hiddenCols": [${cols.map(c => s"\"$c\"").mkString(", ")}],\n""")

  /**
   * A header cell's key text: its display text — an uncached formula shows its stored expression,
   * having no value to show.
   */
  private def headerText(record: CellRecord): String = record.formula match
    case Some(f) if !f.cached => f.expression
    case _ => record.formatted

  /**
   * One value as the `{type, value, formatted}` triple the `view --format json` cells carry, as
   * JSON TEXT for the number-carrying `--json` payloads (`eval`): the same type names, raw-value
   * rules and display formatting as a rendered cell — a whole number prints every digit, never a
   * `Double` — and a formula projects its cached value (`null`/`""` when uncached).
   */
  def valueJson(value: CellValue, numFmt: NumFmt): String =
    val record = CellRecord.of(
      SheetName.unsafe("eval"),
      ARef.from0(0, 0),
      value,
      Some(CellStyle.default.withNumFmt(numFmt)),
      hidden = false,
      mergedInto = None
    )
    s"""{"type": "${record.kind.name}", "value": ${record.rawJson}, "formatted": ${Escape.json(
        record.formatted
      )}}"""
