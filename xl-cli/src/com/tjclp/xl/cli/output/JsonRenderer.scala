package com.tjclp.xl.cli.output

import fs2.Stream

import com.tjclp.xl.addressing.{ARef, CellRange, Column, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.contract.StreamedBody
import com.tjclp.xl.cli.read.{CellRecord, InMemorySource, RecordGrid, RecordWindow}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.numfmt.NumFmt

/**
 * JSON renderer for xl CLI output.
 *
 * Produces structured JSON suitable for LLM consumption with cell references, types, raw values,
 * and formatted values — the same rows whether they came from a loaded sheet or the streaming
 * reader — and writes text rather than a JSON tree so every number lexeme is exactly the record's
 * ([[CellRecord.rawJson]]). The document streams (GH-635): [[head]] is its text up to the `rows`
 * (or `records`) array's opening bracket, [[elements]] one array element per drawn row as the rows
 * arrive, [[tail]] the text from the closing bracket on; [[render]] is their composition over a
 * grid held in memory ([[StreamedBody.jsonArray]], the layout `view --format json` has always
 * printed).
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
    val window = grid.window
    val body = elements(window, Stream.emits(grid.rows), header, skipEmpty, skipHidden).toList
    StreamedBody.jsonArray(
      head(window, header.isDefined, truncatedTotalRows, truncatedTotalCols),
      body.toVector,
      tail
    )

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

  /**
   * The document up to and including the array's opening bracket: `{`, `sheet`, `range`, the
   * truncation fields when `--limit`/`--max-cols` clipped the window (GH-351), the hidden-line
   * fields when the window holds hidden lines (GH-474), then `"rows": [` — `"records": [` under
   * `--header-row`.
   */
  def head(
    window: RecordWindow,
    records: Boolean,
    truncatedTotalRows: Option[Int],
    truncatedTotalCols: Option[Int]
  ): String =
    val sb = new StringBuilder
    sb.append("{\n")
    sb.append(s"""  "sheet": ${Escape.json(window.sheet.value)},\n""")
    sb.append(s"""  "range": "${window.range.toA1}",\n""")
    appendTruncationFields(sb, truncatedTotalRows, truncatedTotalCols)
    appendHiddenFields(sb, window)
    sb.append(if records then """  "records": [""" else """  "rows": [""")
    sb.toString

  /** The document from the array's closing bracket on. */
  val tail: String = "]\n}"

  /**
   * One array element per drawn row, as JSON text on its own line — `{"row": n, "cells": [...]}`,
   * or under `header` the `{name: value}` record keyed by the header row's texts (the header row
   * itself is never a record). GH-474: hidden rows render unless `skipHidden`; `skipEmpty` drops
   * empty cells, and a row left with none.
   */
  def elements[F[_]](
    window: RecordWindow,
    rows: Stream[F, Vector[CellRecord]],
    header: Option[(Int, Vector[CellRecord])],
    skipEmpty: Boolean,
    skipHidden: Boolean
  ): Stream[F, String] =
    val visibleCols = window.renderedCols(skipHidden)
    val firstRow = window.range.start.row.index0
    header match
      case Some((headerRowIdx, headerRecords)) =>
        val names = headerNames(window, visibleCols, headerRecords)
        rows.zipWithIndex.map { (row, i) =>
          val rowIdx = firstRow + i.toInt
          if rowIdx == headerRowIdx || !window.isRendered(rowIdx, skipHidden) then None
          else recordElement(window, row, visibleCols, names, skipEmpty)
        }.unNone
      case None =>
        rows.zipWithIndex.map { (row, i) =>
          val rowIdx = firstRow + i.toInt
          if !window.isRendered(rowIdx, skipHidden) then None
          else rowElement(window, row, rowIdx, visibleCols, skipEmpty)
        }.unNone

  /** `{"row": n, "cells": [...]}`, or None when `skipEmpty` leaves the row without a cell. */
  private def rowElement(
    window: RecordWindow,
    row: Vector[CellRecord],
    rowIdx: Int,
    visibleCols: Vector[Int],
    skipEmpty: Boolean
  ): Option[String] =
    val cellJsons = visibleCols.flatMap { colIdx =>
      window.cell(row, colIdx) match
        case Some(record) if skipEmpty && record.isEmpty => None
        case Some(record) => Some(record.toJson(legacyKeys = true))
        case None =>
          if skipEmpty then None
          else
            Some(
              CellRecord
                .empty(window.sheet, ARef.from0(colIdx, rowIdx), hidden = None, None)
                .toJson(legacyKeys = true)
            )
    }
    // Skip entire row if all cells are empty (when skipEmpty is true)
    if skipEmpty && cellJsons.isEmpty then None
    else Some(s"""    {"row": ${rowIdx + 1}, "cells": [${cellJsons.mkString(", ")}]}""")

  /** `{name: value, ...}` keyed by the header texts, or None when `skipEmpty` leaves no field. */
  private def recordElement(
    window: RecordWindow,
    row: Vector[CellRecord],
    visibleCols: Vector[Int],
    names: Map[Int, String],
    skipEmpty: Boolean
  ): Option[String] =
    val fields = visibleCols.flatMap { colIdx =>
      val headerName = names.getOrElse(colIdx, Column.from0(colIdx).toLetter)
      window.cell(row, colIdx) match
        case Some(record) if skipEmpty && record.isEmpty => None
        case Some(record) => Some(s"${Escape.json(headerName)}: ${record.rawJson}")
        case None => if skipEmpty then None else Some(s"${Escape.json(headerName)}: null")
    }
    // Skip entire record if all fields are empty
    if skipEmpty && fields.isEmpty then None
    else Some(s"    {${fields.mkString(", ")}}")

  /** Header names by column: the header cell's display text, the column letter when it is blank. */
  private def headerNames(
    window: RecordWindow,
    visibleCols: Vector[Int],
    headerRecords: Vector[CellRecord]
  ): Map[Int, String] =
    val firstCol = window.range.start.col.index0
    visibleCols.flatMap { colIdx =>
      headerRecords.lift(colIdx - firstCol).map { record =>
        val headerName = headerText(record)
        val name = if headerName.trim.isEmpty then Column.from0(colIdx).toLetter else headerName
        colIdx -> name
      }
    }.toMap

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
  private def appendHiddenFields(sb: StringBuilder, window: RecordWindow): Unit =
    val rows = window.hiddenRowNumbers
    val cols = window.hiddenColLetters
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
      hidden = None,
      mergedInto = None
    )
    s"""{"type": "${record.kind.name}", "value": ${record.rawJson}, "formatted": ${Escape.json(
        record.formatted
      )}}"""
