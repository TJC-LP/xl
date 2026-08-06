package com.tjclp.xl.cli.output

import com.tjclp.xl.addressing.{ARef, Column, Row}
import com.tjclp.xl.cells.{Cell, CellValue, FormulaKind}
import com.tjclp.xl.sheets.Sheet

/**
 * Shared utilities for all renderers (Markdown, JSON, CSV).
 *
 * Eliminates duplicate implementations of formatEvalError, isCellEmpty, and visibility filtering
 * across the three renderers.
 */
object RendererCommon:

  /**
   * Format evaluation errors as Excel-style error codes.
   *
   * Maps error messages to standard Excel error codes for consistent display across all output
   * formats.
   */
  def formatEvalError(message: String): String =
    if message.toLowerCase.contains("circular") then "#CIRC!"
    else if message.toLowerCase.contains("division") || message.toLowerCase.contains("div") then
      "#DIV/0!"
    else if message.toLowerCase.contains("parse") || message.toLowerCase.contains("unknown") then
      "#NAME?"
    else if message.toLowerCase.contains("ref") then "#REF!"
    else "#ERROR!"

  /**
   * Formula display text: leading `=` always; `{=...}` braces when the cell carries a non-Normal
   * CT_CellFormula record (array/dataTable) — Excel's own formula-bar convention (GH-430).
   */
  def formulaDisplay(expr: String, kind: FormulaKind): String =
    val withEquals = if expr.startsWith("=") then expr else s"=$expr"
    kind match
      case _: FormulaKind.Normal => withEquals
      case _ => s"{$withEquals}"

  /**
   * Human-readable truncation notice emitted when --limit clips output (GH-351).
   *
   * Rendered as a trailer line after markdown tables, or on stderr for machine-parseable formats
   * (CSV) so stdout stays clean.
   */
  def truncationNotice(shown: Int, total: Int, noun: String = "rows"): String =
    s"… showing $shown of $total $noun (use --limit to raise; --limit 0 = no limit)"

  /**
   * Check if a cell is effectively empty.
   *
   * Handles: - Empty cells - Whitespace-only text - Formulas returning empty values
   */
  def isCellEmpty(cell: Cell): Boolean =
    cell.value match
      case CellValue.Empty => true
      case CellValue.Text(s) if s.trim.isEmpty => true
      case CellValue.Formula(_, Some(CellValue.Empty), _) => true
      case CellValue.Formula(_, Some(CellValue.Text(s)), _) if s.trim.isEmpty => true
      case _ => false

  /**
   * Check if a cell at given position is empty.
   *
   * @param sheet
   *   Sheet to check
   * @param col
   *   0-based column index
   * @param row
   *   0-based row index
   * @return
   *   true if cell doesn't exist or is empty
   */
  def isCellEmptyAt(sheet: Sheet, col: Int, row: Int): Boolean =
    sheet.cells.get(ARef.from0(col, row)) match
      case None => true
      case Some(cell) => isCellEmpty(cell)

  /**
   * Get visible column indices (filtering out hidden columns).
   *
   * @param sheet
   *   Sheet to check column properties
   * @param startCol
   *   0-based start column index
   * @param endCol
   *   0-based end column index (inclusive)
   * @return
   *   Sequence of visible column indices
   */
  def visibleColumns(sheet: Sheet, startCol: Int, endCol: Int): IndexedSeq[Int] =
    (startCol to endCol).filterNot { col =>
      sheet.getColumnProperties(Column.from0(col)).hidden
    }

  /**
   * Get visible row indices (filtering out hidden rows).
   *
   * @param sheet
   *   Sheet to check row properties
   * @param startRow
   *   0-based start row index
   * @param endRow
   *   0-based end row index (inclusive)
   * @return
   *   Sequence of visible row indices
   */
  def visibleRows(sheet: Sheet, startRow: Int, endRow: Int): IndexedSeq[Int] =
    (startRow to endRow).filterNot { row =>
      sheet.getRowProperties(Row.from0(row)).hidden
    }

  /**
   * GH-474: column indices to render for a span — every column unless `skipHidden` asked for the
   * visible-only view.
   */
  def renderedColumns(
    sheet: Sheet,
    startCol: Int,
    endCol: Int,
    skipHidden: Boolean
  ): IndexedSeq[Int] =
    if skipHidden then visibleColumns(sheet, startCol, endCol) else startCol to endCol

  /** GH-474: row indices to render for a span (see [[renderedColumns]]). */
  def renderedRows(sheet: Sheet, startRow: Int, endRow: Int, skipHidden: Boolean): IndexedSeq[Int] =
    if skipHidden then visibleRows(sheet, startRow, endRow) else startRow to endRow

  /** 0-based indices of hidden columns inside a span. */
  def hiddenColumns(sheet: Sheet, startCol: Int, endCol: Int): IndexedSeq[Int] =
    (startCol to endCol).filter(col => sheet.getColumnProperties(Column.from0(col)).hidden)

  /** 0-based indices of hidden rows inside a span. */
  def hiddenRows(sheet: Sheet, startRow: Int, endRow: Int): IndexedSeq[Int] =
    (startRow to endRow).filter(row => sheet.getRowProperties(Row.from0(row)).hidden)

  /** Cap long index listings so the notice stays one readable line. */
  private def summarizeIndices(label: String, items: IndexedSeq[String]): String =
    val shown = items.take(10).mkString(", ")
    if items.size > 10 then s"$label $shown, … (${items.size} total)" else s"$label $shown"

  /**
   * GH-474: the hidden-line marker for a rendered range, or None when the range holds no hidden
   * rows/columns.
   *
   * A viewer that drops data from a range the caller named reads as corruption (`search` finds the
   * cell, `cell` reads it, `view` showed nothing). Hidden lines are therefore rendered by default
   * and MARKED; `--skip-hidden` opts back into elision and the marker then names what was dropped.
   *
   * Rendered as a trailer line after markdown tables, on stderr for machine-parseable formats
   * (CSV), and as structured fields in JSON.
   */
  def hiddenNotice(
    sheet: Sheet,
    startCol: Int,
    endCol: Int,
    startRow: Int,
    endRow: Int,
    skipHidden: Boolean
  ): Option[String] =
    val rows = hiddenRows(sheet, startRow, endRow).map(r => (r + 1).toString)
    val cols = hiddenColumns(sheet, startCol, endCol).map(c => Column.from0(c).toLetter)
    if rows.isEmpty && cols.isEmpty then None
    else
      val parts = List(
        Option.when(rows.nonEmpty)(summarizeIndices("row(s)", rows)),
        Option.when(cols.nonEmpty)(summarizeIndices("column(s)", cols))
      ).flatten.mkString(" and ")
      Some(
        if skipHidden then
          s"note: omitted hidden $parts from the requested range (drop --skip-hidden to include them)"
        else
          s"note: range includes hidden $parts — shown because the range was explicitly requested (use --skip-hidden to omit)"
      )

  /**
   * GH-474: `--skip-hidden` cannot be honored under `--stream`.
   *
   * The streaming reader yields cell values only — it never parses row/column properties — so it
   * has no way to know which lines are hidden and renders all of them. Accepting the flag and doing
   * nothing is the silent no-op this issue is about, so the streaming view says so on stderr
   * instead.
   */
  val streamingSkipHiddenNotice: String =
    "note: --skip-hidden is ignored with --stream — the streaming reader does not read row/column " +
      "properties, so hidden rows/columns were rendered. Re-run without --stream to omit them."

  /**
   * Filter columns to only include non-empty ones.
   *
   * @param sheet
   *   Sheet to check
   * @param cols
   *   Column indices to filter
   * @param rows
   *   Row indices to consider when checking emptiness
   * @return
   *   Column indices that have at least one non-empty cell
   */
  def nonEmptyColumns(sheet: Sheet, cols: IndexedSeq[Int], rows: IndexedSeq[Int]): IndexedSeq[Int] =
    cols.filter(col => rows.exists(row => !isCellEmptyAt(sheet, col, row)))

  /**
   * Filter rows to only include non-empty ones.
   *
   * @param sheet
   *   Sheet to check
   * @param rows
   *   Row indices to filter
   * @param cols
   *   Column indices to consider when checking emptiness
   * @return
   *   Row indices that have at least one non-empty cell
   */
  def nonEmptyRows(sheet: Sheet, rows: IndexedSeq[Int], cols: IndexedSeq[Int]): IndexedSeq[Int] =
    rows.filter(row => cols.exists(col => !isCellEmptyAt(sheet, col, row)))
