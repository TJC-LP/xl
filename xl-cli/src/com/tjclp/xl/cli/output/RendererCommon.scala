package com.tjclp.xl.cli.output

import com.tjclp.xl.addressing.{ARef, Column, Row}
import com.tjclp.xl.cells.{Cell, CellValue}
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
   * Check if a cell is effectively empty.
   *
   * Handles: - Empty cells - Whitespace-only text - Formulas returning empty values
   */
  def isCellEmpty(cell: Cell): Boolean =
    cell.value match
      case CellValue.Empty => true
      case CellValue.Text(s) if s.trim.isEmpty => true
      case CellValue.Formula(_, Some(CellValue.Empty)) => true
      case CellValue.Formula(_, Some(CellValue.Text(s))) if s.trim.isEmpty => true
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
