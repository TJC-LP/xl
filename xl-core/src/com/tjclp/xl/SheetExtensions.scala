package com.tjclp.xl

import com.tjclp.xl.style.CellStyle

// ========== Style Application Extensions ==========

extension (sheet: Sheet)

  /**
   * Apply a CellStyle to a cell, registering it automatically.
   *
   * Registers the style in the sheet's styleRegistry and applies the resulting index to the cell.
   * If the style is already registered, reuses the existing index.
   */
  @annotation.targetName("withCellStyleExt")
  def withCellStyle(ref: ARef, style: CellStyle): Sheet =
    val (newRegistry, styleId) = sheet.styleRegistry.register(style)
    val cell = sheet(ref).withStyle(styleId)
    sheet.copy(
      styleRegistry = newRegistry,
      cells = sheet.cells.updated(ref, cell)
    )

  /** Apply a CellStyle to all cells in a range. */
  @annotation.targetName("withRangeStyleExt")
  def withRangeStyle(range: CellRange, style: CellStyle): Sheet =
    val (newRegistry, styleId) = sheet.styleRegistry.register(style)
    val updatedCells = range.cells.foldLeft(sheet.cells) { (cells, ref) =>
      val cell = cells.getOrElse(ref, Cell.empty(ref)).withStyle(styleId)
      cells.updated(ref, cell)
    }
    sheet.copy(
      styleRegistry = newRegistry,
      cells = updatedCells
    )

  /** Get the CellStyle for a cell (if it has one). */
  @annotation.targetName("getCellStyleExt")
  def getCellStyle(ref: ARef): Option[CellStyle] =
    sheet(ref).styleId.flatMap(sheet.styleRegistry.get)

  /**
   * Export a cell range to HTML table.
   *
   * Generates an HTML `<table>` element with cells rendered as `<td>` elements. Rich text
   * formatting and cell styles are preserved as HTML tags and inline CSS.
   *
   * @param range
   *   The cell range to export
   * @param includeStyles
   *   Whether to include inline CSS for cell styles (default: true)
   * @return
   *   HTML table string
   */
  @annotation.targetName("toHtmlExt")
  def toHtml(range: CellRange, includeStyles: Boolean = true): String =
    com.tjclp.xl.html.HtmlRenderer.toHtml(sheet, range, includeStyles)

  // ========== Range Combinators ==========

  /**
   * Fill a range of cells using a function that takes column and row coordinates.
   *
   * Cells are filled in row-major order (left-to-right, top-to-bottom) for deterministic behavior.
   *
   * @param range
   *   The cell range to fill
   * @param f
   *   Function that generates cell value from column and row
   * @return
   *   Updated sheet with filled cells
   */
  def fillBy(range: CellRange)(f: (Column, Row) => CellValue): Sheet =
    val newCells = range.cells.map { ref => Cell(ref, f(ref.col, ref.row)) }.toVector
    sheet.putAll(newCells)

  /**
   * Fill a range of cells using a function that takes 0-based indices.
   *
   * Similar to fillBy but uses array-style indexing instead of Excel coordinates.
   *
   * @param range
   *   The cell range to fill
   * @param f
   *   Function that generates cell value from (colIndex, rowIndex) where both are 0-based
   * @return
   *   Updated sheet with filled cells
   */
  def tabulate(range: CellRange)(f: (Int, Int) => CellValue): Sheet =
    val newCells = range.cells.map { ref =>
      Cell(ref, f(ref.col.index0, ref.row.index0))
    }.toVector
    sheet.putAll(newCells)

  /**
   * Put cells in a range in row-major order.
   *
   * The number of supplied values must exactly match the range size to prevent silent data loss
   * from truncation or unintended empty cells from padding. Returns ValueCountMismatch error if
   * counts don't match.
   *
   * @param range
   *   The cell range to populate
   * @param values
   *   Cell values in row-major order (must match range size exactly)
   * @return
   *   Updated sheet or error if value count doesn't match range size
   */
  def putRange(range: CellRange, values: Iterable[CellValue]): XLResult[Sheet] =
    val refs = range.cells.toVector
    val supplied = values.toVector
    val expected = refs.size
    val actual = supplied.size
    if expected != actual then
      Left(XLError.ValueCountMismatch(expected, actual, s"range ${range.toA1}"))
    else
      val newCells = refs.zip(supplied).map { case (ref, value) => Cell(ref, value) }
      Right(sheet.putAll(newCells))

  /**
   * Put a sequence of values in a row starting from a given column.
   *
   * Values are placed left-to-right starting at (row, startCol).
   *
   * @param row
   *   The row to populate
   * @param startCol
   *   The starting column
   * @param values
   *   The cell values to place
   * @return
   *   Updated sheet
   */
  def putRow(row: Row, startCol: Column, values: Iterable[CellValue]): XLResult[Sheet] =
    val buffered = values.toVector
    if buffered.isEmpty then Right(sheet)
    else
      // Check if writing all values would exceed Excel's column limit
      val requiredColumns = buffered.size
      val maxAllowedColumns = Column.MaxIndex0 - startCol.index0 + 1
      if requiredColumns > maxAllowedColumns then
        val maxCol = Column.from0(Column.MaxIndex0).toLetter
        Left(
          XLError.OutOfBounds(
            s"row ${row.index1}",
            s"Cannot write $requiredColumns values starting at ${startCol.toLetter} (Excel limit: column $maxCol)"
          )
        )
      else
        val cells = buffered.zipWithIndex.map { case (value, idx) =>
          Cell(ARef.from0(startCol.index0 + idx, row.index0), value)
        }
        Right(sheet.putAll(cells))

  /**
   * Put a sequence of values in a column starting from a given row.
   *
   * Values are placed top-to-bottom starting at (startRow, col).
   *
   * @param col
   *   The column to populate
   * @param startRow
   *   The starting row
   * @param values
   *   The cell values to place
   * @return
   *   Updated sheet
   */
  def putCol(col: Column, startRow: Row, values: Iterable[CellValue]): XLResult[Sheet] =
    val buffered = values.toVector
    if buffered.isEmpty then Right(sheet)
    else
      // Check if writing all values would exceed Excel's row limit
      val requiredRows = buffered.size
      val maxAllowedRows = Row.MaxIndex0 - startRow.index0 + 1
      if requiredRows > maxAllowedRows then
        Left(
          XLError.OutOfBounds(
            s"column ${col.toLetter}",
            s"Cannot write $requiredRows values starting at row ${startRow.index1} (Excel limit: row ${Row.MaxIndex0 + 1})"
          )
        )
      else
        val cells = buffered.zipWithIndex.map { case (value, idx) =>
          Cell(ARef.from0(col.index0, startRow.index0 + idx), value)
        }
        Right(sheet.putAll(cells))

  // ========== Deterministic Iteration Helpers ==========

  /**
   * Get all cells sorted in row-major order (left-to-right, top-to-bottom).
   *
   * Provides deterministic iteration order matching the canonical write order.
   *
   * @return
   *   Vector of cells sorted by (row, column)
   */
  def cellsSorted: Vector[Cell] =
    sheet.cells.values.toVector.sortBy(c => (c.row.index0, c.col.index0))

  /**
   * Get cells grouped by row, sorted by row index.
   *
   * Each row's cells are sorted left-to-right.
   *
   * @return
   *   Vector of (row, cells) pairs sorted by row index
   */
  def rowsSorted: Vector[(Row, Vector[Cell])] =
    sheet.cells.values
      .groupBy(_.row)
      .toVector
      .sortBy(_._1.index0)
      .map { case (row, cells) => (row, cells.toVector.sortBy(_.col.index0)) }

  /**
   * Get cells grouped by column, sorted by column index.
   *
   * Each column's cells are sorted top-to-bottom.
   *
   * @return
   *   Vector of (column, cells) pairs sorted by column index
   */
  def columnsSorted: Vector[(Column, Vector[Cell])] =
    sheet.cells.values
      .groupBy(_.col)
      .toVector
      .sortBy(_._1.index0)
      .map { case (col, cells) => (col, cells.toVector.sortBy(_.row.index0)) }
