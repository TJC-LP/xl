package com.tjclp.xl

import scala.collection.immutable.{Map, Set}
import com.tjclp.xl.style.{CellStyle, StyleId, StyleRegistry}

/** Properties for columns */
case class ColumnProperties(
  width: Option[Double] = None,
  hidden: Boolean = false,
  styleId: Option[StyleId] = None
)

/** Properties for rows */
case class RowProperties(
  height: Option[Double] = None,
  hidden: Boolean = false,
  styleId: Option[StyleId] = None
)

/**
 * A worksheet containing cells, merged ranges, and properties.
 *
 * Immutable design: all operations return new Sheet instances. Uses persistent data structures for
 * efficient updates.
 */
case class Sheet(
  name: SheetName,
  cells: Map[ARef, Cell] = Map.empty,
  mergedRanges: Set[CellRange] = Set.empty,
  columnProperties: Map[Column, ColumnProperties] = Map.empty,
  rowProperties: Map[Row, RowProperties] = Map.empty,
  defaultColumnWidth: Option[Double] = None,
  defaultRowHeight: Option[Double] = None,
  styleRegistry: StyleRegistry = StyleRegistry.default
):

  /** Get cell at reference (returns empty cell if not present) */
  def apply(ref: ARef): Cell =
    cells.getOrElse(ref, Cell.empty(ref))

  /** Get cell at A1 notation */
  def apply(a1: String): XLResult[Cell] =
    ARef
      .parse(a1)
      .left
      .map(err => XLError.InvalidCellRef(a1, err))
      .map(apply)

  /** Check if cell exists (not empty) */
  def contains(ref: ARef): Boolean =
    cells.contains(ref)

  /** Put cell at reference */
  def put(cell: Cell): Sheet =
    copy(cells = cells.updated(cell.ref, cell))

  /** Put value at reference */
  def put(ref: ARef, value: CellValue): Sheet =
    put(Cell(ref, value))

  /** Put multiple cells */
  def putAll(newCells: Iterable[Cell]): Sheet =
    copy(cells = cells ++ newCells.map(c => c.ref -> c))

  /** Remove cell at reference */
  def remove(ref: ARef): Sheet =
    copy(cells = cells.removed(ref))

  /** Remove all cells in range */
  def removeRange(range: CellRange): Sheet =
    val toRemove = range.cells.toSet
    copy(cells = cells.filterNot((ref, _) => toRemove.contains(ref)))

  /** Get all cells in a range */
  def getRange(range: CellRange): Iterable[Cell] =
    range.cells.flatMap(ref => cells.get(ref)).toSeq

  /** Put cells in a range (row-major order) */
  def putRange(range: CellRange, values: Iterable[CellValue]): Sheet =
    val newCells = range.cells.zip(values).map((ref, value) => Cell(ref, value))
    putAll(newCells.toSeq)

  /** Merge cells in range */
  def merge(range: CellRange): Sheet =
    copy(mergedRanges = mergedRanges + range)

  /** Unmerge cells in range */
  def unmerge(range: CellRange): Sheet =
    copy(mergedRanges = mergedRanges - range)

  /** Check if cell is part of a merged range */
  def isMerged(ref: ARef): Boolean =
    mergedRanges.exists(_.contains(ref))

  /** Get merged range containing ref (if any) */
  def getMergedRange(ref: ARef): Option[CellRange] =
    mergedRanges.find(_.contains(ref))

  /** Set column properties */
  def setColumnProperties(col: Column, props: ColumnProperties): Sheet =
    copy(columnProperties = columnProperties.updated(col, props))

  /** Get column properties */
  def getColumnProperties(col: Column): ColumnProperties =
    columnProperties.getOrElse(col, ColumnProperties())

  /** Set row properties */
  def setRowProperties(row: Row, props: RowProperties): Sheet =
    copy(rowProperties = rowProperties.updated(row, props))

  /** Get row properties */
  def getRowProperties(row: Row): RowProperties =
    rowProperties.getOrElse(row, RowProperties())

  /** Get all non-empty cells */
  def nonEmptyCells: Iterable[Cell] =
    cells.values.filter(_.nonEmpty)

  /** Get used range (bounding box of all non-empty cells) */
  def usedRange: Option[CellRange] =
    val nonEmpty = nonEmptyCells
    if nonEmpty.isEmpty then None
    else
      val refs = nonEmpty.map(_.ref)
      val minCol = refs.map(_.col.index0).min
      val maxCol = refs.map(_.col.index0).max
      val minRow = refs.map(_.row.index0).min
      val maxRow = refs.map(_.row.index0).max
      Some(
        CellRange(
          ARef.from0(minCol, minRow),
          ARef.from0(maxCol, maxRow)
        )
      )

  /** Count of non-empty cells */
  def cellCount: Int = cells.size

  /** Clear all cells */
  def clearCells: Sheet =
    copy(cells = Map.empty)

  /** Clear all merged ranges */
  def clearMerged: Sheet =
    copy(mergedRanges = Set.empty)

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
  def putRow(row: Row, startCol: Column, values: Iterable[CellValue]): Sheet =
    val cells = values.zipWithIndex.map { case (value, idx) =>
      Cell(ARef.from0(startCol.index0 + idx, row.index0), value)
    }
    sheet.putAll(cells)

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
  def putCol(col: Column, startRow: Row, values: Iterable[CellValue]): Sheet =
    val cells = values.zipWithIndex.map { case (value, idx) =>
      Cell(ARef.from0(col.index0, startRow.index0 + idx), value)
    }
    sheet.putAll(cells)

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

object Sheet:
  /** Create empty sheet with name */
  def apply(name: String): XLResult[Sheet] =
    SheetName(name).left
      .map(err => XLError.InvalidSheetName(name, err))
      .map(sn => Sheet(sn))

  /** Create empty sheet with validated name */
  def apply(name: SheetName): Sheet =
    Sheet(name, Map.empty, Set.empty, Map.empty, Map.empty, None, None, StyleRegistry.default)
