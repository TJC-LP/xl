package com.tjclp.xl.tables

import com.tjclp.xl.addressing.CellRange

/**
 * Column definition within an Excel table.
 *
 * Each column has a unique ID (1-indexed) and a display name shown in the header row.
 *
 * @param id
 *   Column identifier (1-indexed, unique within table)
 * @param name
 *   Column display name (shown in header)
 * @since 0.5.0
 */
final case class TableColumn(id: Long, name: String) derives CanEqual

/**
 * AutoFilter configuration for Excel tables.
 *
 * When enabled, Excel displays filter dropdown buttons in the header row.
 *
 * @param enabled
 *   Whether AutoFilter is active
 * @since 0.5.0
 */
final case class TableAutoFilter(enabled: Boolean) derives CanEqual

/**
 * Table visual style (predefined Excel themes).
 *
 * Excel provides built-in table styles with banding, headers, and color schemes. These match the
 * Excel UI table style gallery.
 *
 * @since 0.5.0
 */
enum TableStyle derives CanEqual:
  /** No table style applied */
  case None

  /** Light table styles (Light1-Light21) */
  case Light(number: Int)

  /** Medium table styles (Medium1-Medium28) */
  case Medium(number: Int)

  /** Dark table styles (Dark1-Dark11) */
  case Dark(number: Int)

object TableStyle:
  /** Default table style (Medium 2) - Excel default */
  val default: TableStyle = Medium(2)

/**
 * Excel table specification (structured data range).
 *
 * Tables provide structured references, AutoFilter, totals rows, and visual banding. They are a
 * first-class Excel feature for managing tabular data.
 *
 * '''Example: Creating a table'''{{{ import com.tjclp.xl.* import com.tjclp.xl.tables.*
 *
 * val table = TableSpec( name = "SalesData", displayName = "Sales Data", range = ref"A1:D100",
 * columns = Vector( TableColumn(1, "Product"), TableColumn(2, "Quantity"), TableColumn(3, "Price"),
 * TableColumn(4, "Total") ), autoFilter = Some(TableAutoFilter(enabled = true)), style =
 * TableStyle.Medium(9) )
 *
 * val sheet = Sheet("Q1 Sales").withTable(table) }}}
 *
 * @param name
 *   Unique table identifier (used internally, no spaces allowed)
 * @param displayName
 *   User-visible table name (shown in Excel Name Manager)
 * @param range
 *   Table data range (includes header row)
 * @param columns
 *   Column definitions (count must match range width)
 * @param showHeaderRow
 *   Whether to display the header row (default: true)
 * @param showTotalsRow
 *   Whether to display the totals row at bottom (default: false)
 * @param autoFilter
 *   AutoFilter configuration (None = no filters)
 * @param style
 *   Visual table style
 *
 * @since 0.5.0
 */
final case class TableSpec(
  name: String,
  displayName: String,
  range: CellRange,
  columns: Vector[TableColumn],
  showHeaderRow: Boolean = true,
  showTotalsRow: Boolean = false,
  autoFilter: Option[TableAutoFilter] = None,
  style: TableStyle = TableStyle.default
) derives CanEqual:

  /**
   * Validates that column count matches range width.
   *
   * @return
   *   true if valid, false otherwise
   */
  def isValid: Boolean =
    val rangeWidth = range.end.col.index0 - range.start.col.index0 + 1
    columns.size == rangeWidth

  /**
   * Gets the data range (excluding header and totals rows).
   *
   * @return
   *   Range containing only data rows
   */
  def dataRange: CellRange =
    import com.tjclp.xl.addressing.ARef
    val startRow = if showHeaderRow then range.start.row + 1 else range.start.row
    val endRow = if showTotalsRow then range.end.row - 1 else range.end.row
    val startRef = ARef(range.start.col, startRow)
    val endRef = ARef(range.end.col, endRow)
    new CellRange(startRef, endRef)

object TableSpec:
  /**
   * Create table with auto-generated column IDs.
   *
   * Column IDs are assigned sequentially starting from 1. This is a convenience method when you
   * just have column names and want IDs auto-generated.
   *
   * @param name
   *   Table identifier
   * @param displayName
   *   Display name
   * @param range
   *   Table range
   * @param columnNames
   *   Column names (order matches columns left-to-right)
   * @return
   *   TableSpec with auto-generated column IDs and default settings
   */
  def fromColumnNames(
    name: String,
    displayName: String,
    range: CellRange,
    columnNames: Vector[String]
  ): TableSpec =
    val columns = columnNames.zipWithIndex.map { case (colName, idx) =>
      TableColumn(idx.toLong + 1, colName)
    }
    TableSpec(
      name = name,
      displayName = displayName,
      range = range,
      columns = columns
    )
