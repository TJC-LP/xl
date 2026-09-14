package com.tjclp.xl.tables

import com.tjclp.xl.addressing.CellRange
import com.tjclp.xl.error.{XLError, XLResult}

/**
 * The aggregate a table column's totals-row cell computes — Excel's totals-row dropdown (ECMA-376
 * `ST_TotalsRowFunction`, §18.18.86). `Custom` carries the formula Excel stores as the column's
 * `<totalsRowFormula>` (without the leading `=`).
 *
 * @since 0.23.0
 */
enum TotalsRowFunction derives CanEqual:
  case Sum, Min, Max, Average, Count, CountNums, StdDev, Var
  case Custom(formula: String)

object TotalsRowFunction:
  /** The `ST_TotalsRowFunction` token Excel writes as `totalsRowFunction`. */
  def token(function: TotalsRowFunction): String = function match
    case Sum => "sum"
    case Min => "min"
    case Max => "max"
    case Average => "average"
    case Count => "count"
    case CountNums => "countNums"
    case StdDev => "stdDev"
    case Var => "var"
    case Custom(_) => "custom"

  /**
   * The function a `totalsRowFunction` token names, with the column's `<totalsRowFormula>` for
   * `custom`; `none`, an unknown token, or `custom` without its formula is `None` (no aggregate).
   */
  def fromToken(token: String, formula: Option[String]): Option[TotalsRowFunction] = token match
    case "sum" => Some(Sum)
    case "min" => Some(Min)
    case "max" => Some(Max)
    case "average" => Some(Average)
    case "count" => Some(Count)
    case "countNums" => Some(CountNums)
    case "stdDev" => Some(StdDev)
    case "var" => Some(Var)
    case "custom" => formula.map(Custom.apply)
    case _ => None

/**
 * Column definition within an Excel table.
 *
 * Each column has a unique ID (1-indexed), a display name shown in the header row and, when the
 * table shows a totals row, what its totals-row cell holds: a label (Excel writes "Total" in the
 * first column) or an aggregate (`SUBTOTAL` in Excel; the [[TotalsRowFunction]] the dropdown
 * offers) — both optional, an unlabelled cell without an aggregate is empty.
 *
 * @param id
 *   Column identifier (1-indexed, unique within table)
 * @param name
 *   Column display name (shown in header)
 * @param totalsRowLabel
 *   The text of this column's totals-row cell (`totalsRowLabel`)
 * @param totalsRowFunction
 *   The aggregate this column's totals-row cell computes (`totalsRowFunction`)
 * @since 0.5.0
 */
final case class TableColumn(
  id: Long,
  name: String,
  totalsRowLabel: Option[String] = None,
  totalsRowFunction: Option[TotalsRowFunction] = None
) derives CanEqual

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
  /**
   * Create validated TableSpec with explicit error handling.
   *
   * Validates:
   *   - name: non-empty, alphanumeric + underscore only
   *   - displayName: non-empty, no spaces (Excel requirement)
   *   - range: at least 2 rows (header + data)
   *   - columns: non-empty, no duplicates, count matches range width
   *
   * @return
   *   Either[XLError, TableSpec] with validation errors
   */
  def create(
    name: String,
    displayName: String,
    range: CellRange,
    columns: Vector[TableColumn],
    showHeaderRow: Boolean = true,
    showTotalsRow: Boolean = false,
    autoFilter: Option[TableAutoFilter] = None,
    style: TableStyle = TableStyle.default
  ): XLResult[TableSpec] =
    import com.tjclp.xl.error.XLError

    // Validate name
    if name.isEmpty then Left(XLError.InvalidTableName(name, "Table name cannot be empty"))
    else if !name.matches("^[a-zA-Z0-9_]+$") then
      Left(XLError.InvalidTableName(name, "Table name must be alphanumeric and underscores only"))
    // Validate displayName
    else if displayName.isEmpty then
      Left(XLError.InvalidTableDisplayName(displayName, "Display name cannot be empty"))
    else if displayName.contains(' ') then
      Left(
        XLError.InvalidTableDisplayName(
          displayName,
          "Display name cannot contain spaces (Excel requirement). Use underscores instead."
        )
      )
    else if !displayName.matches("^[a-zA-Z0-9_]+$") then
      Left(
        XLError.InvalidTableDisplayName(
          displayName,
          "Display name must be alphanumeric and underscores only"
        )
      )
    // Validate range
    else if range.height < 2 then
      Left(
        XLError.InvalidTableRange(
          range.toA1,
          s"Table range must have at least 2 rows (header + data), got ${range.height}"
        )
      )
    // Validate columns
    else if columns.isEmpty then
      Left(XLError.InvalidTableColumns("Table must have at least one column"))
    else
      val duplicateNames = columns.groupBy(_.name).filter(_._2.size > 1).keys
      if duplicateNames.nonEmpty then
        Left(
          XLError.InvalidTableColumns(s"Duplicate column names: ${duplicateNames.mkString(", ")}")
        )
      else if columns.size != range.width then
        Left(
          XLError.InvalidTableColumns(
            s"Column count (${columns.size}) must match range width (${range.width})"
          )
        )
      else
        Right(
          TableSpec(
            name = name,
            displayName = displayName,
            range = range,
            columns = columns,
            showHeaderRow = showHeaderRow,
            showTotalsRow = showTotalsRow,
            autoFilter = autoFilter,
            style = style
          )
        )

  /**
   * Convenience factory: create table from column names with validation.
   *
   * Auto-generates column IDs (1-indexed) from provided names. Uses TableSpec.create for validation
   * (returns Either for explicit error handling).
   *
   * @return
   *   Either[XLError, TableSpec] with validation errors
   */
  def fromColumnNames(
    name: String,
    displayName: String,
    range: CellRange,
    columnNames: Vector[String]
  ): XLResult[TableSpec] =
    val columns = columnNames.zipWithIndex.map { case (colName, idx) =>
      TableColumn(idx.toLong + 1, colName)
    }
    create(
      name = name,
      displayName = displayName,
      range = range,
      columns = columns
    )

  /**
   * Unsafe version of fromColumnNames for test convenience. Throws if validation fails.
   *
   * WARNING: Only use in tests. Production code should use fromColumnNames and handle Either.
   */
  def unsafeFromColumnNames(
    name: String,
    displayName: String,
    range: CellRange,
    columnNames: Vector[String]
  ): TableSpec =
    fromColumnNames(name, displayName, range, columnNames).getOrElse(
      throw new IllegalArgumentException(s"Invalid table: $name")
    )
