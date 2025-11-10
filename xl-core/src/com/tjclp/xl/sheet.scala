package com.tjclp.xl

import scala.collection.immutable.{Map, Set}

/** Properties for columns */
case class ColumnProperties(
  width: Option[Double] = None,
  hidden: Boolean = false,
  styleId: Option[Int] = None
)

/** Properties for rows */
case class RowProperties(
  height: Option[Double] = None,
  hidden: Boolean = false,
  styleId: Option[Int] = None
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
   * Generates an HTML `<table>` element with cells rendered as `<td>` elements. Rich text formatting and cell styles
   * are preserved as HTML tags and inline CSS.
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

object Sheet:
  /** Create empty sheet with name */
  def apply(name: String): XLResult[Sheet] =
    SheetName(name).left
      .map(err => XLError.InvalidSheetName(name, err))
      .map(sn => Sheet(sn))

  /** Create empty sheet with validated name */
  def apply(name: SheetName): Sheet =
    Sheet(name, Map.empty, Set.empty, Map.empty, Map.empty, None, None, StyleRegistry.default)

/** Workbook metadata */
case class WorkbookMetadata(
  creator: Option[String] = None,
  created: Option[java.time.LocalDateTime] = None,
  modified: Option[java.time.LocalDateTime] = None,
  lastModifiedBy: Option[String] = None,
  application: Option[String] = Some("XL - Pure Scala 3.7 Excel Library"),
  appVersion: Option[String] = Some("0.1.0-SNAPSHOT")
)

/**
 * An Excel workbook containing multiple sheets.
 *
 * Immutable design with efficient persistent data structures.
 */
case class Workbook(
  sheets: Vector[Sheet] = Vector.empty,
  metadata: WorkbookMetadata = WorkbookMetadata(),
  activeSheetIndex: Int = 0
):

  /** Get sheet by index */
  def apply(index: Int): XLResult[Sheet] =
    if index >= 0 && index < sheets.size then Right(sheets(index))
    else Left(XLError.OutOfBounds(s"sheet[$index]", s"Valid range: 0 to ${sheets.size - 1}"))

  /** Get sheet by name */
  def apply(name: SheetName): XLResult[Sheet] =
    sheets
      .find(_.name == name)
      .toRight(XLError.SheetNotFound(name.value))

  /** Get sheet by name string */
  @annotation.targetName("applyByString")
  def apply(name: String): XLResult[Sheet] =
    SheetName(name).left
      .map(err => XLError.InvalidSheetName(name, err): XLError)
      .flatMap(apply)

  /** Add sheet at end */
  def addSheet(sheet: Sheet): XLResult[Workbook] =
    if sheets.exists(_.name == sheet.name) then Left(XLError.DuplicateSheet(sheet.name.value))
    else Right(copy(sheets = sheets :+ sheet))

  /** Insert sheet at index */
  def insertSheet(index: Int, sheet: Sheet): XLResult[Workbook] =
    if index < 0 || index > sheets.size then
      Left(XLError.OutOfBounds(s"insert[$index]", s"Valid range: 0 to ${sheets.size}"))
    else if sheets.exists(_.name == sheet.name) then Left(XLError.DuplicateSheet(sheet.name.value))
    else
      val (before, after) = sheets.splitAt(index)
      Right(copy(sheets = before ++ (sheet +: after)))

  /** Update sheet by index */
  def updateSheet(index: Int, sheet: Sheet): XLResult[Workbook] =
    if index >= 0 && index < sheets.size then
      // Check for duplicate names (excluding the sheet being updated)
      val hasDuplicate = sheets.zipWithIndex
        .exists { case (s, i) => i != index && s.name == sheet.name }
      if hasDuplicate then Left(XLError.DuplicateSheet(sheet.name.value))
      else Right(copy(sheets = sheets.updated(index, sheet)))
    else Left(XLError.OutOfBounds(s"sheet[$index]", s"Valid range: 0 to ${sheets.size - 1}"))

  /** Update sheet by name */
  def updateSheet(name: SheetName, f: Sheet => Sheet): XLResult[Workbook] =
    sheets.indexWhere(_.name == name) match
      case -1 => Left(XLError.SheetNotFound(name.value))
      case index =>
        val updated = f(sheets(index))
        updateSheet(index, updated)

  /** Remove sheet by index */
  def removeSheet(index: Int): XLResult[Workbook] =
    if sheets.size <= 1 then Left(XLError.InvalidWorkbook("Cannot remove last sheet"))
    else if index >= 0 && index < sheets.size then
      val newSheets = sheets.patch(index, Nil, 1)
      val newActiveIndex =
        if activeSheetIndex >= newSheets.size then newSheets.size - 1 else activeSheetIndex
      Right(copy(sheets = newSheets, activeSheetIndex = newActiveIndex))
    else Left(XLError.OutOfBounds(s"sheet[$index]", s"Valid range: 0 to ${sheets.size - 1}"))

  /** Remove sheet by name */
  def removeSheet(name: SheetName): XLResult[Workbook] =
    sheets.indexWhere(_.name == name) match
      case -1 => Left(XLError.SheetNotFound(name.value))
      case index => removeSheet(index)

  /** Rename sheet */
  def renameSheet(oldName: SheetName, newName: SheetName): XLResult[Workbook] =
    sheets.indexWhere(_.name == oldName) match
      case -1 => Left(XLError.SheetNotFound(oldName.value))
      case index =>
        if sheets.exists(s => s.name == newName && s.name != oldName) then
          Left(XLError.DuplicateSheet(newName.value))
        else
          val updated = sheets(index).copy(name = newName)
          Right(copy(sheets = sheets.updated(index, updated)))

  /** Set active sheet index */
  def setActiveSheet(index: Int): XLResult[Workbook] =
    if index >= 0 && index < sheets.size then Right(copy(activeSheetIndex = index))
    else Left(XLError.OutOfBounds(s"active[$index]", s"Valid range: 0 to ${sheets.size - 1}"))

  /** Get active sheet */
  def activeSheet: XLResult[Sheet] = apply(activeSheetIndex)

  /** Get sheet names */
  def sheetNames: Seq[SheetName] = sheets.map(_.name)

  /** Number of sheets */
  def sheetCount: Int = sheets.size

object Workbook:
  /** Create workbook with a single empty sheet */
  def apply(sheetName: String): XLResult[Workbook] =
    for sheet <- Sheet(sheetName)
    yield Workbook(Vector(sheet))

  /** Create empty workbook (requires at least one sheet) */
  def empty: XLResult[Workbook] =
    Workbook("Sheet1")
