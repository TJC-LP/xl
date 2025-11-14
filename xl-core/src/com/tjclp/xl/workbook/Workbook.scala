package com.tjclp.xl.workbook

import com.tjclp.xl.addressing.{RefType, SheetName}
import com.tjclp.xl.cell.Cell
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.sheet.Sheet

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
      .flatMap((sheetName: SheetName) => apply(sheetName))

  /**
   * Access cell(s) using unified reference type.
   *
   * For qualified refs (Sales!A1), looks up sheet first. For unqualified refs (A1), returns error
   * since workbook needs sheet qualification.
   *
   * Returns Cell for single refs, Iterable[Cell] for ranges.
   */
  @annotation.targetName("applyRefType")
  def apply(ref: RefType): XLResult[Cell | Iterable[Cell]] =
    ref match
      case RefType.Cell(_) | RefType.Range(_) =>
        Left(
          XLError.InvalidReference("Workbook access requires sheet-qualified ref (e.g., Sales!A1)")
        )
      case RefType.QualifiedCell(sheetName, cellRef) =>
        apply(sheetName).map(sheet => sheet(cellRef))
      case RefType.QualifiedRange(sheetName, range) =>
        apply(sheetName).map(sheet => sheet.getRange(range))

  /**
   * Put sheet (add-or-replace by name).
   *
   * If a sheet with the same name exists, replaces it in-place. Otherwise, adds at end. This is the
   * preferred method for adding/updating sheets.
   *
   * Example:
   * {{{
   * val sales = Sheet("Sales").map(_.put(ref"A1" -> "Revenue"))
   * val wb2 = wb.put(sales).unsafe  // Add or replace "Sales" sheet
   * }}}
   */
  def put(sheet: Sheet): XLResult[Workbook] =
    sheets.indexWhere(_.name == sheet.name) match
      case -1 =>
        // Sheet doesn't exist → add at end
        Right(copy(sheets = sheets :+ sheet))
      case index =>
        // Sheet exists → replace in-place (preserves order)
        Right(copy(sheets = sheets.updated(index, sheet)))

  /**
   * Put sheet with explicit name (rename if needed, then add-or-replace).
   *
   * Useful for renaming sheets during updates.
   */
  def put(name: SheetName, sheet: Sheet): XLResult[Workbook] =
    put(sheet.copy(name = name))

  /**
   * Put sheet with explicit name (string variant).
   */
  @annotation.targetName("putWithStringName")
  def put(name: String, sheet: Sheet): XLResult[Workbook] =
    SheetName(name).left
      .map(err => XLError.InvalidSheetName(name, err))
      .flatMap(sn => put(sn, sheet))

  /**
   * Put multiple sheets atomically (batch operation).
   *
   * Transactional semantics: All sheets must be added successfully or the operation fails. If any
   * sheet cannot be added (e.g., validation error), the entire batch fails and the workbook is
   * unchanged. This ensures consistency - you never get a workbook with partial updates.
   *
   * Example:
   * {{{
   * wb.put(Sheet("Sales"), Sheet("Marketing"), Sheet("Finance")) match
   *   case Right(updated) => updated  // All 3 sheets added
   *   case Left(err) => original      // None added, workbook unchanged
   * }}}
   *
   * For partial success semantics, add sheets individually and accumulate results.
   */
  def put(firstSheet: Sheet, restSheets: Sheet*): XLResult[Workbook] =
    (firstSheet +: restSheets).foldLeft(Right(this): XLResult[Workbook]) { (acc, sheet) =>
      acc.flatMap(_.put(sheet))
    }

  /** Add sheet at end */
  @deprecated("Use put(sheet) instead (add-or-replace semantic)", "0.2.0")
  def addSheet(sheet: Sheet): XLResult[Workbook] =
    if sheets.exists(_.name == sheet.name) then Left(XLError.DuplicateSheet(sheet.name.value))
    else Right(copy(sheets = sheets :+ sheet))

  /** Remove sheet by name (preferred method) */
  def remove(name: SheetName): XLResult[Workbook] =
    sheets.indexWhere(_.name == name) match
      case -1 => Left(XLError.SheetNotFound(name.value))
      case index => removeAt(index)

  /** Remove sheet by name (string variant) */
  @annotation.targetName("removeByString")
  def remove(name: String): XLResult[Workbook] =
    SheetName(name).left
      .map(err => XLError.InvalidSheetName(name, err))
      .flatMap(remove)

  /** Remove sheet by index */
  def removeAt(index: Int): XLResult[Workbook] =
    if sheets.size <= 1 then Left(XLError.InvalidWorkbook("Cannot remove last sheet"))
    else if index >= 0 && index < sheets.size then
      val newSheets = sheets.patch(index, Nil, 1)
      val newActiveIndex =
        if activeSheetIndex >= newSheets.size then newSheets.size - 1 else activeSheetIndex
      Right(copy(sheets = newSheets, activeSheetIndex = newActiveIndex))
    else Left(XLError.OutOfBounds(s"sheet[$index]", s"Valid range: 0 to ${sheets.size - 1}"))

  /** Rename sheet */
  def rename(oldName: SheetName, newName: SheetName): XLResult[Workbook] =
    sheets.indexWhere(_.name == oldName) match
      case -1 => Left(XLError.SheetNotFound(oldName.value))
      case index =>
        if sheets.exists(s => s.name == newName && s.name != oldName) then
          Left(XLError.DuplicateSheet(newName.value))
        else
          val updated = sheets(index).copy(name = newName)
          Right(copy(sheets = sheets.updated(index, updated)))

  /**
   * Update sheet by applying a function to it.
   *
   * Convenience method for modify-in-place pattern. Extracts sheet, applies function, puts back.
   *
   * Example:
   * {{{
   * workbook.update("Sales", _.put(ref"A1" -> "Revenue"))
   * }}}
   */
  def update(name: SheetName, f: Sheet => Sheet): XLResult[Workbook] =
    apply(name).flatMap(sheet => put(f(sheet)))

  /**
   * Update sheet by applying a function (string variant).
   */
  @annotation.targetName("updateByString")
  def update(name: String, f: Sheet => Sheet): XLResult[Workbook] =
    SheetName(name).left
      .map(err => XLError.InvalidSheetName(name, err))
      .flatMap(sn => update(sn, f))

  /** Insert sheet at specific index (explicit positioning - rarely needed) */
  def insertAt(index: Int, sheet: Sheet): XLResult[Workbook] =
    if index < 0 || index > sheets.size then
      Left(XLError.OutOfBounds(s"insert[$index]", s"Valid range: 0 to ${sheets.size}"))
    else if sheets.exists(_.name == sheet.name) then Left(XLError.DuplicateSheet(sheet.name.value))
    else
      val (before, after) = sheets.splitAt(index)
      Right(copy(sheets = before ++ (sheet +: after)))

  // ========== Deprecated Methods (Removed in v0.3.0) ==========

  @deprecated("Use put(sheet) instead (add-or-replace semantic)", "0.2.0")
  def insertSheet(index: Int, sheet: Sheet): XLResult[Workbook] =
    insertAt(index, sheet)

  @deprecated("Use put(sheet) instead", "0.2.0")
  def updateSheet(index: Int, sheet: Sheet): XLResult[Workbook] =
    put(sheet)

  @deprecated("Use update(name, f) instead", "0.2.0")
  def updateSheet(name: SheetName, f: Sheet => Sheet): XLResult[Workbook] =
    update(name, f)

  @deprecated("Use remove(name) instead", "0.2.0")
  def removeSheet(name: SheetName): XLResult[Workbook] =
    remove(name)

  @deprecated("Use removeAt(index) instead", "0.2.0")
  def removeSheet(index: Int): XLResult[Workbook] =
    removeAt(index)

  @deprecated("Use rename(oldName, newName) instead", "0.2.0")
  def renameSheet(oldName: SheetName, newName: SheetName): XLResult[Workbook] =
    rename(oldName, newName)

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
