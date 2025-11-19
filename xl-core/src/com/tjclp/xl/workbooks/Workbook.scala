package com.tjclp.xl.workbooks

import com.tjclp.xl.SourceContext
import com.tjclp.xl.addressing.{RefType, SheetName}
import com.tjclp.xl.cells.Cell
import com.tjclp.xl.errors.{XLError, XLResult}
import com.tjclp.xl.sheets.Sheet

/**
 * An Excel workbooks containing multiple sheets.
 *
 * Immutable design with efficient persistent data structures.
 */
case class Workbook(
  sheets: Vector[Sheet] = Vector.empty,
  metadata: WorkbookMetadata = WorkbookMetadata(),
  activeSheetIndex: Int = 0,
  sourceContext: Option[SourceContext] = None
):

  /** Get sheets by index */
  def apply(index: Int): XLResult[Sheet] =
    if index >= 0 && index < sheets.size then Right(sheets(index))
    else Left(XLError.OutOfBounds(s"sheets[$index]", s"Valid range: 0 to ${sheets.size - 1}"))

  /** Get sheets by name */
  def apply(name: SheetName): XLResult[Sheet] =
    sheets
      .find(_.name == name)
      .toRight(XLError.SheetNotFound(name.value))

  /** Get sheets by name string */
  @annotation.targetName("applyByString")
  def apply(name: String): XLResult[Sheet] =
    SheetName(name).left
      .map(err => XLError.InvalidSheetName(name, err): XLError)
      .flatMap((sheetName: SheetName) => apply(sheetName))

  /**
   * Access cell(s) using unified reference type.
   *
   * For qualified refs (Sales!A1), looks up sheets first. For unqualified refs (A1), returns errors
   * since workbooks needs sheets qualification.
   *
   * Returns Cell for single refs, Iterable[Cell] for ranges.
   */
  @annotation.targetName("applyRefType")
  def apply(ref: RefType): XLResult[Cell | Iterable[Cell]] =
    ref match
      case RefType.Cell(_) | RefType.Range(_) =>
        Left(
          XLError.InvalidReference("Workbook access requires sheets-qualified ref (e.g., Sales!A1)")
        )
      case RefType.QualifiedCell(sheetName, cellRef) =>
        apply(sheetName).map(sheet => sheet(cellRef))
      case RefType.QualifiedRange(sheetName, range) =>
        apply(sheetName).map(sheet => sheet.getRange(range))

  /**
   * Put sheets (add-or-replace by name).
   *
   * If a sheets with the same name exists, replaces it in-place and marks as modified for surgical
   * writes. Otherwise, adds at end. This is the preferred method for adding/updating sheets.
   *
   * Example:
   * {{{
   * val sales = Sheet("Sales").map(_.put(ref"A1" -> "Revenue"))
   * val wb2 = wb.put(sales).unsafe  // Add or replace "Sales" sheets
   * }}}
   */
  def put(sheet: Sheet): XLResult[Workbook] =
    sheets.indexWhere(_.name == sheet.name) match
      case -1 =>
        // Sheet doesn't exist → add at end (no tracking needed for new sheets)
        Right(copy(sheets = sheets :+ sheet))
      case index =>
        // Sheet exists → replace in-place and track modification
        val updatedContext = sourceContext.map(_.markSheetModified(index))
        Right(copy(sheets = sheets.updated(index, sheet), sourceContext = updatedContext))

  /**
   * Put sheets with explicit name (rename if needed, then add-or-replace).
   *
   * Useful for renaming sheets during updates.
   */
  def put(name: SheetName, sheet: Sheet): XLResult[Workbook] =
    put(sheet.copy(name = name))

  /**
   * Put sheets with explicit name (string variant).
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
   * sheets cannot be added (e.g., validation errors), the entire batch fails and the workbooks is
   * unchanged. This ensures consistency - you never get a workbooks with partial updates.
   *
   * Example:
   * {{{
   * wb.put(Sheet("Sales"), Sheet("Marketing"), Sheet("Finance")) match
   *   case Right(updated) => updated  // All 3 sheets added
   *   case Left(err) => original      // None added, workbooks unchanged
   * }}}
   *
   * For partial success semantics, add sheets individually and accumulate results.
   */
  def put(firstSheet: Sheet, restSheets: Sheet*): XLResult[Workbook] =
    (firstSheet +: restSheets).foldLeft(Right(this): XLResult[Workbook]) { (acc, sheet) =>
      acc.flatMap(_.put(sheet))
    }

  /** Add sheets at end */
  @deprecated("Use put(sheets) instead (add-or-replace semantic)", "0.2.0")
  def addSheet(sheet: Sheet): XLResult[Workbook] =
    if sheets.exists(_.name == sheet.name) then Left(XLError.DuplicateSheet(sheet.name.value))
    else Right(copy(sheets = sheets :+ sheet))

  /** Remove sheets by name (preferred method) */
  def remove(name: SheetName): XLResult[Workbook] =
    sheets.indexWhere(_.name == name) match
      case -1 => Left(XLError.SheetNotFound(name.value))
      case index => removeAt(index)

  /** Remove sheets by name (string variant) */
  @annotation.targetName("removeByString")
  def remove(name: String): XLResult[Workbook] =
    SheetName(name).left
      .map(err => XLError.InvalidSheetName(name, err))
      .flatMap(remove)

  /** Remove sheets by index (Excel requires at least one sheets to remain). */
  def removeAt(index: Int): XLResult[Workbook] =
    // Validation: Excel workbooks must have at least one sheets
    if sheets.size <= 1 then Left(XLError.InvalidWorkbook("Cannot remove last sheets"))
    else if index >= 0 && index < sheets.size then
      val newSheets = sheets.patch(index, Nil, 1)
      val newActiveIndex =
        if activeSheetIndex >= newSheets.size then newSheets.size - 1 else activeSheetIndex
      val updatedContext = sourceContext.map(_.markSheetDeleted(index))
      Right(
        copy(sheets = newSheets, activeSheetIndex = newActiveIndex, sourceContext = updatedContext)
      )
    else Left(XLError.OutOfBounds(s"sheets[$index]", s"Valid range: 0 to ${sheets.size - 1}"))

  /** Rename sheets (marks metadata as modified since sheets names live in workbook.xml). */
  def rename(oldName: SheetName, newName: SheetName): XLResult[Workbook] =
    sheets.indexWhere(_.name == oldName) match
      case -1 => Left(XLError.SheetNotFound(oldName.value))
      case index =>
        if sheets.exists(s => s.name == newName && s.name != oldName) then
          Left(XLError.DuplicateSheet(newName.value))
        else
          val updated = sheets(index).copy(name = newName)
          val updatedContext = sourceContext.map(_.markMetadataModified)
          Right(copy(sheets = sheets.updated(index, updated), sourceContext = updatedContext))

  /**
   * Update sheets by applying a function to it.
   *
   * Convenience method for modify-in-place pattern. Extracts sheets, applies function, puts back.
   *
   * Example:
   * {{{
   * workbooks.update("Sales", _.put(ref"A1" -> "Revenue"))
   * }}}
   */
  def update(name: SheetName, f: Sheet => Sheet): XLResult[Workbook] =
    sheets.indexWhere(_.name == name) match
      case -1 => Left(XLError.SheetNotFound(name.value))
      case idx => updateAt(idx, f)

  /**
   * Update sheets by applying a function (string variant).
   */
  @annotation.targetName("updateByString")
  def update(name: String, f: Sheet => Sheet): XLResult[Workbook] =
    SheetName(name).left
      .map(err => XLError.InvalidSheetName(name, err))
      .flatMap(sn => update(sn, f))

  /** Update sheets by index while tracking modification state. */
  def updateAt(idx: Int, f: Sheet => Sheet): XLResult[Workbook] =
    if idx < 0 || idx >= sheets.size then
      Left(XLError.OutOfBounds(s"sheets[$idx]", s"Valid range: 0 to ${sheets.size - 1}"))
    else
      val updatedSheet = f(sheets(idx))
      val newSheets = sheets.updated(idx, updatedSheet)
      val updatedContext = sourceContext.map(_.markSheetModified(idx))
      Right(copy(sheets = newSheets, sourceContext = updatedContext))

  /** Delete sheets by name while tracking modification state. */
  def delete(name: SheetName): XLResult[Workbook] =
    sheets.indexWhere(_.name == name) match
      case -1 => Left(XLError.SheetNotFound(name.value))
      case idx => removeAt(idx)

  /** Delete sheets by name (string variant). */
  @annotation.targetName("deleteByString")
  def delete(name: String): XLResult[Workbook] =
    SheetName(name).left
      .map(err => XLError.InvalidSheetName(name, err))
      .flatMap(sn => delete(sn))

  /** Reorder sheets to the provided order while tracking modifications. */
  def reorder(newOrder: Vector[SheetName]): XLResult[Workbook] =
    if newOrder.size != sheets.size || newOrder.toSet != sheets.map(_.name).toSet then
      Left(XLError.InvalidWorkbook("Sheet names must match existing set"))
    else
      val nameToSheet = sheets.map(sheet => sheet.name -> sheet).toMap
      val reordered = newOrder.map(nameToSheet)
      val activeName = sheets.lift(activeSheetIndex).map(_.name)
      val newActiveIndex = activeName
        .flatMap(name =>
          newOrder.indexWhere(_ == name) match
            case -1 => None
            case idx => Some(idx)
        )
        .getOrElse(activeSheetIndex)
      // Sheet reordering only modifies workbook.xml (order metadata), not individual sheets files.
      // Therefore we mark reordered but don't mark individual sheets as modified.
      val updatedContext = sourceContext.map(_.markReordered)
      Right(
        copy(sheets = reordered, activeSheetIndex = newActiveIndex, sourceContext = updatedContext)
      )

  /** Insert sheets at specific index (explicit positioning - rarely needed) */
  def insertAt(index: Int, sheet: Sheet): XLResult[Workbook] =
    if index < 0 || index > sheets.size then
      Left(XLError.OutOfBounds(s"insert[$index]", s"Valid range: 0 to ${sheets.size}"))
    else if sheets.exists(_.name == sheet.name) then Left(XLError.DuplicateSheet(sheet.name.value))
    else
      val (before, after) = sheets.splitAt(index)
      Right(copy(sheets = before ++ (sheet +: after)))

  // ========== Deprecated Methods (Removed in v0.3.0) ==========

  @deprecated("Use put(sheets) instead (add-or-replace semantic)", "0.2.0")
  def insertSheet(index: Int, sheet: Sheet): XLResult[Workbook] =
    insertAt(index, sheet)

  @deprecated("Use remove(name) instead", "0.2.0")
  def removeSheet(name: SheetName): XLResult[Workbook] =
    remove(name)

  @deprecated("Use removeAt(index) instead", "0.2.0")
  def removeSheet(index: Int): XLResult[Workbook] =
    removeAt(index)

  @deprecated("Use rename(oldName, newName) instead", "0.2.0")
  def renameSheet(oldName: SheetName, newName: SheetName): XLResult[Workbook] =
    rename(oldName, newName)

  /** Set active sheets index */
  def setActiveSheet(index: Int): XLResult[Workbook] =
    if index >= 0 && index < sheets.size then Right(copy(activeSheetIndex = index))
    else Left(XLError.OutOfBounds(s"active[$index]", s"Valid range: 0 to ${sheets.size - 1}"))

  /** Get active sheets */
  def activeSheet: XLResult[Sheet] = apply(activeSheetIndex)

  /** Get sheets names */
  def sheetNames: Seq[SheetName] = sheets.map(_.name)

  /** Number of sheets */
  def sheetCount: Int = sheets.size

object Workbook:
  /** Create workbooks with a single empty sheets */
  def apply(sheetName: String): XLResult[Workbook] =
    for sheet <- Sheet(sheetName)
    yield Workbook(Vector(sheet))

  /** Create empty workbooks (requires at least one sheets) */
  def empty: XLResult[Workbook] =
    Workbook("Sheet1")
