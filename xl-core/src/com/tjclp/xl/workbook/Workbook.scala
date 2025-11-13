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
