package com.tjclp.xl.workbooks

import com.tjclp.xl.context.SourceContext
import com.tjclp.xl.addressing.{RefType, SheetName}
import com.tjclp.xl.cells.Cell
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.sheets.Sheet

/**
 * An Excel workbook containing multiple sheets.
 *
 * Immutable design with efficient persistent data structures.
 */
final case class Workbook(
  sheets: Vector[Sheet] = Vector.empty,
  metadata: WorkbookMetadata = WorkbookMetadata(),
  activeSheetIndex: Int = 0,
  sourceContext: Option[SourceContext] = None
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

  /**
   * Get sheet by name string.
   *
   * When called with a string literal, the name format is validated at compile time. Invalid
   * literals like "Invalid:Name" fail to compile. Runtime strings are validated at runtime.
   *
   * Always returns XLResult[Sheet] since sheet existence is runtime-dependent.
   */
  @annotation.targetName("applyByString")
  transparent inline def apply(inline name: String): XLResult[Sheet] =
    ${ com.tjclp.xl.macros.WorkbookMacros.applyImpl('{ this }, 'name) }

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
   * If a sheet with the same name exists, replaces it in-place and marks as modified for surgical
   * writes. Otherwise, adds at end. This is the preferred method for adding/updating sheets.
   *
   * This operation is infallible - it always succeeds for valid Sheet inputs.
   *
   * Example:
   * {{{
   * val sales = Sheet("Sales").put(ref"A1" -> "Revenue")
   * val wb2 = wb.put(sales)  // Add or replace "Sales" sheet
   * }}}
   */
  def put(sheet: Sheet): Workbook =
    sheets.indexWhere(_.name == sheet.name) match
      case -1 =>
        // Sheet doesn't exist → add at end and mark metadata as modified (workbook.xml changes)
        val updatedContext = sourceContext.map(_.markMetadataModified)
        copy(sheets = sheets :+ sheet, sourceContext = updatedContext)
      case index =>
        // Sheet exists → replace in-place and track modification
        val updatedContext = sourceContext.map(_.markSheetModified(index))
        copy(sheets = sheets.updated(index, sheet), sourceContext = updatedContext)

  /**
   * Put sheet with explicit name (rename if needed, then add-or-replace).
   *
   * Useful for renaming sheets during updates.
   */
  def put(name: SheetName, sheet: Sheet): Workbook =
    put(sheet.copy(name = name))

  /**
   * Put sheet with explicit name (string variant).
   *
   * When called with a string literal, the name format is validated at compile time and returns
   * Workbook directly. Invalid literals like "Invalid:Name" fail to compile. Runtime strings return
   * XLResult[Workbook].
   */
  @annotation.targetName("putWithStringName")
  transparent inline def put(inline name: String, sheet: Sheet): Workbook | XLResult[Workbook] =
    ${ com.tjclp.xl.macros.WorkbookMacros.putStringNameImpl('{ this }, 'name, 'sheet) }

  /**
   * Put multiple sheets (batch operation).
   *
   * Adds or replaces each sheet by name. This operation is infallible since each individual put is
   * infallible.
   *
   * Example:
   * {{{
   * val updated = wb.put(Sheet("Sales"), Sheet("Marketing"), Sheet("Finance"))
   * }}}
   */
  def put(firstSheet: Sheet, restSheets: Sheet*): Workbook =
    (firstSheet +: restSheets).foldLeft(this) { (acc, sheet) =>
      acc.put(sheet)
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

  /**
   * Remove sheet by name (string variant).
   *
   * When called with a string literal, the name format is validated at compile time. Invalid
   * literals fail to compile. Always returns XLResult since sheet existence is runtime-dependent.
   */
  @annotation.targetName("removeByString")
  transparent inline def remove(inline name: String): XLResult[Workbook] =
    ${ com.tjclp.xl.macros.WorkbookMacros.removeImpl('{ this }, 'name) }

  /** Remove sheet by index (Excel requires at least one sheet to remain). */
  def removeAt(index: Int): XLResult[Workbook] =
    // Validation: Excel workbooks must have at least one sheet
    if sheets.size <= 1 then Left(XLError.InvalidWorkbook("Cannot remove last sheet"))
    else if index >= 0 && index < sheets.size then
      val newSheets = sheets.patch(index, Nil, 1)
      val newActiveIndex =
        if activeSheetIndex >= newSheets.size then newSheets.size - 1 else activeSheetIndex
      val updatedContext = sourceContext.map(_.markSheetDeleted(index))
      Right(
        copy(sheets = newSheets, activeSheetIndex = newActiveIndex, sourceContext = updatedContext)
      )
    else Left(XLError.OutOfBounds(s"sheet[$index]", s"Valid range: 0 to ${sheets.size - 1}"))

  /** Rename sheet (marks metadata as modified since sheet names live in workbook.xml). */
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
    sheets.indexWhere(_.name == name) match
      case -1 => Left(XLError.SheetNotFound(name.value))
      case idx => updateAt(idx, f)

  /**
   * Update sheet by applying a function (string variant).
   *
   * When called with a string literal, the name format is validated at compile time. Invalid
   * literals fail to compile. Always returns XLResult since sheet existence is runtime-dependent.
   */
  @annotation.targetName("updateByString")
  transparent inline def update(inline name: String, f: Sheet => Sheet): XLResult[Workbook] =
    ${ com.tjclp.xl.macros.WorkbookMacros.updateImpl('{ this }, 'name, 'f) }

  /** Update sheet by index while tracking modification state. */
  def updateAt(idx: Int, f: Sheet => Sheet): XLResult[Workbook] =
    if idx < 0 || idx >= sheets.size then
      Left(XLError.OutOfBounds(s"sheet[$idx]", s"Valid range: 0 to ${sheets.size - 1}"))
    else
      val updatedSheet = f(sheets(idx))
      val newSheets = sheets.updated(idx, updatedSheet)
      val updatedContext = sourceContext.map(_.markSheetModified(idx))
      Right(copy(sheets = newSheets, sourceContext = updatedContext))

  /** Delete sheet by name while tracking modification state. */
  def delete(name: SheetName): XLResult[Workbook] =
    sheets.indexWhere(_.name == name) match
      case -1 => Left(XLError.SheetNotFound(name.value))
      case idx => removeAt(idx)

  /**
   * Delete sheet by name (string variant).
   *
   * When called with a string literal, the name format is validated at compile time. Invalid
   * literals fail to compile. Always returns XLResult since sheet existence is runtime-dependent.
   */
  @annotation.targetName("deleteByString")
  transparent inline def delete(inline name: String): XLResult[Workbook] =
    ${ com.tjclp.xl.macros.WorkbookMacros.deleteImpl('{ this }, 'name) }

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
      // Sheet reordering only modifies workbook.xml (order metadata), not individual sheet files.
      // Therefore we mark reordered but don't mark individual sheets as modified.
      val updatedContext = sourceContext.map(_.markReordered)
      Right(
        copy(sheets = reordered, activeSheetIndex = newActiveIndex, sourceContext = updatedContext)
      )

  /** Insert sheet at specific index (explicit positioning - rarely needed) */
  def insertAt(index: Int, sheet: Sheet): XLResult[Workbook] =
    if index < 0 || index > sheets.size then
      Left(XLError.OutOfBounds(s"insert[$index]", s"Valid range: 0 to ${sheets.size}"))
    else if sheets.exists(_.name == sheet.name) then Left(XLError.DuplicateSheet(sheet.name.value))
    else
      val (before, after) = sheets.splitAt(index)
      val updatedContext = sourceContext.map(_.markMetadataModified)
      Right(copy(sheets = before ++ (sheet +: after), sourceContext = updatedContext))

  // ========== Deprecated Methods (Removed in v0.3.0) ==========

  @deprecated("Use put(sheet) instead (add-or-replace semantic)", "0.2.0")
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
  /**
   * Create workbook from a single sheet.
   *
   * Example:
   * {{{
   * val sheet = Sheet("Sales").put(ref"A1" -> "Revenue").unsafe
   * val workbook = Workbook(sheet)
   * }}}
   */
  def apply(sheet: Sheet): Workbook = Workbook(Vector(sheet))

  /**
   * Create workbook from multiple sheets.
   *
   * Example:
   * {{{
   * val workbook = Workbook(salesSheet, marketingSheet, financeSheet)
   * }}}
   */
  def apply(first: Sheet, second: Sheet, rest: Sheet*): Workbook =
    Workbook((first +: second +: rest).toVector)

  /**
   * Create workbook with a single empty sheet, with compile-time or runtime validation.
   *
   * Uses `transparent inline` to enable '''type narrowing''' based on the argument:
   *   - '''String literal''' ("Sales"): Validated at compile time, returns `Workbook` directly
   *   - '''Runtime expression''' (variable): Validated at runtime, returns `XLResult[Workbook]`
   *
   * The union return type `Workbook | XLResult[Workbook]` allows the compiler to narrow to the
   * appropriate type at each call site, providing both type safety and ergonomics.
   *
   * @example
   *   {{{
   *   // Literal string → compile-time validation → Workbook
   *   val wb1: Workbook = Workbook("Sales")
   *
   *   // Runtime string → runtime validation → XLResult[Workbook]
   *   val name = getUserInput()
   *   val wb2: XLResult[Workbook] = Workbook(name)
   *   wb2 match
   *     case Right(wb) => // valid
   *     case Left(err) => // invalid sheet name
   *   }}}
   */
  transparent inline def apply(inline sheetName: String): Workbook | XLResult[Workbook] =
    ${ com.tjclp.xl.macros.WorkbookMacros.createSingleImpl('sheetName) }

  /**
   * Create workbook with multiple empty sheets, with compile-time or runtime validation.
   *
   * Uses `transparent inline` to enable '''type narrowing''' based on all arguments:
   *   - '''All string literals''': Validated at compile time, returns `Workbook` directly
   *   - '''Any runtime expression''': Validated at runtime, returns `XLResult[Workbook]`
   *
   * @example
   *   {{{
   *   // All literals → Workbook
   *   val wb1: Workbook = Workbook("Sales", "Marketing", "Finance")
   *
   *   // Any runtime value → XLResult[Workbook]
   *   val dept = getDepartment()
   *   val wb2: XLResult[Workbook] = Workbook("Sales", dept)
   *   }}}
   */
  transparent inline def apply(
    inline first: String,
    inline second: String,
    inline rest: String*
  ): Workbook | XLResult[Workbook] =
    ${ com.tjclp.xl.macros.WorkbookMacros.createMultiImpl('first, 'second, 'rest) }

  /** Create empty workbook with a single sheet named "Sheet1" */
  def empty: Workbook =
    Workbook(Vector(Sheet(SheetName.unsafe("Sheet1"))))
