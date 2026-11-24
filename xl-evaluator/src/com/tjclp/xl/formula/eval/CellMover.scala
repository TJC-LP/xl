package com.tjclp.xl.formula.eval

import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cells.{CellValue, FormulaKind}
import com.tjclp.xl.error.XLResult
import com.tjclp.xl.formula.printer.FormulaOps
import com.tjclp.xl.sheets.{Sheet, SheetEdits}
import com.tjclp.xl.workbooks.Workbook

/**
 * GH-935: Excel's Merge Cells with the reference half xl-core cannot do. When the upper-left cell
 * of the merged range is empty, Excel moves the first non-empty cell of the range (row-major) into
 * it as a cut and paste moves a cell: the value or formula verbatim, its format, and every formula
 * that read the moved cell now reads the upper-left cell (checked against Excel for Mac: with `E1 =
 * B1+1`, merging `A1:B1` over an empty A1 rewrites E1 to `=A1+1`).
 *
 * The rewrite reaches single-cell references and a range that is exactly that one cell, on every
 * sheet (local references on the moved cell's sheet, qualified ones anywhere), in defined names and
 * in typed conditional-format and data-validation formulas ([[FormulaOps.relocateCell]]). Text that
 * does not parse is left as it is. Changed sheets go back through `Workbook.put` (ADR-017 invariant
 * 1).
 */
object CellMover:

  /**
   * Point every reference to `sheet!from` at `sheet!to`, workbook-wide. The cell itself is not
   * moved. Total; the workbook is returned unchanged when nothing reads `from`.
   */
  def relocateReferences(wb: Workbook, sheet: SheetName, from: ARef, to: ARef): Workbook =
    if from == to then wb
    else
      val canonical = wb(sheet).fold(_ => sheet, _.name)
      val withSheets = wb.sheets.foldLeft(wb) { (acc, s) =>
        val rewritten = relocateIn(s, Some(s.name), canonical, from, to)
        if rewritten eq s then acc else acc.put(rewritten)
      }
      val names = withSheets.metadata.definedNames.map { dn =>
        val text = StructuralEditor
          .splitTopLevelCommas(dn.formula)
          .map(FormulaOps.relocateCell(_, None, canonical, from, to))
          .mkString(",")
        if text == dn.formula then dn else dn.copy(formula = text)
      }
      if names == withSheets.metadata.definedNames then withSheets
      else
        withSheets.copy(
          metadata = withSheets.metadata.copy(definedNames = names),
          sourceContext = withSheets.sourceContext.map(_.markMetadataModified)
        )

  /** [[relocateReferences]] limited to the formulas on `sheet` itself (its cells, CF and DV). */
  def relocateOnSheet(sheet: Sheet, from: ARef, to: ARef): Sheet =
    relocateIn(sheet, Some(sheet.name), sheet.name, from, to)

  /**
   * Merge `range` on `sheetName` as Excel does: [[relocateReferences]] for the cell an empty
   * upper-left cell receives, then [[DynamicArrayAuthoring.merge]] (move, release hidden dynamic
   * arrays, clear the rest). `Left` only when the sheet does not exist.
   */
  def merge(wb: Workbook, sheetName: SheetName, range: CellRange): XLResult[Workbook] =
    wb(sheetName).flatMap { sheet =>
      val relocated = SheetEdits
        .mergeWinner(sheet, range)
        .fold(wb)(w =>
          relocateReferences(wb, sheet.name, w.ref, SheetEdits.mergeAnchor(sheet, range))
        )
      relocated(sheet.name).map(s => relocated.put(DynamicArrayAuthoring.merge(s, range)))
    }

  /**
   * The sheet-level merge: as [[merge]], but only the moved cell's readers on this sheet are
   * rewritten (a sheet does not see the workbook's other sheets or its defined names).
   */
  def mergeSheet(sheet: Sheet, range: CellRange): Sheet =
    val relocated = SheetEdits
      .mergeWinner(sheet, range)
      .fold(sheet)(w => relocateOnSheet(sheet, w.ref, SheetEdits.mergeAnchor(sheet, range)))
    DynamicArrayAuthoring.merge(relocated, range)

  /** The sheet with its readers of `target!from` rewritten; the same instance when none changed. */
  private def relocateIn(
    sheet: Sheet,
    host: Option[SheetName],
    target: SheetName,
    from: ARef,
    to: ARef
  ): Sheet =
    def rewrite(text: String): String = FormulaOps.relocateCell(text, host, target, from, to)
    val changed = sheet.cells.iterator.flatMap { cell =>
      cell.value match
        // GH-430: a data-table record's TABLE(...) text is derived from local geometry
        case CellValue.Formula(_, _, _: FormulaKind.DataTable) => None
        case CellValue.Formula(text, cached, kind) =>
          val next = rewrite(text)
          Option.when(next != text)(cell.copy(value = CellValue.Formula(next, cached, kind)))
        case _ => None
    }.toVector
    val cfs = sheet.conditionalFormats.map(cf =>
      SheetRenamer.rewriteCf(cf, t => Right(rewrite(t))).fold(_ => cf, identity)
    )
    val dvs = sheet.dataValidations.map(dv =>
      SheetRenamer.rewriteDv(dv, t => Right(rewrite(t))).fold(_ => dv, identity)
    )
    val cfChanged = cfs != sheet.conditionalFormats
    val dvChanged = dvs != sheet.dataValidations
    if changed.isEmpty && !cfChanged && !dvChanged then sheet
    else
      sheet.copy(
        cells = sheet.cells.putAll(changed),
        conditionalFormats = if cfChanged then cfs else sheet.conditionalFormats,
        dataValidations = if dvChanged then dvs else sheet.dataValidations
      )
