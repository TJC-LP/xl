package com.tjclp.xl.formula.eval

import com.tjclp.xl.workbooks.Workbook
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.parser.FormulaParser
import com.tjclp.xl.formula.printer.{FormulaPrinter, FormulaShifter}

/**
 * GH-128 / GH-129: workbook-level structural editing (insert/delete rows & columns) WITH formula
 * rewriting.
 *
 * The pure cell/merge/property shift lives in `xl-core` (`Sheet.insertRows`, ...). This layer adds
 * what xl-core cannot do (it has no formula parser): after shifting cells on the edited sheet, it
 * rewrites the formula strings of EVERY sheet so references track the edit. A formula that
 * references a fully-deleted cell/range becomes `#REF!` (`CellValue.Error(Ref)`); partially
 * overlapped ranges shrink. Cross-sheet references to the edited sheet are rewritten too.
 *
 * Determinism: the edit is a pure `Workbook => Workbook`; re-running on identical input yields
 * identical output.
 */
object StructuralEditor:

  /** Insert `count` rows at 0-based row index `at` on `sheet`. */
  def insertRows(wb: Workbook, sheet: SheetName, at: Int, count: Int): Workbook =
    edit(wb, sheet, isRow = true, at = at, delta = count)

  /** Delete `count` rows starting at 0-based row index `at` on `sheet`. */
  def deleteRows(wb: Workbook, sheet: SheetName, at: Int, count: Int): Workbook =
    edit(wb, sheet, isRow = true, at = at, delta = -count)

  /** Insert `count` columns at 0-based column index `at` on `sheet`. */
  def insertColumns(wb: Workbook, sheet: SheetName, at: Int, count: Int): Workbook =
    edit(wb, sheet, isRow = false, at = at, delta = count)

  /** Delete `count` columns starting at 0-based column index `at` on `sheet`. */
  def deleteColumns(wb: Workbook, sheet: SheetName, at: Int, count: Int): Workbook =
    edit(wb, sheet, isRow = false, at = at, delta = -count)

  private def edit(
    wb: Workbook,
    target: SheetName,
    isRow: Boolean,
    at: Int,
    delta: Int
  ): Workbook =
    val editedName = target.value
    val updatedSheets = wb.sheets.map { s =>
      // 1. Pure cell/merge/property shift — only on the edited sheet.
      val shifted =
        if s.name == target then
          (isRow, delta >= 0) match
            case (true, true) => s.insertRows(at, delta)
            case (true, false) => s.deleteRows(at, -delta)
            case (false, true) => s.insertColumns(at, delta)
            case (false, false) => s.deleteColumns(at, -delta)
        else s
      // 2. Rewrite formula references on every sheet (local refs only on the edited sheet;
      //    sheet-qualified refs to the edited sheet on all sheets).
      rewriteFormulas(shifted, shiftLocal = s.name == target, editedName, isRow, at, delta)
    }
    wb.copy(sheets = updatedSheets)

  private def rewriteFormulas(
    sheet: Sheet,
    shiftLocal: Boolean,
    editedSheet: String,
    isRow: Boolean,
    at: Int,
    delta: Int
  ): Sheet =
    val updatedCells = sheet.cells.map { case (ref, cell) =>
      cell.value match
        case CellValue.Formula(formulaStr, _) =>
          FormulaParser.parse(formulaStr) match
            case Right(expr) =>
              FormulaShifter.shiftStructural(expr, shiftLocal, editedSheet, isRow, at, delta) match
                case Some(shiftedExpr) =>
                  val newStr = FormulaPrinter.print(shiftedExpr, includeEquals = true)
                  (ref, cell.copy(value = CellValue.Formula(newStr, None)))
                case None =>
                  (ref, cell.copy(value = CellValue.Error(CellError.Ref)))
            // Unparseable formula: leave untouched rather than guess.
            case Left(_) => (ref, cell)
        case _ => (ref, cell)
    }
    sheet.copy(cells = updatedCells)

  /** Ergonomic workbook extensions. */
  extension (wb: Workbook)
    def insertRowsShifted(sheet: SheetName, at: Int, count: Int): Workbook =
      StructuralEditor.insertRows(wb, sheet, at, count)
    def deleteRowsShifted(sheet: SheetName, at: Int, count: Int): Workbook =
      StructuralEditor.deleteRows(wb, sheet, at, count)
    def insertColumnsShifted(sheet: SheetName, at: Int, count: Int): Workbook =
      StructuralEditor.insertColumns(wb, sheet, at, count)
    def deleteColumnsShifted(sheet: SheetName, at: Int, count: Int): Workbook =
      StructuralEditor.deleteColumns(wb, sheet, at, count)
