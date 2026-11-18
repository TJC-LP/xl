package com.tjclp.xl.formula

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.formula.eval.{DependentRecalculation, RecalcResult}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook

/**
 * #891: the after-edit pass (`DependentRecalculation.recalculateAfterEdit`) in the shape the
 * removed `recalculateDependents` extensions had, so a spec reads `sheet.afterEdit(Set(ref))`.
 */
object AfterEditSyntax:

  extension (sheet: Sheet)
    /** The sheet after recalculating the cone of an edit to `refs` on it (a one-sheet book). */
    def afterEdit(refs: Set[ARef], clock: Clock = Clock.system): Sheet =
      Workbook(sheet).afterEditResult(sheet.name, refs, clock).workbook(sheet.name).getOrElse(sheet)

  extension (wb: Workbook)
    /** The workbook after recalculating the cone of an edit to `refs` on `sheet`. */
    def afterEdit(sheet: SheetName, refs: Set[ARef]): Workbook =
      wb.afterEditResult(sheet, refs, Clock.system).workbook

    /** The whole result, errors included. */
    def afterEditResult(sheet: SheetName, refs: Set[ARef], clock: Clock): RecalcResult =
      DependentRecalculation.recalculateAfterEdit(wb, sheet, refs, clock)
