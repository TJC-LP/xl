package com.tjclp.xl.formula.eval

import com.tjclp.xl.addressing.CellRange
import com.tjclp.xl.formula.Clock
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook

/**
 * The live values of a picture of a window (`view --eval` in html, svg and the raster formats). The
 * window's formulas are evaluated together with every formula the window's conditional formatting
 * reads, plus the precedents of both. That includes a colour scale's, data bar's or top-N rule's
 * whole block, and the cells a formula rule reads at each window cell. So a cell's paint comes from
 * the values the picture draws, the same whatever window is asked for, never from a mix of live
 * cells and the stale caches outside the window.
 */
private[xl] object LiveRender:

  /**
   * [[SheetEvaluator.evaluateForRangePerCell]] over `window` and what the conditional formatting of
   * `sheet` over `window` reads. The values cover the formulas of both. The failures and blocked
   * cells are those of their joint closure, since the picture shows all of them, in its cells or in
   * its paint.
   */
  def evaluate(
    sheet: Sheet,
    window: CellRange,
    clock: Clock,
    workbook: Option[Workbook]
  ): RangeEvalResult =
    val reads =
      if sheet.conditionalFormats.isEmpty then Vector.empty else CfEvaluator.reads(sheet, window)
    SheetEvaluator.evaluateForRangesPerCell(sheet, window +: reads, clock, workbook)
