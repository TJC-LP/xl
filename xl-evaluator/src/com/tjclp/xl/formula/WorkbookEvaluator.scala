package com.tjclp.xl.formula

import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.workbooks.Workbook

// Import SheetEvaluator extension methods
import SheetEvaluator.*

/**
 * Extension methods for evaluating and caching formula values in workbooks.
 *
 * Usage:
 * {{{
 * import com.tjclp.xl.{*, given}
 *
 * val wb = Workbook(sheets).withCachedFormulas()
 * Excel.write(wb, "output.xlsx")
 * }}}
 */
object WorkbookEvaluator:

  extension (wb: Workbook)

    /**
     * Evaluate all formulas and cache their computed values.
     *
     * Useful for:
     *   - CLI tools that need to display computed values
     *   - Files opened by tools other than Excel
     *   - Faster Excel open times (no recalculation needed)
     *
     * Formulas that fail to evaluate (unsupported functions, circular refs) are left without cached
     * values - Excel will recalculate on open.
     *
     * @param clock
     *   Clock for volatile functions (TODAY, NOW). Defaults to system clock.
     * @return
     *   Workbook with formula cells containing cached values
     */
    def withCachedFormulas(clock: Clock = Clock.system): Workbook =
      val updatedSheets = wb.sheets.map { sheet =>
        sheet.evaluateWithDependencyCheck(clock) match
          case Right(results) =>
            // Update formula cells with cached values
            results.foldLeft(sheet) { case (s, (ref, computedValue)) =>
              s.cells.get(ref) match
                case Some(cell) =>
                  cell.value match
                    case CellValue.Formula(expr, _) =>
                      s.put(ref, CellValue.Formula(expr, Some(computedValue)))
                    case _ => s
                case None => s
            }
          case Left(_) =>
            // Circular reference or evaluation error - leave formulas uncached
            // Excel will recalculate on open
            sheet
      }
      wb.copy(sheets = updatedSheets)
