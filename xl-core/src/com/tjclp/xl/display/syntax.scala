package com.tjclp.xl.display

import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.sheets.Sheet

import scala.language.implicitConversions

/**
 * Extension methods for explicit display formatting.
 *
 * Provides both automatic (via implicit conversion) and explicit display methods:
 * {{{
 * import com.tjclp.xl.display.{*, given}
 * given Sheet = mySheet
 *
 * // Automatic (via implicit conversion)
 * println(s"Revenue: \${ref"A1"}")
 *
 * // Explicit (via extension method)
 * println(s"Revenue: \${mySheet.display(ref"A1")}")
 *
 * // Force raw formula display
 * println(s"Formula: \${mySheet.displayFormula(ref"A1")}")
 * }}}
 *
 * @since 0.2.0
 */
object syntax:

  extension (sheet: Sheet)
    /**
     * Get formatted display value for a cell reference.
     *
     * This method explicitly displays a cell with formatting applied. The formula display behavior
     * depends on the active FormulaDisplayStrategy.
     *
     * @param ref
     *   The cell reference
     * @param fds
     *   The formula display strategy (implicit, defaults to raw text)
     * @return
     *   DisplayWrapper with formatted string
     */
    def display(ref: ARef)(using fds: FormulaDisplayStrategy): DisplayWrapper =
      // Use the implicit conversion - sheet is provided by extension context
      import DisplayConversions.given
      given Sheet = sheet
      val conv = summon[Conversion[ARef, DisplayWrapper]]
      conv.apply(ref)

    /**
     * Get raw formula text for a cell reference.
     *
     * This always shows the formula expression (e.g., "=SUM(A1:A10)"), even if an evaluating
     * FormulaDisplayStrategy is active. Useful for debugging or documentation.
     *
     * @param ref
     *   The cell reference
     * @return
     *   Formula string with leading =, or formatted value if not a formula
     */
    def displayFormula(ref: ARef)(using fds: FormulaDisplayStrategy): String =
      val cell = sheet.cells.get(ref).getOrElse(Cell.empty(ref))
      cell.value match
        case CellValue.Formula(expr) =>
          // Expression already includes "=" prefix
          if expr.startsWith("=") then expr else s"=$expr"
        case _ => display(ref).formatted
