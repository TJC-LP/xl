package com.tjclp.xl.display

import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.Cell
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.numfmt.NumFmt

import scala.language.implicitConversions

/**
 * Implicit conversions for automatic display formatting in string interpolation.
 *
 * Enables ergonomic syntax:
 * {{{
 * import com.tjclp.xl.display.{*, given}
 * given Sheet = mySheet
 * println(s"Revenue: \${ref"A1"}")  // Automatic formatted display!
 * }}}
 *
 * Conversions require a given Sheet in scope for context. Formula display behavior depends on the
 * active FormulaDisplayStrategy (raw text in xl-core, evaluation in xl-evaluator).
 *
 * @since 0.2.0
 */
object DisplayConversions:

  /**
   * Implicit conversion: ARef → DisplayWrapper.
   *
   * Converts a cell reference to a display wrapper with formatted toString. Requires given Sheet
   * and FormulaDisplayStrategy in scope.
   *
   * @param sheet
   *   The sheet context
   * @param fds
   *   The formula display strategy (defaults to raw formula text)
   */
  given arefToDisplay(using
    sheet: Sheet,
    fds: FormulaDisplayStrategy
  ): Conversion[ARef, DisplayWrapper] with
    def apply(ref: ARef): DisplayWrapper =
      val cell = sheet.cells.get(ref).getOrElse(Cell.empty(ref))
      formatCell(cell, sheet, fds)

  /**
   * Implicit conversion: Cell → DisplayWrapper.
   *
   * Converts a cell to a display wrapper with formatted toString. Requires given Sheet and
   * FormulaDisplayStrategy in scope.
   *
   * @param sheet
   *   The sheet context (for style registry and formula evaluation)
   * @param fds
   *   The formula display strategy
   */
  given cellToDisplay(using
    sheet: Sheet,
    fds: FormulaDisplayStrategy
  ): Conversion[Cell, DisplayWrapper] with
    def apply(cell: Cell): DisplayWrapper =
      formatCell(cell, sheet, fds)

  /**
   * Format a cell for display.
   *
   * @param cell
   *   The cell to format
   * @param sheet
   *   The sheet context
   * @param fds
   *   The formula display strategy
   * @return
   *   DisplayWrapper with formatted string
   */
  private def formatCell(
    cell: Cell,
    sheet: Sheet,
    fds: FormulaDisplayStrategy
  ): DisplayWrapper =
    import com.tjclp.xl.cells.CellValue

    // Get the cell's number format from its style
    val numFmt = cell.styleId
      .flatMap(sheet.styleRegistry.get)
      .map(_.numFmt)
      .getOrElse(NumFmt.General)

    // Format based on value type
    val formatted = cell.value match
      case CellValue.Formula(expr) =>
        // Use formula display strategy (raw text or evaluation)
        fds.format(expr, sheet)

      case other =>
        // Use NumFmt formatter for non-formula values
        NumFmtFormatter.formatValue(other, numFmt)

    DisplayWrapper(formatted)
