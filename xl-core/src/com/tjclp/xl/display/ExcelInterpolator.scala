package com.tjclp.xl.display

import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.Cell
import com.tjclp.xl.sheets.Sheet

import scala.language.implicitConversions

/**
 * Custom string interpolator for automatic Excel-formatted display.
 *
 * The `excel` interpolator enables automatic formatted display of cell references and values:
 * {{{
 * import com.tjclp.xl.display.{*, given}
 * import com.tjclp.xl.display.ExcelInterpolator.*
 * given Sheet = mySheet
 *
 * println(excel"Revenue: \${ref"A1"}")        // "Revenue: $1,000,000"
 * println(excel"Margin: \${ref"B19"}")        // "Margin: 60.0%"
 * }}}
 *
 * Unlike standard `s"..."` interpolation, the `excel` interpolator applies implicit conversions to
 * ARef and Cell values, formatting them according to their NumFmt styles.
 *
 * @since 0.2.0
 */
object ExcelInterpolator:

  extension (sc: StringContext)
    /**
     * Excel string interpolator with automatic formatted display.
     *
     * Formats interpolated values according to Excel conventions:
     *   - ARef: Looks up cell and formats according to its NumFmt
     *   - Cell: Formats according to sheet's style registry
     *   - Other: Uses toString
     *
     * @param args
     *   Interpolated values
     * @param sheet
     *   The sheet context (implicit)
     * @param fds
     *   The formula display strategy (implicit, determines evaluation behavior)
     * @return
     *   Formatted string with Excel-style display values
     */
    @SuppressWarnings(Array("org.wartremover.warts.While"))
    def excel(args: Any*)(using sheet: Sheet, fds: FormulaDisplayStrategy): String =
      import DisplayConversions.given // Import conversions in interpolator scope

      val parts = sc.parts.iterator
      val values = args.iterator
      val result = new StringBuilder()

      while parts.hasNext do
        result.append(parts.next())
        if values.hasNext then
          values.next() match
            case ref: ARef =>
              // Apply ARef → DisplayWrapper conversion
              val conv = summon[Conversion[ARef, DisplayWrapper]]
              result.append(conv.apply(ref).formatted)

            case cell: Cell =>
              // Apply Cell → DisplayWrapper conversion
              val conv = summon[Conversion[Cell, DisplayWrapper]]
              result.append(conv.apply(cell).formatted)

            case other =>
              // Non-Excel values: use default toString
              result.append(other.toString)

      result.toString
