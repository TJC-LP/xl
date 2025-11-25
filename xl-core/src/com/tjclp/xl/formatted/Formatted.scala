package com.tjclp.xl.formatted

import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.styles.numfmt.NumFmt

/**
 * Formatted cell value with associated number format.
 *
 * Used by formatted literal macros (money"", percent"", date"", etc.) to preserve both the parsed
 * value and the intended display format.
 *
 * The unified Sheet.put() method handles Formatted values directly:
 * {{{
 * sheet.put(ref"A1", money"$$1,234.56")  // Preserves Currency format
 * sheet.put(ref"B1", percent"15%")       // Preserves Percent format
 * }}}
 */
final case class Formatted(value: CellValue, numFmt: NumFmt):
  /** Extract just the value */
  def toCellValue: CellValue = value

  /** Extract just the format */
  def toNumFmt: NumFmt = numFmt

object Formatted:
  /** Create from value and format */
  def apply(value: CellValue, numFmt: NumFmt): Formatted =
    new Formatted(value, numFmt)

  /** Automatic conversion to CellValue (loses format info - prefer using put() directly) */
  given Conversion[Formatted, CellValue] = _.value
