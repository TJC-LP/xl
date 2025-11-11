package com.tjclp.xl

import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cell.CellValue
import com.tjclp.xl.style.NumFmt

/**
 * Formatted cell value with associated number format.
 *
 * Used by formatted literal macros (money"", percent"", date"", etc.) to preserve both the parsed
 * value and the intended display format.
 */
case class Formatted(value: CellValue, numFmt: NumFmt):
  /** Extract just the value */
  def toCellValue: CellValue = value

  /** Extract just the format */
  def toNumFmt: NumFmt = numFmt

object Formatted:
  /** Create from value and format */
  def apply(value: CellValue, numFmt: NumFmt): Formatted =
    new Formatted(value, numFmt)

  /** Automatic conversion to CellValue (loses format info) */
  given Conversion[Formatted, CellValue] = _.value

  /** Extension methods for Sheet to handle Formatted values */
  extension (sheet: Sheet)
    /** Put formatted value (uses format for future style integration) */
    @annotation.targetName("putFormattedExt")
    def putFormatted(ref: ARef, formatted: Formatted): Sheet =
      // For now, just put the value
      // Future: integrate with style system to apply numFmt
      sheet.put(ref, formatted.value)

    /** Put multiple formatted values */
    @annotation.targetName("putFormattedBatchExt")
    def putFormatted(pairs: (ARef, Formatted)*): Sheet =
      pairs.foldLeft(sheet) { (s, pair) =>
        s.putFormatted(pair._1, pair._2)
      }
