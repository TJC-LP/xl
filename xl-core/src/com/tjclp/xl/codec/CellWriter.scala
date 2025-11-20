package com.tjclp.xl.codec

import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.styles.CellStyle

/**
 * Write a typed value to produce cell data with optional style hint.
 *
 * Returns (CellValue, Optional[CellStyle]) where the style is auto-inferred based on the value type
 * (e.g., DateTime gets date format, BigDecimal gets decimal format).
 *
 * Used by Easy Mode extensions for generic put() operations with automatic NumFmt inference.
 */
trait CellWriter[A]:
  def write(a: A): (CellValue, Option[CellStyle])

object CellWriter:
  /** Summon the writer instance for type A */
  def apply[A](using w: CellWriter[A]): CellWriter[A] = w
