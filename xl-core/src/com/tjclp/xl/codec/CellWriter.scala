package com.tjclp.xl.codec

import com.tjclp.xl.cell.CellValue
import com.tjclp.xl.style.CellStyle

/**
 * Write a typed value to produce cell data with optional style hint.
 *
 * Returns (CellValue, Optional[CellStyle]) where the style is auto-inferred based on the value type
 * (e.g., DateTime gets date format, BigDecimal gets decimal format).
 */
trait CellWriter[A]:
  def write(a: A): (CellValue, Option[CellStyle])
