package com.tjclp.xl.codecs

import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.styles.CellStyle

/**
 * Write a typed value to produce cell data with optional styles hint.
 *
 * Returns (CellValue, Optional[CellStyle]) where the styles is auto-inferred based on the value
 * type (e.g., DateTime gets date format, BigDecimal gets decimal format).
 */
trait CellWriter[A]:
  def write(a: A): (CellValue, Option[CellStyle])
