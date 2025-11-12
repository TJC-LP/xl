package com.tjclp.xl

import com.tjclp.xl.cell.CellValue
import java.time.LocalDateTime
import scala.language.implicitConversions

/**
 * Automatic conversions from common Scala types to CellValue.
 *
 * Import with: `import com.tjclp.xl.conversions.given`
 *
 * Enables clean syntax:
 * {{{
 * import com.tjclp.xl.conversions.given
 *
 * sheet.put(cell"A1", "Hello")      // String → CellValue.Text
 *   .put(cell"B1", 42)              // Int → CellValue.Number
 *   .put(cell"C1", 3.14)            // Double → CellValue.Number
 *   .put(cell"D1", true)            // Boolean → CellValue.Bool
 * }}}
 *
 * All conversions are inline for zero runtime overhead.
 */

object conversions:
  /** Convert String to CellValue.Text */
  given Conversion[String, CellValue] = CellValue.Text(_)

  /** Convert Int to CellValue.Number */
  given Conversion[Int, CellValue] = i => CellValue.Number(BigDecimal(i))

  /** Convert Long to CellValue.Number */
  given Conversion[Long, CellValue] = l => CellValue.Number(BigDecimal(l))

  /** Convert Double to CellValue.Number */
  given Conversion[Double, CellValue] = d => CellValue.Number(BigDecimal(d))

  /** Convert BigDecimal to CellValue.Number */
  given Conversion[BigDecimal, CellValue] = CellValue.Number(_)

  /** Convert Boolean to CellValue.Bool */
  given Conversion[Boolean, CellValue] = CellValue.Bool(_)

  /** Convert LocalDateTime to CellValue.DateTime */
  given Conversion[LocalDateTime, CellValue] = CellValue.DateTime(_)

export conversions.*
