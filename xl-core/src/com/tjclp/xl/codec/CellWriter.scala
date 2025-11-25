package com.tjclp.xl.codec

import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formatted.Formatted
import com.tjclp.xl.richtext.{RichText, TextRun}
import com.tjclp.xl.styles.CellStyle

import java.time.{LocalDate, LocalDateTime}

/**
 * Closed union of all types that can be written to cells.
 *
 * This union enables compile-time type safety for `Sheet.put` methods while preserving ergonomic
 * heterogeneous batch operations like `sheet.put(ref"A1" -> "hello", ref"B1" -> 42)`.
 *
 * When the compiler infers the type parameter for `put[A: CellWriter]`, it computes the union of
 * all argument types (e.g., `String | Int`). Since `CellWriter` is contravariant, the instance for
 * `CellWritable` (the full union) satisfies any sub-union.
 *
 * Note: `TextRun` is included to support the RichText DSL where `"text".bold` produces a TextRun.
 */
type CellWritable = String | Int | Long | Double | BigDecimal | Boolean | LocalDate |
  LocalDateTime | RichText | TextRun | CellValue | Formatted

/**
 * Write a typed value to produce cell data with optional style hint.
 *
 * Returns (CellValue, Optional[CellStyle]) where the style is auto-inferred based on the value type
 * (e.g., DateTime gets date format, BigDecimal gets decimal format).
 *
 * Used by Easy Mode extensions for generic put() operations with automatic NumFmt inference.
 *
 * '''Contravariance''': The `-A` variance annotation enables the master `CellWriter[CellWritable]`
 * instance to satisfy any sub-union (e.g., `CellWriter[String | Int]`).
 */
trait CellWriter[-A]:
  def write(a: A): (CellValue, Option[CellStyle])

object CellWriter:
  /** Summon the writer instance for type A */
  def apply[A](using w: CellWriter[A]): CellWriter[A] = w
