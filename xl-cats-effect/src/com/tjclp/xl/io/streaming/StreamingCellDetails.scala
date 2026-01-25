package com.tjclp.xl.io.streaming

import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.{CellValue, Comment}
import com.tjclp.xl.styles.CellStyle

/**
 * Result of streaming cell details query.
 *
 * Contains all available information about a cell that can be extracted in O(1) worksheet memory.
 * Pre-loaded components (styles, shared strings, comments) are loaded once from their respective
 * XML parts, then the worksheet is streamed until the target cell is found.
 *
 * @param ref
 *   Cell reference (e.g., A1, B5)
 * @param value
 *   Cell value (may be formula with cached value)
 * @param style
 *   Optional resolved CellStyle (from styles.xml)
 * @param comment
 *   Optional comment (from comments{N}.xml)
 * @param dependencies
 *   Parsed formula dependencies (from formula text, if present)
 * @param dependentsUnavailable
 *   True because computing dependents would require full workbook scan
 */
final case class StreamingCellDetails(
  ref: ARef,
  value: CellValue,
  style: Option[CellStyle],
  comment: Option[Comment],
  dependencies: Vector[String],
  dependentsUnavailable: Boolean = true
)
