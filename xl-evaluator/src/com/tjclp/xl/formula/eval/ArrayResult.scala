package com.tjclp.xl.formula.eval

import com.tjclp.xl.cells.CellValue

/**
 * Result from array-returning functions like TRANSPOSE, SEQUENCE.
 *
 * Represents a 2D grid of cell values that can be "spilled" into adjacent cells when evaluated as
 * an array formula.
 *
 * Design principles:
 *   - Immutable value type
 *   - Row-major ordering (values(row)(col))
 *   - Empty arrays have size 0x0
 *
 * Example:
 * {{{
 * // 2x3 array (2 rows, 3 columns)
 * val arr = ArrayResult(Vector(
 *   Vector(CellValue.Number(1), CellValue.Number(2), CellValue.Number(3)),
 *   Vector(CellValue.Number(4), CellValue.Number(5), CellValue.Number(6))
 * ))
 * arr.rows   // 2
 * arr.cols   // 3
 * arr.transpose.rows  // 3
 * arr.transpose.cols  // 2
 * }}}
 */
final case class ArrayResult(values: Vector[Vector[CellValue]]):
  /** Number of rows in the array. */
  def rows: Int = values.size

  /** Number of columns in the array. */
  def cols: Int = values.headOption.map(_.size).getOrElse(0)

  /** Transpose the array (swap rows and columns). */
  def transpose: ArrayResult =
    if values.isEmpty then this
    else ArrayResult(values.transpose)

  /** Get value at (row, col), both 0-indexed. Returns Empty if out of bounds. */
  def apply(row: Int, col: Int): CellValue =
    if row >= 0 && row < rows && col >= 0 && col < cols then values(row)(col)
    else CellValue.Empty

  /** Check if array is empty (0x0). */
  def isEmpty: Boolean = values.isEmpty || values.forall(_.isEmpty)

object ArrayResult:
  /** Empty array result. */
  val empty: ArrayResult = ArrayResult(Vector.empty)

  /** Create a 1x1 array from a single value. */
  def single(value: CellValue): ArrayResult =
    ArrayResult(Vector(Vector(value)))

  /** Create an array from a flat sequence with specified dimensions. */
  def fromFlat(values: Seq[CellValue], rows: Int, cols: Int): ArrayResult =
    require(
      values.size == rows * cols,
      s"Expected ${rows * cols} values, got ${values.size}"
    )
    ArrayResult(
      values.grouped(cols).map(_.toVector).toVector
    )
