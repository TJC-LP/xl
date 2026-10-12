package com.tjclp.xl.formula.eval

import com.tjclp.xl.cells.CellValue

/**
 * Result from array-returning functions like TRANSPOSE, SEQUENCE.
 *
 * Represents a 2D grid of cell values that can be "spilled" into adjacent cells when evaluated as
 * an array formula.
 *
 * Design principles:
 *   - Immutable value type, read through its accessors ([[apply]], [[runs]], [[map]], [[slice]],
 *     [[transpose]], [[toDense]]) rather than its storage, so the storage can become sparse (a
 *     populated box plus a [[fill]] value for the cells outside it, #750) without callers
 *     materialising arrays they only walk
 *   - Row-major ordering: `apply(row, col)`, and [[runs]] and [[map]] visit cells row by row
 *   - Empty arrays have size 0x0
 *   - Equality is cell by cell: two arrays are equal when they have the same shape and the same
 *     value in every cell, whatever their storage
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
 * arr(1, 2)  // CellValue.Number(6)
 * arr.transpose.rows  // 3
 * arr.transpose.cols  // 2
 * }}}
 */
final class ArrayResult private (private val grid: Vector[Vector[CellValue]]):

  /** Number of rows in the array. */
  def rows: Int = grid.size

  /** Number of columns in the array. */
  def cols: Int = grid.headOption.map(_.size).getOrElse(0)

  /**
   * Number of cells in the array (`rows * cols`), as a Long so a whole-sheet shape cannot overflow.
   */
  def cellCount: Long = rows.toLong * cols.toLong

  /**
   * The value of every cell outside the populated box. A dense array's box is the whole array, so
   * its fill is never read and is `CellValue.Empty`.
   */
  def fill: CellValue = CellValue.Empty

  /** Get value at (row, col), both 0-indexed. Returns Empty if out of bounds. */
  def apply(row: Int, col: Int): CellValue =
    if row >= 0 && row < rows && col >= 0 && col < cols then grid(row)(col)
    else CellValue.Empty

  /** Check if array is empty (0x0). */
  def isEmpty: Boolean = rows == 0 || cols == 0

  /**
   * The cells in row-major order as runs of `(value, length)`. A dense array yields every cell as a
   * run of 1 (equal neighbours are not coalesced); a sparse one yields fill cells as longer runs.
   * Walking runs in order visits values in the same order as a row-major cell walk, so the first
   * error found is the same.
   */
  def runs: Iterator[(CellValue, Long)] =
    grid.iterator.flatMap(_.iterator).map(cv => (cv, 1L))

  /**
   * Every row, top to bottom, each materialised in full. Unlike [[toDense]] this has no size cap:
   * it is for the consumers that already build the whole grid whatever its size (spill writes,
   * authoring) until #750 moves them onto [[toDense]]; new code should use [[toDense]], [[runs]] or
   * [[apply]].
   */
  def rowsIterator: Iterator[Vector[CellValue]] = grid.iterator

  /** Every cell in row-major order: [[runs]] expanded, one value per cell. */
  def cells: Iterator[CellValue] =
    runs.flatMap((cv, n) => Iterator.fill(n.toInt)(cv))

  /**
   * Apply `f` to every cell in row-major order; the first `Left` stops the walk and is returned.
   */
  def map(f: CellValue => Either[EvalError, CellValue]): Either[EvalError, ArrayResult] =
    ArrayResult.traverse(grid)(row => ArrayResult.traverse(row)(f)).map(new ArrayResult(_))

  /** Apply a total `f` to every cell. */
  def mapPure(f: CellValue => CellValue): ArrayResult =
    new ArrayResult(grid.map(_.map(f)))

  /**
   * The `rowCount` x `colCount` window whose top-left cell is (`rowFrom`, `colFrom`), read through
   * [[apply]] (cells outside this array are Empty). Negative counts are treated as 0.
   */
  def slice(rowFrom: Int, rowCount: Int, colFrom: Int, colCount: Int): ArrayResult =
    val nr = math.max(rowCount, 0)
    val nc = math.max(colCount, 0)
    if nr == 0 || nc == 0 then ArrayResult.empty
    else
      new ArrayResult(
        Vector.tabulate(nr)(r => Vector.tabulate(nc)(c => apply(rowFrom + r, colFrom + c)))
      )

  /** Transpose the array (swap rows and columns). */
  def transpose: ArrayResult =
    if grid.isEmpty then this
    else new ArrayResult(grid.transpose)

  /**
   * The whole array as rows of cells, for consumers that need every cell (sorting, spilling,
   * printing). `label` names the consumer in the error returned when the array holds more than
   * [[ArrayResult.MaxCells]] cells.
   */
  def toDense(label: String): Either[EvalError, Vector[Vector[CellValue]]] =
    if cellCount > ArrayResult.MaxCells then
      Left(
        EvalError.EvalFailed(
          s"$label: array of ${rows}×${cols} ($cellCount cells) exceeds the 1,048,576-cell limit",
          None
        )
      )
    else Right(grid)

  override def equals(that: Any): Boolean = that match
    case other: ArrayResult =>
      (this eq other) || (
        rows == other.rows && cols == other.cols &&
          (0 until rows).forall(r => (0 until cols).forall(c => apply(r, c) == other(r, c)))
      )
    case _ => false

  override def hashCode: Int = (rows, cols, apply(0, 0)).hashCode

  override def toString: String =
    if cellCount <= 64 then
      grid.map(_.mkString(", ")).mkString(s"ArrayResult(${rows}x$cols: [", "; ", "])")
    else s"ArrayResult(${rows}x$cols, ${grid.headOption.flatMap(_.headOption).getOrElse(fill)}, …)"

object ArrayResult:

  /**
   * The most cells one array may hold: one full column (1,048,576 rows). SEQUENCE and INDIRECT
   * refuse a larger result, and so does dynamic-array authoring (GH-714): a spill of at most this
   * many cells — a full column — is authored, a larger one refused before it is written.
   */
  val MaxCells: Long = 1048576L

  /** Map `f` over `vec` in order, stopping at the first `Left` (no call of `f` after it). */
  private def traverse[A, B](vec: Vector[A])(
    f: A => Either[EvalError, B]
  ): Either[EvalError, Vector[B]] =
    @scala.annotation.tailrec
    def loop(
      it: Iterator[A],
      acc: scala.collection.immutable.VectorBuilder[B]
    ): Either[EvalError, Vector[B]] =
      if !it.hasNext then Right(acc.result())
      else
        f(it.next()) match
          case Right(b) => loop(it, acc += b)
          case Left(e) => Left(e)
    loop(vec.iterator, new scala.collection.immutable.VectorBuilder[B])

  /** A dense array from its rows (row-major: `values(row)(col)`). */
  def apply(values: Vector[Vector[CellValue]]): ArrayResult = new ArrayResult(values)

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
