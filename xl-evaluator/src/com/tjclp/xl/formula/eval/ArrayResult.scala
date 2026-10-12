package com.tjclp.xl.formula.eval

import com.tjclp.xl.cells.CellValue

/**
 * Result from array-returning functions like TRANSPOSE, SEQUENCE, and the value of a range or an
 * array expression in array mode.
 *
 * Represents a 2D grid of cell values that can be "spilled" into adjacent cells when evaluated as
 * an array formula.
 *
 * Storage is sparse (#750): a dense, populated '''box''' of cells at (`boxTop`, `boxLeft`) and one
 * [[fill]] value for every cell outside it. A range read in array mode keeps only its populated
 * cells (`R:KHI` is 1,048,576 × 7,636 cells, almost all of them blank), and an elementwise
 * operation computes its result over the box and once for the fill (`A:B=""` is TRUE everywhere
 * outside the populated rows). So a whole-axis array evaluates exactly, in time and memory
 * proportional to its populated cells. A dense array is one whose box is the whole array.
 *
 * Design principles:
 *   - Immutable value type, read through its accessors ([[apply]], [[runs]], [[map]], [[slice]],
 *     [[transpose]], [[toDense]]) rather than its storage
 *   - Row-major ordering: `apply(row, col)`, and [[runs]] and [[map]] visit cells row by row
 *   - Empty arrays have size 0x0
 *   - Equality is cell by cell: two arrays are equal when they have the same shape and the same
 *     value in every cell, whatever their storage
 *   - Only [[toDense]] builds every cell, and it refuses an array of more than
 *     [[ArrayResult.MaxCells]] cells
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
final class ArrayResult private (
  val rows: Int,
  val cols: Int,
  private val top: Int,
  private val left: Int,
  private val box: Vector[Vector[CellValue]],
  private val outsideValue: CellValue
):
  private val boxHeight: Int = box.size
  private val boxWidth: Int = box.headOption.map(_.size).getOrElse(0)

  /**
   * Number of cells in the array (`rows * cols`), as a Long so a whole-sheet shape cannot overflow.
   */
  def cellCount: Long = rows.toLong * cols.toLong

  /** Number of cells the populated box holds (every cell of a dense array). */
  private[formula] def boxArea: Long = boxHeight.toLong * boxWidth.toLong

  /** Number of cells outside the box: each holds [[fill]]. */
  private[formula] def outsideCount: Long = cellCount - boxArea

  /** Whether every cell is stored (the box is the whole array). */
  private[formula] def isDense: Boolean = outsideCount == 0L

  /** The box's first row and column, and its height and width. */
  private[formula] def boxTop: Int = top
  private[formula] def boxLeft: Int = left
  private[formula] def boxRows: Int = boxHeight
  private[formula] def boxCols: Int = boxWidth

  /** One cell of the box, by its position within the box. */
  private[formula] def boxCell(row: Int, col: Int): CellValue = box(row)(col)

  /**
   * The value of every cell outside the populated box. A dense array has no such cell, and its fill
   * is `CellValue.Empty`.
   */
  def fill: CellValue = outsideValue

  /** Get value at (row, col), both 0-indexed. Returns Empty if out of bounds. */
  def apply(row: Int, col: Int): CellValue =
    if row >= 0 && row < rows && col >= 0 && col < cols then
      val r = row - top
      val c = col - left
      if r >= 0 && r < boxHeight && c >= 0 && c < boxWidth then box(r)(c) else outsideValue
    else CellValue.Empty

  /** Check if array is empty (0x0). */
  def isEmpty: Boolean = rows == 0 || cols == 0

  /**
   * The cells in row-major order as runs of `(value, length)`. Box cells come as runs of 1 (equal
   * neighbours are not coalesced); the fill cells between and around them as one run each, so a
   * sparse array has a few more runs than box cells whatever its size. Walking runs in order visits
   * values in the same order as a row-major cell walk, so the first error found is the same.
   */
  def runs: Iterator[(CellValue, Long)] =
    if isEmpty then Iterator.empty
    else if boxHeight == 0 || boxWidth == 0 then Iterator.single((outsideValue, cellCount))
    else
      val colsL = cols.toLong
      val before = top.toLong * colsL
      val after = (rows - top - boxHeight).toLong * colsL
      val rightGap = (cols - left - boxWidth).toLong
      // the fill run that precedes box row `i`: what is left of the row above plus this row's
      // cells left of the box (and, for the first box row, every row above the box)
      def leading(i: Int): Long = (if i == 0 then before else rightGap) + left.toLong
      val inner = box.iterator.zipWithIndex.flatMap { (row, i) =>
        val gap = leading(i)
        (if gap > 0 then Iterator.single((outsideValue, gap)) else Iterator.empty) ++
          row.iterator.map(cv => (cv, 1L))
      }
      val trailing = rightGap + after
      inner ++ (if trailing > 0 then Iterator.single((outsideValue, trailing)) else Iterator.empty)

  /**
   * Every cell in row-major order: [[runs]] expanded, one value per cell. This visits all
   * [[cellCount]] cells, so it is for arrays known to be small; walk [[runs]] otherwise.
   */
  def cells: Iterator[CellValue] =
    runs.flatMap((cv, n) => Iterator.iterate(n)(_ - 1L).takeWhile(_ > 0L).map(_ => cv))

  /**
   * The row-major position (`row * cols + col`) of the first cell outside the box, or
   * `Long.MaxValue` when every cell is in it. Ranks a failure on the fill against one in the box.
   */
  private[formula] def firstOutside: Long =
    if outsideCount == 0L then Long.MaxValue
    else if boxHeight == 0 || boxWidth == 0 || top > 0 || left > 0 then 0L
    else if boxWidth < cols then boxWidth.toLong
    else boxHeight.toLong * cols.toLong

  /**
   * Apply `f` to every cell in row-major order; the first `Left` (in that order) is returned. `f`
   * is applied to the fill once, standing for every cell outside the box, so it must be pure.
   */
  def map(f: CellValue => Either[EvalError, CellValue]): Either[EvalError, ArrayResult] =
    val fillResult = if outsideCount > 0L then Some(f(outsideValue)) else None
    ArrayResult.mapBox(box, top, left, cols)(f) match
      case Left((err, at)) =>
        fillResult match
          case Some(Left(fillErr)) if firstOutside < at => Left(fillErr)
          case _ => Left(err)
      case Right(mapped) =>
        fillResult match
          case Some(Left(fillErr)) => Left(fillErr)
          case Some(Right(newFill)) =>
            Right(new ArrayResult(rows, cols, top, left, mapped, newFill))
          case None => Right(new ArrayResult(rows, cols, top, left, mapped, CellValue.Empty))

  /** Apply a total `f` to every cell (to the fill once). */
  def mapPure(f: CellValue => CellValue): ArrayResult =
    new ArrayResult(
      rows,
      cols,
      top,
      left,
      box.map(_.map(f)),
      if outsideCount > 0L then f(outsideValue) else CellValue.Empty
    )

  /**
   * The `rowCount` x `colCount` window whose top-left cell is (`rowFrom`, `colFrom`), read through
   * [[apply]] (cells outside this array are Empty). Negative counts are treated as 0. A window
   * inside the array keeps its storage sparse; one that runs past its edge is built cell by cell.
   */
  def slice(rowFrom: Int, rowCount: Int, colFrom: Int, colCount: Int): ArrayResult =
    val nr = math.max(rowCount, 0)
    val nc = math.max(colCount, 0)
    if nr == 0 || nc == 0 then ArrayResult.empty
    else
      val inside = rowFrom >= 0 && colFrom >= 0 &&
        rowFrom.toLong + nr <= rows && colFrom.toLong + nc <= cols
      if inside || outsideValue == CellValue.Empty then
        // the box's part of the window, in window coordinates; outside it the window reads the
        // fill (or Empty past the array's edge, the same value when the fill is Empty)
        val r0 = math.max(top, rowFrom)
        val r1 = math.min(top + boxHeight, math.min(rows, rowFrom + nr))
        val c0 = math.max(left, colFrom)
        val c1 = math.min(left + boxWidth, math.min(cols, colFrom + nc))
        if r0 >= r1 || c0 >= c1 then ArrayResult.constant(nr, nc, outsideValue)
        else
          ArrayResult.sparse(
            nr,
            nc,
            r0 - rowFrom,
            c0 - colFrom,
            box.slice(r0 - top, r1 - top).map(_.slice(c0 - left, c1 - left)),
            outsideValue
          )
      else
        new ArrayResult(
          nr,
          nc,
          0,
          0,
          Vector.tabulate(nr)(r => Vector.tabulate(nc)(c => apply(rowFrom + r, colFrom + c))),
          CellValue.Empty
        )

  /** Transpose the array (swap rows and columns). */
  def transpose: ArrayResult =
    if isEmpty then ArrayResult.empty
    else new ArrayResult(cols, rows, left, top, box.transpose, outsideValue)

  /**
   * The whole array as rows of cells, for consumers that need every cell (sorting, spilling,
   * printing). `label` names the consumer in the error returned when the array holds more than
   * [[ArrayResult.MaxCells]] cells.
   */
  def toDense(label: String): Either[EvalError, Vector[Vector[CellValue]]] =
    if cellCount > ArrayResult.MaxCells then Left(ArrayResult.tooLarge(label, rows, cols))
    else if isDense && top == 0 && left == 0 && boxHeight == rows then Right(box)
    else Right(Vector.tabulate(rows)(r => Vector.tabulate(cols)(c => apply(r, c))))

  override def equals(that: Any): Boolean = that match
    case other: ArrayResult =>
      (this eq other) || (rows == other.rows && cols == other.cols && ArrayResult
        .sameRuns(runs, other.runs))
    case _ => false

  override def hashCode: Int = (rows, cols, apply(0, 0)).hashCode

  override def toString: String =
    if cellCount <= 64 then
      (0 until rows)
        .map(r => (0 until cols).map(c => apply(r, c)).mkString(", "))
        .mkString(s"ArrayResult(${rows}x$cols: [", "; ", "])")
    else if isDense then s"ArrayResult(${rows}x$cols, ${apply(0, 0)}, …)"
    else
      s"ArrayResult(${rows}x$cols, box ${boxHeight}x$boxWidth at ($top, $left), fill $outsideValue)"

object ArrayResult:

  /**
   * The most cells one array may hold: one full column (1,048,576 rows). SEQUENCE and INDIRECT
   * refuse a larger result, and so does dynamic-array authoring (GH-714): a spill of at most this
   * many cells — a full column — is authored, a larger one refused before it is written. A sparse
   * array may declare any shape; only [[ArrayResult.toDense]] and a box larger than this are
   * refused.
   */
  val MaxCells: Long = 1048576L

  private def count(n: Long): String = String.format(java.util.Locale.ROOT, "%,d", n)

  /** The refusal of an array too large to build cell by cell. */
  private[formula] def tooLarge(label: String, rows: Long, cols: Long): EvalError =
    EvalError.EvalFailed(
      s"$label: array of ${rows}×$cols (${count(rows * cols)} cells) exceeds the " +
        s"${count(MaxCells)}-cell limit",
      None
    )

  /**
   * Map `f` over the box row-major, stopping at the first `Left`, which comes back with its
   * row-major position in an array of `cols` columns whose box is at (`top`, `left`).
   */
  private def mapBox(box: Vector[Vector[CellValue]], top: Int, left: Int, cols: Int)(
    f: CellValue => Either[EvalError, CellValue]
  ): Either[(EvalError, Long), Vector[Vector[CellValue]]] =
    val out = new scala.collection.immutable.VectorBuilder[Vector[CellValue]]
    @scala.annotation.tailrec
    def rowLoop(i: Int): Either[(EvalError, Long), Vector[Vector[CellValue]]] =
      if i >= box.size then Right(out.result())
      else
        val row = box(i)
        val cellsOut = new scala.collection.immutable.VectorBuilder[CellValue]
        @scala.annotation.tailrec
        def cellLoop(j: Int): Option[(EvalError, Long)] =
          if j >= row.size then None
          else
            f(row(j)) match
              case Right(v) =>
                cellsOut += v
                cellLoop(j + 1)
              case Left(e) => Some((e, (top + i).toLong * cols.toLong + (left + j).toLong))
        cellLoop(0) match
          case Some(failure) => Left(failure)
          case None =>
            out += cellsOut.result()
            rowLoop(i + 1)
    rowLoop(0)

  /** Whether two run walks describe the same cells (their run boundaries may differ). */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def sameRuns(a: Iterator[(CellValue, Long)], b: Iterator[(CellValue, Long)]): Boolean =
    var av: CellValue = CellValue.Empty
    var an = 0L
    var bv: CellValue = CellValue.Empty
    var bn = 0L
    var same = true
    while same && (an > 0L || a.hasNext) && (bn > 0L || b.hasNext) do
      if an == 0L then
        val (v, n) = a.next()
        av = v
        an = n
      if bn == 0L then
        val (v, n) = b.next()
        bv = v
        bn = n
      if an > 0L && bn > 0L then
        if av != bv then same = false
        else
          val step = math.min(an, bn)
          an -= step
          bn -= step
    same && an == 0L && bn == 0L && !a.hasNext && !b.hasNext

  /**
   * A dense array from its rows (row-major: `values(row)(col)`). Its width is the first row's; a
   * shorter row reads Empty past its end.
   */
  def apply(values: Vector[Vector[CellValue]]): ArrayResult =
    val width = values.headOption.map(_.size).getOrElse(0)
    val rect =
      if values.forall(_.size == width) then values
      else values.map(row => row.take(width) ++ Vector.fill(width - row.size)(CellValue.Empty))
    new ArrayResult(values.size, width, 0, 0, rect, CellValue.Empty)

  /**
   * A `rows` x `cols` array whose cells are `box` (rectangular) at (`top`, `left`) and `fill`
   * everywhere else. The box must lie inside the array.
   */
  private[formula] def sparse(
    rows: Int,
    cols: Int,
    top: Int,
    left: Int,
    box: Vector[Vector[CellValue]],
    fill: CellValue
  ): ArrayResult =
    if rows <= 0 || cols <= 0 then empty
    else
      val height = box.size
      val width = box.headOption.map(_.size).getOrElse(0)
      if height == 0 || width == 0 then constant(rows, cols, fill)
      else
        val covers = top == 0 && left == 0 && height == rows && width == cols
        new ArrayResult(rows, cols, top, left, box, if covers then CellValue.Empty else fill)

  /** A `rows` x `cols` array holding `value` in every cell, stored as that one value. */
  private[formula] def constant(rows: Int, cols: Int, value: CellValue): ArrayResult =
    if rows <= 0 || cols <= 0 then empty
    else new ArrayResult(rows, cols, 0, 0, Vector.empty, value)

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
