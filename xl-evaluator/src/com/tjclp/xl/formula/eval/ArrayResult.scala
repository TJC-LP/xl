package com.tjclp.xl.formula.eval

import com.tjclp.xl.cells.CellValue

/**
 * Result from array-returning functions like TRANSPOSE, SEQUENCE, and the value of a range or an
 * array expression in array mode.
 *
 * Represents a 2D grid of cell values that can be "spilled" into adjacent cells when evaluated as
 * an array formula.
 *
 * Storage is sparse (#750): a dense, populated '''box''' of cells at (`boxTop`, `boxLeft`) and a
 * [[fill]] describing every cell outside it: one value ([[ArrayResult.Fill.Const]]), or one value
 * per column ([[ArrayResult.Fill.ByCol]]) or per row ([[ArrayResult.Fill.ByRow]]). A range read in
 * array mode keeps only its populated cells (`R:KHI` is 1,048,576 × 7,636 cells, almost all of them
 * blank), and an elementwise operation computes its result over the box and once per fill entry
 * (`A:B=""` is TRUE everywhere outside the populated rows; `A:A*{1,2,3}` is 0, 0, 0 by column
 * there, and `{1,2,3}` itself is one value per column whatever the row). So a whole-axis array
 * evaluates exactly, in time and memory proportional to its populated cells and its axis lengths. A
 * dense array is one whose box is the whole array.
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
  private val outside: ArrayResult.Fill
):
  import ArrayResult.Fill

  private val boxHeight: Int = box.size
  private val boxWidth: Int = box.headOption.map(_.size).getOrElse(0)

  /**
   * Number of cells in the array (`rows * cols`), as a Long so a whole-sheet shape cannot overflow.
   */
  def cellCount: Long = rows.toLong * cols.toLong

  /** Number of cells the populated box holds (every cell of a dense array). */
  private[formula] def boxArea: Long = boxHeight.toLong * boxWidth.toLong

  /** Number of cells outside the box: each holds its [[fill]] entry. */
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
   * The cells outside the populated box: one value, or one value per row or per column. A dense
   * array has no such cell, and its fill is `Fill.Const(CellValue.Empty)`; an entry for a row or
   * column that has no cell outside the box is `CellValue.Empty` too.
   */
  def fill: ArrayResult.Fill = outside

  /** Get value at (row, col), both 0-indexed. Returns Empty if out of bounds. */
  def apply(row: Int, col: Int): CellValue =
    if row >= 0 && row < rows && col >= 0 && col < cols then
      val r = row - top
      val c = col - left
      if r >= 0 && r < boxHeight && c >= 0 && c < boxWidth then box(r)(c) else outside.at(row, col)
    else CellValue.Empty

  /** Check if array is empty (0x0). */
  def isEmpty: Boolean = rows == 0 || cols == 0

  /**
   * The cells in row-major order as runs of `(value, length)`. Box cells come as runs of 1 (equal
   * neighbours are not coalesced). A constant fill's cells between and around the box come as one
   * run each, so such an array has a few more runs than box cells whatever its size; a fill by row
   * gives a run per row segment, and a fill by column a run per cell outside the box (`rows × cols`
   * runs at worst: fold with [[tally]] when order matters only for the first error). Walking runs
   * in order visits values in the same order as a row-major cell walk, so the first error found is
   * the same.
   */
  def runs: Iterator[(CellValue, Long)] = ArrayResult.merge(pieces(compact = false))

  /**
   * The cells as `(value, count)` groups for an order-insensitive fold (a sum, a count, AND/OR):
   * the counts are the cells', and every value is met for the first time in row-major order, so a
   * fold that stops at its first error stops at the error a row-major walk finds first. Equal to
   * [[runs]] except for a fill by column, whose rows past the first of each block of rows outside
   * the box come as one group per column: `A:A*{1,2,3}` is its box, three runs and three groups.
   */
  private[formula] def tally: Iterator[(CellValue, Long)] =
    ArrayResult.merge(pieces(compact = true))

  /**
   * Row-major pieces `(value, count, key)`: box cells with key -1, fill cells with the key of their
   * fill entry (0 for a constant fill), so [[ArrayResult.merge]] joins neighbouring cells of one
   * entry. `compact` groups a fill by column's repeated rows (see [[tally]]).
   */
  private def pieces(compact: Boolean): Iterator[(CellValue, Long, Int)] =
    if isEmpty then Iterator.empty
    else
      val hasBox = boxHeight > 0 && boxWidth > 0
      val boxEnd = if hasBox then top + boxHeight else 0
      val colsL = cols.toLong
      def columnPieces(from: Int, until: Int, vs: Vector[CellValue], n: Long) =
        (from until until).iterator.map(c => (Fill.entry(vs, c), n, c))
      // rows [from, until), every cell outside the box
      def outsideRows(from: Int, until: Int): Iterator[(CellValue, Long, Int)] =
        if from >= until then Iterator.empty
        else
          outside match
            case Fill.Const(v) => Iterator.single((v, (until - from).toLong * colsL, 0))
            case Fill.ByRow(vs) =>
              (from until until).iterator.map(r => (Fill.entry(vs, r), colsL, r))
            case Fill.ByCol(vs) =>
              val first = columnPieces(0, cols, vs, 1L)
              val more = until - from - 1
              if more == 0 then first
              else if compact then first ++ columnPieces(0, cols, vs, more.toLong)
              else first ++ (1 to more).iterator.flatMap(_ => columnPieces(0, cols, vs, 1L))
      // the cells of row `r` in columns [from, until), outside the box
      def gap(r: Int, from: Int, until: Int): Iterator[(CellValue, Long, Int)] =
        if from >= until then Iterator.empty
        else
          outside match
            case Fill.Const(v) => Iterator.single((v, (until - from).toLong, 0))
            case Fill.ByRow(vs) => Iterator.single((Fill.entry(vs, r), (until - from).toLong, r))
            case Fill.ByCol(vs) => columnPieces(from, until, vs, 1L)
      val right = left + boxWidth
      val groupGaps = compact && (outside match
        case Fill.ByCol(_) => true
        case _ => false)
      val boxRowsIt =
        if !hasBox then Iterator.empty
        else
          box.iterator.zipWithIndex.flatMap { (row, i) =>
            val cells = row.iterator.map(cv => (cv, 1L, -1))
            if groupGaps && i > 0 then cells
            else gap(top + i, 0, left) ++ cells ++ gap(top + i, right, cols)
          } ++ (outside match
            case Fill.ByCol(vs) if groupGaps && boxHeight > 1 =>
              val more = (boxHeight - 1).toLong
              columnPieces(0, left, vs, more) ++ columnPieces(right, cols, vs, more)
            case _ => Iterator.empty)
      if hasBox then outsideRows(0, top) ++ boxRowsIt ++ outsideRows(boxEnd, rows)
      else outsideRows(0, rows)

  /**
   * Every cell in row-major order: [[runs]] expanded, one value per cell. This visits all
   * [[cellCount]] cells, so it is for arrays known to be small; walk [[runs]] otherwise.
   */
  def cells: Iterator[CellValue] =
    runs.flatMap((cv, n) => Iterator.iterate(n)(_ - 1L).takeWhile(_ > 0L).map(_ => cv))

  private def geometry: ArrayResult.Geometry =
    ArrayResult.Geometry(rows, cols, top, left, boxHeight, boxWidth)

  /**
   * Apply `f` to every cell in row-major order; the first `Left` (in that order) is returned. `f`
   * is applied once per fill entry, standing for every cell outside the box that entry covers, so
   * it must be pure.
   */
  def map(f: CellValue => Either[EvalError, CellValue]): Either[EvalError, ArrayResult] =
    val g = geometry
    val mappedFill = ArrayResult.mapFill(g, outside)(f)
    ArrayResult.mapBox(box, top, left, cols)(f) match
      case Left((err, at)) =>
        mappedFill match
          case Left((fillErr, fillAt)) if fillAt < at => Left(fillErr)
          case _ => Left(err)
      case Right(mapped) =>
        mappedFill match
          case Left((fillErr, _)) => Left(fillErr)
          case Right(newFill) => Right(ArrayResult.sparse(rows, cols, top, left, mapped, newFill))

  /** Apply a total `f` to every cell (to each fill entry once). */
  def mapPure(f: CellValue => CellValue): ArrayResult =
    val g = geometry
    val newFill = outside match
      case Fill.Const(v) => Fill.Const(if outsideCount > 0L then f(v) else CellValue.Empty)
      case Fill.ByRow(vs) =>
        Fill.ByRow(vs.zipWithIndex.map((v, r) => if g.rowHasOutside(r) then f(v) else v))
      case Fill.ByCol(vs) =>
        Fill.ByCol(vs.zipWithIndex.map((v, c) => if g.colHasOutside(c) then f(v) else v))
    ArrayResult.sparse(rows, cols, top, left, box.map(_.map(f)), newFill)

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
      if inside || outside == Fill.Const(CellValue.Empty) then
        // the box's part of the window, in window coordinates; outside it the window reads the
        // fill (or Empty past the array's edge, the same value when the fill is Empty)
        val windowFill = outside match
          case Fill.Const(v) => Fill.Const(v)
          case Fill.ByRow(vs) => Fill.ByRow(vs.slice(rowFrom, rowFrom + nr))
          case Fill.ByCol(vs) => Fill.ByCol(vs.slice(colFrom, colFrom + nc))
        val r0 = math.max(top, rowFrom)
        val r1 = math.min(top + boxHeight, math.min(rows, rowFrom + nr))
        val c0 = math.max(left, colFrom)
        val c1 = math.min(left + boxWidth, math.min(cols, colFrom + nc))
        if r0 >= r1 || c0 >= c1 then ArrayResult.sparse(nr, nc, 0, 0, Vector.empty, windowFill)
        else
          ArrayResult.sparse(
            nr,
            nc,
            r0 - rowFrom,
            c0 - colFrom,
            box.slice(r0 - top, r1 - top).map(_.slice(c0 - left, c1 - left)),
            windowFill
          )
      else
        new ArrayResult(
          nr,
          nc,
          0,
          0,
          Vector.tabulate(nr)(r => Vector.tabulate(nc)(c => apply(rowFrom + r, colFrom + c))),
          Fill.Const(CellValue.Empty)
        )

  /** Transpose the array (swap rows and columns). */
  def transpose: ArrayResult =
    if isEmpty then ArrayResult.empty
    else
      val swapped = outside match
        case Fill.Const(v) => Fill.Const(v)
        case Fill.ByRow(vs) => Fill.ByCol(vs)
        case Fill.ByCol(vs) => Fill.ByRow(vs)
      new ArrayResult(cols, rows, left, top, box.transpose, swapped)

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
      val described = outside match
        case Fill.Const(v) => s"fill $v"
        case Fill.ByRow(vs) => s"fill by row ${vs.take(3).mkString(", ")}…"
        case Fill.ByCol(vs) => s"fill by column ${vs.take(3).mkString(", ")}…"
      s"ArrayResult(${rows}x$cols, box ${boxHeight}x$boxWidth at ($top, $left), $described)"

object ArrayResult:

  /**
   * The cells of an [[ArrayResult]] outside its populated box (#750). An array holds one of these
   * three, so a whole-axis array whose outside cells differ along one axis (`A:A*{1,2,3}`, a whole
   * column times a row of constants) is still held by its box and one value per column, never cell
   * by cell. An outside that varies along both axes is not representable: the operation that would
   * produce one builds its result cell by cell up to [[MaxCells]] cells and is refused past that.
   */
  enum Fill:
    /** One value for every cell outside the box. */
    case Const(value: CellValue)

    /** The value of row `r`'s cells outside the box is `values(r)`: one entry per row. */
    case ByRow(values: Vector[CellValue])

    /** The value of column `c`'s cells outside the box is `values(c)`: one entry per column. */
    case ByCol(values: Vector[CellValue])

    /** The fill's value at (`row`, `col`), a cell outside the box. */
    def at(row: Int, col: Int): CellValue = this match
      case Const(v) => v
      case ByRow(vs) => Fill.entry(vs, row)
      case ByCol(vs) => Fill.entry(vs, col)

  object Fill:
    /** `values(i)`, Empty past the end. */
    private[eval] def entry(values: Vector[CellValue], i: Int): CellValue =
      if i >= 0 && i < values.size then values(i) else CellValue.Empty

  /**
   * An array's shape and box: `rows` x `cols`, the box `height` x `width` at (`top`, `left`). Finds
   * which fill entries have cells outside the box and where, in row-major order, the first of them
   * is.
   */
  private[formula] final case class Geometry(
    rows: Int,
    cols: Int,
    top: Int,
    left: Int,
    height: Int,
    width: Int
  ):
    private val hasBox = height > 0 && width > 0
    private val bottom = top + height
    private val right = left + width

    def cellCount: Long = rows.toLong * cols.toLong
    def outsideCount: Long = cellCount - (if hasBox then height.toLong * width.toLong else 0L)

    /** Whether row `r` has a cell outside the box. */
    def rowHasOutside(r: Int): Boolean = !hasBox || r < top || r >= bottom || width < cols

    /** Whether column `c` has a cell outside the box. */
    def colHasOutside(c: Int): Boolean = !hasBox || c < left || c >= right || height < rows

    /** The row-major position of the first cell outside the box, Long.MaxValue for none. */
    def firstOutside: Long =
      if outsideCount == 0L then Long.MaxValue
      else if !hasBox || top > 0 || left > 0 then 0L
      else if width < cols then width.toLong
      else height.toLong * cols.toLong

    /** The row-major position of row `r`'s first cell outside the box, Long.MaxValue for none. */
    def firstOutsideInRow(r: Int): Long =
      if !rowHasOutside(r) then Long.MaxValue
      else if !hasBox || r < top || r >= bottom || left > 0 then r.toLong * cols.toLong
      else r.toLong * cols.toLong + right.toLong

    /**
     * The row-major position of column `c`'s first cell outside the box, Long.MaxValue for none.
     */
    def firstOutsideInCol(c: Int): Long =
      if !colHasOutside(c) then Long.MaxValue
      else if !hasBox || c < left || c >= right || top > 0 then c.toLong
      else bottom.toLong * cols.toLong + c.toLong

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

  /**
   * Compute a fill entry by entry: `entry(i)` for every entry that has a cell outside the box of
   * `g` (Empty for one that has none, never computed). The `Left` whose first outside cell comes
   * first in row-major order wins, and comes back with that cell's position. A constant fill is one
   * entry.
   */
  private[formula] def computeFill(g: Geometry, axis: Option[Boolean])(
    entry: Int => Either[EvalError, CellValue]
  ): Either[(EvalError, Long), Fill] =
    axis match
      case None =>
        if g.outsideCount == 0L then Right(Fill.Const(CellValue.Empty))
        else entry(0).map(Fill.Const(_)).left.map(e => (e, g.firstOutside))
      case Some(byRow) =>
        val n = if byRow then g.rows else g.cols
        val out = new scala.collection.immutable.VectorBuilder[CellValue]
        val failure = (0 until n).foldLeft(Option.empty[(EvalError, Long)]) { (worst, i) =>
          val at = if byRow then g.firstOutsideInRow(i) else g.firstOutsideInCol(i)
          if at == Long.MaxValue then
            out += CellValue.Empty
            worst
          else
            entry(i) match
              case Right(v) =>
                out += v
                worst
              case Left(e) =>
                out += CellValue.Empty
                worst match
                  case Some((_, earlier)) if earlier <= at => worst
                  case _ => Some((e, at))
        }
        failure match
          case Some(f) => Left(f)
          case None =>
            val values = out.result()
            Right(if byRow then Fill.ByRow(values) else Fill.ByCol(values))

  /** `f` over a fill's entries ([[computeFill]]: the first failure in row-major order wins). */
  private def mapFill(g: Geometry, fill: Fill)(
    f: CellValue => Either[EvalError, CellValue]
  ): Either[(EvalError, Long), Fill] =
    fill match
      case Fill.Const(v) => computeFill(g, None)(_ => f(v))
      case Fill.ByRow(vs) => computeFill(g, Some(true))(r => f(Fill.entry(vs, r)))
      case Fill.ByCol(vs) => computeFill(g, Some(false))(c => f(Fill.entry(vs, c)))

  /**
   * Join neighbouring pieces of one fill entry (the same key, never -1: a box cell) into one run.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def merge(pieces: Iterator[(CellValue, Long, Int)]): Iterator[(CellValue, Long)] =
    new Iterator[(CellValue, Long)]:
      private val in = pieces.filter(_._2 > 0L).buffered
      def hasNext: Boolean = in.hasNext
      def next(): (CellValue, Long) =
        val (v, n, key) = in.next()
        var total = n
        if key != -1 then while in.hasNext && in.head._3 == key do total += in.next()._2
        (v, total)

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
    new ArrayResult(values.size, width, 0, 0, rect, Fill.Const(CellValue.Empty))

  /**
   * A `rows` x `cols` array whose cells are `box` (rectangular) at (`top`, `left`) and `fill`
   * everywhere else. The box must lie inside the array; a fill by row has `rows` entries and one by
   * column `cols` (a shorter one reads Empty past its end). The fill is kept in its canonical form:
   * an entry with no cell outside the box is Empty, and a fill whose entries that do have such
   * cells are all equal is that one value.
   */
  private[formula] def sparse(
    rows: Int,
    cols: Int,
    top: Int,
    left: Int,
    box: Vector[Vector[CellValue]],
    fill: Fill
  ): ArrayResult =
    if rows <= 0 || cols <= 0 then empty
    else
      val height0 = box.size
      val width0 = box.headOption.map(_.size).getOrElse(0)
      val (height, width, at, from, kept) =
        if height0 == 0 || width0 == 0 then (0, 0, 0, 0, Vector.empty[Vector[CellValue]])
        else (height0, width0, top, left, box)
      val g = Geometry(rows, cols, at, from, height, width)
      new ArrayResult(rows, cols, at, from, kept, canonical(g, fill))

  /** A `rows` x `cols` array holding `value` in every cell, stored as that one value. */
  private[formula] def constant(rows: Int, cols: Int, value: CellValue): ArrayResult =
    sparse(rows, cols, 0, 0, Vector.empty, Fill.Const(value))

  /** See [[sparse]]: irrelevant entries Empty, a uniform fill one value. */
  private def canonical(g: Geometry, fill: Fill): Fill =
    if g.outsideCount == 0L then Fill.Const(CellValue.Empty)
    else
      def axis(n: Int, vs: Vector[CellValue], has: Int => Boolean)(
        make: Vector[CellValue] => Fill
      ): Fill =
        val entries = Vector.tabulate(n)(i => if has(i) then Fill.entry(vs, i) else CellValue.Empty)
        val used = (0 until n).iterator.filter(has).map(entries(_))
        if !used.hasNext then Fill.Const(CellValue.Empty)
        else
          val first = used.next()
          if used.forall(_ == first) then Fill.Const(first) else make(entries)
      fill match
        case c: Fill.Const => c
        case Fill.ByRow(vs) => axis(g.rows, vs, g.rowHasOutside)(Fill.ByRow(_))
        case Fill.ByCol(vs) => axis(g.cols, vs, g.colHasOutside)(Fill.ByCol(_))

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
