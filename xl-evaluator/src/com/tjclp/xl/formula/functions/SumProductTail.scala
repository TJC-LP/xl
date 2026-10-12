package com.tjclp.xl.formula.functions

import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row, SheetName}
import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.eval.{EvalError, WorkbookAudit}
import com.tjclp.xl.sheets.Sheet

/**
 * GH-711/GH-192: the blank tail of SUMPRODUCT over whole columns (or whole rows). The shared bound
 * keeps only the rows inside every involved sheet's used range; every row it drops lies outside all
 * of them, so it is blank in every argument and adds the same product. SUMPRODUCT reads the
 * arguments once over one such row ([[Tail.line]]) and adds `count` × that product to the trimmed
 * sum, so `SUMPRODUCT(--(A:A=""))` counts every blank row, as Excel does.
 */
private[formula] object SumProductTail:

  /**
   * What the bound dropped: how many rows (columns), along which axis, and how to read any one of
   * them, a one-row (one-column) range in place of each whole column (row).
   */
  final case class Tail(count: Long, alongRows: Boolean, line: CellRange => CellRange):
    /** How many rows (columns) the bound kept: the extent the trimmed arguments have. */
    def kept: Long = (if alongRows then Row.MaxIndex0 + 1 else Column.MaxIndex0 + 1).toLong - count

    /** Whether `range` runs the whole trimmed axis. */
    def spans(range: CellRange): Boolean = if alongRows then range.isFullColumn else range.isFullRow

  /**
   * The tail of SUMPRODUCT's `ranges` under its shared bounds (the kept rows, the kept columns):
   * None when nothing was dropped, or when whole columns and whole rows mix (the dropped cells then
   * form no whole rows).
   */
  def of(
    ranges: List[CellRange],
    bounds: (Option[(Row, Row)], Option[(Column, Column)])
  ): Option[Tail] =
    val columns = ranges.exists(_.isFullColumn)
    val rows = ranges.exists(_.isFullRow)
    // the first index outside the kept [lo, hi], and how many of `total` lie outside it
    def outside(kept: Option[(Int, Int)], total: Int): Option[(Int, Long)] =
      kept match
        case None => Some((0, total.toLong))
        case Some((lo, hi)) =>
          val dropped = total.toLong - (hi - lo + 1)
          Option.when(dropped > 0L)((if lo > 0 then 0 else hi + 1, dropped))
    if columns == rows then None
    else if columns then
      outside(bounds._1.map((lo, hi) => (lo.index0, hi.index0)), Row.MaxIndex0 + 1).map {
        (blank, dropped) =>
          Tail(
            dropped,
            alongRows = true,
            r =>
              if r.isFullColumn then
                CellRange(ARef.from0(r.colStart.index0, blank), ARef.from0(r.colEnd.index0, blank))
              else r
          )
      }
    else
      outside(bounds._2.map((lo, hi) => (lo.index0, hi.index0)), Column.MaxIndex0 + 1).map {
        (blank, dropped) =>
          Tail(
            dropped,
            alongRows = false,
            r =>
              if r.isFullRow then
                CellRange(ARef.from0(blank, r.rowStart.index0), ARef.from0(blank, r.rowEnd.index0))
              else r
          )
      }

  /**
   * The volatile functions that draw anew on every evaluation (RAND, RANDBETWEEN): reading the
   * blank line would be a second draw. TODAY and NOW are volatile across recalculations but read
   * the evaluation's one clock, so a second read is the same value and they keep the tail
   * (`SUMPRODUCT(--(A:A<TODAY()))` counts the blank rows as `--(A:A<45000)` does). A volatile spec
   * outside [[ClockReads]] counts as a draw until classified; a drift law pins both sets against
   * the registry's volatile specs.
   */
  val ClockReads: Set[String] = Set("TODAY", "NOW")

  /** Whether `spec` draws: volatile, and not one of the [[ClockReads]]. */
  def draws(spec: FunctionSpec[?]): Boolean =
    spec.flags.volatile && !ClockReads.contains(spec.name.toUpperCase)

  /**
   * The arguments to read over the blank line, decided before anything is evaluated: each
   * expression with its trimmed whole columns on the line, a range location as is (its caller maps
   * it). None — no tail, the trimmed sum stands — unless all hold:
   *   - the trimmed arguments' extent along the axis (`trimmedExtent`) is the kept one, so no
   *     argument came back at another height.
   *   - every argument reads the axis only through the line: a range location that runs it
   *     (`spansLocation`), or an expression that reads at least one whole column (row) as a value
   *     and has every one it reads as a value moved onto the line by its transform
   *     ([[TExpr.collectValueRanges]]). A whole column left whole in an expression position
   *     (`ROW(A:A)`, which an array context may spread over every row) could vary from row to row,
   *     so it rules the tail out on its own. A whole column in a range slot (SUM's, COUNTIF's
   *     range, a lookup's table) is one fixed reference the function consumes, the same for every
   *     row, so it may stay: `1/COUNTIF(A:A,A:A)` moves its criterion and keeps its range.
   *     `SUMPRODUCT(SUM(A:A))` reads no whole column as a value, so it has no tail.
   *   - no argument also reads a bounded range running more than one row (column) along the axis as
   *     a value: beside a whole column it is another height, which a selector may pick
   *     (`IF(A1>0,B:B,C1:C5)`) and no single blank row stands for.
   *   - no argument reads a reference's position along the axis (a call flagged
   *     [[FunctionFlags.boundsReferenceAxis]]: ROW, COLUMN). SUMPRODUCT's shared extent reaches
   *     into ROW(A:A) (GH-712), so its whole column moves onto the line like any other, but the
   *     line's one row stands for no blank row's number: `(A:A="")*ROW(A:A)` varies from row to
   *     row, so it keeps the trimmed sum.
   *   - no argument draws a random number ([[draws]]), directly or through a defined name
   *     (`drawsViaName`): reading the line evaluates every non-range part of the argument a second
   *     time, and `RAND()` must draw once.
   */
  def plan(
    tail: Tail,
    args: List[ArgSpec.SumProductArg],
    trimmedExtent: Int,
    spansLocation: TExpr.RangeLocation => Boolean,
    drawsViaName: (Option[SheetName], String) => Boolean = (_, _) => false
  ): Option[List[ArgSpec.SumProductArg]] =
    def extent(r: CellRange): Int = if tail.alongRows then r.height else r.width
    def readsLineOnly(expr: TExpr[?], line: TExpr[?]): Boolean =
      val before = TExpr.collectValueRanges(expr).map(_._2)
      val after = TExpr.collectValueRanges(line).map(_._2)
      before.exists(tail.spans) && !after.exists(tail.spans) && !after.exists(extent(_) > 1)
    val lines = args.map {
      case Left(location) => (spansLocation(location), Left(location))
      case Right(expr) =>
        val line = TExpr.transformRanges(expr, (_, r) => tail.line(r))
        (
          readsLineOnly(expr, line) && !WorkbookAudit.callsWhere(expr, drawsViaName, draws) &&
            !WorkbookAudit.callsWhere(expr, (_, _) => false, _.flags.boundsReferenceAxis),
          Right(line)
        )
    }
    Option.when(trimmedExtent.toLong == tail.kept && lines.forall(_._1))(lines.map(_._2))

  /**
   * The tail SUMPRODUCT adds to its trimmed sum: [[of]] its ranges under their shared bounds,
   * [[plan]]ned against the trimmed arguments' shape (`trimmedRows` × `trimmedCols`), then the
   * planned arguments read over the blank line, multiplied element-wise, summed and counted once
   * per dropped row. 0 when there is no tail, the plan declines, or the line does not come back as
   * one row (column) at the trimmed width. Each range location resolves once (`resolve`); `line`
   * reads a planned expression as SUMPRODUCT reads an argument, with the bound it gives a computed
   * reference (the trimmed read's shared extent) replaced by [[Tail.line]]; `rangeArray` reads a
   * range location on the line, `shape` an array's rows × columns, and `sumOfProducts` sums the
   * products over a shape.
   *
   * The plan moves only written whole columns onto the line. A reference-returning call over one
   * (`INDEX(B:C,0,1)`) keeps its range slot whole and computes its whole column at evaluation, so
   * the line bounds that column to the blank row, as the trimmed read bounds it to the kept rows.
   * Otherwise it would read the column's used rows: several rows fail the aligned check (the
   * trimmed sum stood), and one row passed it, so every blank row read that populated cell
   * (`SUMPRODUCT((A:A="")*(INDEX(B:C,0,1)>0))` over A1 = 4, B1 = 1 was 1048575, not 0).
   */
  def sum[A](
    ranges: List[CellRange],
    bounds: (Option[(Row, Row)], Option[(Column, Column)]),
    args: List[ArgSpec.SumProductArg],
    trimmedRows: Int,
    trimmedCols: Int,
    resolve: TExpr.RangeLocation => Either[EvalError, (Sheet, CellRange)],
    drawsViaName: (Option[SheetName], String) => Boolean,
    line: (TExpr[Any], CellRange => CellRange) => Either[EvalError, A],
    rangeArray: (Sheet, CellRange) => A,
    shape: A => (Int, Int),
    sumOfProducts: (List[A], Int, Int) => Either[EvalError, BigDecimal]
  ): Either[EvalError, BigDecimal] =
    val none: Either[EvalError, BigDecimal] = Right(BigDecimal(0))
    of(ranges, bounds).fold(none) { tail =>
      val located: Map[TExpr.RangeLocation, (Sheet, CellRange)] =
        args
          .collect { case Left(loc) => loc }
          .distinct
          .flatMap { loc =>
            resolve(loc).toOption.map(loc -> _)
          }
          .toMap
      plan(
        tail,
        args,
        if tail.alongRows then trimmedRows else trimmedCols,
        loc => located.get(loc).exists((_, range) => tail.spans(range)),
        drawsViaName
      ).fold(none) { lineArgs =>
        lineArgs
          .foldLeft[Either[EvalError, List[A]]](Right(List.empty)) {
            case (Left(err), _) => Left(err)
            // the plan accepted only locations that resolved and run the axis
            case (Right(acc), Left(loc)) =>
              located
                .get(loc)
                .toRight(EvalError.EvalFailed("SUMPRODUCT: unresolved blank line", None))
                .map((sheet, range) => acc :+ rangeArray(sheet, tail.line(range)))
            // the line evaluates every non-range part of the argument a second time (a lookup, a
            // criteria count): once per SUMPRODUCT, not per row, and never a draw, which the plan
            // rules out
            case (Right(acc), Right(expr)) => line(expr, tail.line).map(acc :+ _)
          }
          .flatMap { lines =>
            val (lineRows, lineCols) = lines.headOption.map(shape).getOrElse((0, 0))
            val aligned =
              lines.forall(shape(_) == (lineRows, lineCols)) &&
                (if tail.alongRows then
                   lineRows == 1 && (trimmedRows == 0 || lineCols == trimmedCols)
                 else lineCols == 1 && (trimmedCols == 0 || lineRows == trimmedRows))
            if aligned then sumOfProducts(lines, lineRows, lineCols).map(_ * BigDecimal(tail.count))
            else none
          }
      }
    }
