package com.tjclp.xl.formula.functions

import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row}
import com.tjclp.xl.formula.ast.TExpr

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
   *   - no argument calls a volatile function (`FunctionFlags.volatile`): reading the line
   *     evaluates every non-range part of the argument a second time, and `RAND()` must draw once.
   *     A defined name is not looked through.
   */
  def plan(
    tail: Tail,
    args: List[ArgSpec.SumProductArg],
    trimmedExtent: Int,
    spansLocation: TExpr.RangeLocation => Boolean
  ): Option[List[ArgSpec.SumProductArg]] =
    def spanning(expr: TExpr[?]): Int =
      TExpr.collectValueRanges(expr).count((_, r) => tail.spans(r))
    def volatile(expr: TExpr[?]): Boolean =
      com.tjclp.xl.formula.eval.WorkbookAudit.callsVolatile(expr, (_, _) => false)
    val lines = args.map {
      case Left(location) => (spansLocation(location), Left(location))
      case Right(expr) =>
        val line = TExpr.transformRanges(expr, (_, r) => tail.line(r))
        (spanning(expr) > 0 && spanning(line) == 0 && !volatile(expr), Right(line))
    }
    Option.when(trimmedExtent.toLong == tail.kept && lines.forall(_._1))(lines.map(_._2))
