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
   * it). None — no tail — unless both hold:
   *   - the trimmed arguments' extent along the axis (`trimmedExtent`) is the kept one. A
   *     shape-sensitive call keeps its whole column whole (`ROW(A:A)`), so its argument already
   *     read every row at the full height (or mismatched the others, an error before any tail):
   *     reading it again would rebuild the million rows only to discard them.
   *   - every argument reads a trimmed whole column as an element source: a range location that
   *     runs the axis (`spansLocation`), or an expression whose transform moved at least one such
   *     range onto the line. `SUMPRODUCT(SUM(A:A))` moves none: its one value is no row's.
   */
  def plan(
    tail: Tail,
    args: List[ArgSpec.SumProductArg],
    trimmedExtent: Int,
    spansLocation: TExpr.RangeLocation => Boolean
  ): Option[List[ArgSpec.SumProductArg]] =
    def spanning(expr: TExpr[?]): Int = TExpr.collectRanges(expr).count((_, r) => tail.spans(r))
    val lines = args.map {
      case Left(location) => (spansLocation(location), Left(location))
      case Right(expr) =>
        val line = TExpr.transformRanges(expr, (_, r) => tail.line(r))
        (spanning(line) < spanning(expr), Right(line))
    }
    Option.when(trimmedExtent.toLong == tail.kept && lines.forall(_._1))(lines.map(_._2))
