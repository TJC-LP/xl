package com.tjclp.xl.sheets

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.Hyperlink

/**
 * One hyperlink over a range of cells: the `<hyperlink ref="A1:A5">` Excel writes when a link is
 * applied to a selection (#931).
 *
 * The cells the range covers do not carry the link themselves: [[Sheet.apply]] and
 * [[Sheet.hyperlinkAt]] report it for every covered cell without an own `Cell.hyperlink`, so a
 * whole-column link never materialises a million cells, and a read-modify-write of a covered cell
 * keeps the range intact.
 *
 * `excluded` holds the parts of `range` a removal cut out (Excel splits the range when one covered
 * cell loses its link). A cell with its own hyperlink overrides the range at that cell (#925).
 *
 * The reader builds these from a source worksheet; the writer keeps the source element for an entry
 * equal to it (range, link, tooltip and display), carving out the excluded parts and the cells with
 * own links, and writes any other entry as a new `<hyperlink>` over what is left.
 *
 * @param range
 *   the cells the link was applied to
 * @param link
 *   the target: a URL, or an internal location
 * @param tooltip
 *   the screen tip, when the link has one
 * @param display
 *   the `display` attribute, when the link has one
 * @param excluded
 *   parts of `range` the link no longer covers
 */
final case class RangeHyperlink(
  range: CellRange,
  link: Hyperlink,
  tooltip: Option[String] = None,
  display: Option[String] = None,
  excluded: Vector[CellRange] = Vector.empty
) derives CanEqual:

  /** Whether the link covers `ref`: inside `range` and outside every excluded part. */
  def covers(ref: ARef): Boolean =
    range.contains(ref) && !excluded.exists(_.contains(ref))

  /** The same link, ignoring what was excluded: the identity the writer matches a source by. */
  def sameLinkAs(other: RangeHyperlink): Boolean =
    range == other.range && link == other.link && tooltip == other.tooltip &&
      display == other.display

  /** This link with `cut` removed from its coverage, or None when nothing of it is left. */
  def without(cut: CellRange): Option[RangeHyperlink] =
    range.intersect(cut) match
      case None => Some(this)
      case Some(part) if part == range => None
      case Some(part) =>
        val rest = copy(excluded = excluded :+ part)
        Option.when(RangeHyperlink.remaining(range, rest.excluded).nonEmpty)(rest)

  /** The covered cells as disjoint rectangles in row-major order of their top-left corner. */
  def coverage: Vector[CellRange] = RangeHyperlink.remaining(range, excluded)

object RangeHyperlink:

  /**
   * `range` minus every rectangle in `holes`, as disjoint rectangles. Each subtraction splits a
   * rectangle into at most four (full-width bands above and below, runs left and right), so a
   * whole-column range never visits its cells.
   */
  def remaining(range: CellRange, holes: Vector[CellRange]): Vector[CellRange] =
    holes.foldLeft(Vector(range))((parts, hole) => parts.flatMap(subtract(_, hole)))

  private def subtract(part: CellRange, hole: CellRange): Vector[CellRange] =
    part.intersect(hole) match
      case None => Vector(part)
      case Some(cut) =>
        val left = part.colStart.index0
        val right = part.colEnd.index0
        val top = part.rowStart.index0
        val bottom = part.rowEnd.index0
        def rect(c1: Int, r1: Int, c2: Int, r2: Int): Option[CellRange] =
          Option.when(c1 <= c2 && r1 <= r2)(CellRange(ARef.from0(c1, r1), ARef.from0(c2, r2)))
        Vector(
          rect(left, top, right, cut.rowStart.index0 - 1),
          rect(left, cut.rowStart.index0, cut.colStart.index0 - 1, cut.rowEnd.index0),
          rect(cut.colEnd.index0 + 1, cut.rowStart.index0, right, cut.rowEnd.index0),
          rect(left, cut.rowEnd.index0 + 1, right, bottom)
        ).flatten
