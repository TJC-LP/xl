package com.tjclp.xl.sheets

import com.tjclp.xl.styles.units.StyleId

/**
 * Properties for rows.
 *
 * @param height
 *   Row height in points (default Excel height is ~15)
 * @param hidden
 *   Whether the row is hidden
 * @param styleId
 *   Default style for cells in this row
 * @param outlineLevel
 *   Grouping/outline level (0-7) for collapsible sections
 * @param collapsed
 *   Whether this outline group is collapsed
 */
final case class RowProperties(
  height: Option[Double] = None,
  hidden: Boolean = false,
  styleId: Option[StyleId] = None,
  outlineLevel: Option[Int] = None,
  collapsed: Boolean = false
):
  require(
    outlineLevel.forall(l => l >= 0 && l <= 7),
    s"Outline level must be 0-7, got: ${outlineLevel.getOrElse(-1)}"
  )
