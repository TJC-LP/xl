package com.tjclp.xl.sheets

import com.tjclp.xl.styles.units.StyleId

/**
 * Properties for columns.
 *
 * @param width
 *   Column width in character units (default Excel width is ~8.43)
 * @param hidden
 *   Whether the column is hidden
 * @param styleId
 *   Default style for cells in this column
 * @param outlineLevel
 *   Grouping/outline level (0-7) for collapsible sections
 * @param collapsed
 *   Whether this outline group is collapsed
 * @param customWidth
 *   OOXML `customWidth` (#869 review): whether [[width]] was set on purpose (`col --width`,
 *   `autofit`, the API) rather than by the widening a `put`/`putf` of a formatted number does, as
 *   Excel's entry does. Only a set width is written as `customWidth="1"`. A width that is not
 *   custom may be widened again by a later entry, never narrowed; a custom one is never changed.
 *   Set a width with [[withWidth]], which marks it custom.
 */
final case class ColumnProperties(
  width: Option[Double] = None,
  hidden: Boolean = false,
  styleId: Option[StyleId] = None,
  outlineLevel: Option[Int] = None,
  collapsed: Boolean = false,
  customWidth: Boolean = true
):
  require(
    outlineLevel.forall(l => l >= 0 && l <= 7),
    s"Outline level must be 0-7, got: ${outlineLevel.getOrElse(-1)}"
  )

  /** These properties with `w` as a width set on purpose ([[customWidth]]). */
  def withWidth(w: Double): ColumnProperties = copy(width = Some(w), customWidth = true)
