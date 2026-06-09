package com.tjclp.xl.styles.border

import com.tjclp.xl.styles.color.Color

/** Cell borders (all four sides) */
final case class Border(
  left: BorderSide = BorderSide.none,
  right: BorderSide = BorderSide.none,
  top: BorderSide = BorderSide.none,
  bottom: BorderSide = BorderSide.none
):
  def withLeft(side: BorderSide): Border = copy(left = side)
  def withRight(side: BorderSide): Border = copy(right = side)
  def withTop(side: BorderSide): Border = copy(top = side)
  def withBottom(side: BorderSide): Border = copy(bottom = side)

  /**
   * Merge another border into this one, side by side (additive overlay).
   *
   * Sides that are set (non-none) in `overlay` replace the corresponding side of this border; sides
   * that are `BorderSide.none` in `overlay` leave the existing side untouched. This mirrors the
   * CLI's additive border semantics (`--border-top` only touches the top side).
   *
   * Laws: associative, idempotent, with `Border.none` as identity.
   */
  def merge(overlay: Border): Border =
    Border(
      left = if overlay.left == BorderSide.none then left else overlay.left,
      right = if overlay.right == BorderSide.none then right else overlay.right,
      top = if overlay.top == BorderSide.none then top else overlay.top,
      bottom = if overlay.bottom == BorderSide.none then bottom else overlay.bottom
    )

object Border:
  val none: Border = Border()

  /** Create uniform border on all sides */
  def all(style: BorderStyle, color: Option[Color] = None): Border =
    val side = BorderSide(style, color)
    Border(side, side, side, side)
