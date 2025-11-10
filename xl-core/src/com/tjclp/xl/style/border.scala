package com.tjclp.xl.style

/**
 * Cell border styles and configurations.
 */

// ========== Border ==========

/** Border line style */
enum BorderStyle:
  case None, Thin, Medium, Thick
  case Dashed, Dotted, Double
  case Hair, DashDot, DashDotDot
  case SlantDashDot, MediumDashed, MediumDashDot, MediumDashDotDot

/** Single border side */
case class BorderSide(
  style: BorderStyle = BorderStyle.None,
  color: Option[Color] = None
)

object BorderSide:
  val none: BorderSide = BorderSide()
  def apply(style: BorderStyle): BorderSide = BorderSide(style, None)
  def apply(style: BorderStyle, color: Color): BorderSide = BorderSide(style, Some(color))

/** Cell borders (all four sides) */
case class Border(
  left: BorderSide = BorderSide.none,
  right: BorderSide = BorderSide.none,
  top: BorderSide = BorderSide.none,
  bottom: BorderSide = BorderSide.none
):
  def withLeft(side: BorderSide): Border = copy(left = side)
  def withRight(side: BorderSide): Border = copy(right = side)
  def withTop(side: BorderSide): Border = copy(top = side)
  def withBottom(side: BorderSide): Border = copy(bottom = side)

object Border:
  val none: Border = Border()

  /** Create uniform border on all sides */
  def all(style: BorderStyle, color: Option[Color] = None): Border =
    val side = BorderSide(style, color)
    Border(side, side, side, side)
