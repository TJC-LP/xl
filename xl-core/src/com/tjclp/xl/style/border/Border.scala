package com.tjclp.xl.style.border

import com.tjclp.xl.style.color.Color

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
