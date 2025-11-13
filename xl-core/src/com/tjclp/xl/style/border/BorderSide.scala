package com.tjclp.xl.style.border

import com.tjclp.xl.style.color.Color

/** Single border side */
case class BorderSide(
  style: BorderStyle = BorderStyle.None,
  color: Option[Color] = None
)

object BorderSide:
  val none: BorderSide = BorderSide()
  def apply(style: BorderStyle): BorderSide = BorderSide(style, None)
  def apply(style: BorderStyle, color: Color): BorderSide = BorderSide(style, Some(color))
