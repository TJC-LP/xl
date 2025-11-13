package com.tjclp.xl.style.font

import com.tjclp.xl.style.color.Color

/** Font styling for cell text */
case class Font(
  name: String = "Calibri",
  sizePt: Double = 11.0,
  bold: Boolean = false,
  italic: Boolean = false,
  underline: Boolean = false,
  color: Option[Color] = None
):
  require(sizePt > 0, s"Font size must be positive, got: $sizePt")
  require(name.nonEmpty, "Font name cannot be empty")

  def withName(n: String): Font = copy(name = n)
  def withSize(size: Double): Font = copy(sizePt = size)
  def withBold(b: Boolean = true): Font = copy(bold = b)
  def withItalic(i: Boolean = true): Font = copy(italic = i)
  def withUnderline(u: Boolean = true): Font = copy(underline = u)
  def withColor(c: Color): Font = copy(color = Some(c))
  def clearColor: Font = copy(color = None)

object Font:
  val default: Font = Font()
