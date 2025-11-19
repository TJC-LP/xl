package com.tjclp.xl.styles.alignment

/** Cell alignment settings */
case class Align(
  horizontal: HAlign = HAlign.Left,
  vertical: VAlign = VAlign.Bottom,
  wrapText: Boolean = false,
  indent: Int = 0
):
  require(indent >= 0, s"Indent must be non-negative, got: $indent")

  def withHAlign(h: HAlign): Align = copy(horizontal = h)
  def withVAlign(v: VAlign): Align = copy(vertical = v)
  def withWrap(w: Boolean = true): Align = copy(wrapText = w)
  def withIndent(i: Int): Align = copy(indent = i)

object Align:
  val default: Align = Align()
