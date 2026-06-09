package com.tjclp.xl.styles.alignment

/**
 * Cell alignment settings.
 *
 * @param horizontal
 *   Horizontal alignment (Excel default: General, content-aware)
 * @param vertical
 *   Vertical alignment (Excel default: Bottom)
 * @param wrapText
 *   Whether text wraps within the cell
 * @param indent
 *   Excel's alignment indent level (~3 characters per level); 0 means no indent and is omitted from
 *   OOXML output. Indentation travels with the style rather than the value, unlike leading spaces.
 */
final case class Align(
  horizontal: HAlign = HAlign.General,
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
