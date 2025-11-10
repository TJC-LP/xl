package com.tjclp.xl.style

/**
 * Complete cell style combining all formatting aspects.
 */

// ========== CellStyle ==========

/** Complete cell style combining all formatting aspects */
case class CellStyle(
  font: Font = Font.default,
  fill: Fill = Fill.default,
  border: Border = Border.none,
  numFmt: NumFmt = NumFmt.General,
  align: Align = Align.default
):
  def withFont(f: Font): CellStyle = copy(font = f)
  def withFill(f: Fill): CellStyle = copy(fill = f)
  def withBorder(b: Border): CellStyle = copy(border = b)
  def withNumFmt(n: NumFmt): CellStyle = copy(numFmt = n)
  def withAlign(a: Align): CellStyle = copy(align = a)

object CellStyle:
  val default: CellStyle = CellStyle()

  /**
   * Generate canonical key for style deduplication.
   *
   * Two styles with the same key are structurally equivalent and should map to the same style index
   * in styles.xml.
   */
  def canonicalKey(style: CellStyle): String =
    val fontKey =
      s"F:${style.font.name},${style.font.sizePt},${style.font.bold},${style.font.italic},${style.font.underline},${style.font.color}"
    val fillKey = s"FL:${style.fill}"
    val borderKey =
      s"B:${style.border.left},${style.border.right},${style.border.top},${style.border.bottom}"
    val numFmtKey = s"N:${style.numFmt}"
    val alignKey =
      s"A:${style.align.horizontal},${style.align.vertical},${style.align.wrapText},${style.align.indent}"
    s"$fontKey|$fillKey|$borderKey|$numFmtKey|$alignKey"
