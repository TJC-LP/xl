package com.tjclp.xl.styles

import com.tjclp.xl.styles.alignment.Align
import com.tjclp.xl.styles.border.Border
import com.tjclp.xl.styles.fill.Fill
import com.tjclp.xl.styles.font.Font
import com.tjclp.xl.styles.numfmt.NumFmt

/**
 * Complete cell style combining all formatting aspects.
 */

/**
 * Complete cell style combining all formatting aspects
 *
 * @param font
 *   Font styling
 * @param fill
 *   Background fill
 * @param border
 *   Cell borders
 * @param numFmt
 *   Number format (semantic type for programmatic use)
 * @param numFmtId
 *   Raw Excel format ID (for byte-perfect preservation from source files)
 * @param align
 *   Cell alignment
 */
final case class CellStyle(
  font: Font = Font.default,
  fill: Fill = Fill.default,
  border: Border = Border.none,
  numFmt: NumFmt = NumFmt.General,
  numFmtId: Option[Int] = None,
  align: Align = Align.default
):
  def withFont(f: Font): CellStyle = copy(font = f)
  def withFill(f: Fill): CellStyle = copy(fill = f)
  def withBorder(b: Border): CellStyle = copy(border = b)
  def withNumFmt(n: NumFmt): CellStyle = copy(numFmt = n, numFmtId = None)
  def withAlign(a: Align): CellStyle = copy(align = a)

  /** Set explicit numFmt ID with semantic type (advanced use) */
  def withNumFmtId(id: Int, fmt: NumFmt): CellStyle = copy(numFmt = fmt, numFmtId = Some(id))

  /**
   * Memoized canonical key for style deduplication.
   *
   * Two styles with the same key are structurally equivalent and should map to the same style index
   * in styles.xml. Computed once and cached for performance.
   */
  lazy val canonicalKey: String =
    val fontKey =
      s"F:${font.name},${font.sizePt},${font.bold},${font.italic},${font.underline},${font.color}"
    val fillKey = s"FL:${fill}"
    val borderKey =
      s"B:${border.left},${border.right},${border.top},${border.bottom}"
    val numFmtKey = s"N:${numFmt}"
    val alignKey =
      s"A:${align.horizontal},${align.vertical},${align.wrapText},${align.indent}"
    s"$fontKey|$fillKey|$borderKey|$numFmtKey|$alignKey"

object CellStyle:
  val default: CellStyle = CellStyle()

  /**
   * Generate canonical key for style deduplication (backward compatibility).
   *
   * Delegates to the memoized instance method for performance.
   */
  def canonicalKey(style: CellStyle): String = style.canonicalKey
