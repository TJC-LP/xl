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
 * @param textRotation
 *   Raw OOXML ST_TextRotation value: 0 is horizontal (omitted from output), 1-90 rotates
 *   counter-clockwise by that many degrees (90 = straight up), 91-180 encodes DOWNWARD rotation as
 *   90 + the clockwise angle (135 = 45 degrees down), and 255 stacks letters vertically. The
 *   [[Align.normalizeRotation]] helper (and the `.rotated` style DSL) accepts Excel-UI-style
 *   negative degrees and maps them onto this encoding.
 */
final case class Align(
  horizontal: HAlign = HAlign.General,
  vertical: VAlign = VAlign.Bottom,
  wrapText: Boolean = false,
  indent: Int = 0,
  textRotation: Int = 0
):
  require(indent >= 0, s"Indent must be non-negative, got: $indent")
  require(
    (textRotation >= 0 && textRotation <= 180) || textRotation == Align.VerticalTextRotation,
    s"textRotation must be 0-180 or 255 (ST_TextRotation), got: $textRotation"
  )

  def withHAlign(h: HAlign): Align = copy(horizontal = h)
  def withVAlign(v: VAlign): Align = copy(vertical = v)
  def withWrap(w: Boolean = true): Align = copy(wrapText = w)
  def withIndent(i: Int): Align = copy(indent = i)
  def withRotation(deg: Int): Align = copy(textRotation = deg)

object Align:
  val default: Align = Align()

  /** ST_TextRotation sentinel for vertically stacked letters. */
  val VerticalTextRotation: Int = 255

  /**
   * Normalize a user-facing rotation into a valid ST_TextRotation value (total, never throws).
   *
   * Excel-UI-style degrees are accepted: positive 1-90 rotates counter-clockwise (90 = straight
   * up), negative -1 to -90 rotates clockwise and is encoded per OOXML as `90 + |deg|` (91-180),
   * and 255 stacks letters vertically. Raw values 91-180 (the downward encoding) pass through
   * unchanged; anything else is clamped to the nearest valid rotation.
   */
  def normalizeRotation(deg: Int): Int =
    if deg == VerticalTextRotation then VerticalTextRotation
    else
      val clamped = math.max(-90, math.min(180, deg))
      if clamped < 0 then 90 - clamped else clamped
