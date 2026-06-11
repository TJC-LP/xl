package com.tjclp.xl.styles

import com.tjclp.xl.styles.border.Border
import com.tjclp.xl.styles.color.Color
import com.tjclp.xl.styles.fill.Fill
import com.tjclp.xl.styles.numfmt.NumFmt

/**
 * Differential format (styles.xml `<dxf>`, CT_Dxf): present fields override the base cell style;
 * absent fields inherit. Used by conditional formatting (GH-136) — a matching rule applies its dxf
 * on top of the cell's existing style.
 *
 * Modeled subset: font deltas, solid fills, the four plain border sides, and number formats.
 * Alignment, protection, and gradient fills are deliberately out of scope; a dxf using them rides
 * through as a Preserved rule on the OOXML layer.
 */
final case class Dxf(
  font: Option[DxfFont] = None,
  fill: Option[Fill] = None,
  border: Option[Border] = None,
  numFmt: Option[NumFmt] = None
) derives CanEqual

/**
 * Differential font. The core [[com.tjclp.xl.styles.font.Font]] REQUIRES name/sizePt and has no
 * strike field, so it cannot express "bold only, inherit the rest" — hence all-Option by design.
 * `Some(false)` is force-off (`<b val="0"/>`), `None` inherits from the base style.
 */
final case class DxfFont(
  bold: Option[Boolean] = None,
  italic: Option[Boolean] = None,
  strike: Option[Boolean] = None,
  underline: Option[Boolean] = None,
  color: Option[Color] = None
) derives CanEqual

object Dxf:
  /** Dxf overriding only the fill with a solid color (the common highlight case). */
  def fill(color: Color): Dxf = Dxf(fill = Some(Fill.Solid(color)))

  /** Dxf overriding only the font deltas. */
  def font(f: DxfFont): Dxf = Dxf(font = Some(f))

  /** Dxf overriding a solid fill plus font deltas. */
  def fillAndFont(fillColor: Color, f: DxfFont): Dxf =
    Dxf(font = Some(f), fill = Some(Fill.Solid(fillColor)))
