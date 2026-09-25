package com.tjclp.xl.styles

import com.tjclp.xl.styles.border.{Border, BorderSide, BorderStyle}
import com.tjclp.xl.styles.color.Color
import com.tjclp.xl.styles.fill.Fill
import com.tjclp.xl.styles.font.Underline
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
) derives CanEqual:

  /**
   * This delta over a lower-precedence one, property by property (GH-497): each font attribute, the
   * fill, each border side and the number format come from this dxf when it sets them, else from
   * `lower`. How Excel composes the true conditional-format rules of one cell — the
   * higher-precedence rule wins a conflict and non-conflicting properties combine.
   *
   * A border side whose style is `None` counts as absent: the reader cannot tell an absent side
   * from an explicit `none` one (both parse to `BorderSide.none`).
   *
   * Laws: associative, with `Dxf()` as the identity on both sides.
   */
  def orElse(lower: Dxf): Dxf =
    Dxf(
      font = Dxf.both(font, lower.font)(_.orElse(_)),
      fill = fill.orElse(lower.fill),
      border = Dxf.both(border, lower.border)(Dxf.overlayBorder),
      numFmt = numFmt.orElse(lower.numFmt)
    )

  /**
   * Lay this delta over `base` (GH-497): a set property replaces the base's, an absent one keeps
   * it. `Some(false)` and `Some(Underline.None)` force the attribute off; the font's name and size
   * always come from the base (a dxf cannot express them); a number format clears the base's
   * `numFmtId`. Strike has no `Font` field, so renderers read it from the dxf itself.
   *
   * Law: a homomorphism from [[orElse]] — `(hi orElse lo).applyTo(s) == hi.applyTo(lo.applyTo(s))`
   * and `Dxf().applyTo(s) == s`.
   */
  def applyTo(base: CellStyle): CellStyle =
    val withFont = font.fold(base) { f =>
      base.withFont(
        base.font.copy(
          bold = f.bold.getOrElse(base.font.bold),
          italic = f.italic.getOrElse(base.font.italic),
          underline = f.underline.getOrElse(base.font.underline),
          color = f.color.orElse(base.font.color)
        )
      )
    }
    val withFill = fill.fold(withFont)(withFont.withFill)
    val withBorder =
      border.fold(withFill)(b => withFill.withBorder(Dxf.overlayBorder(b, withFill.border)))
    numFmt.fold(withBorder)(withBorder.withNumFmt)

object Dxf:
  /** Dxf overriding only the fill with a solid color (the common highlight case). */
  def fill(color: Color): Dxf = Dxf(fill = Some(Fill.Solid(color)))

  /** Dxf overriding only the font deltas. */
  def font(f: DxfFont): Dxf = Dxf(font = Some(f))

  /** Dxf overriding a solid fill plus font deltas. */
  def fillAndFont(fillColor: Color, f: DxfFont): Dxf =
    Dxf(font = Some(f), fill = Some(Fill.Solid(fillColor)))

  private def both[A](hi: Option[A], lo: Option[A])(merge: (A, A) => A): Option[A] =
    (hi, lo) match
      case (Some(h), Some(l)) => Some(merge(h, l))
      case _ => hi.orElse(lo)

  /** `hi`'s sides where their style is set, `lo`'s elsewhere. */
  private def overlayBorder(hi: Border, lo: Border): Border =
    def side(h: BorderSide, l: BorderSide): BorderSide =
      if h.style == BorderStyle.None then l else h
    Border(
      left = side(hi.left, lo.left),
      right = side(hi.right, lo.right),
      top = side(hi.top, lo.top),
      bottom = side(hi.bottom, lo.bottom)
    )

/**
 * Differential font. The core [[com.tjclp.xl.styles.font.Font]] REQUIRES name/sizePt and has no
 * strike field, so it cannot express "bold only, inherit the rest" — hence all-Option by design.
 * `Some(false)` is force-off (`<b val="0"/>`), `None` inherits from the base style. For underline,
 * `Some(Underline.None)` is force-off (`<u val="none"/>`) and `None` inherits.
 */
final case class DxfFont(
  bold: Option[Boolean] = None,
  italic: Option[Boolean] = None,
  strike: Option[Boolean] = None,
  underline: Option[Underline] = None,
  color: Option[Color] = None
) derives CanEqual:

  /** These deltas over lower-precedence ones, attribute by attribute (see [[Dxf.orElse]]). */
  def orElse(lower: DxfFont): DxfFont =
    DxfFont(
      bold = bold.orElse(lower.bold),
      italic = italic.orElse(lower.italic),
      strike = strike.orElse(lower.strike),
      underline = underline.orElse(lower.underline),
      color = color.orElse(lower.color)
    )
