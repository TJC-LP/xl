package com.tjclp.xl.ops

import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.alignment.{Align, HAlign, VAlign}
import com.tjclp.xl.styles.border.{Border, BorderSide}
import com.tjclp.xl.styles.color.Color
import com.tjclp.xl.styles.fill.Fill
import com.tjclp.xl.styles.font.{Font, Underline}
import com.tjclp.xl.styles.numfmt.NumFmt

/**
 * A partial cell style: every field `Some` is set, every field `None` is left as the target has it
 * (ADR-017 §2.12). `Option` fields make "un-bold" expressible, which the CLI's flag form never
 * could; border colour is per-side data ([[com.tjclp.xl.styles.border.BorderSide]] carries its own
 * colour), never "the colour applies to every side being set" — that rule broke the action law.
 *
 * Laws (EditModelSpec): `++` is a right-biased monoid with [[StyleOverlay.empty]] as identity, and
 * it acts on styles — `(a ++ b).applyTo(s) == b.applyTo(a.applyTo(s))`; `of(s).applyTo(default)
 * == s` for every style (`numFmtId` excepted: it is writer-assigned and outside the canonical key).
 */
final case class StyleOverlay(
  fontName: Option[String] = None,
  fontSize: Option[Double] = None,
  bold: Option[Boolean] = None,
  italic: Option[Boolean] = None,
  underline: Option[Underline] = None,
  fontColor: Option[Color] = None,
  fill: Option[Fill] = None,
  numFmt: Option[NumFmt] = None,
  hAlign: Option[HAlign] = None,
  vAlign: Option[VAlign] = None,
  wrap: Option[Boolean] = None,
  indent: Option[Int] = None,
  textRotation: Option[Int] = None,
  borderTop: Option[BorderSide] = None,
  borderRight: Option[BorderSide] = None,
  borderBottom: Option[BorderSide] = None,
  borderLeft: Option[BorderSide] = None
) derives CanEqual:

  /** Right-biased merge: `other`'s set fields win. */
  infix def ++(other: StyleOverlay): StyleOverlay =
    StyleOverlay(
      fontName = other.fontName.orElse(fontName),
      fontSize = other.fontSize.orElse(fontSize),
      bold = other.bold.orElse(bold),
      italic = other.italic.orElse(italic),
      underline = other.underline.orElse(underline),
      fontColor = other.fontColor.orElse(fontColor),
      fill = other.fill.orElse(fill),
      numFmt = other.numFmt.orElse(numFmt),
      hAlign = other.hAlign.orElse(hAlign),
      vAlign = other.vAlign.orElse(vAlign),
      wrap = other.wrap.orElse(wrap),
      indent = other.indent.orElse(indent),
      textRotation = other.textRotation.orElse(textRotation),
      borderTop = other.borderTop.orElse(borderTop),
      borderRight = other.borderRight.orElse(borderRight),
      borderBottom = other.borderBottom.orElse(borderBottom),
      borderLeft = other.borderLeft.orElse(borderLeft)
    )

  def isEmpty: Boolean = this == StyleOverlay.empty

  /** Overlay onto `style`: set fields replace, unset fields keep the style's value. */
  def applyTo(style: CellStyle): CellStyle =
    val f = style.font
    val font = Font(
      name = fontName.getOrElse(f.name),
      sizePt = fontSize.getOrElse(f.sizePt),
      bold = bold.getOrElse(f.bold),
      italic = italic.getOrElse(f.italic),
      underline = underline.getOrElse(f.underline),
      color = fontColor.orElse(f.color)
    )
    val a = style.align
    val align = Align(
      horizontal = hAlign.getOrElse(a.horizontal),
      vertical = vAlign.getOrElse(a.vertical),
      wrapText = wrap.getOrElse(a.wrapText),
      indent = indent.getOrElse(a.indent),
      textRotation = textRotation.getOrElse(a.textRotation)
    )
    val b = style.border
    val border = Border(
      left = borderLeft.getOrElse(b.left),
      right = borderRight.getOrElse(b.right),
      top = borderTop.getOrElse(b.top),
      bottom = borderBottom.getOrElse(b.bottom)
    )
    val withParts = style.withFont(font).withFill(fill.getOrElse(style.fill)).withBorder(border)
    numFmt.fold(withParts)(withParts.withNumFmt).withAlign(align)

  /**
   * The guards the model enforces with `require`, checked ahead of [[applyTo]] so an interpreter
   * never throws: a positive font size, a non-empty font name, a non-negative indent, a rotation in
   * ST_TextRotation (0-180 or 255).
   */
  def validate: XLResult[Unit] =
    def refuse(reason: String): XLResult[Unit] = Left(XLError.InvalidArgument("style", reason))
    if fontSize.exists(_ <= 0) then
      refuse(s"font size must be positive, got: ${fontSize.getOrElse(0.0)}")
    else if fontName.exists(_.isEmpty) then refuse("font name cannot be empty")
    else if indent.exists(_ < 0) then
      refuse(s"indent must be non-negative, got: ${indent.getOrElse(0)}")
    else if textRotation.exists(r => !((r >= 0 && r <= 180) || r == Align.VerticalTextRotation))
    then refuse(s"textRotation must be 0-180 or 255, got: ${textRotation.getOrElse(0)}")
    else Right(())

object StyleOverlay:
  /** Sets nothing: the identity of `++` and of `applyTo`. */
  val empty: StyleOverlay = StyleOverlay()

  /** Every field set from `style`, so `of(style).applyTo(x) == style` (up to `numFmtId`). */
  def of(style: CellStyle): StyleOverlay =
    StyleOverlay(
      fontName = Some(style.font.name),
      fontSize = Some(style.font.sizePt),
      bold = Some(style.font.bold),
      italic = Some(style.font.italic),
      underline = Some(style.font.underline),
      fontColor = style.font.color,
      fill = Some(style.fill),
      numFmt = Some(style.numFmt),
      hAlign = Some(style.align.horizontal),
      vAlign = Some(style.align.vertical),
      wrap = Some(style.align.wrapText),
      indent = Some(style.align.indent),
      textRotation = Some(style.align.textRotation),
      borderTop = Some(style.border.top),
      borderRight = Some(style.border.right),
      borderBottom = Some(style.border.bottom),
      borderLeft = Some(style.border.left)
    )

  /**
   * The sides of `border` that are set (not `BorderSide.none`), as an overlay — the additive
   * semantics of `Patch.MergeBorder` / `Border.merge`.
   */
  def ofBorder(border: Border): StyleOverlay =
    def side(s: BorderSide): Option[BorderSide] = Option.when(s != BorderSide.none)(s)
    StyleOverlay(
      borderTop = side(border.top),
      borderRight = side(border.right),
      borderBottom = side(border.bottom),
      borderLeft = side(border.left)
    )
