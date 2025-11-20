package com.tjclp.xl.richtext

import com.tjclp.xl.styles.color.Color
import com.tjclp.xl.styles.font.Font

/**
 * Text run with optional font formatting.
 *
 * Maps directly to OOXML `<r>` (text run) element with `<rPr>` (run properties). Supports bold,
 * italic, underline, colors, font size, and font family.
 *
 * @param text
 *   The text content of this run
 * @param font
 *   Optional font formatting (if None, inherits from cell style or defaults)
 * @param rawRPrXml
 *   Preserved raw `<rPr>` XML as string for byte-perfect round-tripping. Used during surgical
 *   modification to preserve properties not in Font model (vertAlign, rFont, family, underline
 *   styles). Takes precedence over `font` during serialization.
 */
final case class TextRun(
  text: String,
  font: Option[Font] = None,
  rawRPrXml: Option[String] = None
):
  /** Append another TextRun or String to create RichText */
  def +(other: TextRun): RichText =
    RichText(Vector(this, other))

  /** Append a String to create RichText */
  def +(s: String): RichText =
    RichText(Vector(this, TextRun(s)))

  /** Append RichText to create combined RichText */
  def +(other: RichText): RichText =
    RichText(Vector(this) ++ other.runs)

  /** Create a new run with bold formatting */
  def bold: TextRun =
    copy(font = Some(font.getOrElse(Font.default).withBold(true)))

  /** Create a new run with italic formatting */
  def italic: TextRun =
    copy(font = Some(font.getOrElse(Font.default).withItalic(true)))

  /** Create a new run with underline formatting */
  def underline: TextRun =
    copy(font = Some(font.getOrElse(Font.default).withUnderline(true)))

  /** Create a new run with the specified font size in points */
  def size(pt: Double): TextRun =
    copy(font = Some(font.getOrElse(Font.default).withSize(pt)))

  /** Create a new run with the specified font family */
  def fontFamily(name: String): TextRun =
    copy(font = Some(font.getOrElse(Font.default).withName(name)))

  /** Create a new run with the specified color */
  def withColor(c: Color): TextRun =
    copy(font = Some(font.getOrElse(Font.default).withColor(c)))

  /** Create a new run with red color */
  def red: TextRun =
    Color.fromHex("#FF0000").map(withColor).getOrElse(this)

  /** Create a new run with green color */
  def green: TextRun =
    Color.fromHex("#00FF00").map(withColor).getOrElse(this)

  /** Create a new run with blue color */
  def blue: TextRun =
    Color.fromHex("#0000FF").map(withColor).getOrElse(this)

  /** Create a new run with black color */
  def black: TextRun =
    Color.fromHex("#000000").map(withColor).getOrElse(this)

  /** Create a new run with white color */
  def white: TextRun =
    Color.fromHex("#FFFFFF").map(withColor).getOrElse(this)
