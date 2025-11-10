package com.tjclp.xl

import com.tjclp.xl.style.{Font, Color}

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
 */
case class TextRun(
  text: String,
  font: Option[Font] = None
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

/**
 * Rich text cell value with multiple formatting runs.
 *
 * Allows different formatting within a single cell (e.g., "Bold text and normal text" where "Bold"
 * is bold but "and normal text" is not). Maps to OOXML `<is>` (inline string) with multiple `<r>`
 * (run) elements.
 *
 * Example:
 * {{{
 * import com.tjclp.xl.RichText.*
 * val text = "Bold".bold.red + " and " + "Italic".italic.blue
 * }}}
 *
 * @param runs
 *   Vector of text runs with individual formatting
 */
case class RichText(runs: Vector[TextRun]):
  /** Append another RichText (concatenate runs) */
  def +(other: RichText): RichText =
    RichText(runs ++ other.runs)

  /** Append a single TextRun */
  def +(run: TextRun): RichText =
    RichText(runs :+ run)

  /** Append a String (creates plain TextRun) */
  def +(s: String): RichText =
    RichText(runs :+ TextRun(s))

  /** Convert to plain text (strips all formatting) */
  def toPlainText: String =
    runs.map(_.text).mkString

  /** Check if this is effectively plain text (all runs have no formatting) */
  def isPlainText: Boolean =
    runs.forall(_.font.isEmpty)

object RichText:
  /** Create RichText from multiple runs */
  def apply(runs: TextRun*): RichText =
    RichText(runs.toVector)

  /** Create RichText from a single text run */
  def single(run: TextRun): RichText =
    RichText(Vector(run))

  /** Create plain RichText from a string */
  def plain(text: String): RichText =
    RichText(Vector(TextRun(text)))

  // ========== Given Conversions for Ergonomics ==========

  /** Convert String to TextRun (for DSL ergonomics) */
  given Conversion[String, TextRun] = TextRun(_)

  /** Convert TextRun to RichText (for composition) */
  given Conversion[TextRun, RichText] = r => RichText(Vector(r))

  // ========== String Extensions for DSL ==========

  extension (s: String)
    /** Create a bold text run */
    def bold: TextRun = TextRun(s).bold

    /** Create an italic text run */
    def italic: TextRun = TextRun(s).italic

    /** Create an underlined text run */
    def underline: TextRun = TextRun(s).underline

    /** Create a text run with the specified size in points */
    def size(pt: Double): TextRun = TextRun(s).size(pt)

    /** Create a text run with the specified font family */
    def fontFamily(name: String): TextRun = TextRun(s).fontFamily(name)

    /** Create a text run with the specified color */
    def withColor(c: Color): TextRun = TextRun(s).withColor(c)

    /** Create a red text run */
    def red: TextRun = TextRun(s).red

    /** Create a green text run */
    def green: TextRun = TextRun(s).green

    /** Create a blue text run */
    def blue: TextRun = TextRun(s).blue

    /** Create a black text run */
    def black: TextRun = TextRun(s).black

    /** Create a white text run */
    def white: TextRun = TextRun(s).white
