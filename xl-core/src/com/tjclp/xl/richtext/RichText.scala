package com.tjclp.xl.richtext

import com.tjclp.xl.styles.color.Color

/**
 * Rich text cell value with multiple formatting runs.
 *
 * Allows different formatting within a single cell (e.g., "Bold text and normal text" where "Bold"
 * is bold but "and normal text" is not). Maps to OOXML `<is>` (inline string) with multiple `<r>`
 * (run) elements.
 *
 * Example:
 * {{{
 * import com.tjclp.xl.richtext.RichText.*
 * val text = "Bold".bold.red + " and " + "Italic".italic.blue
 * }}}
 *
 * @param runs
 *   Vector of text runs with individual formatting
 */
final case class RichText(runs: Vector[TextRun]):
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
