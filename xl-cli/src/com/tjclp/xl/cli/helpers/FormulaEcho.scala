package com.tjclp.xl.cli.helpers

import com.tjclp.xl.formula.ParseError

/**
 * The formula a putf gate quotes back when the parser refuses it (GH-681): the parser's own
 * diagnostic ([[ParseError.formatWithContext]]: the formula, a caret under the offending position,
 * the reason) with the formula capped at [[SampleChars]] — the sample `xl lint` quotes for an
 * unparseable `<f>` — so an 8193-character op is not an 8 KB message line plus an 8 KB caret line.
 *
 *   - within the sample: the parser's rendering, unchanged
 *   - `FormulaTooLong`: the first [[SampleChars]] characters and the reason, no caret (the length
 *     is the fault; no position in the text is)
 *   - longer, with a caret: the [[SampleChars]]-character window around the caret, `…` marking each
 *     cut side, the caret moved with it so it stays under the same character
 *
 * The caret column is read from the parser's rendering rather than recomputed, so the rendering's
 * own positioning stays the one rule.
 */
object FormulaEcho:

  /** The echo cap: the 80 characters `WorkbookLint` quotes of an unparseable `<f>`. */
  val SampleChars: Int = 80

  private val Ellipsis = "…"

  def diagnostic(error: ParseError, fullFormula: String): String =
    val rendered = ParseError.formatWithContext(error, fullFormula)
    error match
      case _: ParseError.FormulaTooLong =>
        s"${sample(fullFormula)}\n${ParseError.describe(error)}"
      case _ if fullFormula.length <= SampleChars => rendered
      // a rendering that does not lead with the formula: quote the head, never the whole text
      case _ if !rendered.startsWith(fullFormula + "\n") =>
        s"${sample(fullFormula)}\n${ParseError.describe(error)}"
      case _ =>
        val below = rendered.stripPrefix(fullFormula + "\n")
        below.split("\n", 2).toList match
          case pointer :: rest if pointer.nonEmpty && pointer.trim == "^" =>
            val column = pointer.length - 1
            val start = whole(
              fullFormula,
              math.max(0, math.min(column - SampleChars / 2, fullFormula.length - SampleChars))
            )
            val end = whole(fullFormula, math.min(start + SampleChars, fullFormula.length))
            val lead = if start > 0 then Ellipsis else ""
            val tail = if end < fullFormula.length then Ellipsis else ""
            val shown = lead + fullFormula.substring(start, end) + tail
            val caret = " " * (column - start + lead.length) + "^"
            (shown :: caret :: rest).mkString("\n")
          case _ => s"${sample(fullFormula)}\n$below"

  /** The first [[SampleChars]] characters, `…` when that cut the text. */
  def sample(text: String): String =
    if text.length > SampleChars then text.take(whole(text, SampleChars)) + Ellipsis else text

  /**
   * `cut` moved back one char when it falls between the halves of a surrogate pair, so a quoted
   * emoji is whole or absent — never a lone surrogate that the console encodes as `?`.
   */
  private def whole(text: String, cut: Int): Int =
    val splitsPair = cut > 0 && cut < text.length &&
      Character.isHighSurrogate(text.charAt(cut - 1)) &&
      Character.isLowSurrogate(text.charAt(cut))
    if splitsPair then cut - 1 else cut
