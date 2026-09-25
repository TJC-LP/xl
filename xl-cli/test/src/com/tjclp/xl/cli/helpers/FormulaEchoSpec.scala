package com.tjclp.xl.cli.helpers

import munit.FunSuite

import com.tjclp.xl.formula.{FormulaParser, ParseError}

/**
 * GH-681: the putf gates echoed the whole formula under the caret, so an 8193-character op made an
 * 8 KB message line plus an 8 KB caret line. The echo is capped at the lint rule's 80-character
 * sample; a long formula shows the window around the caret, and `FormulaTooLong` shows no caret.
 */
class FormulaEchoSpec extends FunSuite:

  private def failure(full: String): ParseError =
    FormulaParser.parse(full).swap.getOrElse(fail(s"the parser accepted $full"))

  /** The caret's column in a rendering, and the line above it. */
  private def caret(rendered: String): Option[(String, Int)] =
    rendered.split("\n", -1).toList.sliding(2).collectFirst {
      case List(above, pointer) if pointer.nonEmpty && pointer.trim == "^" =>
        (above, pointer.length - 1)
    }

  test("a formula within the sample renders exactly as the parser's own diagnostic") {
    val full = "=SUM(A1:A2"
    val error = failure(full)
    assertEquals(FormulaEcho.diagnostic(error, full), ParseError.formatWithContext(error, full))
    val unknown = "=FOOBAR(1)"
    assertEquals(
      FormulaEcho.diagnostic(failure(unknown), unknown),
      ParseError.formatWithContext(failure(unknown), unknown)
    )
  }

  test("FormulaTooLong: the first 80 characters and the reason, no caret") {
    val full = "=SUM(" + ("A1," * 2800) + "A1)"
    val error = failure(full)
    error match
      case _: ParseError.FormulaTooLong => ()
      case other => fail(s"expected FormulaTooLong, got $other")
    val rendered = FormulaEcho.diagnostic(error, full)
    assertEquals(
      rendered,
      s"${full.take(FormulaEcho.SampleChars)}…\n${ParseError.describe(error)}"
    )
    assertEquals(caret(rendered), None)
  }

  test("a long formula shows the 80-character window around the caret, pointing where it did") {
    // the unbalanced `)` sits near the end of a 300-character formula
    val full = "=" + ("A1+" * 100) + "1)"
    val error = failure(full)
    val (_, column) =
      caret(ParseError.formatWithContext(error, full)).getOrElse(fail("the parser gave no caret"))
    val rendered = FormulaEcho.diagnostic(error, full)
    val (shown, windowed) = caret(rendered).getOrElse(fail(s"no caret in\n$rendered"))
    assert(shown.startsWith("…"), shown)
    assert(shown.length <= FormulaEcho.SampleChars + 2, shown)
    // the caret sits under the same character of the formula as the uncapped rendering's did
    assertEquals(shown.lift(windowed), full.lift(column))
    assert(rendered.endsWith(ParseError.describe(error)), rendered)
    assert(rendered.length < 300, rendered)
  }

  test("a long formula failing near its start keeps the head, with a trailing ellipsis") {
    val full = "=FOOBAR(" + ("A1," * 100) + "A1)"
    val error = failure(full)
    val (_, column) =
      caret(ParseError.formatWithContext(error, full)).getOrElse(fail("the parser gave no caret"))
    val (shown, windowed) =
      caret(FormulaEcho.diagnostic(error, full)).getOrElse(fail("no caret"))
    assertEquals(shown, full.take(FormulaEcho.SampleChars) + "…")
    assertEquals(shown.lift(windowed), full.lift(column))
  }

  test("a cut never splits a surrogate pair: an emoji is quoted whole or not at all") {
    def lone(text: String): Boolean =
      text.indices.exists { i =>
        val c = text(i)
        (Character.isHighSurrogate(c) && !text.lift(i + 1).exists(Character.isLowSurrogate)) ||
        (Character.isLowSurrogate(c) && !text.lift(i - 1).exists(Character.isHighSurrogate))
      }
    val emoji = "😀" * 100
    // either parity of padding puts one of the cuts between a high and a low surrogate
    for pad <- Vector("", "a") do
      val tail = "=\"" + pad + emoji + "\"&)"
      val head = "=FOOBAR(\"" + pad + emoji + "\")"
      val tooLong = "=\"" + pad + ("😀" * 4200) + "\""
      Vector(tail, head, tooLong).foreach { full =>
        val error = failure(full)
        val rendered = FormulaEcho.diagnostic(error, full)
        assert(!lone(rendered), s"a lone surrogate in\n$rendered")
        caret(ParseError.formatWithContext(error, full)).zip(caret(rendered)).foreach {
          case ((_, column), (shown, windowed)) =>
            assertEquals(shown.lift(windowed), full.lift(column))
        }
      }
      assert(!lone(FormulaEcho.sample(pad + emoji)), pad)
  }
