package com.tjclp.xl.xml

import munit.FunSuite
import OwnedToken.*

class NormalizationSpec extends FunSuite:

  private def ok(xml: String): Vector[OwnedToken] =
    TestTokens.parse(xml) match
      case Right(ts) => ts
      case Left(e) => fail(s"expected success, got: ${e.render}")

  // --- line-end normalization (XML 1.0 §2.11) ------------------------------------------------------

  test("CRLF normalizes to LF in element content") {
    assertEquals(
      ok("<a>x\r\ny</a>"),
      Vector(Start("a", "", "a", Vector.empty, 1), Text("x\ny"), End("a"))
    )
  }

  test("lone CR normalizes to LF in element content") {
    assertEquals(
      ok("<a>x\ry</a>"),
      Vector(Start("a", "", "a", Vector.empty, 1), Text("x\ny"), End("a"))
    )
  }

  test("CRCRLF normalizes to two LFs") {
    assertEquals(
      ok("<a>x\r\r\ny</a>"),
      Vector(Start("a", "", "a", Vector.empty, 1), Text("x\n\ny"), End("a"))
    )
  }

  test("CRLF normalizes inside CDATA too") {
    assertEquals(
      ok("<a><![CDATA[x\r\ny]]></a>"),
      Vector(Start("a", "", "a", Vector.empty, 1), CData("x\ny"), End("a"))
    )
  }

  test("charref &#xD; survives normalization (decoded after line-end handling)") {
    assertEquals(
      ok("<a>x&#xD;y</a>"),
      Vector(Start("a", "", "a", Vector.empty, 1), Text("x\ry"), End("a"))
    )
  }

  test("the ST_Xstring contract: a raw CR cannot round-trip, the charref form can") {
    // This is exactly why escapeXstring writes _x000D_ (xml.scala GH-288): raw \r comes back \n.
    val raw = ok("<a>x\ry</a>")
    assertEquals(raw, Vector(Start("a", "", "a", Vector.empty, 1), Text("x\ny"), End("a")))
  }

  // --- attribute-value normalization (XML 1.0 §3.3.3) ------------------------------------------------

  test("raw tab in attribute value becomes a space") {
    assertEquals(
      ok("<a x=\"p\tq\"/>").headOption,
      Some(Start("a", "", "a", Vector("x" -> "p q"), 1))
    )
  }

  test("raw LF in attribute value becomes a space") {
    assertEquals(
      ok("<a x=\"p\nq\"/>").headOption,
      Some(Start("a", "", "a", Vector("x" -> "p q"), 1))
    )
  }

  test("raw CR in attribute value becomes a single space") {
    assertEquals(
      ok("<a x=\"p\rq\"/>").headOption,
      Some(Start("a", "", "a", Vector("x" -> "p q"), 1))
    )
  }

  test("raw CRLF in attribute value becomes ONE space (line-end normalization first)") {
    assertEquals(
      ok("<a x=\"p\r\nq\"/>").headOption,
      Some(Start("a", "", "a", Vector("x" -> "p q"), 1))
    )
  }

  test("charref-produced whitespace is exempt from attribute normalization (GH-429 premise)") {
    assertEquals(
      ok("""<a x="p&#x9;&#xA;&#xD;q"/>""").headOption,
      Some(Start("a", "", "a", Vector("x" -> "p\t\n\rq"), 1))
    )
  }

  // --- position tracking across normalization ---------------------------------------------------------

  test("CRLF counts as one line break for error positions") {
    val e = TestTokens.parse("<a>\r\n\r\n<b>\r\n</a>") match
      case Left(e) => e
      case Right(ts) => fail(s"expected failure, got $ts")
    assertEquals(e.line, 4)
  }

  test("lone CR counts as one line break for error positions") {
    val e = TestTokens.parse("<a>\r\r<b>\r</a>") match
      case Left(e) => e
      case Right(ts) => fail(s"expected failure, got $ts")
    assertEquals(e.line, 4)
  }
