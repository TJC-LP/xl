package com.tjclp.xl.xml

import munit.FunSuite
import OwnedToken.*

class EntityCharrefSpec extends FunSuite:

  private def textOf(xml: String): String =
    TestTokens.parse(xml) match
      case Right(Vector(Start(_, _, _, _, _), Text(v), End(_))) => v
      case other => fail(s"expected single text doc, got: $other")

  private def attrOf(xml: String): String =
    TestTokens.parse(xml) match
      case Right(ts) =>
        ts.headOption match
          case Some(Start(_, _, _, attrs, _)) =>
            attrs.headOption.map(_._2).getOrElse(fail("no attribute"))
          case other => fail(s"expected start, got $other")
      case Left(e) => fail(s"expected success: ${e.render}")

  private def bad(xml: String): XmlParseError =
    TestTokens.parse(xml) match
      case Left(e) => e
      case Right(ts) => fail(s"expected failure, got tokens: $ts")

  // --- the five predefined entities -------------------------------------------------------------------

  test("five-entity matrix in element content") {
    assertEquals(textOf("<a>&amp;&lt;&gt;&apos;&quot;</a>"), "&<>'\"")
  }

  test("five-entity matrix in attribute values") {
    assertEquals(attrOf("""<a x="&amp;&lt;&gt;&apos;&quot;"/>"""), "&<>'\"")
  }

  test("entities decode inside a longer run and coalesce with surrounding text") {
    assertEquals(textOf("<a>1 &lt; 2 &amp;&amp; 3 &gt; 2</a>"), "1 < 2 && 3 > 2")
  }

  // --- character references ---------------------------------------------------------------------------

  test("decimal and hex character references") {
    assertEquals(textOf("<a>&#65;&#x42;&#x63;</a>"), "ABc")
  }

  test("hex digits are case-insensitive, 'x' marker is not") {
    assertEquals(textOf("<a>&#xAb;&#xaB;</a>"), "««")
    val e = bad("<a>&#X41;</a>")
    assert(e.message.nonEmpty)
  }

  test("supplementary-plane character reference produces a surrogate pair") {
    assertEquals(textOf("<a>&#x1F600;</a>"), "😀")
  }

  test("boundary Chars accepted: 0x9 0xA 0xD 0x20 0xD7FF 0xE000 0xFFFD 0x10FFFF") {
    for cp <- List(0x9, 0xa, 0xd, 0x20, 0xd7ff, 0xe000, 0xfffd, 0x10ffff) do
      val r = TestTokens.parse(s"<a>&#x${cp.toHexString};</a>")
      assert(r.isRight, s"&#x${cp.toHexString}; should be accepted: $r")
  }

  test("&#0; is rejected (not an XML Char)") {
    val e = bad("<a>&#0;</a>")
    assert(e.message.contains("invalid character reference"), e.message)
  }

  test("control characters below 0x20 (other than tab/LF/CR) are rejected") {
    for cp <- List(0x1, 0x8, 0xb, 0xc, 0xe, 0x1f) do
      val r = TestTokens.parse(s"<a>&#$cp;</a>")
      assert(r.isLeft, s"&#$cp; must be rejected")
  }

  test("surrogate code points are rejected: &#xD800; &#xDFFF;") {
    assert(bad("<a>&#xD800;</a>").message.contains("invalid character reference"))
    assert(bad("<a>&#xDFFF;</a>").message.contains("invalid character reference"))
  }

  test("0xFFFE and 0xFFFF are rejected") {
    assert(TestTokens.parse("<a>&#xFFFE;</a>").isLeft)
    assert(TestTokens.parse("<a>&#xFFFF;</a>").isLeft)
  }

  test("beyond U+10FFFF is rejected: &#x110000;") {
    val e = bad("<a>&#x110000;</a>")
    assert(e.message.contains("invalid character reference"), e.message)
  }

  test("absurdly long digit strings terminate with an error, no overflow") {
    val e = bad(s"<a>&#${"9" * 100};</a>")
    assert(e.message.contains("invalid character reference"), e.message)
  }

  // --- malformed references ---------------------------------------------------------------------------

  test("undeclared entity is rejected: &nbsp;") {
    val e = bad("<a>&nbsp;</a>")
    assert(e.message.contains("undeclared entity '&nbsp;'"), e.message)
  }

  test("entity names are case-sensitive: &AMP; is undeclared") {
    val e = bad("<a>&AMP;</a>")
    assert(e.message.contains("undeclared"), e.message)
  }

  test("bare '&' is rejected with guidance") {
    val e = bad("<a>a & b</a>")
    assert(e.message.contains("&amp;"), e.message)
  }

  test("'&;' is rejected") {
    val e = bad("<a>&;</a>")
    assert(e.message.contains("&"), e.message)
  }

  test("'&#;' is rejected") {
    val e = bad("<a>&#;</a>")
    assert(e.message.contains("character reference"), e.message)
  }

  test("'&#xZZ;' is rejected") {
    val e = bad("<a>&#xZZ;</a>")
    assert(e.message.contains("character reference"), e.message)
  }

  test("unterminated entity at EOF is rejected") {
    val e = bad("<a>&amp")
    assert(e.message.contains("entity"), e.message)
  }

  test("unterminated entity (missing ';') is rejected") {
    val e = bad("<a>&amp b</a>")
    assert(e.message.contains(";"), e.message)
  }

  test("references are rejected in attribute values too") {
    assert(TestTokens.parse("""<a x="&nbsp;"/>""").isLeft)
    assert(TestTokens.parse("""<a x="&#xD800;"/>""").isLeft)
    assert(TestTokens.parse("""<a x="a & b"/>""").isLeft)
  }

  test("error position points at the reference") {
    val e = bad("<a>ok &bogus; tail</a>")
    assertEquals(e.line, 1)
    assertEquals(e.col, 7)
  }
