package com.tjclp.xl.xml

import munit.FunSuite
import OwnedToken.*

class TokenizerSpec extends FunSuite:

  private def ok(xml: String): Vector[OwnedToken] =
    TestTokens.parse(xml) match
      case Right(ts) => ts
      case Left(e) => fail(s"expected success, got: ${e.render}")

  private def bad(xml: String): XmlParseError =
    TestTokens.parse(xml) match
      case Left(e) => e
      case Right(ts) => fail(s"expected failure, got tokens: $ts")

  // --- basic documents ---------------------------------------------------------------------------

  test("minimal document") {
    assertEquals(
      ok("<a>hi</a>"),
      Vector(Start("a", "", "a", Vector.empty, 1), Text("hi"), End("a"))
    )
  }

  test("XML declaration is consumed silently") {
    assertEquals(
      ok("<?xml version=\"1.0\" encoding=\"UTF-8\"?><a/>"),
      Vector(Start("a", "", "a", Vector.empty, 1), End("a"))
    )
  }

  test("XML declaration with standalone") {
    assertEquals(
      ok("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><a/>"),
      Vector(Start("a", "", "a", Vector.empty, 1), End("a"))
    )
  }

  test("XML declaration with single quotes") {
    assertEquals(ok("<?xml version='1.0'?><a/>").size, 2)
  }

  test("unsupported XML version is rejected") {
    val e = bad("<?xml version=\"1.1\"?><a/>")
    assert(e.message.contains("1.0"), e.message)
  }

  test("empty element yields Start then synthesized End") {
    assertEquals(ok("<e/>"), Vector(Start("e", "", "e", Vector.empty, 1), End("e")))
  }

  test("start/end pair with no children yields no Text token") {
    assertEquals(ok("<a></a>"), Vector(Start("a", "", "a", Vector.empty, 1), End("a")))
  }

  test("nested elements track depth") {
    assertEquals(
      ok("<a><b><c/></b></a>"),
      Vector(
        Start("a", "", "a", Vector.empty, 1),
        Start("b", "", "b", Vector.empty, 2),
        Start("c", "", "c", Vector.empty, 3),
        End("c"),
        End("b"),
        End("a")
      )
    )
  }

  test("whitespace-only text is always delivered (nothing is ignorable without a DTD)") {
    assertEquals(
      ok("<a> <b/> </a>"),
      Vector(
        Start("a", "", "a", Vector.empty, 1),
        Text(" "),
        Start("b", "", "b", Vector.empty, 2),
        End("b"),
        Text(" "),
        End("a")
      )
    )
  }

  // --- attributes --------------------------------------------------------------------------------

  test("attributes preserve document order") {
    assertEquals(
      ok("""<c r="A1" s="3" t="s"/>"""),
      Vector(
        Start("c", "", "c", Vector("r" -> "A1", "s" -> "3", "t" -> "s"), 1),
        End("c")
      )
    )
  }

  test("single-quoted attribute values") {
    assertEquals(ok("<a x='1'/>"), Vector(Start("a", "", "a", Vector("x" -> "1"), 1), End("a")))
  }

  test("attribute value may contain the other quote kind") {
    assertEquals(
      ok("""<a x="it's"/>""").headOption,
      Some(Start("a", "", "a", Vector("x" -> "it's"), 1))
    )
  }

  test("whitespace around = is allowed") {
    assertEquals(ok("<a x = \"1\"/>").headOption, Some(Start("a", "", "a", Vector("x" -> "1"), 1)))
  }

  test("unquoted attribute value is rejected") {
    val e = bad("<a x=1/>")
    assert(e.message.contains("quoted"), e.message)
  }

  test("attribute without value is rejected") {
    val e = bad("<a x/>")
    assert(e.message.contains("="), e.message)
  }

  test("duplicate attribute is rejected") {
    val e = bad("""<a x="1" x="2"/>""")
    assert(e.message.contains("duplicate"), e.message)
  }

  test("'<' in attribute value is rejected") {
    val e = bad("""<a x="a<b"/>""")
    assert(e.message.contains("'<'"), e.message)
  }

  test("missing whitespace between attributes is rejected") {
    val e = bad("""<a x="1"y="2"/>""")
    assert(e.message.contains("whitespace"), e.message)
  }

  // --- text, comments, PIs -------------------------------------------------------------------------

  test("comments do not split a text run (coalesced across)") {
    assertEquals(
      ok("<a>x<!-- c -->y</a>"),
      Vector(Start("a", "", "a", Vector.empty, 1), Text("xy"), End("a"))
    )
  }

  test("processing instructions do not split a text run") {
    assertEquals(
      ok("<a>x<?pi some data?>y</a>"),
      Vector(Start("a", "", "a", Vector.empty, 1), Text("xy"), End("a"))
    )
  }

  test("PI with no data") {
    assertEquals(ok("<a><?pi?></a>").size, 2)
  }

  test("comments and PIs allowed in prolog and epilog") {
    assertEquals(ok("<!-- pre --><?pi?><a/><?post?><!-- done -->").size, 2)
  }

  test("'--' inside a comment is rejected") {
    val e = bad("<a><!-- -- --></a>")
    assert(e.message.contains("--"), e.message)
  }

  test("unterminated comment is rejected") {
    val e = bad("<a><!-- never ends")
    assert(e.message.contains("comment"), e.message)
  }

  test("reserved PI target 'xml' after start is rejected") {
    val e = bad("<a><?xml version=\"1.0\"?></a>")
    assert(e.message.contains("reserved"), e.message)
  }

  test("reserved PI target is case-insensitive") {
    val e = bad("<a><?XML stuff?></a>")
    assert(e.message.contains("reserved"), e.message)
  }

  // --- CDATA ---------------------------------------------------------------------------------------

  test("CDATA is its own token with raw content") {
    assertEquals(
      ok("<a><![CDATA[<b>&amp;</b>]]></a>"),
      Vector(Start("a", "", "a", Vector.empty, 1), CData("<b>&amp;</b>"), End("a"))
    )
  }

  test("empty CDATA section") {
    assertEquals(
      ok("<a><![CDATA[]]></a>"),
      Vector(Start("a", "", "a", Vector.empty, 1), CData(""), End("a"))
    )
  }

  test("CDATA with ']]' inside and ']' before the terminator") {
    assertEquals(
      ok("<a><![CDATA[a]]b]]></a>"),
      Vector(Start("a", "", "a", Vector.empty, 1), CData("a]]b"), End("a"))
    )
    assertEquals(
      ok("<a><![CDATA[]]]></a>"),
      Vector(Start("a", "", "a", Vector.empty, 1), CData("]"), End("a"))
    )
  }

  test("CDATA does not merge with adjacent text") {
    assertEquals(
      ok("<a>x<![CDATA[y]]>z</a>"),
      Vector(
        Start("a", "", "a", Vector.empty, 1),
        Text("x"),
        CData("y"),
        Text("z"),
        End("a")
      )
    )
  }

  test("unterminated CDATA is rejected") {
    val e = bad("<a><![CDATA[never")
    assert(e.message.contains("CDATA"), e.message)
  }

  test("CDATA in prolog is rejected") {
    val e = bad("<![CDATA[x]]><a/>")
    assert(e.message.contains("prolog"), e.message)
  }

  // --- well-formedness ------------------------------------------------------------------------------

  test("mismatched end tag is rejected") {
    val e = bad("<a><b></a></b>")
    assert(e.message.contains("mismatched"), e.message)
  }

  test("content after the root element is rejected") {
    val e = bad("<a/>text")
    assert(e.message.contains("after the root"), e.message)
  }

  test("second root element is rejected") {
    val e = bad("<a/><b/>")
    assert(e.message.contains("root"), e.message)
  }

  test("unclosed element at EOF is rejected") {
    val e = bad("<a><b>")
    assert(e.message.contains("not closed"), e.message)
  }

  test("empty input is rejected") {
    val e = bad("")
    assert(e.message.contains("no root element"), e.message)
  }

  test("whitespace-only input is rejected") {
    val e = bad("   \n\t  ")
    assert(e.message.contains("no root element"), e.message)
  }

  test("']]>' in character data is rejected") {
    val e = bad("<a>x]]>y</a>")
    assert(e.message.contains("]]>"), e.message)
  }

  test("bare '<!x' markup is rejected") {
    val e = bad("<a><!x></a>")
    assert(e.message.contains("markup"), e.message)
  }

  test("end tag in prolog is rejected") {
    val e = bad("</a>")
    assert(e.message.contains("prolog"), e.message)
  }

  test("text in prolog is rejected") {
    val e = bad("hello<a/>")
    assert(e.message.contains("prolog"), e.message)
  }

  test("malformed end tag is rejected") {
    val e = bad("<a></a b>")
    assert(e.message.contains("end tag"), e.message)
  }

  // --- positions ------------------------------------------------------------------------------------

  test("error position is 1-based line/col") {
    val e = bad("<a>\n  <b>\n</a>")
    assertEquals(e.line, 3)
    assert(e.col > 0)
  }

  test("render matches the parseSafe error shape") {
    val e = bad("<a><b></a>")
    assert(e.render.startsWith(s"XML parse error at line ${e.line}, column ${e.col}: "), e.render)
  }

  test("token positions: parser cursor reports the token start") {
    val p = XmlPullParser.fromString("<a>hello</a>")
    assertEquals(p.next(), Right(XmlToken.StartElement))
    assertEquals((p.line, p.col), (1, 1))
    assertEquals(p.next(), Right(XmlToken.Text))
    assertEquals((p.line, p.col), (1, 4))
  }

  // --- cursor accessors -------------------------------------------------------------------------------

  test("appendTextTo accumulates without materializing") {
    val p = XmlPullParser.fromString("<a>one</a>")
    assertEquals(p.next(), Right(XmlToken.StartElement))
    assertEquals(p.next(), Right(XmlToken.Text))
    val sb = new java.lang.StringBuilder("x:")
    p.appendTextTo(sb)
    assertEquals(sb.toString, "x:one")
  }

  test("isWhitespaceText") {
    val p = XmlPullParser.fromString("<a> \t\n</a>")
    assertEquals(p.next(), Right(XmlToken.StartElement))
    assertEquals(p.next(), Right(XmlToken.Text))
    assert(p.isWhitespaceText)
    val p2 = XmlPullParser.fromString("<a> x </a>")
    assertEquals(p2.next(), Right(XmlToken.StartElement))
    assertEquals(p2.next(), Right(XmlToken.Text))
    assert(!p2.isWhitespaceText)
  }

  test("AttrList lookup surface") {
    val p = XmlPullParser.fromString("""<c r="A1" t="s"/>""")
    assertEquals(p.next(), Right(XmlToken.StartElement))
    val a = p.attrs
    assertEquals(a.size, 2)
    assertEquals(a.indexOf("r"), 0)
    assertEquals(a.indexOf("t"), 1)
    assertEquals(a.indexOf("zz"), -1)
    assertEquals(a.get("r"), Some("A1"))
    assertEquals(a.get("zz"), None)
    assertEquals(a.name(0), "r")
    assertEquals(a.value(1), "s")
    assertEquals(a.name(5), "")
    assertEquals(a.value(-1), "")
    assertEquals(a.toVector, Vector("r" -> "A1", "t" -> "s"))
  }

  test("AttrList instance is reused across elements (aliasing contract)") {
    val p = XmlPullParser.fromString("""<a x="1"><b y="2"/></a>""")
    assertEquals(p.next(), Right(XmlToken.StartElement))
    val first = p.attrs
    assertEquals(p.next(), Right(XmlToken.StartElement))
    assert(first eq p.attrs)
    assertEquals(p.attrs.get("y"), Some("2"))
    assertEquals(p.attrs.get("x"), None)
  }
