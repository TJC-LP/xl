package com.tjclp.xl.xml

import munit.FunSuite

class NamespaceSpec extends FunSuite:

  private val xmlNs = "http://www.w3.org/XML/1998/namespace"

  private def bad(xml: String): XmlParseError =
    TestTokens.parse(xml) match
      case Left(e) => e
      case Right(ts) => fail(s"expected failure, got tokens: $ts")

  test("prefix declared on an ancestor resolves on descendants") {
    val p = XmlPullParser.fromString("""<a xmlns:p="urn:u"><p:b/></a>""")
    assertEquals(p.next(), Right(XmlToken.StartElement))
    assertEquals(p.resolveNs("p"), Some("urn:u"))
    assertEquals(p.next(), Right(XmlToken.StartElement))
    assertEquals(p.name, "p:b")
    assertEquals(p.prefix, "p")
    assertEquals(p.localName, "b")
    assertEquals(p.elementNs, Some("urn:u"))
  }

  test("declaration on the element itself applies to that element") {
    val p = XmlPullParser.fromString("""<p:a xmlns:p="urn:u"/>""")
    assertEquals(p.next(), Right(XmlToken.StartElement))
    assertEquals(p.elementNs, Some("urn:u"))
  }

  test("shadowing: inner redeclaration wins, restored after end tag") {
    val p = XmlPullParser.fromString("""<a xmlns:p="urn:1"><b xmlns:p="urn:2"/><c/></a>""")
    assertEquals(p.next(), Right(XmlToken.StartElement)) // a
    assertEquals(p.resolveNs("p"), Some("urn:1"))
    assertEquals(p.next(), Right(XmlToken.StartElement)) // b
    assertEquals(p.resolveNs("p"), Some("urn:2"))
    assertEquals(p.next(), Right(XmlToken.EndElement)) // /b (synthesized)
    assertEquals(p.next(), Right(XmlToken.StartElement)) // c
    assertEquals(p.resolveNs("p"), Some("urn:1"))
  }

  test("default namespace declare and undeclare") {
    val p = XmlPullParser.fromString("""<a xmlns="urn:d"><b xmlns=""/></a>""")
    assertEquals(p.next(), Right(XmlToken.StartElement)) // a
    assertEquals(p.resolveNs(""), Some("urn:d"))
    assertEquals(p.elementNs, Some("urn:d"))
    assertEquals(p.next(), Right(XmlToken.StartElement)) // b
    assertEquals(p.resolveNs(""), None)
    assertEquals(p.elementNs, None)
    assertEquals(p.next(), Right(XmlToken.EndElement))
    assertEquals(p.next(), Right(XmlToken.EndElement)) // /a
    assertEquals(p.next(), Right(XmlToken.EndDocument))
  }

  test("default namespace does not apply to attributes (unprefixed attrs have no namespace)") {
    // We only assert the AttrList keeps the plain name; namespace interpretation is the consumer's.
    val p = XmlPullParser.fromString("""<a xmlns="urn:d" x="1"/>""")
    assertEquals(p.next(), Right(XmlToken.StartElement))
    assertEquals(p.attrs.get("x"), Some("1"))
  }

  test("'xml' prefix is implicitly bound") {
    val p = XmlPullParser.fromString("""<a xml:space="preserve"/>""")
    assertEquals(p.next(), Right(XmlToken.StartElement))
    assertEquals(p.resolveNs("xml"), Some(xmlNs))
    assertEquals(p.attrs.get("xml:space"), Some("preserve"))
  }

  test("xmlns declarations stay visible in the AttrList verbatim") {
    val p = XmlPullParser.fromString("""<a xmlns="urn:d" xmlns:p="urn:u"/>""")
    assertEquals(p.next(), Right(XmlToken.StartElement))
    assertEquals(p.attrs.toVector, Vector("xmlns" -> "urn:d", "xmlns:p" -> "urn:u"))
  }

  test("unbound element prefix is an error (JAXP-strict)") {
    val e = bad("<p:a/>")
    assert(e.message.contains("unbound"), e.message)
  }

  test("unbound attribute prefix is an error") {
    val e = bad("""<a p:x="1"/>""")
    assert(e.message.contains("unbound"), e.message)
  }

  test("prefix bound on the same element is not 'unbound' regardless of attribute order") {
    val p = XmlPullParser.fromString("""<a p:x="1" xmlns:p="urn:u"/>""")
    assertEquals(p.next(), Right(XmlToken.StartElement))
    assertEquals(p.resolveNs("p"), Some("urn:u"))
  }

  test("prefixed namespace undeclaration is rejected (Namespaces 1.0)") {
    val e = bad("""<a xmlns:p=""/>""")
    assert(e.message.contains("undeclared") || e.message.contains("cannot"), e.message)
  }

  test("declaring the 'xmlns' prefix is rejected") {
    val e = bad("""<a xmlns:xmlns="urn:u"/>""")
    assert(e.message.contains("xmlns"), e.message)
  }

  test("binding 'xml' to a different namespace is rejected") {
    val e = bad("""<a xmlns:xml="urn:not-xml"/>""")
    assert(e.message.contains("xml"), e.message)
  }

  test("binding 'xml' to the correct XML namespace is accepted") {
    val p = XmlPullParser.fromString(s"""<a xmlns:xml="$xmlNs"/>""")
    assertEquals(p.next(), Right(XmlToken.StartElement))
  }

  test("multiple colons in a name are rejected") {
    val e = bad("""<a:b:c xmlns:a="u"/>""")
    assert(e.message.contains("qualified name"), e.message)
  }

  test("name ending with a colon is rejected") {
    val e = bad("<a: />")
    assert(e.message.contains("qualified name"), e.message)
  }

  test("name starting with a colon is rejected") {
    val e = bad("<:a/>")
    assert(e.message.contains("name"), e.message)
  }

  test("unprefixed element with no default namespace has elementNs None") {
    val p = XmlPullParser.fromString("<a/>")
    assertEquals(p.next(), Right(XmlToken.StartElement))
    assertEquals(p.elementNs, None)
  }

  test("resolveNs of an unknown prefix is None") {
    val p = XmlPullParser.fromString("<a/>")
    assertEquals(p.next(), Right(XmlToken.StartElement))
    assertEquals(p.resolveNs("nope"), None)
  }

  test("EndElement carries the element name (prefix/local split)") {
    val p = XmlPullParser.fromString("""<p:a xmlns:p="urn:u"></p:a>""")
    assertEquals(p.next(), Right(XmlToken.StartElement))
    assertEquals(p.next(), Right(XmlToken.EndElement))
    assertEquals(p.name, "p:a")
    assertEquals(p.prefix, "p")
    assertEquals(p.localName, "a")
  }

  test("OOXML-shaped worksheet root namespaces resolve") {
    val xml =
      """<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>"""
    val p = XmlPullParser.fromString(xml)
    assertEquals(p.next(), Right(XmlToken.StartElement))
    assertEquals(
      p.elementNs,
      Some("http://schemas.openxmlformats.org/spreadsheetml/2006/main")
    )
    assertEquals(
      p.resolveNs("r"),
      Some("http://schemas.openxmlformats.org/officeDocument/2006/relationships")
    )
  }
