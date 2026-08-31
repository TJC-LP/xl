package com.tjclp.xl.xml

import munit.FunSuite
import OwnedToken.*

/**
 * DOCTYPE skip-benign semantics — the token-level port of every GH-350 stripper case
 * (SecuritySpec), replacing stripLeadingDoctype/stripLeadingDoctypeStream in the portable engine.
 */
class DoctypeSpec extends FunSuite:

  private def ok(xml: String): Vector[OwnedToken] =
    TestTokens.parse(xml) match
      case Right(ts) => ts
      case Left(e) => fail(s"expected success, got: ${e.render}")

  private def bad(xml: String): XmlParseError =
    TestTokens.parse(xml) match
      case Left(e) => e
      case Right(ts) => fail(s"expected failure, got tokens: $ts")

  test("GH-350: SYSTEM-form doctype in the prolog is skipped") {
    assertEquals(
      ok("<?xml version=\"1.0\"?>\n<!DOCTYPE workbook SYSTEM \"wb.dtd\">\n<workbook/>"),
      Vector(Start("workbook", "", "workbook", Vector.empty, 1), End("workbook"))
    )
  }

  test("GH-350: quotes and bracket nesting in an internal subset are tracked") {
    assertEquals(
      ok("""<?xml version="1.0"?><!DOCTYPE x [ <!ENTITY e "tricky > ] value"> ]><x/>"""),
      Vector(Start("x", "", "x", Vector.empty, 1), End("x"))
    )
  }

  test("GH-350: prolog comments before the doctype are fine") {
    assertEquals(
      ok("<?xml version=\"1.0\"?><!-- produced by tool --><!DOCTYPE x SYSTEM \"x.dtd\"><x/>").size,
      2
    )
  }

  test("GH-350: comment inside the internal subset (apostrophe) does not open quote state") {
    assertEquals(ok("""<?xml version="1.0"?><!DOCTYPE x [ <!-- don't --> ]><x/>""").size, 2)
  }

  test("GH-350: comment inside the internal subset (bracket) does not close the subset early") {
    assertEquals(ok("""<?xml version="1.0"?><!DOCTYPE x [ <!-- ] --> ]><x/>""").size, 2)
  }

  test("GH-350: comment-laden internal subset with declarations parses") {
    assertEquals(
      ok("<?xml version=\"1.0\"?>\n<!DOCTYPE x [ <!-- don't ] --> <!ELEMENT x ANY> ]>\n<x/>").size,
      2
    )
  }

  test("GH-350: PI span inside the internal subset is skipped wholesale") {
    assertEquals(ok("""<!DOCTYPE x [ <?pi tricky > and ] ?> ]><x/>""").size, 2)
  }

  test("GH-350: BOM + doctype parses") {
    val bom = Array(0xef.toByte, 0xbb.toByte, 0xbf.toByte)
    val xml = "<?xml version=\"1.0\"?>\n<!DOCTYPE x SYSTEM \"http://example.invalid/x.dtd\">\n<x/>"
    assertEquals(TestTokens.parseBytes(bom ++ internal.Utf8.encode(xml)).map(_.size), Right(2))
  }

  test("GH-350: entity declared in a skipped subset is NEVER honored") {
    val e = bad("""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE x [<!ENTITY xxe SYSTEM "file:///etc/passwd">]>
<x>&xxe;</x>""")
    assert(e.message.contains("undeclared entity '&xxe;'"), e.message)
  }

  test("GH-350: internally-declared value entity is also never honored") {
    val e = bad("""<!DOCTYPE x [<!ENTITY e "v">]><x>&e;</x>""")
    assert(e.message.contains("undeclared entity '&e;'"), e.message)
  }

  test("parameter entities are structurally impossible") {
    val e = bad("""<!DOCTYPE x [<!ENTITY % pe "v">]><x>&pe;</x>""")
    assert(e.message.contains("undeclared"), e.message)
  }

  test("doctype after the root element is rejected") {
    val e = bad("<x/><!DOCTYPE x []>")
    assert(e.message.contains("DOCTYPE"), e.message)
  }

  test("a second doctype is rejected") {
    val e = bad("<!DOCTYPE x []><!DOCTYPE y []><x/>")
    assert(e.message.contains("multiple DOCTYPE"), e.message)
  }

  test("unterminated doctype is a positioned error at EOF") {
    val e = bad("<?xml version=\"1.0\"?><!DOCTYPE x [ <!ENTITY e \"v\">")
    assert(e.message.contains("unterminated DOCTYPE"), e.message)
    assertEquals(e.line, 1)
  }

  test("doctype spanning newlines keeps line numbers accurate for later errors") {
    // error (mismatched tag) is on line 6
    val e = bad("<?xml version=\"1.0\"?>\n<!DOCTYPE x [\n<!ELEMENT x ANY>\n]>\n<x>\n<y></x>")
    assertEquals(e.line, 6)
  }

  test("doctype in element content is rejected") {
    val e = bad("<x><!DOCTYPE y []></x>")
    assert(e.message.contains("markup"), e.message)
  }

  test("DOCTYPE keyword is case-sensitive: <!doctype is not recognized") {
    val e = bad("<!doctype x []><x/>")
    assert(e.message.contains("markup"), e.message)
  }

  test("doctype-looking text inside CDATA is content, not markup") {
    assertEquals(
      ok("<x><![CDATA[<!DOCTYPE y [ ]>]]></x>"),
      Vector(Start("x", "", "x", Vector.empty, 1), CData("<!DOCTYPE y [ ]>"), End("x"))
    )
  }

  test("large internal subset is bounded only by maxTotalChars") {
    // A subset far beyond the old 16KB MaxPrologScanBytes stream-stripper window (GH-350 weakness).
    val big = "<!-- " + ("x" * 40000) + " -->"
    assertEquals(ok(s"<!DOCTYPE x [ $big ]><x/>").size, 2)
  }
