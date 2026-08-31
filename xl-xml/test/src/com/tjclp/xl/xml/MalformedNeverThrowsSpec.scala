package com.tjclp.xl.xml

import org.scalacheck.{Arbitrary, Gen}
import org.scalacheck.Prop.forAll

/**
 * S1: totality — for ANY input bytes, the next() loop reaches Left or EndDocument, never throws,
 * and terminates within a linear step budget under limits. The try/catch lives in this harness
 * only; the parser itself must never throw.
 */
class MalformedNeverThrowsSpec extends munit.ScalaCheckSuite:

  override def scalaCheckTestParameters =
    super.scalaCheckTestParameters.withMinSuccessfulTests(300)

  private val harnessLimits = XmlLimits(maxTotalChars = 1000000L)

  /** Drive the parser to a terminal; returns (threw, terminatedInBudget, result). */
  private def drive(
    bytes: Array[Byte],
    limits: XmlLimits
  ): (Option[Throwable], Boolean, Option[Either[XmlParseError, XmlToken]]) =
    val budget = 2 * bytes.length + 64
    try
      val p = XmlPullParser.fromArray(bytes, limits)
      var calls = 0
      var r = p.next()
      calls += 1
      while r.isRight && r != Right(XmlToken.EndDocument) && calls <= budget do
        r = p.next()
        calls += 1
      (None, calls <= budget, Some(r))
    catch case t: Throwable => (Some(t), false, None)

  private def assertTotal(bytes: Array[Byte], limits: XmlLimits = harnessLimits): Unit =
    val (threw, inBudget, result) = drive(bytes, limits)
    threw.foreach(t => fail(s"parser threw ${t.getClass.getName}: ${t.getMessage}"))
    assert(inBudget, s"parser exceeded its step budget on ${bytes.length} bytes")
    result match
      case Some(Left(e)) =>
        assert(e.line >= 1 && e.col >= 1, s"error must be positioned: $e")
        assert(e.message.nonEmpty)
      case Some(Right(XmlToken.EndDocument)) => ()
      case other => fail(s"non-terminal result: $other")

  private def utf8(s: String): Array[Byte] = internal.Utf8.encode(s)

  test("fixed malformed corpus: returns Left or EndDocument, never throws, terminates") {
    val corpus: List[Array[Byte]] = List(
      utf8(""),
      utf8("   \n\t "),
      utf8("<"),
      utf8("<a"),
      utf8("<a "),
      utf8("<a x"),
      utf8("<a x="),
      utf8("<a x=\""),
      utf8("<a x=\"v\" /"),
      utf8("<a>"),
      utf8("</a>"),
      utf8("<a></b>"),
      utf8("<a><b></a></b>"),
      utf8("<a/>junk"),
      utf8("<a/><b/>"),
      utf8("<a>&"),
      utf8("<a>&amp"),
      utf8("<a>& </a>"),
      utf8("<a>&bogus;</a>"),
      utf8("<a>&#x110000;</a>"),
      utf8("<a>&#xD800;</a>"),
      utf8("<a>&#</a>"),
      utf8("<!DOCTYPE x [<!ENTITY e 'v'>]><a>&e;</a>"),
      utf8("<!DOCTYPE x ["),
      utf8("<!DOCTYPE"),
      utf8("<!DOCTYP <a/>"),
      utf8("<!-- unterminated"),
      utf8("<a><!-- -- --></a>"),
      utf8("<?pi unterminated"),
      utf8("<?xml version=\"2.0\"?><a/>"),
      utf8("<?xml encoding=\"UTF-8\"?><a/>"),
      utf8("<?xml version=\"1.0\" standalone=\"maybe\"?><a/>"),
      utf8("<a><![CDATA[unterminated"),
      utf8("<a>]]></a>"),
      utf8("<a x=1/>"),
      utf8("<a x='1' x='2'/>"),
      utf8("<a 1x=\"v\"/>"),
      utf8("<p:a/>"),
      utf8("<a:b:c/>"),
      utf8("<a xmlns:p=''/>"),
      Array(0xff.toByte, 0xfe.toByte), // BOM only
      Array(0xef.toByte, 0xbb.toByte, 0xbf.toByte),
      Array(0xc3.toByte), // truncated UTF-8
      Array(0x3c.toByte, 0x00.toByte, 0x61.toByte), // truncated UTF-16LE
      utf8("<a>") ++ Array(0xed.toByte, 0xa0.toByte, 0x80.toByte) ++ utf8("</a>"),
      utf8("<a>") ++ Array(0x00.toByte) ++ utf8("</a>")
    )
    corpus.foreach(assertTotal(_))
  }

  test("fixed malformed corpus under tight limits") {
    val tight =
      XmlLimits(maxDepth = 3, maxAttrsPerElement = 2, maxNameLength = 5, maxTotalChars = 200)
    List(
      utf8("<aaaaaaaaaa/>"),
      utf8("<a><a><a><a><a></a></a></a></a></a>"),
      utf8("<a b=\"1\" c=\"2\" d=\"3\"/>"),
      utf8(s"<a>${"x" * 500}</a>"),
      utf8(s"<!DOCTYPE x [ ${"y" * 500} ]><a/>")
    ).foreach(assertTotal(_, tight))
  }

  property("S1: arbitrary byte arrays never throw and terminate") {
    forAll(Gen.containerOf[Array, Byte](Arbitrary.arbitrary[Byte])) { bytes =>
      assertTotal(bytes)
      true
    }
  }

  property("S1: truncations of a valid document never throw") {
    val doc = utf8(
      """<?xml version="1.0" encoding="UTF-8"?><wb xmlns:r="urn:r"><s r:id="rId1" n="A &amp; B"/><t xml:space="preserve"> x </t><!-- c --><![CDATA[raw]]></wb>"""
    )
    forAll(Gen.choose(0, doc.length)) { n =>
      assertTotal(doc.take(n))
      true
    }
  }

  property("S1: single-byte flips of a valid document never throw") {
    val doc = utf8("""<r a="1"><c t="s"><v>hé &lt;3</v></c><![CDATA[x]]><e/></r>""")
    forAll(Gen.choose(0, doc.length - 1), Arbitrary.arbitrary[Byte]) { (i, b) =>
      val mutated = doc.clone()
      mutated(i) = b
      assertTotal(mutated)
      true
    }
  }

  property("S1: single-byte insertions into a valid document never throw") {
    val doc = utf8("""<r><c r="A1"><f>SUM(A1:B2)</f><v>3</v></c></r>""")
    forAll(Gen.choose(0, doc.length), Arbitrary.arbitrary[Byte]) { (i, b) =>
      val mutated = (doc.take(i) :+ b) ++ doc.drop(i)
      assertTotal(mutated)
      true
    }
  }
