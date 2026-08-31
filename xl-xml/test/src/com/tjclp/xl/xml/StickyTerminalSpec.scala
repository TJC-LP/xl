package com.tjclp.xl.xml

import org.scalacheck.{Arbitrary, Gen}
import org.scalacheck.Prop.forAll

/** P1: terminal results are sticky — the parser can never resurrect after Left or EndDocument. */
class StickyTerminalSpec extends munit.ScalaCheckSuite:

  test("after a Left, every later next() returns the identical Left") {
    val p = XmlPullParser.fromString("<a><b></a>")
    var last = p.next()
    while last.isRight do last = p.next()
    val first = last
    (1 to 5).foreach { _ =>
      val again = p.next()
      assertEquals(again, first)
      assert(again eq first, "sticky Left must be the same cached value, not a re-allocation")
    }
  }

  test("after EndDocument, every later next() returns EndDocument") {
    val p = XmlPullParser.fromString("<a/>")
    assertEquals(p.next(), Right(XmlToken.StartElement))
    assertEquals(p.next(), Right(XmlToken.EndElement))
    val end = p.next()
    assertEquals(end, Right(XmlToken.EndDocument))
    (1 to 5).foreach { _ =>
      val again = p.next()
      assertEquals(again, end)
      assert(again eq end, "sticky EndDocument must be the same cached value")
    }
  }

  test("token Rights are cached values (zero per-token allocation)") {
    val p = XmlPullParser.fromString("<a><b/><b/></a>")
    val t1 = p.next() // Start a
    val t2 = p.next() // Start b
    assert(t1 eq t2)
    val e1 = p.next() // End b (synthesized)
    val s3 = p.next() // Start b
    val e2 = p.next() // End b
    assert(e1 eq e2)
    assert(t1 eq s3)
  }

  property("P1: for arbitrary bytes, once terminal, next() is a fixpoint") {
    forAll(Gen.containerOf[Array, Byte](Arbitrary.arbitrary[Byte])) { bytes =>
      val p = XmlPullParser.fromArray(bytes, XmlLimits(maxTotalChars = 100000L))
      var steps = 0
      var r = p.next()
      while r.isRight && r != Right(XmlToken.EndDocument) && steps < 1000000 do
        r = p.next()
        steps += 1
      val terminal = r
      (1 to 3).forall(_ => p.next() eq terminal)
    }
  }
