package com.tjclp.xl.xml

import org.scalacheck.Gen
import org.scalacheck.Prop.forAll

class LimitsSpec extends munit.ScalaCheckSuite:

  private def nested(n: Int): String =
    (1 to n).map(i => s"<e$i>").mkString + (n to 1 by -1).map(i => s"</e$i>").mkString

  test("maxDepth: passes at the limit, trips at limit + 1, error names the limit") {
    val limits = XmlLimits(maxDepth = 5)
    assert(TestTokens.parse(nested(5), limits).isRight)
    TestTokens.parse(nested(6), limits) match
      case Left(e) =>
        assert(e.message.contains("maxDepth"), e.message)
        assert(e.line >= 1 && e.col >= 1)
      case Right(ts) => fail(s"expected maxDepth failure, got $ts")
  }

  test("maxAttrsPerElement: passes at the limit, trips at limit + 1, error names the limit") {
    def doc(n: Int) = "<a " + (1 to n).map(i => s"""a$i="v"""").mkString(" ") + "/>"
    val limits = XmlLimits(maxAttrsPerElement = 4)
    assert(TestTokens.parse(doc(4), limits).isRight)
    TestTokens.parse(doc(5), limits) match
      case Left(e) => assert(e.message.contains("maxAttrsPerElement"), e.message)
      case Right(ts) => fail(s"expected maxAttrsPerElement failure, got $ts")
  }

  test("maxNameLength: element name at the limit passes, one longer trips") {
    val limits = XmlLimits(maxNameLength = 8)
    assert(TestTokens.parse(s"<${"n" * 8}/>", limits).isRight)
    TestTokens.parse(s"<${"n" * 9}/>", limits) match
      case Left(e) => assert(e.message.contains("maxNameLength"), e.message)
      case Right(ts) => fail(s"expected maxNameLength failure, got $ts")
  }

  test("maxNameLength applies to attribute names too") {
    val limits = XmlLimits(maxNameLength = 8)
    assert(TestTokens.parse(s"""<a ${"n" * 9}="v"/>""", limits).isLeft)
  }

  test("maxTotalChars: input under the cap passes, over trips, error names the limit") {
    val doc = s"<a>${"x" * 100}</a>"
    assert(TestTokens.parse(doc, XmlLimits(maxTotalChars = 200)).isRight)
    TestTokens.parse(doc, XmlLimits(maxTotalChars = 50)) match
      case Left(e) => assert(e.message.contains("maxTotalChars"), e.message)
      case Right(ts) => fail(s"expected maxTotalChars failure, got $ts")
  }

  test("maxTotalChars bounds DOCTYPE skipping (no unbounded prolog scan)") {
    val doc = s"<!DOCTYPE x [ ${"y" * 1000} ]><x/>"
    TestTokens.parse(doc, XmlLimits(maxTotalChars = 100)) match
      case Left(e) => assert(e.message.contains("maxTotalChars"), e.message)
      case Right(ts) => fail(s"expected maxTotalChars failure, got $ts")
  }

  test("defaults mirror the ported posture") {
    assertEquals(XmlLimits.default.maxDepth, 256)
    assertEquals(XmlLimits.default.maxAttrsPerElement, 512)
    assertEquals(XmlLimits.default.maxNameLength, 1000)
    assertEquals(XmlLimits.default.maxTotalChars, 2_000_000_000L)
  }

  // L1: limits monotonicity — success under L implies success under any L' >= L
  property("L1: limits monotonicity") {
    val genDoc =
      for
        depth <- Gen.choose(1, 6)
        attrs <- Gen.choose(0, 4)
        nameLen <- Gen.choose(1, 6)
        text <- Gen.alphaNumStr.map(_.take(20))
      yield
        val name = "n" * nameLen
        val attrStr = (1 to attrs).map(i => s"""a$i="v$i"""").mkString(" ")
        val open = (1 to depth).map(i => s"<$name$i $attrStr>").mkString
        val close = (depth to 1 by -1).map(i => s"</$name$i>").mkString
        s"$open$text$close"
    val genLimits =
      for
        d <- Gen.choose(1, 10)
        a <- Gen.choose(1, 10)
        n <- Gen.choose(1, 20)
        t <- Gen.choose(1L, 2000L)
      yield XmlLimits(d, a, n, t)
    forAll(genDoc, genLimits, genLimits) { (doc, l1, l2) =>
      val lo = XmlLimits(
        math.min(l1.maxDepth, l2.maxDepth),
        math.min(l1.maxAttrsPerElement, l2.maxAttrsPerElement),
        math.min(l1.maxNameLength, l2.maxNameLength),
        math.min(l1.maxTotalChars, l2.maxTotalChars)
      )
      val hi = XmlLimits(
        math.max(l1.maxDepth, l2.maxDepth),
        math.max(l1.maxAttrsPerElement, l2.maxAttrsPerElement),
        math.max(l1.maxNameLength, l2.maxNameLength),
        math.max(l1.maxTotalChars, l2.maxTotalChars)
      )
      // If the doc parses under the LOWER limits it must parse under the HIGHER ones.
      TestTokens.parse(doc, lo).isLeft || TestTokens.parse(doc, hi).isRight
    }
  }
