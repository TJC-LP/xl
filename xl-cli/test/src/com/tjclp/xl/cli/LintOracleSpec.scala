package com.tjclp.xl.cli

import munit.{FunSuite, ScalaCheckSuite}
import org.scalacheck.{Gen, Prop}
import org.scalacheck.Prop.forAll

import com.tjclp.xl.cli.commands.LintCommands
import com.tjclp.xl.formula.parser.UnparseableFormula
import com.tjclp.xl.testing.SmallStack

/**
 * #681 item 5: the `formula-unparseable` oracle behind `xl lint`, pointed at untrusted `<f>` text —
 * its fast-path prefilter must never hide a finding the parse reports, and nothing the parser does
 * may escape the lint's `XLResult` contract (a `StackOverflowError` is not `NonFatal`).
 */
class LintOracleSpec extends FunSuite with ScalaCheckSuite:

  override def scalaCheckTestParameters =
    super.scalaCheckTestParameters.withMinSuccessfulTests(3000)

  /** Run `f` on a thread with a deliberately small stack; its result, or what it threw. */
  private def onSmallStack[A](stackBytes: Long)(f: => A): Either[Throwable, A] =
    SmallStack.attempt(stackBytes, "lint-oracle-small-stack")(f)

  /**
   * Exhaust the stack, then call `check` while unwinding one live frame at a time: the first frame
   * with room for the guard's `try` but not for the parse's descent answers, whatever the JIT tier
   * or the stack size, so the overflow the guard converts is real by construction, never luck.
   *
   * The handler runs at the overflow point on purpose. HotSpot and GraalVM re-arm the stack's guard
   * zone as the overflow unwinds and skip a handler whose frame still lies inside it, so the frame
   * that runs `check` has a re-armed zone below it and the parse's second overflow is thrown, not
   * fatal; green under the interpreter, C1 and C2 on both JVMs.
   */
  @scala.noinline
  private def atStackBoundary(check: () => Option[String]): Option[String] =
    val result =
      try atStackBoundary(check)
      catch case _: StackOverflowError => check()
    java.lang.ref.Reference.reachabilityFence(result)
    result

  test("#681: a StackOverflowError inside the oracle is a finding, never an escape") {
    val overflowing: String => Option[String] = _ => throw new StackOverflowError()
    val verdict = UnparseableFormula.stackGuarded(overflowing)("SUM(((1)))")
    assert(verdict.exists(_.contains("stack")), verdict.toString)
    // only the stack overflow is converted: any other throwable is a bug to surface, not hide
    intercept[IllegalStateException](
      UnparseableFormula.stackGuarded(_ => throw new IllegalStateException("boom"))("1")
    )
  }

  test("#681: pathological <f> text on a small stack never escapes the oracle") {
    // the parser's 128-level budget bounds its depth, but not its frame size: on a small thread
    // stack (a fiber, a native image) the deepest legal nests can overflow. Every text must come
    // back as a verdict, the overflow as a finding, whether or not it overflows on the stack at
    // hand: the frame size of an interpreted parser is several times its compiled size, so which
    // texts overflow depends on which suites warmed the JVM first (#843). The injected-overflow
    // test above pins the conversion itself; this one pins the real lint path.
    val texts = Vector(
      "(" * 127 + "1" + ")" * 127,
      "SUM(" * 127 + "1" + ")" * 127,
      "IF(A1," * 120 + "1" + ")" * 120,
      "-" * 4000 + "1",
      "(" * 4000 + "1]",
      "{" * 3000 + "1" + "}" * 3000,
      "\"" * 8191,
      "'" * 8191,
      "[" * 4000 + "]" * 4000,
      "(" * 8192,
      "SUM(Table1[Amount])" * 400
    )
    // a nest the prefilter admits (an unterminated call is a certain truncation) that descends to
    // the parser's depth limit
    val deep = "SUM(" * 127 + "1"
    assert(UnparseableFormula.needsParse(deep), "the probe must reach the parser")
    // 1. warm the oracle and every class it touches on the normal stack: a class whose initialiser
    //    overflowed on a small thread is poisoned for the JVM (NoClassDefFoundError afterwards,
    //    which no guard converts), and the first touch must not happen there
    (texts :+ deep).foreach(LintCommands.formulaCheck)
    val baseline = LintCommands.formulaCheck(deep)
    assert(
      baseline.exists(f => !f.contains("stack")),
      s"the normal stack must hold the parse: $baseline"
    )
    // 2. the contract: every text is a verdict on a 256 KB stack, a size HotSpot and GraalVM honour
    //    (64 KB is below the JVM's minimum and silently rounded up), whether the parse fits (an
    //    ordinary verdict) or not (the guard's finding)
    val results =
      texts.map(t => t.take(20) -> onSmallStack(256L * 1024)(LintCommands.formulaCheck(t)))
    val escaped = results.collect { case (t, Left(err)) => s"$t…: $err" }
    assertEquals(escaped, Vector.empty[String], escaped.mkString("\n"))
    // 3. and a real overflow is the guard's finding: at the stack boundary the guard's `try` fits
    //    where the 127-level descent cannot, so the first verdict to return names the stack; a
    //    guard that had come loose would unwind until the parse fits and return the ordinary
    //    verdict. Probed on the 256 KB stack and on a 1 MB one.
    Vector(256L * 1024, 1024L * 1024).foreach { stackBytes =>
      val boundary =
        onSmallStack(stackBytes)(atStackBoundary(() => LintCommands.formulaCheck(deep)))
      assertEquals(
        boundary.map(_.exists(_.contains("stack"))),
        Right(true),
        s"at $stackBytes bytes: $boundary"
      )
    }
  }

  // ===== the prefilter: structured and external references skip the parse =====

  test("#681: ordinary structured and external references never pay for a parse") {
    val skipped = Vector(
      "SUM(Table1[Amount])",
      "(Table1[Amount])",
      "Table1[[#This Row],[Amount]]",
      "SUM(Table1[[#Totals],[Amount]])*2",
      "[1]Sheet1!A1",
      "SUM([1]Sheet1!A1:B2)",
      "([1]Sheet1!A1)",
      "'[2]Q1 Report'!B2",
      "{1,2;3,4}",
      "SUM({1,2,3})",
      "(A1:A3={1;2;3})",
      "\"a]b\"",
      "IF(A1,\"}\",\"]\")"
    )
    skipped.foreach(t => assert(!UnparseableFormula.needsParse(t), s"should skip the parse: $t"))
  }

  test("#681: a ] or } closing a ( still reaches the parser") {
    val parsed = Vector("(A1]", "(A1}", "SUM(A1:A2]", "((A1]", "{(1]}", "SUM(T[a]]", "A1]", "x)")
    parsed.foreach(t => assert(UnparseableFormula.needsParse(t), s"should be parsed: $t"))
    assert(LintCommands.formulaCheck("(A1]").isDefined)
    assert(LintCommands.formulaCheck("(A1}").isDefined)
  }

  /** Formula-ish text built from the delimiters and tokens the prefilter reasons about. */
  private val formulaText: Gen[String] =
    val token = Gen.frequency(
      8 -> Gen.oneOf("(", ")", "[", "]", "{", "}", "\"", "'", ",", " ", ":", "!", "+", ";"),
      4 -> Gen.oneOf("A1", "B2:C3", "1", "SUM(", "IF(", "Sheet1!", "#REF!", "x"),
      2 -> Gen.oneOf("Table1[Amount]", "[1]Sheet1!A1", "'Q 1'!A1", "\"s\"", "{1,2}", "''", "\"\"")
    )
    val random = Gen.choose(0, 24).flatMap(n => Gen.listOfN(n, token)).map(_.mkString)
    // near the closer arm: a parenthesized expression built from complete operands (strings,
    // quoted names, external and structured references, arrays), then a closer of any kind
    val operand = Gen.oneOf(
      "A1",
      "1",
      "\"a)b]\"",
      "'Q 1'!A1",
      "[1]Sheet1!A1",
      "Table1[Amount]",
      "{1,2}",
      "SUM(A1)",
      "(B2)"
    )
    val expr = Gen.listOfN(3, operand).flatMap(ops => Gen.choose(1, 3).map(n => ops.take(n)))
    val nearCloser =
      for
        prefix <- Gen.oneOf("", "SUM(", "1+", "IF(A1,", "{", "(")
        inner <- expr.map(_.mkString("+"))
        closer <- Gen.oneOf("]", "}", ")")
        suffix <- Gen.oneOf("", ")", "+1", "]", ")]", "}")
      yield s"$prefix($inner$closer$suffix"
    Gen.frequency(3 -> random, 2 -> nearCloser)

  test("#681: the parity generator reaches the closer finding and the spared references") {
    // guards the two properties below against a vacuous generator: over a fixed-seed sample, some
    // texts must be the `]`/`}`-closes-`(` finding and some must be spared a parse despite a `]`
    val params = Gen.Parameters.default
    val sample =
      (0L until 5000L).flatMap(seed => formulaText(params, org.scalacheck.rng.Seed(seed)))
    val closerFindings = sample.count(t =>
      UnparseableFormula.checkSlow(t).exists(_.startsWith("Unbalanced delimiter"))
    )
    val spared =
      sample.count(t => t.exists(c => c == ']' || c == '}') && !UnparseableFormula.needsParse(t))
    assert(closerFindings >= 20, s"closer findings: $closerFindings")
    assert(spared >= 20, s"spared bracketed texts: $spared")
  }

  property("#681: the prefiltered oracle answers exactly as the parse-backed one") {
    forAll(formulaText) { text =>
      Prop(LintCommands.formulaCheck(text) == UnparseableFormula.checkSlow(text)) :|
        s"text: $text"
    }
  }

  property("#681: the prefilter never hides a finding the full parse reports") {
    forAll(formulaText) { text =>
      Prop(UnparseableFormula.checkSlow(text).isEmpty || UnparseableFormula.needsParse(text)) :|
        s"text: $text"
    }
  }
