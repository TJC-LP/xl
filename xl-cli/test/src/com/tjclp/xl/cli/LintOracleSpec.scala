package com.tjclp.xl.cli

import java.util.concurrent.atomic.AtomicReference

import munit.{FunSuite, ScalaCheckSuite}
import org.scalacheck.{Gen, Prop}
import org.scalacheck.Prop.forAll

import com.tjclp.xl.cli.commands.LintCommands
import com.tjclp.xl.formula.parser.UnparseableFormula

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
    val out = new AtomicReference[Either[Throwable, A]](Left(new IllegalStateException("not run")))
    val thread = Thread
      .ofPlatform()
      .name("lint-oracle-small-stack")
      .stackSize(stackBytes)
      .start(() =>
        out.set(
          try Right(f)
          catch case t: Throwable => Left(t)
        )
      )
    thread.join()
    out.get()

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
    // stack (a fiber, a native image) the deepest legal nests still overflow. Every text must come
    // back as a verdict — the overflow as a finding — on a stack small enough to overflow.
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
    val results =
      texts.map(t => t.take(20) -> onSmallStack(64L * 1024)(LintCommands.formulaCheck(t)))
    val escaped = results.collect { case (t, Left(err)) => s"$t…: $err" }
    assertEquals(escaped, Vector.empty[String], escaped.mkString("\n"))
    // and at least one of them really overflowed — the guard, not luck, answered
    assert(
      results.exists { case (_, verdict) => verdict.exists(_.exists(_.contains("stack"))) },
      results.mkString("\n")
    )
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
