package com.tjclp.xl.formula

import munit.ScalaCheckSuite
import org.scalacheck.Gen
import org.scalacheck.Prop.*

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.CriteriaMatcher.*
import com.tjclp.xl.sheets.Sheet

/**
 * GH-765: how a criterion reads `~`.
 *
 *   - An operand (after any `<>` or `=` prefix) with no `*` and no `?` is compared as written: `~`
 *     is an ordinary character, so `"a~~b"` matches only the text `a~~b`. Desktop Excel (Microsoft
 *     365) did not count a cell holding one `~` against the same key pre-escaped as `~~`.
 *   - An operand with a `*` or `?` keeps the escape grammar: `~x` is a literal `x` for every
 *     character `x` (`~*`, `~?`, `~~`, and `~b` too, as desktop Excel reads `"a~b*"`); a `~` at the
 *     end is a literal tilde.
 *   - XLOOKUP/XMATCH match_mode 2 read every lookup value with the escape grammar (#774); SEARCH is
 *     unchanged.
 */
class CriteriaTildeSpec extends ScalaCheckSuite:

  // The issue's grid: A1:A6 a~b, a~~b, ab, a~bc, x~, x keyed to B1:B6 1, 10, ..., 100000.
  private val keys = Vector("a~b", "a~~b", "ab", "a~bc", "x~", "x")
  private val sheet: Sheet =
    keys.zipWithIndex.foldLeft(Sheet(SheetName.unsafe("S"))) { case (s, (k, i)) =>
      s.put(ref"A1".shift(0, i), CellValue.Text(k))
        .put(ref"B1".shift(0, i), CellValue.Number(BigDecimal(10).pow(i)))
    }

  private def check(formula: String, expected: CellValue)(implicit loc: munit.Location): Unit =
    assertEquals(sheet.evaluateFormula(formula), Right(expected), formula)

  private def num(n: Long): CellValue = CellValue.Number(BigDecimal(n))

  private def text(s: String): CellValue.Text = CellValue.Text(s)

  /** Escapes every `~`, `*` and `?`: the key a criterion needs to match `s` in escape mode. */
  private def escapeAll(s: String): String =
    s.flatMap(c => if c == '~' || c == '*' || c == '?' then s"~$c" else c.toString)

  // Text over `a`, `b` and `~` only: no wildcard, no operator prefix, never numeric.
  private val literalOperand: Gen[String] =
    Gen.listOf(Gen.oneOf('a', 'b', '~')).map(_.mkString)

  // Text over `a`, `b`, `~`, `*` and `?` with at least one `*` or `?`.
  private val wildcardText: Gen[String] =
    for
      prefix <- Gen.listOf(Gen.oneOf('a', 'b', '~', '*', '?'))
      wild <- Gen.oneOf('*', '?')
      suffix <- Gen.listOf(Gen.oneOf('a', 'b', '~', '*', '?'))
    yield (prefix ++ (wild :: suffix)).mkString

  // ===== Rule 1: no `*` or `?` means a literal compare =====

  test("GH-765: parse keeps ~ verbatim in an operand with no * or ?") {
    assertEquals(parse(ExprValue.Text("a~~b")), Exact(ExprValue.Text("a~~b")))
    assertEquals(parse(ExprValue.Text("a~b")), Exact(ExprValue.Text("a~b")))
    assertEquals(parse(ExprValue.Text("~~")), Exact(ExprValue.Text("~~")))
    assertEquals(parse(ExprValue.Text("x~")), Exact(ExprValue.Text("x~")))
    assertEquals(parse(ExprValue.Cell(text("a~~b"))), Exact(ExprValue.Text("a~~b")))
  }

  test("GH-765: <> keeps ~ verbatim in an operand with no * or ?") {
    assertEquals(parse(ExprValue.Text("<>a~~b")), CompareText(CompareOp.Neq, "a~~b"))
    assertEquals(parse(ExprValue.Text("<>~~")), CompareText(CompareOp.Neq, "~~"))
  }

  test("GH-765: = keeps its operand verbatim, ~ and wildcards included") {
    assertEquals(parse(ExprValue.Text("=a~~b")), Exact(ExprValue.Text("a~~b")))
    // xl reads `*` after `=` as a literal (Excel applies the wildcard; docs/LIMITATIONS.md)
    assertEquals(parse(ExprValue.Text("=a~~b*")), Exact(ExprValue.Text("a~~b*")))
    check("=COUNTIF(A1:A6,\"=a~~b\")", num(1))
    check("=SUMIFS(B1:B6,A1:A6,\"=a~b\")", num(1))
  }

  property("GH-765: an operand with no * or ? parses to itself") {
    forAllNoShrink(literalOperand) { s =>
      assertEquals(parse(ExprValue.Text(s)), Exact(ExprValue.Text(s)))
      assertEquals(parse(ExprValue.Text(s"<>$s")), CompareText(CompareOp.Neq, s))
    }
  }

  property("GH-765: a text cell with no * or ? used as its own criterion matches only itself") {
    forAllNoShrink(literalOperand, literalOperand) { (s, other) =>
      val criterion = parse(ExprValue.Cell(text(s)))
      assert(matches(text(s), criterion), s)
      assertEquals(matches(text(other), criterion), other.equalsIgnoreCase(s), s"$s vs $other")
      assertEquals(
        matches(text(other), parse(ExprValue.Text(s"<>$s"))),
        !other.equalsIgnoreCase(s),
        s"<>$s vs $other"
      )
    }
  }

  property("GH-765: an operand with no * or ? never reaches the lookups' wildcard path") {
    forAllNoShrink(literalOperand) { s =>
      parse(ExprValue.Text(s)) match
        case Wildcard(_) => fail(s"$s parsed as a Wildcard")
        case _ => ()
    }
  }

  // ===== Rule 2: an operand with `*` or `?` keeps the escape grammar =====

  test("GH-765: escape mode keeps ~*, ~? and ~~ as escapes") {
    assertEquals(parse(ExprValue.Text("~*")), Exact(ExprValue.Text("*")))
    assertEquals(parse(ExprValue.Text("a~~b~?")), Exact(ExprValue.Text("a~b?")))
    assertEquals(parse(ExprValue.Text("~**")), Wildcard("~**"))
    assertEquals(parse(ExprValue.Text("a~~b*")), Wildcard("a~~b*"))
    assertEquals(parse(ExprValue.Text("<>a~~b~*")), CompareText(CompareOp.Neq, "a~b*"))
    assert(matches(text("a~bc"), Wildcard("a~~b*")))
    assert(!matches(text("a~~bc"), Wildcard("a~~b*")))
  }

  test("escape mode drops ~ before an ordinary character and keeps a final ~ literal") {
    // desktop Excel: SUMIF(r,"a~b*",...) counts "abc" and "ab", not "a~b"
    assert(matches(text("abc"), Wildcard("a~b*")))
    assert(!matches(text("a~bc"), Wildcard("a~b*")))
    assert(matches(text("xy~"), Wildcard("x*~")))
  }

  property("GH-765: text with * or ?, every ~ * ? escaped, matches itself exactly") {
    forAllNoShrink(wildcardText) { s =>
      val criterion = parse(ExprValue.Text(escapeAll(s)))
      assertEquals(criterion, Exact(ExprValue.Text(s)))
      assert(matches(text(s), criterion), s)
    }
  }

  // ===== Callers: the issue's repro grid =====

  test("GH-765: SUMIF over the issue's grid") {
    check("=SUMIF(A1:A6,\"a~~b\",B1:B6)", num(10))
    check("=SUMIF(A1:A6,\"a~b\",B1:B6)", num(1))
    check("=SUMIF(A1:A6,\"a~~b*\",B1:B6)", num(1001))
    check("=SUMIF(A1:A6,\"a~b*\",B1:B6)", num(100)) // escape-mode `~b` is `b`
    check("=SUMIF(A1:A6,\"x~\",B1:B6)", num(10000))
    check("=SUMIF(A1:A6,A2,B1:B6)", num(10))
  }

  test("GH-765: COUNTIF and COUNTIFS compare an operand with no * or ? literally") {
    check("=COUNTIF(A1:A6,\"a~~b\")", num(1))
    check("=COUNTIFS(A1:A6,A2)", num(1))
    check("=SUMPRODUCT((A1:A6=A2)*1)", num(1))
    // each key counts exactly once as its own criterion
    check("=COUNTIFS(A1:A6,A2,B1:B6,10)", num(1))
    keys.indices.foreach(i => check(s"=COUNTIF(A1:A6,A${i + 1})", num(1)))
  }

  test("GH-765: SUMIFS, AVERAGEIF(S), MINIFS and MAXIFS compare literally") {
    check("=SUMIFS(B1:B6,A1:A6,\"a~~b\")", num(10))
    check("=AVERAGEIF(A1:A6,\"a~~b\",B1:B6)", num(10))
    check("=AVERAGEIFS(B1:B6,A1:A6,\"a~~b\")", num(10))
    check("=MINIFS(B1:B6,A1:A6,\"a~~b\")", num(10))
    check("=MAXIFS(B1:B6,A1:A6,\"a~~b\")", num(10))
    check("=MAXIFS(B1:B6,A1:A6,\"x~\")", num(10000))
  }

  test("GH-765: <> compares an operand with no * or ? literally") {
    check("=SUMIF(A1:A6,\"<>a~~b\",B1:B6)", num(111101))
    check("=COUNTIFS(A1:A6,\"<>a~~b\")", num(5))
    check("=SUMIF(A1:A6,\"<>a~~b*\",B1:B6)", num(110110)) // NotWildcard on a~b*
  }

  test("#774: XLOOKUP and XMATCH match_mode 2 always read ~ as an escape") {
    check("=XMATCH(\"a~~b\",A1:A6,2)", num(1))
    check("=XLOOKUP(\"a~~b\",A1:A6,B1:B6,,2)", num(1))
    check("=XMATCH(\"a~~b*\",A1:A6,2)", num(1))
    check("=XMATCH(\"x~\",A1:A6,2)", num(5))
  }

  test("GH-765: SEARCH keeps its ~ escapes") {
    check("=SEARCH(\"~~\",\"a~~b\")", num(2))
    check("=SEARCH(\"~*\",\"a*b\")", num(2))
    check("=SEARCH(\"a~b\",\"xa~b\")", num(2))
  }
