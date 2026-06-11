package com.tjclp.xl.cli.helpers

import munit.FunSuite

import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.helpers.FilterPredicate.{CmpOp, Literal, Pred}

/**
 * Tests for the filter predicate grammar and evaluator (GH-134, phase 1).
 *
 * Grammar: comparisons (= != <> > >= < <=) between a column ref and a literal, AND/OR/NOT with
 * parens, LIKE 'pat%', BETWEEN x AND y, IN (a, b, c), IS [NOT] EMPTY. Total: parse errors are Left,
 * evaluation never throws — type mismatches simply don't match.
 */
class FilterPredicateSpec extends FunSuite:

  // ========== Parsing ==========

  test("parse: numeric comparison") {
    assertEquals(
      FilterPredicate.parse("B > 100"),
      Right(Pred.Cmp("B", CmpOp.Gt, Literal.Num(BigDecimal(100))))
    )
  }

  test("parse: all comparison operators") {
    val ops = Map(
      "=" -> CmpOp.Eq,
      "!=" -> CmpOp.Ne,
      "<>" -> CmpOp.Ne,
      ">" -> CmpOp.Gt,
      ">=" -> CmpOp.Ge,
      "<" -> CmpOp.Lt,
      "<=" -> CmpOp.Le
    )
    ops.foreach { (sym, op) =>
      assertEquals(
        FilterPredicate.parse(s"A $sym 5"),
        Right(Pred.Cmp("A", op, Literal.Num(BigDecimal(5)))),
        s"operator $sym"
      )
    }
  }

  test("parse: string and boolean literals") {
    assertEquals(
      FilterPredicate.parse("A = 'Widget'"),
      Right(Pred.Cmp("A", CmpOp.Eq, Literal.Str("Widget")))
    )
    assertEquals(
      FilterPredicate.parse("A = \"Widget\""),
      Right(Pred.Cmp("A", CmpOp.Eq, Literal.Str("Widget")))
    )
    assertEquals(
      FilterPredicate.parse("C = TRUE"),
      Right(Pred.Cmp("C", CmpOp.Eq, Literal.Bool(true)))
    )
    assertEquals(
      FilterPredicate.parse("C != false"),
      Right(Pred.Cmp("C", CmpOp.Ne, Literal.Bool(false)))
    )
  }

  test("parse: negative numbers") {
    assertEquals(
      FilterPredicate.parse("A < -5.5"),
      Right(Pred.Cmp("A", CmpOp.Lt, Literal.Num(BigDecimal("-5.5"))))
    )
  }

  test("parse: AND binds tighter than OR; NOT tighter than AND") {
    val parsed = FilterPredicate.parse("A > 1 OR B > 2 AND NOT C > 3")
    val expected = Pred.Or(
      Pred.Cmp("A", CmpOp.Gt, Literal.Num(BigDecimal(1))),
      Pred.And(
        Pred.Cmp("B", CmpOp.Gt, Literal.Num(BigDecimal(2))),
        Pred.Not(Pred.Cmp("C", CmpOp.Gt, Literal.Num(BigDecimal(3))))
      )
    )
    assertEquals(parsed, Right(expected))
  }

  test("parse: parentheses override precedence") {
    val parsed = FilterPredicate.parse("(A > 1 OR B > 2) AND C > 3")
    val expected = Pred.And(
      Pred.Or(
        Pred.Cmp("A", CmpOp.Gt, Literal.Num(BigDecimal(1))),
        Pred.Cmp("B", CmpOp.Gt, Literal.Num(BigDecimal(2)))
      ),
      Pred.Cmp("C", CmpOp.Gt, Literal.Num(BigDecimal(3)))
    )
    assertEquals(parsed, Right(expected))
  }

  test("parse: keywords are case-insensitive") {
    val parsed = FilterPredicate.parse("a like 'x%' and b between 1 and 2 or c is not empty")
    assert(parsed.isRight, s"Got $parsed")
  }

  test("parse: LIKE, BETWEEN, IN, IS EMPTY forms") {
    assertEquals(
      FilterPredicate.parse("A LIKE 'Widget%'"),
      Right(Pred.Like("A", "Widget%"))
    )
    assertEquals(
      FilterPredicate.parse("B BETWEEN 10 AND 100"),
      Right(Pred.Between("B", Literal.Num(BigDecimal(10)), Literal.Num(BigDecimal(100))))
    )
    assertEquals(
      FilterPredicate.parse("A IN ('x', 'y', 'z')"),
      Right(Pred.In("A", List(Literal.Str("x"), Literal.Str("y"), Literal.Str("z"))))
    )
    assertEquals(FilterPredicate.parse("A IS EMPTY"), Right(Pred.IsEmpty("A", negated = false)))
    assertEquals(FilterPredicate.parse("A IS NOT EMPTY"), Right(Pred.IsEmpty("A", negated = true)))
  }

  test("parse: doubled quotes escape inside strings") {
    assertEquals(
      FilterPredicate.parse("A = 'it''s'"),
      Right(Pred.Cmp("A", CmpOp.Eq, Literal.Str("it's")))
    )
  }

  test("parse: errors are Left with a message") {
    assert(FilterPredicate.parse("").isLeft, "empty input")
    assert(FilterPredicate.parse("A >").isLeft, "dangling operator")
    assert(FilterPredicate.parse("A > 1 AND").isLeft, "dangling AND")
    assert(FilterPredicate.parse("(A > 1").isLeft, "unbalanced paren")
    assert(FilterPredicate.parse("A > 1 extra").isLeft, "trailing tokens")
    assert(FilterPredicate.parse("A BETWEEN 1 AND 'x'").isLeft, "mixed BETWEEN bound types")
    assert(FilterPredicate.parse("A IN ()").isLeft, "empty IN list")
    assert(FilterPredicate.parse("A LIKE 5").isLeft, "LIKE requires string pattern")
  }

  test("parse: collects referenced columns") {
    val parsed = FilterPredicate.parse("A > 1 AND (price < 2 OR C IS EMPTY)")
    assertEquals(parsed.map(FilterPredicate.columnRefs), Right(Set("A", "price", "C")))
  }

  // ========== Evaluation ==========

  private def eval(
    where: String,
    cells: Map[String, CellValue],
    resolve: Map[String, Int] = Map("A" -> 0, "B" -> 1, "C" -> 2)
  ): Boolean =
    val pred = FilterPredicate.parse(where).toOption.get
    val byIdx = cells.flatMap((name, v) => resolve.get(name).map(_ -> v))
    FilterPredicate.evaluate(pred, name => resolve.get(name), idx => byIdx.get(idx))

  test("evaluate: numeric comparisons on Number cells") {
    assert(eval("A > 100", Map("A" -> CellValue.Number(150))))
    assert(!eval("A > 100", Map("A" -> CellValue.Number(50))))
    assert(eval("A <= 100", Map("A" -> CellValue.Number(100))))
    assert(eval("A = 100", Map("A" -> CellValue.Number(BigDecimal("100.0")))))
    assert(eval("A != 99", Map("A" -> CellValue.Number(100))))
  }

  test("evaluate: type mismatch never matches and never errors") {
    assert(!eval("A > 100", Map("A" -> CellValue.Text("abc"))), "text vs number")
    assert(!eval("A != 100", Map("A" -> CellValue.Text("abc"))), "mismatch fails even !=")
    assert(!eval("A = 'x'", Map("A" -> CellValue.Number(5))), "number vs string")
    assert(!eval("A > 100", Map.empty), "missing cell")
    assert(!eval("A = TRUE", Map("A" -> CellValue.Text("TRUE"))), "text vs boolean")
  }

  test("evaluate: string comparisons are case-insensitive") {
    assert(eval("A = 'widget'", Map("A" -> CellValue.Text("Widget"))))
    assert(eval("A < 'b'", Map("A" -> CellValue.Text("Apple"))))
  }

  test("evaluate: boolean equality") {
    assert(eval("A = TRUE", Map("A" -> CellValue.Bool(true))))
    assert(!eval("A = TRUE", Map("A" -> CellValue.Bool(false))))
  }

  test("evaluate: LIKE with % wildcards, case-insensitive") {
    assert(eval("A LIKE 'Widget%'", Map("A" -> CellValue.Text("widget pro"))))
    assert(eval("A LIKE '%pro'", Map("A" -> CellValue.Text("Widget Pro"))))
    assert(eval("A LIKE '%dge%'", Map("A" -> CellValue.Text("Widget"))))
    assert(!eval("A LIKE 'Widget%'", Map("A" -> CellValue.Text("Pro Widget"))))
    assert(!eval("A LIKE 'W%'", Map("A" -> CellValue.Number(5))), "LIKE on non-text")
  }

  test("evaluate: LIKE escapes regex metacharacters in the pattern") {
    assert(eval("A LIKE 'a.b%'", Map("A" -> CellValue.Text("a.b-c"))))
    assert(!eval("A LIKE 'a.b%'", Map("A" -> CellValue.Text("axb-c"))), ". is literal")
  }

  test("evaluate: BETWEEN is inclusive") {
    assert(eval("A BETWEEN 10 AND 100", Map("A" -> CellValue.Number(10))))
    assert(eval("A BETWEEN 10 AND 100", Map("A" -> CellValue.Number(100))))
    assert(!eval("A BETWEEN 10 AND 100", Map("A" -> CellValue.Number(101))))
    assert(!eval("A BETWEEN 10 AND 100", Map("A" -> CellValue.Text("50"))), "text never between")
  }

  test("evaluate: IN membership") {
    assert(eval("A IN ('x', 'y')", Map("A" -> CellValue.Text("Y"))))
    assert(!eval("A IN ('x', 'y')", Map("A" -> CellValue.Text("z"))))
    assert(eval("A IN (1, 2, 3)", Map("A" -> CellValue.Number(2))))
  }

  test("evaluate: IS EMPTY on missing, Empty, and blank text") {
    assert(eval("A IS EMPTY", Map.empty))
    assert(eval("A IS EMPTY", Map("A" -> CellValue.Empty)))
    assert(eval("A IS EMPTY", Map("A" -> CellValue.Text("  "))))
    assert(!eval("A IS EMPTY", Map("A" -> CellValue.Number(0))))
    assert(eval("A IS NOT EMPTY", Map("A" -> CellValue.Number(0))))
  }

  test("evaluate: formula cells use cached value") {
    assert(eval("A > 100", Map("A" -> CellValue.Formula("=B1*2", Some(CellValue.Number(150))))))
    assert(!eval("A > 100", Map("A" -> CellValue.Formula("=B1*2", None))), "no cache, no match")
    assert(eval("A IS NOT EMPTY", Map("A" -> CellValue.Formula("=B1", None))))
  }

  test("evaluate: AND/OR/NOT combinations") {
    val cells = Map("A" -> CellValue.Number(5), "B" -> CellValue.Text("x"))
    assert(eval("A > 1 AND B = 'x'", cells))
    assert(!eval("A > 10 AND B = 'x'", cells))
    assert(eval("A > 10 OR B = 'x'", cells))
    assert(eval("NOT A > 10", cells))
  }
