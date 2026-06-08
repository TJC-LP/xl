package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.parser.{FormulaParser, ParseError}
import munit.FunSuite

/**
 * GH-56: pathologically nested formulas must be rejected with a `ParseError.NestingTooDeep`
 * (a total `Left`), never a `StackOverflowError`. Capping parse depth also keeps the AST shallow
 * enough that evaluation cannot overflow, so this single guard protects both parser and evaluator.
 */
class RecursionGuardSpec extends FunSuite:

  private val s = new Sheet(name = SheetName.unsafe("S"))

  test("deeply nested parentheses → NestingTooDeep, not StackOverflowError") {
    val formula = "=" + ("(" * 300) + "1" + (")" * 300)
    FormulaParser.parse(formula) match
      case Left(_: ParseError.NestingTooDeep) => ()
      case other => fail(s"expected NestingTooDeep, got $other")
  }

  test("deeply nested unary minus → NestingTooDeep, not StackOverflowError") {
    val formula = "=" + ("-" * 300) + "1"
    FormulaParser.parse(formula) match
      case Left(_: ParseError.NestingTooDeep) => ()
      case other => fail(s"expected NestingTooDeep, got $other")
  }

  test("deeply nested function calls → NestingTooDeep") {
    val formula = "=" + ("ABS(" * 300) + "1" + (")" * 300)
    FormulaParser.parse(formula) match
      case Left(_: ParseError.NestingTooDeep) => ()
      case other => fail(s"expected NestingTooDeep, got $other")
  }

  test("deeply nested formula via evaluateFormula returns Left, never throws") {
    val formula = "=" + ("(" * 300) + "1" + (")" * 300)
    assert(s.evaluateFormula(formula).isLeft)
  }

  test("normal moderately-nested formulas still parse and evaluate") {
    assertEquals(s.evaluateFormula("=((((1+2))))*3"), Right(CellValue.Number(BigDecimal(9))))
    assert(FormulaParser.parse("=SUM(1,ABS(-2),MAX(3,4))").isRight)
    // ~50 nested parens is well under the 256 cap and must still work
    val ok = "=" + ("(" * 50) + "7" + (")" * 50)
    assertEquals(s.evaluateFormula(ok), Right(CellValue.Number(BigDecimal(7))))
  }

  // The binary operators are right-recursive, so chained operators are a distinct overflow
  // vector from parens/unary (caught by the claude-review of PR #250).
  private def assertTooDeep(formula: String): Unit =
    FormulaParser.parse(formula) match
      case Left(_: ParseError.NestingTooDeep) => ()
      case other => fail(s"expected NestingTooDeep, got $other")

  test("deeply chained binary operators → NestingTooDeep, not StackOverflowError") {
    assertTooDeep("=1" + ("=1" * 300)) // comparison chain
    assertTooDeep("=1" + ("&1" * 300)) // concatenation chain
    assertTooDeep("=1" + (" AND 1" * 300)) // logical AND chain
    assertTooDeep("=1" + (" OR 1" * 300)) // logical OR chain
    assertTooDeep("=1" + ("<1" * 300)) // inequality chain
  }

  test("normal short operator chains still parse") {
    assert(FormulaParser.parse("=1=1").isRight)
    assert(FormulaParser.parse("=1&2&3").isRight)
    assert(FormulaParser.parse("=TRUE AND FALSE OR TRUE").isRight)
  }

  test("flat additive/multiplicative chains are bounded (no eval StackOverflow)") {
    // Left-associative chains are built by @tailrec loops (the parser doesn't recurse), but they
    // produce a deep left-nested AST that the evaluator walks recursively. Counting each chain term
    // against the depth budget rejects pathological chains at parse (total Left) instead of letting
    // eval overflow. (GH-56, second review round.)
    assertTooDeep("=1" + ("+1" * 4000))
    assertTooDeep("=1" + ("*1" * 4000))
    // short chains still evaluate fine
    assertEquals(s.evaluateFormula("=1" + ("+1" * 100)), Right(CellValue.Number(BigDecimal(101))))
  }
