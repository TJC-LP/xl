package com.tjclp.xl.formula

import com.tjclp.xl.XLResult
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.eval.SheetEvaluator.*
import com.tjclp.xl.formula.functions.ArgValue
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.syntax.*
import munit.FunSuite

/**
 * GH-603: Excel lets a formula omit an argument by leaving its slot empty — `RATE(nper,,pv,fv,)`,
 * `IF(cond,,x)`, `PMT(r,n,pv,,1)`. The parser emits `TExpr.Missing` for every such slot; an
 * optional slot reads it as absent (the function's default applies), a required typed slot coerces
 * it like a blank cell, an Any position reads it as 0, and the printer keeps interior slots empty
 * while dropping trailing ones.
 */
class OmittedArgumentSpec extends FunSuite:

  private val sheet = Sheet("Test")
    .put(ref"A1", CellValue.Number(BigDecimal(2)))
    .put(ref"A2", CellValue.Number(BigDecimal(3)))
    .put(ref"B1", CellValue.Number(BigDecimal(10)))
    .put(ref"B2", CellValue.Number(BigDecimal(20)))

  private def value(formula: String): XLResult[CellValue] = sheet.evaluateFormula(formula)

  private def number(formula: String): BigDecimal =
    value(formula) match
      case Right(CellValue.Number(n)) => n
      case other => fail(s"$formula: expected a number, got $other")

  private def printed(formula: String): String =
    FormulaParser.parse(formula).map(FormulaPrinter.print(_)).fold(e => fail(e.toString), identity)

  test("RATE(10,,-100,150,) — the dogfood idiom: omitted pmt is 0, omitted type is the default") {
    // Excel: =RATE(10,,-100,150,) is 4.1380% — (150/100)^(1/10) - 1
    assertEqualsDouble(number("=RATE(10,,-100,150,)").toDouble, 0.0413797439, 1e-9)
    assertEquals(number("=RATE(10,,-100,150,)"), number("=RATE(10,0,-100,150,0)"))
  }

  test("IF(TRUE,,5) is 0 and IF(FALSE,5,) is 0 — an omitted branch is Excel's 0") {
    assertEquals(value("=IF(TRUE,,5)"), Right(CellValue.Number(BigDecimal(0))))
    assertEquals(value("=IF(FALSE,5,)"), Right(CellValue.Number(BigDecimal(0))))
    assertEquals(value("=IF(A1>1,,B1)"), Right(CellValue.Number(BigDecimal(0))))
    assertEquals(value("=IF(A1>5,,B1)"), Right(CellValue.Number(BigDecimal(10))))
  }

  test("an omitted required slot coerces like a blank cell: 0 in numbers, \"\" in text") {
    assertEquals(number("=SUM(1,,2)"), BigDecimal(3))
    assertEquals(value("=LEFT(\"abc\",)"), Right(CellValue.Text("")))
    assertEquals(number("=POWER(,3)"), BigDecimal(0))
    assertEquals(number("=ROUND(2.567,)"), BigDecimal(3))
  }

  test("an omitted optional slot is absent: the function's own default applies, not 0") {
    // PMT's fv defaults to 0 and type to 0 — an empty fourth slot before an explicit type keeps
    // the position: PMT(r,n,pv,,1) is PMT(r,n,pv,0,1), not PMT(r,n,pv,1)
    assertEquals(number("=PMT(0.05,10,1000,,1)"), number("=PMT(0.05,10,1000,0,1)"))
    assert(number("=PMT(0.05,10,1000,,1)") != number("=PMT(0.05,10,1000,1)"))
    // RATE's guess stays 10% when its slot is omitted — the same root, not a guess of 0
    assertEquals(number("=RATE(10,,-100,150,,)"), number("=RATE(10,,-100,150)"))
    // VLOOKUP's range_lookup stays TRUE (approximate match) when omitted
    assertEquals(number("=VLOOKUP(2.5,A1:B2,2,)"), BigDecimal(10))
  }

  test("F() is still the zero-argument call; F(,) is two omitted arguments") {
    assert(FormulaParser.parse("=TODAY()").isRight)
    assert(FormulaParser.parse("=RAND()").isRight)
    assertEquals(number("=SUM(,)"), BigDecimal(0))
    FormulaParser.parse("=SUM(,)") match
      case Right(call: TExpr.Call[?]) =>
        assertEquals(
          call.spec.argSpec.toValues(call.args).count(_ == ArgValue.Expr(TExpr.Missing)),
          0,
          "SUM's numeric slots coerce Missing (Coerced wrapper) — it never survives bare"
        )
      case other => fail(s"expected a SUM call, got $other")
  }

  test("the omitted slot is TExpr.Missing in the AST (Any positions keep it bare)") {
    FormulaParser.parse("=IF(TRUE,,5)") match
      case Right(call: TExpr.Call[?]) =>
        val values = call.spec.argSpec.toValues(call.args)
        assertEquals(values.length, 3)
        assertEquals(values(1), ArgValue.Expr(TExpr.Missing))
      case other => fail(s"expected an IF call, got $other")
  }

  test("a range slot refuses an omitted argument with the slot's own message") {
    FormulaParser.parse("=SUMIF(,\">0\")") match
      case Left(_: com.tjclp.xl.formula.parser.ParseError.InvalidArguments) => ()
      case other => fail(s"expected InvalidArguments, got $other")
  }

  test("printing keeps interior omitted slots and drops trailing ones; parse ∘ print = id") {
    assertEquals(printed("=PMT(0.05,10,1000,,1)"), "=PMT(0.05, 10, 1000,, 1)")
    assertEquals(printed("=IF(TRUE,,5)"), "=IF(TRUE,, 5)")
    assertEquals(printed("=RATE(10,,-100,150,)"), "=RATE(10,, -100, 150)")
    assertEquals(printed("=A1+INDEX(A1:B2,,2)"), "=A1+INDEX(A1:B2,, 2)")
    List(
      "=PMT(0.05,10,1000,,1)",
      "=IF(TRUE,,5)",
      "=RATE(10,,-100,150,)",
      "=SUM(1,,2)",
      "=INDEX(A1:B2,,2)",
      "=LEFT(\"abc\",)"
    ).foreach { f =>
      val parsed = FormulaParser.parse(f)
      assert(parsed.isRight, s"$f: $parsed")
      val reprinted = parsed.map(FormulaPrinter.print(_))
      assertEquals(reprinted.flatMap(FormulaParser.parse), parsed, s"$f via $reprinted")
    }
    // a trailing omitted OPTIONAL slot is absent after reprint — the same tree either way
    assertEquals(
      FormulaParser.parse("=RATE(10,,-100,150,)"),
      FormulaParser.parse("=RATE(10,,-100,150)")
    )
  }

  test("the file form keeps the bare-comma spelling Excel writes") {
    val parsed = FormulaParser.parse("=PMT(0.05,10,1000,,1)").fold(e => fail(e.toString), identity)
    assertEquals(FormulaPrinter.printFileForm(parsed), "PMT(0.05,10,1000,,1)")
  }

  test("omitted slots shift with the formula like any other argument") {
    val parsed = FormulaParser.parse("=PMT(A1,B1,A2,,1)").fold(e => fail(e.toString), identity)
    val shifted = FormulaShifter.shift(parsed, colDelta = 1, rowDelta = 1)
    assertEquals(FormulaPrinter.print(shifted), "=PMT(B2, C2, B3,, 1)")
  }
