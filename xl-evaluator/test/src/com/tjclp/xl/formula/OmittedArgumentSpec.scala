package com.tjclp.xl.formula

import com.tjclp.xl.XLResult
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.ast.BindingCoercion
import com.tjclp.xl.formula.eval.SheetEvaluator.*
import com.tjclp.xl.formula.functions.ArgValue
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.syntax.*
import munit.FunSuite

/**
 * GH-603: Excel lets a formula omit an argument by leaving its slot empty — `RATE(nper,,pv,fv,)`,
 * `IF(cond,,x)`, `PMT(r,n,pv,,1)`. The parser emits `TExpr.Missing` for every such slot — Excel's
 * blank: a typed slot coerces it like a blank cell, a value slot reads it as 0, and an optional
 * slot holds it as a PRESENT blank (GH-654: `VLOOKUP(x,rng,2,)` is an exact match), except in the
 * dynamic-array functions, which read it as omitted the way Excel does. Every slot prints back
 * empty, trailing ones included. Expected values verified against LibreOffice's recalculation.
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

  test("GH-654: an empty optional slot is a present blank — exact-match lookups, a guess of 0") {
    // PMT's fv: an empty fourth slot before an explicit type is fv 0 (PMT(r,n,pv,0,1))
    assertEquals(number("=PMT(0.05,10,1000,,1)"), number("=PMT(0.05,10,1000,0,1)"))
    assert(number("=PMT(0.05,10,1000,,1)") != number("=PMT(0.05,10,1000,1)"))
    // VLOOKUP's range_lookup: the empty slot is FALSE — the `,)` idiom every model uses for an
    // exact match — so 2.5 finds nothing where the approximate match returned 10
    val vlookup = value("=VLOOKUP(2.5,A1:B2,2,)")
    assert(vlookup.swap.exists(_.toString.contains("exact match not found")), vlookup.toString)
    assertEquals(number("=VLOOKUP(2.5,A1:B2,2)"), BigDecimal(10), "the absent slot is TRUE")
    assertEquals(number("=VLOOKUP(2,A1:B2,2,)"), BigDecimal(10))
    // MATCH's match_type: the empty slot is 0 (exact), not the default 1
    val matched = value("=MATCH(2.5,A1:A2,)")
    assert(matched.swap.exists(_.toString.contains("no match")), matched.toString)
    assertEquals(number("=MATCH(2.5,A1:A2)"), BigDecimal(1), "the absent slot is 1")
    assertEquals(number("=MATCH(3,A1:A2,)"), BigDecimal(2))
    // RATE's guess is 0 rather than 10% — Newton still reaches the same root
    assertEqualsDouble(number("=RATE(10,,-100,150,,)").toDouble, 0.0413797439, 1e-9)
    // LOG's base is 0 — an error, as in Excel (#NUM!) and LibreOffice (Err:502), never 10
    assert(value("=LOG(10,)").isLeft || value("=LOG(10,)").exists(_.isInstanceOf[CellValue.Error]))
  }

  test("GH-654: the dynamic-array and reference functions read the empty slot as omitted") {
    // Excel tests presence in these slots (an explicit 0 is an error): the empty slot is the default
    assertEquals(number("=SUM(OFFSET(A1,0,0,,2))"), BigDecimal(12))
    assertEquals(number("=SUM(SEQUENCE(3,,5))"), BigDecimal(18))
    assertEquals(number("=SUM(SEQUENCE(2,2,,))"), BigDecimal(10))
    assertEquals(number("=SUM(SORT(A1:A2,,-1))"), BigDecimal(5))
    val (sorted, _) = sheet.evaluateArrayFormula("=SORT(A1:A2,,-1)", ref"E1").toOption.get
    assertEquals(sorted(ref"E1").value, CellValue.Number(BigDecimal(3)), "sorted by column 1, desc")
    assertEquals(number("=SUM(UNIQUE(A1:A2,,TRUE))"), BigDecimal(5))
    // XLOOKUP's if_not_found: the empty slot is #N/A on no match, not a blank result
    val notFound = value("=XLOOKUP(99,A1:A2,B1:B2,,0)")
    assert(
      notFound.isLeft || notFound.exists(_.isInstanceOf[CellValue.Error]),
      s"empty if_not_found is omitted, so no match is #N/A: $notFound"
    )
    assertEquals(value("=XLOOKUP(99,A1:A2,B1:B2,\"nf\",)"), Right(CellValue.Text("nf")))
    assertEquals(number("=XLOOKUP(3,A1:A2,B1:B2,,)"), BigDecimal(20))
    // FILTER's if_empty: the empty slot is omitted, so no match is the #N/A of the two-arg form
    assertEquals(value("=FILTER(A1:A2,A1:A2>100,)"), value("=FILTER(A1:A2,A1:A2>100)"))
  }

  test("GH-654: INDEX with an empty (or 0) position selects the whole row or column") {
    assertEquals(number("=SUM(INDEX(A1:B2,,2))"), BigDecimal(30))
    assertEquals(number("=SUM(INDEX(A1:B2,0,2))"), BigDecimal(30))
    assertEquals(number("=SUM(INDEX(A1:B2,2,0))"), BigDecimal(23))
    assertEquals(number("=SUM(INDEX(A1:B2,,))"), BigDecimal(35))
    // the two-argument form on a 2-D array is row_num with the whole row
    assertEquals(number("=SUM(INDEX(A1:B2,2))"), BigDecimal(23))
    // standalone, the whole column spills
    val (s, range) = sheet.evaluateArrayFormula("=INDEX(A1:B2,,2)", ref"E1").toOption.get
    assertEquals(range.height, 2)
    assertEquals(range.width, 1)
    assertEquals(s(ref"E1").value, CellValue.Number(BigDecimal(10)))
    assertEquals(s(ref"E2").value, CellValue.Number(BigDecimal(20)))
    // a scalar position still collapses to the top-left cell, the value it always returned
    assertEquals(number("=INDEX(A1:B2,,2)"), BigDecimal(10))
    assertEquals(number("=INDEX(A1:B2,2,2)"), BigDecimal(20))
    // a whole-column array bounds the whole-axis selection to the used range
    assertEquals(number("=SUM(INDEX(A:B,0,2))"), BigDecimal(30))
    assertEquals(number("=SUM(INDEX(A:A,,))"), BigDecimal(5))
  }

  test("GH-654: a value slot reads the empty argument as 0 on the scalar and the array route") {
    // A1 = 2, A2 = 3: the TRUE branch is the omitted one
    val (s, _) = sheet.evaluateArrayFormula("=IF(A1:A2>2,,5)", ref"E1").toOption.get
    assertEquals(s(ref"E1").value, CellValue.Number(BigDecimal(5)))
    assertEquals(s(ref"E2").value, CellValue.Number(BigDecimal(0)))
    val (t, _) = sheet.evaluateArrayFormula("=IF(A1:A2>2,5,)", ref"E1").toOption.get
    assertEquals(t(ref"E1").value, CellValue.Number(BigDecimal(0)))
    assertEquals(t(ref"E2").value, CellValue.Number(BigDecimal(5)))
    assertEquals(number("=CHOOSE(1,,2)"), BigDecimal(0))
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

  test("the omitted slot is TExpr.Missing in the AST, coerced to 0 in a value position") {
    FormulaParser.parse("=IF(TRUE,,5)") match
      case Right(call: TExpr.Call[?]) =>
        val values = call.spec.argSpec.toValues(call.args)
        assertEquals(values.length, 3)
        values(1) match
          case ArgValue.Expr(TExpr.Coerced(TExpr.Missing, BindingCoercion.Numeric)) => ()
          case other => fail(s"expected the empty slot coerced to 0, got $other")
      case other => fail(s"expected an IF call, got $other")
  }

  test("a range slot refuses an omitted argument with the slot's own message") {
    FormulaParser.parse("=SUMIF(,\">0\")") match
      case Left(_: com.tjclp.xl.formula.parser.ParseError.InvalidArguments) => ()
      case other => fail(s"expected InvalidArguments, got $other")
  }

  test("printing keeps every empty slot, trailing ones included; parse ∘ print = id") {
    assertEquals(printed("=PMT(0.05,10,1000,,1)"), "=PMT(0.05, 10, 1000,, 1)")
    assertEquals(printed("=IF(TRUE,,5)"), "=IF(TRUE,, 5)")
    // GH-654: a trailing empty slot is kept — dropping it would turn VLOOKUP's exact match into
    // an approximate one on every reprint (drag, insert-rows, rename-sheet)
    assertEquals(printed("=RATE(10,,-100,150,)"), "=RATE(10,, -100, 150,)")
    assertEquals(printed("=VLOOKUP(2.5,A1:B2,2,)"), "=VLOOKUP(2.5, A1:B2, 2,)")
    assertEquals(printed("=A1+INDEX(A1:B2,,2)"), "=A1+INDEX(A1:B2,, 2)")
    List(
      "=PMT(0.05,10,1000,,1)",
      "=IF(TRUE,,5)",
      "=RATE(10,,-100,150,)",
      "=RATE(10,,-100,150,,)",
      "=SUM(1,,2)",
      "=INDEX(A1:B2,,2)",
      "=LEFT(\"abc\",)",
      "=LOG(,)",
      "=TRUNC(,)",
      "=MATCH($A2,Data!$A:$A,)",
      "=XLOOKUP(1,A1:A2,B1:B2,,0)",
      "=SORT(A1:B2,,-1)"
    ).foreach { f =>
      val parsed = FormulaParser.parse(f)
      assert(parsed.isRight, s"$f: $parsed")
      val reprinted = parsed.map(FormulaPrinter.print(_))
      assertEquals(reprinted.flatMap(FormulaParser.parse), parsed, s"$f via $reprinted")
    }
    // the empty slot and the absent argument are different trees with different meanings
    assertNotEquals(
      FormulaParser.parse("=VLOOKUP(2.5,A1:B2,2,)"),
      FormulaParser.parse("=VLOOKUP(2.5,A1:B2,2)")
    )
  }

  test("the file form keeps the bare-comma spelling Excel writes, trailing slot included") {
    val parsed = FormulaParser.parse("=PMT(0.05,10,1000,,1)").fold(e => fail(e.toString), identity)
    assertEquals(FormulaPrinter.printFileForm(parsed), "PMT(0.05,10,1000,,1)")
    val exact =
      FormulaParser.parse("=MATCH($A2,Data!$A:$A,)").fold(e => fail(e.toString), identity)
    assertEquals(FormulaPrinter.printFileForm(exact), "MATCH($A2,Data!$A:$A,)")
    // the shifter (drags, structural edits) preserves it too
    assertEquals(
      FormulaPrinter.printFileForm(FormulaShifter.shift(exact, colDelta = 0, rowDelta = 1)),
      "MATCH($A3,Data!$A:$A,)"
    )
  }

  test("omitted slots shift with the formula like any other argument") {
    val parsed = FormulaParser.parse("=PMT(A1,B1,A2,,1)").fold(e => fail(e.toString), identity)
    val shifted = FormulaShifter.shift(parsed, colDelta = 1, rowDelta = 1)
    assertEquals(FormulaPrinter.print(shifted), "=PMT(B2, C2, B3,, 1)")
  }
