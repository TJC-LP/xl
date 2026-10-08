package com.tjclp.xl.formula

import munit.FunSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.eval.{EvalError, Evaluator}
import com.tjclp.xl.sheets.Sheet

/**
 * #692: the lookup family after #670 — one error channel for every miss (XLOOKUP raises the typed
 * `#N/A` through `lookupNotFound` like the rest, with its call as the diagnostic), XLOOKUP and
 * XMATCH agree that a match_mode or search_mode outside its table is Excel's `#VALUE!`, and MATCH
 * and XLOOKUP read their lookup range through the evaluating reader, so an uncached formula key is
 * computed rather than skipped.
 */
class LookupChannelSpec extends FunSuite:

  // A1:A5 10..50 keyed to B1:B5 a..e; G1 blank; W1:W2 two misses
  private val sheet: Sheet =
    (0 until 5)
      .foldLeft(Sheet(SheetName.unsafe("S"))) { (s, i) =>
        s.put(ref"A1".shift(0, i), num(10 * (i + 1)))
          .put(ref"B1".shift(0, i), text(("a".charAt(0) + i).toChar.toString))
      }
      .put(ref"W1", num(5))
      .put(ref"W2", num(55))

  private def num(n: Int): CellValue = CellValue.Number(BigDecimal(n))
  private def text(s: String): CellValue = CellValue.Text(s)
  private def err(e: CellError): CellValue = CellValue.Error(e)

  private def check(on: Sheet, formula: String, expected: CellValue)(implicit
    loc: munit.Location
  ): Unit =
    assertEquals(on.evaluateFormula(formula), Right(expected), formula)

  private def leftOf(formula: String)(implicit loc: munit.Location): EvalError =
    FormulaParser.parse(formula) match
      case Right(expr) =>
        Evaluator.instance.eval(expr, sheet) match
          case Left(e) => e
          case Right(v) => fail(s"$formula: expected a Left, got $v")
      case Left(e) => fail(s"$formula: $e")

  // ===== (e) one error channel =====

  test("#692(e): an XLOOKUP miss is the typed #N/A through lookupNotFound, call as diagnostic") {
    check(sheet, "=XLOOKUP(25,A1:A5,B1:B5)", err(CellError.NA))
    check(sheet, "=IFNA(XLOOKUP(25,A1:A5,B1:B5),\"miss\")", text("miss"))
    check(sheet, "=XLOOKUP(25,A1:A5,B1:B5,\"nf\")", text("nf"))
    leftOf("=XLOOKUP(25,A1:A5,B1:B5)") match
      case EvalError.ErrorValue(CellError.NA, Some(ctx)) =>
        assertEquals(
          ctx,
          "XLOOKUP: no match found for lookup value: XLOOKUP(25, A1:A5, B1:B5, 0, 1)"
        )
      case other => fail(s"expected ErrorValue(NA, ctx), got $other")
    leftOf("=XLOOKUP(\"z\",A1:A5,B1:B5,,-1,-2)") match
      case EvalError.ErrorValue(CellError.NA, Some(ctx)) =>
        assert(ctx.endsWith("XLOOKUP(z, A1:A5, B1:B5, -1, -2)"), ctx)
      case other => fail(s"expected ErrorValue(NA, ctx), got $other")
  }

  test("#692(e): every lookup's miss travels the same channel as NA()") {
    val misses = List(
      "VLOOKUP(25,A1:B5,2,FALSE)",
      "HLOOKUP(25,A1:A5,1,FALSE)",
      "MATCH(25,A1:A5,0)",
      "LOOKUP(5,A1:A5,B1:B5)",
      "XMATCH(25,A1:A5)",
      "XLOOKUP(25,A1:A5,B1:B5)"
    )
    misses.foreach { miss =>
      leftOf(s"=$miss") match
        case EvalError.ErrorValue(CellError.NA, Some(_)) => ()
        case other => fail(s"$miss: expected ErrorValue(NA, ctx), got $other")
      check(sheet, s"=ERROR.TYPE($miss)", num(7))
      check(sheet, s"=ISNA($miss)", CellValue.Bool(true))
    }
    // the lift still answers per element: two misses are two #N/A, not one host failure
    check(sheet, "=SUMPRODUCT(--ISNA(XLOOKUP(W1:W2,A1:A5,B1:B5)))", num(2))
    check(sheet, "=SUMPRODUCT(--ISNA(XLOOKUP(G1:G2,A1:A5,B1:B5)))", num(2))
  }

  // ===== XLOOKUP and XMATCH modes =====

  test("#692: XLOOKUP and XMATCH answer #VALUE! for a mode outside its table") {
    List("3", "-2", "99").foreach { mode =>
      check(sheet, s"=XLOOKUP(30,A1:A5,B1:B5,,$mode)", err(CellError.Value))
      check(sheet, s"=XMATCH(30,A1:A5,$mode)", err(CellError.Value))
    }
    List("0", "3", "-3").foreach { mode =>
      check(sheet, s"=XLOOKUP(30,A1:A5,B1:B5,,0,$mode)", err(CellError.Value))
      check(sheet, s"=XMATCH(30,A1:A5,0,$mode)", err(CellError.Value))
    }
    // if_not_found answers a miss, not a bad mode; the diagnostic names the mode
    check(sheet, "=XLOOKUP(30,A1:A5,B1:B5,\"nf\",3)", err(CellError.Value))
    check(sheet, "=ERROR.TYPE(XLOOKUP(30,A1:A5,B1:B5,,3))", num(3))
    check(sheet, "=ERROR.TYPE(XMATCH(30,A1:A5,3))", num(3))
    leftOf("=XLOOKUP(30,A1:A5,B1:B5,,3)") match
      case EvalError.ErrorValue(CellError.Value, Some(ctx)) =>
        assertEquals(
          ctx,
          "XLOOKUP: match_mode 3 is not -1, 0, 1 or 2: XLOOKUP(30, A1:A5, B1:B5, 3, 1)"
        )
      case other => fail(s"expected ErrorValue(Value, ctx), got $other")
    leftOf("=XMATCH(30,A1:A5,0,0)") match
      case EvalError.ErrorValue(CellError.Value, Some(ctx)) =>
        assertEquals(ctx, "XMATCH: search_mode 0 is not 1, -1, 2 or -2: XMATCH(30, A1:A5, 0, 0)")
      case other => fail(s"expected ErrorValue(Value, ctx), got $other")
    // the modes in the tables still answer
    check(sheet, "=XLOOKUP(30,A1:A5,B1:B5,,2,-2)", text("c"))
    check(sheet, "=XLOOKUP(25,A1:A5,B1:B5,,1,2)", text("c"))
    check(sheet, "=XMATCH(25,A1:A5,-1,-1)", num(2))
  }

  // ===== the evaluating reader =====

  test("#692: MATCH and XLOOKUP find an uncached formula cell in the lookup range") {
    // A2 is a formula with no cached value: the reader must compute 5*4 before comparing
    val keys = Sheet(SheetName.unsafe("S"))
      .put(ref"A1", num(10))
      .put(ref"A2", CellValue.Formula("5*4", None))
      .put(ref"A3", num(30))
      .put(ref"B1", text("a"))
      .put(ref"B2", CellValue.Formula("\"b\"&\"2\"", None))
      .put(ref"B3", text("c"))
    check(keys, "=MATCH(20,A1:A3,0)", num(2))
    check(keys, "=MATCH(25,A1:A3,1)", num(2))
    check(keys, "=XLOOKUP(20,A1:A3,B1:B3)", text("b2"))
    check(keys, "=XLOOKUP(25,A1:A3,B1:B3,,-1)", text("b2"))
    check(keys, "=XLOOKUP(20,A1:A3,B1:B3,,0,-1)", text("b2"))
    // the readers agree with LOOKUP and XMATCH, which already evaluated
    check(keys, "=XMATCH(20,A1:A3)", num(2))
    check(keys, "=LOOKUP(20,A1:A3,B1:B3)", text("b2"))
  }

  test("#692: MATCH keeps RANGE positions over a range wider than the used range") {
    // A1:B2 is the used range; C1 is blank, so 3 (A2) sits at row-major position 4 of A1:C2
    val grid = Sheet(SheetName.unsafe("S"))
      .put(ref"A1", num(1))
      .put(ref"B1", num(2))
      .put(ref"A2", num(3))
      .put(ref"B2", num(4))
    check(grid, "=MATCH(3,A1:C2,0)", num(4))
    check(grid, "=MATCH(4,A1:C2,0)", num(5))
    check(grid, "=MATCH(3,A1:A9,0)", num(2))
    check(grid, "=XLOOKUP(3,A1:A9,B1:B9)", num(4))
    check(grid, "=XLOOKUP(9,A1:A9,B1:B9)", err(CellError.NA))
  }
