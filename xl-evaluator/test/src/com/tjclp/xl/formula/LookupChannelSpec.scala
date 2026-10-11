package com.tjclp.xl.formula

import munit.FunSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.eval.{EvalError, Evaluator}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook

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

  test("#692: an exact lookup stops at the first hit and never evaluates the cells past it") {
    // A5 is a formula that cannot even parse: reading it is a host failure (a Left), so a search
    // that reached it would fail the whole lookup. The hit at A2 comes first in a forward walk.
    val keys = Sheet(SheetName.unsafe("S"))
      .put(ref"A1", num(10))
      .put(ref"A2", num(20))
      .put(ref"A3", num(30))
      .put(ref"A5", CellValue.Formula("1+", None))
      .put(ref"B1", text("a"))
      .put(ref"B2", text("b"))
      .put(ref"B3", text("c"))
      .put(ref"B5", text("e"))
    check(keys, "=XLOOKUP(20,A1:A5,B1:B5)", text("b"))
    check(keys, "=XLOOKUP(20,A1:A5,B1:B5,,0,2)", text("b"))
    check(keys, "=XLOOKUP(\"b*\",B1:B5,A1:A5,,2)", num(20))
    check(keys, "=MATCH(20,A1:A5,0)", num(2))
    check(keys, "=XMATCH(20,A1:A5)", num(2))
    check(keys, "=XMATCH(\"b*\",B1:B5,2)", num(2))
    // a reverse search reads A5 first, and the first evaluation failure ends the search
    assert(keys.evaluateFormula("=XLOOKUP(30,A1:A5,B1:B5,,0,-1)").isLeft)
    assert(keys.evaluateFormula("=XMATCH(30,A1:A5,0,-1)").isLeft)
    // so does a miss, which has to read every key; and the approximate modes read them all
    assert(keys.evaluateFormula("=XLOOKUP(99,A1:A5,B1:B5)").isLeft)
    assert(keys.evaluateFormula("=MATCH(25,A1:A5,1)").isLeft)
    assert(keys.evaluateFormula("=XLOOKUP(25,A1:A5,B1:B5,,-1)").isLeft)
    // the walk is row-major over a 2-D lookup_array, positions the range's own
    val grid = Sheet(SheetName.unsafe("S"))
      .put(ref"A1", num(1))
      .put(ref"B1", num(2))
      .put(ref"A2", num(3))
      .put(ref"B2", num(4))
      .put(ref"A3", CellValue.Formula("1+", None))
      .put(ref"D2", text("dd"))
    check(grid, "=MATCH(4,A1:B3,0)", num(4))
    // XLOOKUP's keys are one row or column; preserve each orientation's own position.
    check(grid, "=XLOOKUP(4,A2:B2,C2:D2)", text("dd"))
    check(grid, "=XLOOKUP(4,B1:B2,D1:D2)", text("dd"))
    assert(grid.evaluateFormula("=MATCH(9,A1:B3,0)").isLeft)
  }

  test("#692: XLOOKUP over a horizontal pair walks the columns in search_mode's direction") {
    // A1:E1 keys with 20 twice, A2:E2 and A4:E4 returns: a reversed walk (search_mode -1) reads
    // E1 first and lands on the LAST 20; the hit's position maps to return_array by `idx % width`
    val keys = (0 until 5).foldLeft(Sheet(SheetName.unsafe("S"))) { (s, c) =>
      s.put(ref"A1".shift(c, 0), num(Vector(10, 20, 30, 20, 50)(c)))
        .put(ref"A2".shift(c, 0), text(("a".charAt(0) + c).toChar.toString))
        .put(ref"A4".shift(c, 0), num(100 * (c + 1)))
    }
    check(keys, "=XLOOKUP(20,A1:E1,A2:E2)", text("b"))
    check(keys, "=XLOOKUP(20,A1:E1,A2:E2,,0,-1)", text("d"))
    check(keys, "=XLOOKUP(20,A1:E1,A4:E4,,0,-1)", num(400))
    check(keys, "=XLOOKUP(50,A1:E1,A4:E4,,0,-1)", num(500))
    check(keys, "=XLOOKUP(10,A1:E1,A4:E4,,0,-1)", num(100))
    check(keys, "=XLOOKUP(\"?\",A2:E2,A1:E1,,2,-1)", num(50))
    check(keys, "=XMATCH(20,A1:E1,0,-1)", num(4))
    // the reversed column walk stops at its first hit: an unparseable key in A1 is never read
    val broken = keys.put(ref"A1", CellValue.Formula("1+", None))
    check(broken, "=XLOOKUP(20,A1:E1,A2:E2,,0,-1)", text("d"))
    check(broken, "=XLOOKUP(50,A1:E1,A2:E2,,0,-1)", text("e"))
    assert(broken.evaluateFormula("=XLOOKUP(20,A1:E1,A2:E2)").isLeft)
    // a miss walks every column and raises through the one channel
    check(keys, "=XLOOKUP(99,A1:E1,A2:E2,,0,-1)", err(CellError.NA))
  }

  test("#692: XLOOKUP reads an uncached formula cell in return_array through the reader") {
    // B2 is a formula with no cached value: the hit's return cell must be computed, not the
    // raw formula record (the old `unwrapCachedValue` read left an uncached record as it was)
    val keys = Sheet(SheetName.unsafe("S"))
      .put(ref"A1", num(10))
      .put(ref"A2", num(20))
      .put(ref"A3", num(30))
      .put(ref"B1", text("a"))
      .put(ref"B2", CellValue.Formula("\"b\"&\"2\"", None))
      .put(ref"B3", CellValue.Formula("A3*2", None))
    check(keys, "=XLOOKUP(20,A1:A3,B1:B3)", text("b2"))
    check(keys, "=XLOOKUP(30,A1:A3,B1:B3,,0,-1)", num(60))
    check(keys, "=XLOOKUP(25,A1:A3,B1:B3,,-1)", text("b2"))
    check(keys, "=XLOOKUP(25,A1:A3,B1:B3,,1)", num(60))
    check(keys, "=XLOOKUP(\"b*\",B1:B3,A1:A3,,2)", num(20))
    // a return cell that computes an error is that error value, not a host failure
    val erring = keys.put(ref"B2", CellValue.Formula("1/0", None))
    check(erring, "=XLOOKUP(20,A1:A3,B1:B3)", err(CellError.Div0))
    // and on another sheet
    val wb =
      Workbook(Vector(Sheet(SheetName.unsafe("Main")), keys.copy(name = SheetName.unsafe("Keys"))))
    assertEquals(
      wb.sheets.headOption
        .map(_.evaluateFormula("=XLOOKUP(30,Keys!A1:A3,Keys!B1:B3)", workbook = Some(wb))),
      Some(Right(num(60)))
    )
  }

  test("#692: a reverse exact search evaluates the cells below the hit, never the ones above") {
    // A1 and B1 cannot parse: a reverse walk (search_mode -1, and -2 the same walk) reads row 5,
    // 4, 3 and stops at the hit in row 3, so row 1 is never evaluated — while a forward walk
    // reads it first
    val keys = Sheet(SheetName.unsafe("S"))
      .put(ref"A1", CellValue.Formula("1+", None))
      .put(ref"A2", num(20))
      .put(ref"A3", num(30))
      .put(ref"A4", num(40))
      .put(ref"A5", num(50))
      .put(ref"B1", CellValue.Formula("1+", None))
      .put(ref"B2", text("b"))
      .put(ref"B3", text("c"))
      .put(ref"B4", text("d"))
      .put(ref"B5", text("e"))
    List(-1, -2).foreach { mode =>
      check(keys, s"=XLOOKUP(30,A1:A5,B1:B5,,0,$mode)", text("c"))
      check(keys, s"=XLOOKUP(\"c*\",B1:B5,A1:A5,,2,$mode)", num(30))
      check(keys, s"=XMATCH(30,A1:A5,0,$mode)", num(3))
      check(keys, s"=XMATCH(\"c*\",B1:B5,2,$mode)", num(3))
      // a hit in row 2 is reached before row 1 too, so row 1 is still never read
      check(keys, s"=XLOOKUP(20,A1:A5,B1:B5,,0,$mode)", text("b"))
    }
    // the broken cell AFTER the hit in reverse order, i.e. below it, ends the search instead
    val broken = keys
      .put(ref"A1", num(10))
      .put(ref"B1", text("a"))
      .put(ref"A5", CellValue.Formula("1+", None))
      .put(ref"B5", CellValue.Formula("1+", None))
    List(-1, -2).foreach { mode =>
      assert(broken.evaluateFormula(s"=XLOOKUP(30,A1:A5,B1:B5,,0,$mode)").isLeft, mode)
      assert(broken.evaluateFormula(s"=XLOOKUP(\"c*\",B1:B5,A1:A5,,2,$mode)").isLeft, mode)
      assert(broken.evaluateFormula(s"=XMATCH(30,A1:A5,0,$mode)").isLeft, mode)
      assert(broken.evaluateFormula(s"=XMATCH(\"c*\",B1:B5,2,$mode)").isLeft, mode)
    }
    // while the forward walk over the same sheet stops at row 3 and never reads row 5
    check(broken, "=XLOOKUP(30,A1:A5,B1:B5)", text("c"))
    check(broken, "=XMATCH(\"c*\",B1:B5,2)", num(3))
    check(broken, "=MATCH(30,A1:A5,0)", num(3))
    // the forward walk over the same keys reads A1 first and fails there
    assert(keys.evaluateFormula("=XLOOKUP(30,A1:A5,B1:B5)").isLeft)
    assert(keys.evaluateFormula("=XMATCH(30,A1:A5)").isLeft)
    assert(keys.evaluateFormula("=MATCH(30,A1:A5,0)").isLeft)
  }

  test("#692: a forward exact search surfaces a host-failing key before the hit, as XMATCH did") {
    // the old MATCH/XLOOKUP readers skipped a formula they could not evaluate; now it fails the
    // search with its own error, in the exact modes only when it sits before the hit
    def failing(formula: String): Sheet =
      Sheet(SheetName.unsafe("S"))
        .put(ref"A1", num(10))
        .put(ref"A2", CellValue.Formula(formula, None))
        .put(ref"A3", num(30))
        .put(ref"B1", text("a"))
        .put(ref"B2", text("b"))
        .put(ref"B3", text("c"))
    List("1+", "NOSUCHFUNCTION(1)").foreach { formula =>
      val keys = failing(formula)
      assert(keys.evaluateFormula("=MATCH(30,A1:A3,0)").isLeft, formula)
      assert(keys.evaluateFormula("=XLOOKUP(30,A1:A3,B1:B3)").isLeft, formula)
      // the return cell past the broken key is read without evaluating the key column's A2
      check(keys, "=XLOOKUP(\"c\",B1:B3,A1:A3,,2)", num(30))
      assert(keys.evaluateFormula("=XMATCH(30,A1:A3)").isLeft, formula)
      // the hit before the broken cell is still found
      check(keys, "=MATCH(10,A1:A3,0)", num(1))
      check(keys, "=XLOOKUP(10,A1:A3,B1:B3)", text("a"))
      // the approximate modes read the whole used range
      assert(keys.evaluateFormula("=MATCH(10,A1:A3,1)").isLeft, formula)
      assert(keys.evaluateFormula("=XLOOKUP(10,A1:A3,B1:B3,,-1)").isLeft, formula)
    }
    // a key that EVALUATES to an error is a value the search skips past, as Excel does
    val errorKey = failing("1/0")
    check(errorKey, "=MATCH(30,A1:A3,0)", num(3))
    check(errorKey, "=XLOOKUP(30,A1:A3,B1:B3)", text("c"))
  }

  test("#692: MATCH and XLOOKUP compute an uncached formula key on another sheet") {
    val keys = Sheet(SheetName.unsafe("Keys"))
      .put(ref"A1", num(10))
      .put(ref"A2", CellValue.Formula("5*4", None))
      .put(ref"A3", num(30))
      .put(ref"B1", text("a"))
      .put(ref"B2", CellValue.Formula("\"b\"&\"2\"", None))
      .put(ref"B3", text("c"))
    val main = Sheet(SheetName.unsafe("S"))
    val wb = Workbook(Vector(main, keys))
    def across(formula: String, expected: CellValue)(implicit loc: munit.Location): Unit =
      assertEquals(main.evaluateFormula(formula, workbook = Some(wb)), Right(expected), formula)
    across("=MATCH(20,Keys!A1:A3,0)", num(2))
    across("=MATCH(25,Keys!A1:A3,1)", num(2))
    across("=XLOOKUP(20,Keys!A1:A3,Keys!B1:B3)", text("b2"))
    across("=XLOOKUP(20,Keys!A1:A3,Keys!B1:B3,,0,-1)", text("b2"))
    across("=XLOOKUP(25,Keys!A1:A3,Keys!B1:B3,,-1)", text("b2"))
    across("=XMATCH(20,Keys!A1:A3,0,-1)", num(2))
    across("=XMATCH(\"b2\",Keys!B1:B3)", num(2))
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
