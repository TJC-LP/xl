package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.eval.{EvalError, Evaluator}
import com.tjclp.xl.formula.parser.FormulaParser
import munit.FunSuite

/**
 * GH-715: MATCH, VLOOKUP, HLOOKUP and XLOOKUP read their keys the way a reference reads its cell,
 * lazily and in search order. An uncached formula key evaluates; a key whose evaluation fails is a
 * non-match, so a broken or circular cell after the match changes nothing; a whole column costs the
 * data, not a million reads.
 */
class LookupUncachedKeysSpec extends FunSuite:
  private def num(n: Int): CellValue = CellValue.Number(BigDecimal(n))

  /** A1:A3 = 1, `=3+4` uncached, 9; B1:B3 = 10, 20, 30. */
  private val column = Sheet("T")
    .put(ref"A1", num(1))
    .put(ref"A2", CellValue.Formula("=3+4"))
    .put(ref"A3", num(9))
    .put(ref"B1", num(10))
    .put(ref"B2", num(20))
    .put(ref"B3", num(30))

  /** The transpose: A1:C1 keys, A2:C2 results. */
  private val row = Sheet("T")
    .put(ref"A1", num(1))
    .put(ref"B1", CellValue.Formula("=3+4"))
    .put(ref"C1", num(9))
    .put(ref"A2", num(10))
    .put(ref"B2", num(20))
    .put(ref"C2", num(30))

  private def eval(sheet: Sheet, formula: String): CellValue =
    sheet.evaluateFormula(formula).fold(e => fail(s"$formula: $e"), identity)

  /** `formula` as the plain cell Z50 evaluates it, its failure the evaluator's own EvalError. */
  private def evalRaw(
    sheet: Sheet,
    formula: String,
    workbook: Option[Workbook] = None
  ): Either[EvalError, Any] =
    FormulaParser
      .parse(formula)
      .fold(
        error => fail(s"$formula: $error"),
        expr => Evaluator.instance.eval(expr, sheet, Clock.system, workbook, Some(ref"Z50"))
      )

  test("an uncached formula key evaluates and matches, with no placement involved") {
    assertEquals(eval(column, "=VLOOKUP(7,A1:B3,2,FALSE)"), num(20))
    assertEquals(eval(row, "=HLOOKUP(7,A1:C2,2,FALSE)"), num(20))
    assertEquals(eval(column, "=MATCH(7,A1:A3,0)"), num(2))
    assertEquals(eval(column, "=XLOOKUP(7,A1:A3,B1:B3)"), num(20))
    assertEquals(eval(column, "=XLOOKUP(7,A1:A3,B1:B3,,0,-1)"), num(20))
  }

  // A4 fails to evaluate: an unknown function, or a formula that reads itself (no positioned
  // evaluation, so it is the recursion guard's failure, not a CircularRef).
  List("=NOSUCHFN(1)", "=A4+1").foreach { broken =>
    val withBroken = column.put(ref"A4", CellValue.Formula(broken)).put(ref"B4", num(40))

    test(s"a key after the match that fails to evaluate ($broken) changes nothing") {
      assertEquals(eval(withBroken, "=VLOOKUP(7,A1:B4,2,FALSE)"), num(20))
      assertEquals(eval(withBroken, "=MATCH(7,A1:A4,0)"), num(2))
      assertEquals(eval(withBroken, "=XLOOKUP(7,A1:A4,B1:B4)"), num(20))
      assertEquals(
        eval(
          row.put(ref"D1", CellValue.Formula(broken.replace("A4", "D1"))),
          "=HLOOKUP(7,A1:D2,2,FALSE)"
        ),
        num(20)
      )
    }

    test(s"a failing key ($broken) is a non-match, so a miss is #N/A, not a failure") {
      assertEquals(eval(withBroken, "=IFNA(MATCH(99,A1:A4,0),-1)"), num(-1))
      assertEquals(eval(withBroken, "=IFNA(VLOOKUP(99,A1:B4,2,FALSE),-1)"), num(-1))
      assertEquals(eval(withBroken, "=XLOOKUP(99,A1:A4,B1:B4,-1)"), num(-1))
    }
  }

  /** A1:A100 = 1..100 ascending, B = twice that. */
  private val tall = (0 until 100).foldLeft(Sheet("T")) { (s, i) =>
    s.put(ARef.from0(0, i), num(i + 1)).put(ARef.from0(1, i), num(2 * (i + 1)))
  }

  test("a whole column searches the used range and still answers every mode") {
    assertEquals(eval(tall, "=MATCH(7,A:A,0)"), num(7))
    assertEquals(eval(tall, "=MATCH(7.5,A:A,1)"), num(7))
    assertEquals(eval(tall, "=XLOOKUP(42,A:A,B:B)"), num(84))
    assertEquals(eval(tall, "=XLOOKUP(42,A:A,B:B,,0,-1)"), num(84))
    assertEquals(eval(tall, "=XLOOKUP(41.5,A:A,B:B,,1)"), num(84))
    assertEquals(eval(tall, "=VLOOKUP(42,A:B,2,FALSE)"), num(84))
    assertEquals(eval(tall, "=VLOOKUP(42.5,A:B,2,TRUE)"), num(84))
    assertEquals(eval(tall, "=IFNA(MATCH(500,A:A,0),-1)"), num(-1))
    assertEquals(eval(tall, "=XLOOKUP(500,A:A,B:B)"), CellValue.Error(CellError.NA))
  }

  test("a result cell computing an Excel error shows that error, as the cell would") {
    // Excel: a lookup returning a cell that holds #DIV/0! shows #DIV/0!
    assertEquals(
      eval(column.put(ref"B2", CellValue.Formula("=1/0")), "=VLOOKUP(7,A1:B3,2,FALSE)"),
      CellValue.Error(CellError.Div0)
    )
    assertEquals(
      eval(row.put(ref"B2", CellValue.Formula("=1/0")), "=HLOOKUP(7,A1:C2,2,FALSE)"),
      CellValue.Error(CellError.Div0)
    )
  }

  test("a result cell failing with no Excel error of its own fails the lookup, as a reference") {
    // An unknown function is a host failure, never laundered into #VALUE!: the lookup fails the
    // way the plain reference =B2 does, so an evaluator failure there is never a spreadsheet error
    val broken = "=NOSUCHFN(1)"
    List(
      column.put(ref"B2", CellValue.Formula(broken)) -> "=VLOOKUP(7,A1:B3,2,FALSE)",
      row.put(ref"B2", CellValue.Formula(broken)) -> "=HLOOKUP(7,A1:C2,2,FALSE)"
    ).foreach { (sheet, formula) =>
      val viaLookup = sheet.evaluateFormula(formula)
      assert(viaLookup.isLeft, s"$formula: $viaLookup")
      assert(sheet.evaluateFormula("=B2").isLeft, "the plain reference fails too")
    }
  }

  test("a result cell in a cycle is #VALUE!, as a key in a cycle is a non-match") {
    // No placement guard: the cycle trips the recursion guard, an EvalFailed the lookup shows as
    // #VALUE! rather than failing (Excel flags the cycle and shows the cycle's value)
    val cyclicColumn = column
      .put(ref"B2", CellValue.Formula("=C2"))
      .put(ref"C2", CellValue.Formula("=B2"))
    assertEquals(eval(cyclicColumn, "=VLOOKUP(7,A1:B3,2,FALSE)"), CellValue.Error(CellError.Value))
    val cyclicRow = row
      .put(ref"B2", CellValue.Formula("=B3"))
      .put(ref"B3", CellValue.Formula("=B2"))
    assertEquals(eval(cyclicRow, "=HLOOKUP(7,A1:C2,2,FALSE)"), CellValue.Error(CellError.Value))
  }

  test("XLOOKUP searching backward over a whole column for an empty text key") {
    // A50 holds the empty string; the blanks past the used range never match a lookup value, so
    // the backward search starts at the used range's end and finds A50
    val withEmpty = tall.put(ref"A50", CellValue.Text("")).put(ref"A70", CellValue.Text(""))
    assertEquals(eval(withEmpty, "=XLOOKUP(\"\",A:A,B:B,,0,-1)"), num(140))
    assertEquals(eval(withEmpty, "=XLOOKUP(\"\",A:A,B:B,,0,1)"), num(100))
    assertEquals(eval(tall, "=XLOOKUP(\"\",A:A,B:B,-1,0,-1)"), num(-1))
  }

  /** A1 and A2 read each other, uncached; A3 = 7; B3 = 70. */
  private val cyclic = Sheet("T")
    .put(ref"A1", CellValue.Formula("=A2"))
    .put(ref"A2", CellValue.Formula("=A1"))
    .put(ref"A3", num(7))
    .put(ref"B3", num(70))

  test("a lookup over a key range holding a cycle reads past the cyclic keys, as Excel does") {
    // Excel flags the cycle and computes the lookup over the cycle's value (0); here each cyclic
    // key trips the recursion guard, a non-match, and the lookup finds A3 the same
    assertEquals(eval(cyclic, "=MATCH(7,A1:A3,0)"), num(3))
    assertEquals(eval(cyclic, "=VLOOKUP(7,A1:B3,2,FALSE)"), num(70))
    assertEquals(eval(cyclic, "=XLOOKUP(7,A1:A3,B1:B3)"), num(70))
  }

  test("a key reading a defined-name cycle, with no placement, is a non-match, not a CircularRef") {
    // Pins the invariant LookupLine relies on: outside `eval --at` no cell read raises a
    // CircularRef, so only the placement guard can make a lookup fail that way
    val names = Vector(
      com.tjclp.xl.workbooks.DefinedName("ping", "=pong"),
      com.tjclp.xl.workbooks.DefinedName("pong", "=ping")
    )
    val sheet = column.put(ref"A1", CellValue.Formula("=ping"))
    val book = Workbook(sheet)
    val named = book.copy(metadata = book.metadata.copy(definedNames = names))
    def evalIn(formula: String): Either[String, CellValue] =
      SheetEvaluator
        .evaluateFormula(sheet)(formula, workbook = Some(named))
        .left
        .map(_.toString)
    List("=MATCH(7,A1:A3,0)" -> num(2), "=VLOOKUP(7,A1:B3,2,FALSE)" -> num(20)).foreach {
      (formula, expected) =>
        val result = evalIn(formula)
        assert(!result.left.exists(_.contains("Circular reference")), s"$formula: $result")
        assertEquals(result, Right(expected), formula)
    }
  }

  test("recalculation leaves a lookup that reads a cycle as the cycle's dependent") {
    // Pinned: recalculation finds the A1-A2 cycle in the dependency graph first and reports every
    // formula reading it, the lookup included, before any lookup runs
    val book = Workbook(cyclic.put(ref"C1", CellValue.Formula("=MATCH(7,A1:A3,0)")))
    val result = book.recalculate()
    assert(result.errors.exists(_.ref == ref"C1"), result.errors.toString)
    assert(result.errors.exists(_.ref == ref"A1"), result.errors.toString)
  }

  test("keys in a cell cycle, with no placement, are non-matches: never a CircularRef") {
    // Pins the invariant LookupLine relies on for cells, as the defined-name test does for names:
    // outside `eval --at` reading a key raises no CircularRef, so a cyclic key cannot fail the
    // lookup. A1 and A2 read each other; every lookup below misses or reads past them.
    val formulas = List(
      "=MATCH(7,A1:A3,0)" -> num(3),
      "=MATCH(99,A1:A2,0)" -> CellValue.Error(CellError.NA),
      "=VLOOKUP(7,A1:B3,2,FALSE)" -> num(70),
      "=VLOOKUP(99,A1:B2,2,FALSE)" -> CellValue.Error(CellError.NA),
      "=XLOOKUP(99,A1:A2,B1:B2)" -> CellValue.Error(CellError.NA)
    )
    formulas.foreach { (formula, expected) =>
      val raw = evalRaw(cyclic, formula)
      assert(
        !raw.left.exists {
          case _: EvalError.CircularRef => true
          case _ => false
        },
        s"$formula: $raw"
      )
      assertEquals(eval(cyclic, formula), expected, formula)
    }
  }

  test("each way a cycle fails without a placement is recognised as a cycle failure") {
    // Evaluator.isCycleFailure matches the messages of three raise sites; a rewording that loses
    // the match would turn a cyclic result cell's #VALUE! into a hard failure, so pin each site
    def failure(result: Either[EvalError, Any], site: String): EvalError =
      result.fold(identity, value => fail(s"$site: expected a failure, got $value"))
    // cells reading each other: the cross-sheet recursion depth guard
    val cells = failure(evalRaw(cyclic, "=A1"), "cells")
    // spill anchors reading each other's spill: the spill recursion depth guard
    val spills = Sheet("T")
      .put(ref"A1", CellValue.Formula("=B1#"))
      .put(ref"B1", CellValue.Formula("=A1#"))
    val spill = failure(evalRaw(spills, "=A1#"), "spill")
    // defined names referring to each other
    val names = Vector(
      com.tjclp.xl.workbooks.DefinedName("ping", "=pong"),
      com.tjclp.xl.workbooks.DefinedName("pong", "=ping")
    )
    val book = Workbook(column)
    val named = book.copy(metadata = book.metadata.copy(definedNames = names))
    val name = failure(evalRaw(column, "=ping", Some(named)), "names")
    List("cells" -> cells, "spill" -> spill, "names" -> name).foreach { (site, error) =>
      assert(Evaluator.isCycleFailure(error), s"$site: $error")
    } // a cell holding an unknown function fails too, but not as a cycle
    val broken = column.put(ref"A9", CellValue.Formula("=NOSUCHFN(1)"))
    val unknown = failure(evalRaw(broken, "=A9"), "unknown function")
    assert(!Evaluator.isCycleFailure(unknown), unknown.toString)
  }

  test("XLOOKUP evaluates an uncached formula in its return cell, as VLOOKUP does (#742)") {
    // XLOOKUP returns a reference (GH-713), whose cells read as a reference reads them, so an
    // uncached `=3+4` in the return cell is 7, as VLOOKUP's result cell is
    val book = column.put(ref"B2", CellValue.Formula("=3+4"))
    assertEquals(eval(book, "=VLOOKUP(7,A1:B3,2,FALSE)"), num(7))
    assertEquals(eval(book, "=XLOOKUP(7,A1:A3,B1:B3)"), num(7))
  }
