package com.tjclp.xl.formula.eval

import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.macros.ref
import com.tjclp.xl.sheets.Sheet
import munit.FunSuite

/** #829: the per-pass memo's in-progress marker and the two placement guards. */
@SuppressWarnings(Array("org.wartremover.warts.Var"))
class EvalMemoSpec extends FunSuite:

  private val sheet = Sheet(SheetName.unsafe("S"))
  private def num(i: Int): CellValue = CellValue.Number(BigDecimal(i))

  test("a re-entrant read of a cell being computed is the dynamic cycle, before any recursion") {
    val memo = new Evaluator.EvalMemo()
    var inner: Either[EvalError, CellValue] = Right(CellValue.Empty)
    var computed = 0
    val outer = memo.getOrCompute(sheet, ref"A1") {
      computed += 1
      inner = memo.getOrCompute(sheet, ref"A1")(fail("the inner compute must not run"))
      inner
    }
    assertEquals(computed, 1)
    assertEquals(inner, Left(Evaluator.dynamicCycle(sheet.name, ref"A1")))
    assert(inner.left.exists(Evaluator.isDynamicCycle), inner.toString)
    assert(inner.left.exists(Evaluator.isCycleFailure), inner.toString)
    assertEquals(outer, inner)
    // the re-entry failure memoizes as the cell's failure: a later read is served, not recomputed
    assertEquals(memo.getOrCompute(sheet, ref"A1")(fail("must be served from the memo")), outer)
    assertEquals(computed, 1)
  }

  test("a completed cell is served; a compute that throws leaves no marker behind") {
    val memo = new Evaluator.EvalMemo()
    assertEquals(memo.getOrCompute(sheet, ref"B1")(Right(num(2))), Right(num(2)))
    assertEquals(memo.getOrCompute(sheet, ref"B1")(fail("memoized")), Right(num(2)))
    intercept[IllegalStateException](
      memo.getOrCompute(sheet, ref"C1")(throw new IllegalStateException("boom"))
    )
    assertEquals(memo.getOrCompute(sheet, ref"C1")(Right(num(3))), Right(num(3)))
  }

  test("a spill computed again while it is being computed is the dynamic cycle of its anchor") {
    val memo = new Evaluator.EvalMemo()
    val inner = memo.spill(sheet, ref"A1")(memo.spill(sheet, ref"A1")(Right(num(1))))
    assertEquals(inner, Left(Evaluator.dynamicCycle(sheet.name, ref"A1", spill = true)))
    assert(inner.left.exists(Evaluator.isCycleFailure), inner.toString)
    assertEquals(memo.spill(sheet, ref"A1")(Right(num(1))), Right(num(1)))
  }

  test("PlacementGuard.placed raises the CLI's CircularRef; evaluating raises the dynamic cycle") {
    val placed = Evaluator.PlacementGuard.placed(sheet.name, ref"D5")
    assertEquals(placed.check(sheet, ref"D5"), Some(EvalError.CircularRef(List(ref"D5"))))
    assertEquals(placed.check(sheet, ref"D6"), None)
    assert(placed.read)
    assert(placed.positioned, "the CLI's placement reads no slice")
    val evaluating = Evaluator.PlacementGuard.evaluating(sheet.name, ref"D5")
    assert(!evaluating.positioned, "the pass's own cell keeps its slices")
    assertEquals(
      evaluating.check(sheet, ref"D5"),
      Some(Evaluator.dynamicCycle(sheet.name, ref"D5"))
    )
    assertEquals(evaluating.check(Sheet(SheetName.unsafe("Other")), ref"D5"), None)
    assertEquals(
      Evaluator.dynamicCycle(SheetName.unsafe("My Sheet"), ref"A1"),
      EvalError.EvalFailed(
        "Circular reference through 'My Sheet'!A1: the cell is read again while its own evaluation is in progress",
        None
      )
    )
  }
