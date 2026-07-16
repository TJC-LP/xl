package com.tjclp.xl.formula

import munit.FunSuite

import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.eval.SheetEvaluator.*
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.syntax.*

/**
 * GH-339: array-path IF/IFS evaluate both branches eagerly (CSE semantics require it for the
 * broadcast), so an error produced by an entirely-unselected branch used to poison the whole
 * formula.
 *
 * With GH-337 elementwise error carriage, a whole-branch evaluation failure demotes to a 1x1
 * carried-error element: broadcastIf then selects it away at positions where the other branch wins,
 * and it only surfaces (as an error element, or loudly via an aggregate) where actually selected.
 * Scalar IF/IFS keep lazy branch selection — re-pinned here adjacent to the array semantics.
 *
 * Accepted micro-divergence (by design): a broadcast-dimension mismatch INSIDE a branch expression
 * also demotes to a 1x1 #VALUE! element instead of staying a whole-formula Left.
 */
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial"))
class IfBranchErrorSpec extends FunSuite:

  private val div0 = CellValue.Error(CellError.Div0)

  private def num(n: Int): CellValue = CellValue.Number(BigDecimal(n))

  // The GH-339 headline: every condition element is TRUE, so the erroring FALSE branch is
  // never selected and must not fail the formula.
  private val allPositive = Sheet("Test")
    .put(ref"A1", num(5))
    .put(ref"A2", num(3))

  private val mixed = Sheet("Test")
    .put(ref"A1", num(5))
    .put(ref"A2", num(-3))

  test("GH-339: =SUM(IF(A1:A2>0,A1:A2,1/0)) succeeds when every condition is TRUE") {
    assertEquals(allPositive.evaluateFormula("=SUM(IF(A1:A2>0,A1:A2,1/0))"), Right(num(8)))

    val arrayEntry = allPositive.evaluateArrayFormula("=SUM(IF(A1:A2>0,A1:A2,1/0))", ref"D1")
    assert(arrayEntry.isRight, s"expected Right, got $arrayEntry")
    assertEquals(arrayEntry.toOption.get._1(ref"D1").value, num(8))
  }

  test("GH-339: mixed conditions surface the branch error only at FALSE positions") {
    val result = mixed.evaluateArrayFormula("=IF(A1:A2>0,A1:A2,1/0)", ref"C1")
    assert(result.isRight, s"expected Right, got $result")
    val (updated, spill) = result.toOption.get
    assertEquals(spill.height, 2)
    assertEquals(updated(ref"C1").value, num(5))
    assertEquals(updated(ref"C2").value, div0)
  }

  test("GH-339: =IF(A1:A2>0,1/0,2) with every condition FALSE yields {2;2}") {
    val allNegative = Sheet("Test")
      .put(ref"A1", num(-5))
      .put(ref"A2", num(-3))
    val result = allNegative.evaluateArrayFormula("=IF(A1:A2>0,1/0,2)", ref"C1")
    assert(result.isRight, s"expected Right, got $result")
    val (updated, _) = result.toOption.get
    assertEquals(updated(ref"C1").value, num(2))
    assertEquals(updated(ref"C2").value, num(2))
  }

  test("GH-339: IFS carries an erroring later condition as elements at unmatched positions") {
    // A1=5 matches the first pair; A2=-3 falls through to the second pair, whose condition
    // divides by zero — that failure carries as #DIV/0! only where the fallthrough is selected.
    val result = mixed.evaluateArrayFormula("=IFS(A1:A2>0,A1:A2,1/0>0,99)", ref"C1")
    assert(result.isRight, s"expected Right, got $result")
    val (updated, _) = result.toOption.get
    assertEquals(updated(ref"C1").value, num(5))
    assertEquals(updated(ref"C2").value, div0)
  }

  test("GH-339: IFS carries an erroring matched VALUE as elements at matched positions") {
    val result = mixed.evaluateArrayFormula("=IFS(A1:A2>0,1/0,TRUE,5)", ref"C1")
    assert(result.isRight, s"expected Right, got $result")
    val (updated, _) = result.toOption.get
    assertEquals(updated(ref"C1").value, div0)
    assertEquals(updated(ref"C2").value, num(5))
  }

  test("GH-339: the CURRENT pair's condition failure stays fatal (scalar condition shape)") {
    // A scalar condition that fails to evaluate is not a branch — it aborts the formula.
    // GH-344: the abort surfaces as Excel's error VALUE at the boundary (=IF(1/0>0,…) is
    // #DIV/0! in Excel), still failing the whole formula rather than selecting a branch.
    assertEquals(
      allPositive.evaluateFormula("=SUM(IF(1/0>0,A1:A2,0))"),
      Right(CellValue.Error(CellError.Div0))
    )
    assertEquals(
      allPositive.evaluateFormula("=IFS(1/0>0,1,TRUE,2)"),
      Right(CellValue.Error(CellError.Div0))
    )
  }

  // ===== Scalar laziness re-pins (adjacent to the array semantics they contrast with) =====

  test("GH-339: scalar IF keeps lazy branch selection (untaken 1/0 never evaluates)") {
    assertEquals(allPositive.evaluateFormula("=IF(A1>0,42,1/0)"), Right(num(42)))
    assertEquals(allPositive.evaluateFormula("=IF(A1<0,1/0,42)"), Right(num(42)))
  }

  test("GH-339: scalar IFS keeps lazy pair selection (later erroring pairs never evaluate)") {
    assertEquals(allPositive.evaluateFormula("=IFS(TRUE,42,1/0>0,99)"), Right(num(42)))
    assertEquals(
      allPositive.evaluateFormula("=IFS(1>2,1)"),
      Right(CellValue.Error(CellError.NA))
    )
  }
