package com.tjclp.xl.formula

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.eval.{ArrayArithmetic, ArrayResult, EvalError, Evaluator}
import com.tjclp.xl.formula.parser.FormulaParser
import com.tjclp.xl.sheets.Sheet
import munit.FunSuite

/** #735: preserve public callback execution and SUMPRODUCT's raw-range first error. */
class Review735RegressionSpec extends FunSuite:
  test("#735: public binary callback output is stored rather than replayed") {
    var calls = 0
    val axis = ArrayResult.sparse(
      3,
      1,
      0,
      0,
      Vector.empty,
      ArrayResult.Fill.ByRow(Vector(1, 2, 3).map(n => CellValue.Number(BigDecimal(n))))
    )
    val result = ArrayArithmetic.broadcast(
      ArrayArithmetic.ArrayOperand.Array(axis),
      ArrayArithmetic.ArrayOperand.Scalar(BigDecimal(1)),
      (_, _) => { calls += 1; Right(BigDecimal(calls)) }
    )
    result match
      case Right(ArrayArithmetic.ArrayOperand.Array(out)) =>
        val first = out(0, 0)
        val afterFirst = calls
        val second = out(0, 0)
        println(
          s"callback reproduction: first=$first second=$second calls=$calls afterFirst=$afterFirst"
        )
        assertEquals(second, first)
        assertEquals(calls, afterFirst)
      case other => fail(s"expected array: $other")
  }

  test("#735: raw SUMPRODUCT ranges preserve row-major first error beside positions") {
    val sheet = Sheet(SheetName.unsafe("Errors"))
      .put(ARef.from0(0, 2), CellValue.Error(CellError.Div0))
      .put(ARef.from0(1, 0), CellValue.Error(CellError.NA))
    val expr = FormulaParser
      .parse("=SUMPRODUCT(A1:A3,B1:B3,ROW(A1:A3))")
      .fold(e => fail(e.toString), identity)
    val result = Evaluator.arrayInstance.eval(expr, sheet, Clock.system, None, None)
    println(s"first-error reproduction: $result")
    result match
      case Left(EvalError.ErrorValue(code, _)) => assertEquals(code, CellError.NA)
      case other => fail(s"expected NA failure: $other")
  }

  test("#735: public ordered-comparison callback is not replayed on reads") {
    var calls = 0
    val axis = ArrayResult.positions(1, 3, true)
    val result = ArrayArithmetic.broadcastOrderedCompare(
      axis,
      ArrayResult.single(CellValue.Number(BigDecimal(0))),
      _ => {
        calls += 1
        calls % 2 == 0
      }
    )
    result match
      case Right(out) =>
        val first = out(0, 0)
        val afterFirst = calls
        assertEquals(out(0, 0), first)
        assertEquals(calls, afterFirst)
      case Left(err) => fail(s"expected comparison: $err")
  }
