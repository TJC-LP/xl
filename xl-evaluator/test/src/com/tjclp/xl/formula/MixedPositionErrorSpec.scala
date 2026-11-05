package com.tjclp.xl.formula

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.eval.{EvalError, Evaluator}
import com.tjclp.xl.formula.parser.FormulaParser
import com.tjclp.xl.sheets.Sheet
import munit.FunSuite

class MixedPositionErrorSpec extends FunSuite:
  test("#735: ROW plus an existing sparse operand keeps the observed base error") {
    val sheet = Sheet(SheetName.unsafe("Errors"))
      .put(ARef.from0(0, 2), CellValue.Error(CellError.Div0))
      .put(ARef.from0(1, 0), CellValue.Error(CellError.NA))
    val expr = FormulaParser
      .parse("=SUMPRODUCT(A1:A3,B1:B3,ROW(A1:A3)+C1:C3)")
      .fold(e => fail(e.toString), identity)
    val result = Evaluator.arrayInstance.eval(expr, sheet, Clock.system, None, None)
    println(s"mixed-origin reproduction: $result")
    result match
      case Left(EvalError.ErrorValue(code, _)) => assertEquals(code, CellError.NA)
      case other => fail(s"expected NA failure: $other")
  }
