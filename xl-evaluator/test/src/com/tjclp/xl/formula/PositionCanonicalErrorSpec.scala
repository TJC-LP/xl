package com.tjclp.xl.formula

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.eval.{EvalError, Evaluator}
import com.tjclp.xl.formula.parser.FormulaParser
import com.tjclp.xl.sheets.Sheet
import munit.FunSuite

/** #735: canonicalizing a position expression to a constant must preserve error order. */
class PositionCanonicalErrorSpec extends FunSuite:
  test("#735: constant position expressions retain raw-range first error") {
    val sheet = Sheet(SheetName.unsafe("Errors"))
      .put(ARef.from0(0, 2), CellValue.Error(CellError.Div0))
      .put(ARef.from0(1, 0), CellValue.Error(CellError.NA))
    List("ROW(A1:A3)^0", "ROW(A1:A3)*0+1").foreach { positions =>
      val expr = FormulaParser
        .parse(s"=SUMPRODUCT(A1:A3,B1:B3,$positions)")
        .fold(e => fail(e.toString), identity)
      val result = Evaluator.arrayInstance.eval(expr, sheet, Clock.system, None, None)
      println(s"constant-origin reproduction $positions: $result")
      result match
        case Left(EvalError.ErrorValue(code, _)) => assertEquals(code, CellError.NA)
        case other => fail(s"expected NA failure: $other")
    }
  }
