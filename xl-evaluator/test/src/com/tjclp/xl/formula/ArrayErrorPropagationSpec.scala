package com.tjclp.xl.formula

import munit.ScalaCheckSuite

import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.codec.CodecError
import com.tjclp.xl.formula.eval.EvalError
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.syntax.*

/**
 * GH-337: errors propagate ELEMENTWISE through array operations instead of failing the whole
 * formula.
 *
 * Excel semantics: `(A1:A3<5)*1` with a #DIV/0! cell at A2 yields the array {1, #DIV/0!, 1} — the
 * error only surfaces when an aggregation consumes that element. Pre-GH-337 the comparison refused
 * with a whole-formula Left ("cannot compare #REF! error value"), which was total and deterministic
 * but not Excel.
 *
 * The demotion table lives in [[EvalError.toCellError]]; the carriage lives in the ArrayArithmetic
 * broadcasts (compare, arithmetic, IF). CircularRef stays a fatal Left.
 */
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial", "org.wartremover.warts.IterableOps"))
class ArrayErrorPropagationSpec extends ScalaCheckSuite:

  // ===== Commit 1: the EvalError -> CellError demotion table =====

  test("GH-337: toCellError maps each EvalError to its carried Excel error") {
    val a1 = ARef.from0(0, 0)
    val table: List[(EvalError, Option[CellError])] = List(
      EvalError.DivByZero("10", "0") -> Some(CellError.Div0),
      EvalError.RefError(a1, "cell not found") -> Some(CellError.Ref),
      EvalError.TypeMismatch("arithmetic", "number", "text: x") -> Some(CellError.Value),
      EvalError.CodecFailed(a1, CodecError.TypeMismatch("Numeric", CellValue.Text("x"))) -> Some(
        CellError.Value
      ),
      EvalError.EvalFailed("anything", None) -> Some(CellError.Value),
      // CircularRef is a structural failure of the sheet, never an element-local one: it must
      // stay a fatal Left rather than demote to an error element.
      EvalError.CircularRef(List(a1)) -> None
    )
    table.foreach { case (evalError, expected) =>
      assertEquals(EvalError.toCellError(evalError), expected, s"toCellError($evalError)")
    }
  }
