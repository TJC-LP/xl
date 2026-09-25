package com.tjclp.xl.formula

import munit.FunSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{CellRange, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.ast.{BindingCoercion, ExprValue, TExpr}
import com.tjclp.xl.formula.eval.{
  ArrayArithmetic,
  ArrayResult,
  EvalError,
  Evaluator,
  ScalarCoercion
}
import com.tjclp.xl.sheets.Sheet

/**
 * #681 (second-round review of #679): a runtime Double that is NaN or infinite has no Excel value.
 * Where one reaches a text position (`&`, a text-typed argument) or the value plane (ExprValue, the
 * CellValue boundary) it is `#NUM!` — never the literal text "NaN"/"Infinity", and never the
 * NumberFormatException `BigDecimal(Double)` throws on it. Only a programmatic AST can carry such a
 * value (the parser yields BigDecimal literals), so the cases build one.
 */
class NonFiniteValueSpec extends FunSuite:

  private val sheet: Sheet = Sheet(SheetName.unsafe("S"))
    .put(ref"A1", CellValue.Text("a"))
    .put(ref"A2", CellValue.Text("b"))

  private val nonFinite = List(Double.NaN, Double.PositiveInfinity, Double.NegativeInfinity)

  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  private def textLit(d: Double): TExpr[String] = TExpr.Lit[Any](d).asInstanceOf[TExpr[String]]

  private def isNum(result: Either[EvalError, Any]): Boolean = result match
    case Left(EvalError.ErrorValue(CellError.Num, _)) => true
    case _ => false

  test("#681: `&` with a non-finite scalar operand is #NUM!, on either side") {
    nonFinite.foreach { d =>
      val left = Evaluator.instance.eval(TExpr.Concat(textLit(d), TExpr.Lit("x")), sheet)
      val right = Evaluator.instance.eval(TExpr.Concat(TExpr.Lit("x"), textLit(d)), sheet)
      assert(isNum(left), s"$d & \"x\": $left")
      assert(isNum(right), s"\"x\" & $d: $right")
    }
  }

  test("#681: `&` broadcasting a non-finite scalar over a range yields #NUM! elements") {
    val range = TExpr.asStringExpr(TExpr.RangeRef(CellRange(ref"A1", ref"A2")))
    nonFinite.foreach { d =>
      val result: Either[EvalError, Any] =
        Evaluator.arrayInstance.eval(TExpr.Concat(range, textLit(d)), sheet)
      result match
        case Right(ar: ArrayResult) =>
          assertEquals(
            ar,
            ArrayResult(
              Vector(Vector(CellValue.Error(CellError.Num)), Vector(CellValue.Error(CellError.Num)))
            ),
            d.toString
          )
        case other => fail(s"$d: expected an array of #NUM!, got $other")
    }
  }

  test("#681: a non-finite value in a text-typed position is #NUM! (ScalarCoercion)") {
    nonFinite.foreach { d =>
      assert(isNum(ScalarCoercion.coerce("text argument", d, BindingCoercion.Text)), d.toString)
      val coerced = TExpr.Coerced[String](TExpr.Lit[Any](d), BindingCoercion.Text)
      assert(isNum(Evaluator.instance.eval(coerced, sheet)), s"Coerced $d")
    }
  }

  test("#681: ExprValue.from and anyToCellValue are total over non-finite doubles") {
    nonFinite.foreach { d =>
      assertEquals(ExprValue.from(d), ExprValue.Cell(CellValue.Error(CellError.Num)), d.toString)
      assertEquals(ArrayArithmetic.anyToCellValue(d), CellValue.Error(CellError.Num), d.toString)
    }
    assertEquals(ExprValue.from(2.5), ExprValue.Number(BigDecimal(2.5)))
    assertEquals(ArrayArithmetic.anyToCellValue(2.5), CellValue.Number(BigDecimal(2.5)))
  }

  test("#681: a finite Double still renders as General text") {
    assertEquals(
      Evaluator.instance.eval(TExpr.Concat(textLit(2.50), TExpr.Lit("")), sheet),
      Right("2.5")
    )
  }
