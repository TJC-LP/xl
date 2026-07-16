package com.tjclp.xl.formula.functions

import com.tjclp.xl.formula.ast.{TExpr, ExprValue, BindingCoercion}
import com.tjclp.xl.formula.eval.{
  EvalError,
  Evaluator,
  ArrayArithmetic,
  ArrayResult,
  ScalarCoercion
}
import com.tjclp.xl.formula.parser.ParseError
import com.tjclp.xl.formula.{Clock, Arity}
import com.tjclp.xl.cells.{CellError, CellValue}

trait FunctionSpecsConditional extends FunctionSpecsBase:

  /**
   * GH-339: demote a whole-result evaluation failure to a 1x1 carried-error element (via the
   * [[EvalError.toCellError]] table) so an error from an entirely-unselected IF/IFS branch no
   * longer poisons the formula — broadcastIf selects it away positionally, and it only surfaces
   * where selected (as an error element, or loudly via a consuming aggregate). Failures with no
   * element form (CircularRef) stay a fatal Left.
   */
  private def carryLeft(result: Either[EvalError, Any]): Either[EvalError, Any] =
    result match
      case Left(e) =>
        EvalError
          .toCellError(e)
          .map(err => ArrayResult.single(CellValue.Error(err)): Any)
          .toRight(e)
      case right => right

  /** GH-339: evaluate an IF/IFS branch expression, carrying a failure as a 1x1 error element. */
  private def evalBranchCarrying(ctx: EvalContext, expr: TExpr[?]): Either[EvalError, Any] =
    carryLeft(evalMaybeArrayArg(ctx, expr))

  val ifFn: FunctionSpec[Any] { type Args = IfArgs } =
    FunctionSpec.simple[Any, IfArgs]("IF", Arity.three) { (args, ctx) =>
      val (condExpr, ifTrueExpr, ifFalseExpr) = args
      // GH-333: evaluate the condition array-aware. A range-shaped condition (the classic
      // MIN(IF(A1:A10>0,…)) CSE idiom) yields an ArrayResult that must broadcast elementwise —
      // collapsing it to a scalar (the old ctx.evalExpr path) delivered a CellValue.Bool into a
      // Boolean position and crashed with a ClassCastException.
      // A failing CONDITION stays fatal (it is not a branch); GH-339 carriage below applies to
      // the two branch expressions only.
      evalMaybeArrayArg(ctx, condExpr).flatMap {
        case condArr: ArrayResult =>
          // CSE semantics: both branches evaluate (they broadcast elementwise against the
          // condition), so branch laziness only applies to the scalar path below. GH-339: a
          // branch that fails wholesale carries as a 1x1 error element instead of aborting.
          for
            ifTrueVal <- evalBranchCarrying(ctx, ifTrueExpr)
            ifFalseVal <- evalBranchCarrying(ctx, ifFalseExpr)
            result <- ArrayArithmetic.broadcastIf(
              condArr,
              toCellArray(ifTrueVal),
              toCellArray(ifFalseVal)
            )
          yield result
        case scalarCond =>
          // Scalar condition: coerce totally (Excel truthiness) and keep lazy branch selection.
          ScalarCoercion.coerce("IF condition", scalarCond, BindingCoercion.Bool).flatMap {
            case cond: Boolean =>
              if cond then evalAny(ctx, ifTrueExpr) else evalAny(ctx, ifFalseExpr)
            case other =>
              Left(EvalError.TypeMismatch("IF condition", "boolean", other.toString))
          }
      }
    }

  // GH-76 (tier 1): variadic conditional / selection functions. Args are a flat List[TExpr[Any]].

  /** IFS(cond1, val1, cond2, val2, ...) — first TRUE condition's value, else #N/A. */
  val ifs: FunctionSpec[Any] { type Args = List[TExpr[Any]] } =
    FunctionSpec.simple[Any, List[TExpr[Any]]]("IFS", Arity.AtLeast(2)) { (args, ctx) =>
      // GH-338: condition slots evaluate array-aware like IF (GH-333). A scalar condition keeps
      // the lazy pair walk; an array condition switches to CSE semantics — its value and the
      // remaining pairs all evaluate, chaining elementwise through broadcastIf with the no-match
      // #N/A as the final fallback: IFS(c1,v1,c2,v2) = if c1 then v1 else (if c2 then v2 else NA).
      // GH-339: under an array condition, the pair's VALUE and the remaining-pairs fallback are
      // branches — their failures carry as 1x1 error elements. The CURRENT pair's condition
      // failure stays fatal (it is not a branch).
      def loop(pairs: List[TExpr[Any]]): Either[EvalError, Any] =
        pairs match
          case cond :: value :: rest =>
            evalMaybeArrayArg(ctx, cond).flatMap {
              case condArr: ArrayResult =>
                for
                  valueVal <- evalBranchCarrying(ctx, value)
                  restVal <- carryLeft(loop(rest))
                  result <- ArrayArithmetic.broadcastIf(
                    condArr,
                    toCellArray(valueVal),
                    toCellArray(restVal)
                  )
                yield result
              case scalarCond =>
                ScalarCoercion.coerce("IFS condition", scalarCond, BindingCoercion.Bool).flatMap {
                  case condBool: Boolean =>
                    if condBool then evalAny(ctx, value) else loop(rest)
                  case other =>
                    Left(EvalError.TypeMismatch("IFS condition", "boolean", other.toString))
                }
            }
          case _ => Right(CellValue.Error(CellError.NA))
      loop(args)
    }

  /**
   * SWITCH(expr, case1, val1, ..., [default]) — value for the first matching case, else
   * default/#N/A.
   */
  val switchFn: FunctionSpec[Any] { type Args = List[TExpr[Any]] } =
    FunctionSpec.simple[Any, List[TExpr[Any]]]("SWITCH", Arity.AtLeast(3)) { (args, ctx) =>
      // GH-344: an error VALUE in the target or an evaluated case propagates with its code
      // (`=SWITCH(#N/A-cell, …)` is #N/A, not a silent never-match) — errors are never "equal",
      // so without the pre-check the walk would fall through to the wrong #N/A.
      def carried(value: Any): Option[CellError] =
        ArrayArithmetic.carriedError(ArrayArithmetic.anyToCellValue(value))
      args match
        case target :: rest =>
          evalAny(ctx, target).flatMap { targetVal =>
            carried(targetVal) match
              case Some(err) => Left(EvalError.ErrorValue(err, Some("SWITCH")))
              case None =>
                val tcv = ArrayArithmetic.anyToCellValue(targetVal)
                @annotation.tailrec
                def loop(pairs: List[TExpr[Any]]): Either[EvalError, Any] =
                  pairs match
                    case caseExpr :: value :: rest2 =>
                      evalAny(ctx, caseExpr) match
                        case Left(err) => Left(err)
                        case Right(cv) =>
                          carried(cv) match
                            case Some(err) => Left(EvalError.ErrorValue(err, Some("SWITCH")))
                            case None =>
                              if ArrayArithmetic
                                  .cellValueEquals(tcv, ArrayArithmetic.anyToCellValue(cv))
                              then evalAny(ctx, value)
                              else loop(rest2)
                    case default :: Nil => evalAny(ctx, default) // trailing default
                    case _ => Right(CellValue.Error(CellError.NA))
                loop(rest)
          }
        case _ => Right(CellValue.Error(CellError.NA))
    }

  /** CHOOSE(index, val1, val2, ...) — 1-based selection; out-of-range → #VALUE!. */
  val choose: FunctionSpec[Any] { type Args = List[TExpr[Any]] } =
    FunctionSpec.simple[Any, List[TExpr[Any]]]("CHOOSE", Arity.AtLeast(2)) { (args, ctx) =>
      args match
        case idxExpr :: values =>
          ctx.evalExpr(TExpr.asNumericExpr(idxExpr)).flatMap { n =>
            values.lift(n.toInt - 1) match
              case Some(v) => evalAny(ctx, v)
              case None => Right(CellValue.Error(CellError.Value))
          }
        case _ => Right(CellValue.Error(CellError.Value))
    }
