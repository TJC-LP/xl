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
  val ifFn: FunctionSpec[Any] { type Args = IfArgs } =
    FunctionSpec.simple[Any, IfArgs]("IF", Arity.three) { (args, ctx) =>
      val (condExpr, ifTrueExpr, ifFalseExpr) = args
      // GH-333: evaluate the condition array-aware. A range-shaped condition (the classic
      // MIN(IF(A1:A10>0,…)) CSE idiom) yields an ArrayResult that must broadcast elementwise —
      // collapsing it to a scalar (the old ctx.evalExpr path) delivered a CellValue.Bool into a
      // Boolean position and crashed with a ClassCastException.
      evalMaybeArrayArg(ctx, condExpr).flatMap {
        case condArr: ArrayResult =>
          // CSE semantics: both branches evaluate (they broadcast elementwise against the
          // condition), so branch laziness only applies to the scalar path below.
          for
            ifTrueVal <- evalMaybeArrayArg(ctx, ifTrueExpr)
            ifFalseVal <- evalMaybeArrayArg(ctx, ifFalseExpr)
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
      def loop(pairs: List[TExpr[Any]]): Either[EvalError, Any] =
        pairs match
          case cond :: value :: rest =>
            evalMaybeArrayArg(ctx, cond).flatMap {
              case condArr: ArrayResult =>
                for
                  valueVal <- evalMaybeArrayArg(ctx, value)
                  restVal <- loop(rest)
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
      args match
        case target :: rest =>
          evalAny(ctx, target).flatMap { targetVal =>
            val tcv = ArrayArithmetic.anyToCellValue(targetVal)
            @annotation.tailrec
            def loop(pairs: List[TExpr[Any]]): Either[EvalError, Any] =
              pairs match
                case caseExpr :: value :: rest2 =>
                  evalAny(ctx, caseExpr) match
                    case Left(err) => Left(err)
                    case Right(cv) =>
                      if ArrayArithmetic.cellValueEquals(tcv, ArrayArithmetic.anyToCellValue(cv))
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
