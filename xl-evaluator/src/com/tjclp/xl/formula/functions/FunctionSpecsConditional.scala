package com.tjclp.xl.formula.functions

import com.tjclp.xl.formula.ast.{TExpr, ExprValue}
import com.tjclp.xl.formula.eval.{EvalError, Evaluator, ArrayArithmetic}
import com.tjclp.xl.formula.parser.ParseError
import com.tjclp.xl.formula.{Clock, Arity}
import com.tjclp.xl.cells.{CellError, CellValue}

trait FunctionSpecsConditional extends FunctionSpecsBase:
  val ifFn: FunctionSpec[Any] { type Args = IfArgs } =
    FunctionSpec.simple[Any, IfArgs]("IF", Arity.three) { (args, ctx) =>
      val (condExpr, ifTrueExpr, ifFalseExpr) = args
      for
        cond <- ctx.evalExpr(condExpr)
        result <- if cond then evalAny(ctx, ifTrueExpr) else evalAny(ctx, ifFalseExpr)
      yield result
    }

  // GH-76 (tier 1): variadic conditional / selection functions. Args are a flat List[TExpr[Any]].

  /** IFS(cond1, val1, cond2, val2, ...) — first TRUE condition's value, else #N/A. */
  val ifs: FunctionSpec[Any] { type Args = List[TExpr[Any]] } =
    FunctionSpec.simple[Any, List[TExpr[Any]]]("IFS", Arity.AtLeast(2)) { (args, ctx) =>
      @annotation.tailrec
      def loop(pairs: List[TExpr[Any]]): Either[EvalError, Any] =
        pairs match
          case cond :: value :: rest =>
            ctx.evalExpr(TExpr.asBooleanExpr(cond)) match
              case Left(err) => Left(err)
              case Right(true) => evalAny(ctx, value)
              case Right(false) => loop(rest)
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
