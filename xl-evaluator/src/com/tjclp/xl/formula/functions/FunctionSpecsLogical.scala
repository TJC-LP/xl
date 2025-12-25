package com.tjclp.xl.formula.functions

import com.tjclp.xl.formula.ast.{TExpr, ExprValue}
import com.tjclp.xl.formula.eval.{EvalError, Evaluator}
import com.tjclp.xl.formula.parser.ParseError
import com.tjclp.xl.formula.{Clock, Arity}

trait FunctionSpecsLogical extends FunctionSpecsBase:
  val and: FunctionSpec[Boolean] { type Args = BooleanList } =
    FunctionSpec.simple[Boolean, BooleanList]("AND", Arity.atLeastOne) { (args, ctx) =>
      @annotation.tailrec
      def loop(remaining: List[TExpr[Boolean]]): Either[EvalError, Boolean] =
        remaining match
          case Nil => Right(true)
          case head :: tail =>
            ctx.evalExpr(head) match
              case Left(err) => Left(err)
              case Right(value) =>
                if !value then Right(false)
                else loop(tail)
      loop(args)
    }

  val or: FunctionSpec[Boolean] { type Args = BooleanList } =
    FunctionSpec.simple[Boolean, BooleanList]("OR", Arity.atLeastOne) { (args, ctx) =>
      @annotation.tailrec
      def loop(remaining: List[TExpr[Boolean]]): Either[EvalError, Boolean] =
        remaining match
          case Nil => Right(false)
          case head :: tail =>
            ctx.evalExpr(head) match
              case Left(err) => Left(err)
              case Right(value) =>
                if value then Right(true)
                else loop(tail)
      loop(args)
    }

  val not: FunctionSpec[Boolean] { type Args = UnaryBoolean } =
    FunctionSpec.simple[Boolean, UnaryBoolean]("NOT", Arity.one) { (expr, ctx) =>
      ctx.evalExpr(expr).map(value => !value)
    }
