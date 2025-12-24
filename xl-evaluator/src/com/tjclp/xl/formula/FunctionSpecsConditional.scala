package com.tjclp.xl.formula

trait FunctionSpecsConditional extends FunctionSpecsBase:
  val ifFn: FunctionSpec[Any] { type Args = IfArgs } =
    FunctionSpec.simple[Any, IfArgs]("IF", Arity.three) { (args, ctx) =>
      val (condExpr, ifTrueExpr, ifFalseExpr) = args
      for
        cond <- ctx.evalExpr(condExpr)
        result <- if cond then evalAny(ctx, ifTrueExpr) else evalAny(ctx, ifFalseExpr)
      yield result
    }
