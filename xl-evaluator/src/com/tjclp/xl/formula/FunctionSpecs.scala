package com.tjclp.xl.formula

object FunctionSpecs:
  private val numericExpr = ArgSpec.expr[BigDecimal]

  val abs: FunctionSpec[BigDecimal] =
    FunctionSpec.simple[BigDecimal, TExpr[BigDecimal]]("ABS", Arity.one) { (expr, ctx) =>
      ctx.evalExpr(TExpr.Abs(expr))
    }(using numericExpr)

  val sqrt: FunctionSpec[BigDecimal] =
    FunctionSpec.simple[BigDecimal, TExpr[BigDecimal]]("SQRT", Arity.one) { (expr, ctx) =>
      ctx.evalExpr(TExpr.Sqrt(expr))
    }(using numericExpr)
