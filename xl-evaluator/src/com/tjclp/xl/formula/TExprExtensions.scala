package com.tjclp.xl.formula

import com.tjclp.xl.formula.TExpr.*

trait TExprExtensions:
  // Extension methods for ergonomic formula construction

  extension (x: TExpr[BigDecimal])
    /**
     * Arithmetic operators as infix methods.
     *
     * Example: expr1 + expr2
     */
    def +(y: TExpr[BigDecimal]): TExpr[BigDecimal] = Add(x, y)
    def -(y: TExpr[BigDecimal]): TExpr[BigDecimal] = Sub(x, y)
    def *(y: TExpr[BigDecimal]): TExpr[BigDecimal] = Mul(x, y)
    def /(y: TExpr[BigDecimal]): TExpr[BigDecimal] = Div(x, y)

    /**
     * Comparison operators (return Boolean).
     *
     * Example: expr1 < expr2
     */
    def <(y: TExpr[BigDecimal]): TExpr[Boolean] = Lt(x, y)
    def <=(y: TExpr[BigDecimal]): TExpr[Boolean] = Lte(x, y)
    def >(y: TExpr[BigDecimal]): TExpr[Boolean] = Gt(x, y)
    def >=(y: TExpr[BigDecimal]): TExpr[Boolean] = Gte(x, y)

  extension (x: TExpr[Boolean])
    /**
     * Boolean operators as infix methods.
     *
     * Example: expr1 && expr2
     */
    def &&(y: TExpr[Boolean]): TExpr[Boolean] = Call(FunctionSpecs.and, List(x, y))
    def ||(y: TExpr[Boolean]): TExpr[Boolean] = Call(FunctionSpecs.or, List(x, y))
    def unary_! : TExpr[Boolean] = Call(FunctionSpecs.not, x)

  extension [A](x: TExpr[A])
    /**
     * Equality/inequality operators.
     *
     * Example: expr1 === expr2
     */
    def ===(y: TExpr[A]): TExpr[Boolean] = Eq(x, y)
    def !==(y: TExpr[A]): TExpr[Boolean] = Neq(x, y)
