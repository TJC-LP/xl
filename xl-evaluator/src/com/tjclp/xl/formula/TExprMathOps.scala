package com.tjclp.xl.formula

import com.tjclp.xl.formula.TExpr.*

trait TExprMathOps:
  /**
   * ROUND: round number to specified digits.
   *
   * Example: TExpr.round(TExpr.Lit(2.5), TExpr.Lit(0))
   */
  def round(value: TExpr[BigDecimal], numDigits: TExpr[BigDecimal]): TExpr[BigDecimal] =
    Call(FunctionSpecs.round, (value, numDigits))

  /**
   * ROUNDUP: round away from zero.
   *
   * Example: TExpr.roundUp(TExpr.Lit(2.1), TExpr.Lit(0))
   */
  def roundUp(value: TExpr[BigDecimal], numDigits: TExpr[BigDecimal]): TExpr[BigDecimal] =
    Call(FunctionSpecs.roundUp, (value, numDigits))

  /**
   * ROUNDDOWN: round toward zero (truncate).
   *
   * Example: TExpr.roundDown(TExpr.Lit(2.9), TExpr.Lit(0))
   */
  def roundDown(value: TExpr[BigDecimal], numDigits: TExpr[BigDecimal]): TExpr[BigDecimal] =
    Call(FunctionSpecs.roundDown, (value, numDigits))

  /**
   * ABS: absolute value.
   *
   * Example: TExpr.abs(TExpr.Lit(-5))
   */
  def abs(value: TExpr[BigDecimal]): TExpr[BigDecimal] =
    Call(FunctionSpecs.abs, value)

  /**
   * SQRT: square root.
   *
   * Example: TExpr.sqrt(TExpr.Lit(16))
   */
  def sqrt(value: TExpr[BigDecimal]): TExpr[BigDecimal] =
    Call(FunctionSpecs.sqrt, value)

  /**
   * MOD: modulo (remainder after division).
   *
   * Example: TExpr.mod(TExpr.Lit(5), TExpr.Lit(3))
   */
  def mod(number: TExpr[BigDecimal], divisor: TExpr[BigDecimal]): TExpr[BigDecimal] =
    Call(FunctionSpecs.mod, (number, divisor))

  /**
   * POWER: number raised to a power.
   *
   * Example: TExpr.power(TExpr.Lit(2), TExpr.Lit(3))
   */
  def power(number: TExpr[BigDecimal], power: TExpr[BigDecimal]): TExpr[BigDecimal] =
    Call(FunctionSpecs.power, (number, power))

  /**
   * LOG: logarithm to specified base.
   *
   * Example: TExpr.log(TExpr.Lit(100), TExpr.Lit(10))
   */
  def log(number: TExpr[BigDecimal], base: TExpr[BigDecimal]): TExpr[BigDecimal] =
    Call(FunctionSpecs.log, (number, Some(base)))

  /**
   * LN: natural logarithm (base e).
   *
   * Example: TExpr.ln(TExpr.Lit(2.718281828))
   */
  def ln(value: TExpr[BigDecimal]): TExpr[BigDecimal] =
    Call(FunctionSpecs.ln, value)

  /**
   * EXP: e raised to a power.
   *
   * Example: TExpr.exp(TExpr.Lit(1))
   */
  def exp(value: TExpr[BigDecimal]): TExpr[BigDecimal] =
    Call(FunctionSpecs.exp, value)

  /**
   * FLOOR: round down to nearest multiple of significance.
   *
   * Example: TExpr.floor(TExpr.Lit(2.5), TExpr.Lit(1))
   */
  def floor(number: TExpr[BigDecimal], significance: TExpr[BigDecimal]): TExpr[BigDecimal] =
    Call(FunctionSpecs.floor, (number, significance))

  /**
   * CEILING: round up to nearest multiple of significance.
   *
   * Example: TExpr.ceiling(TExpr.Lit(2.5), TExpr.Lit(1))
   */
  def ceiling(number: TExpr[BigDecimal], significance: TExpr[BigDecimal]): TExpr[BigDecimal] =
    Call(FunctionSpecs.ceiling, (number, significance))

  /**
   * TRUNC: truncate to specified number of decimal places.
   *
   * Example: TExpr.trunc(TExpr.Lit(8.9), TExpr.Lit(0))
   */
  def trunc(number: TExpr[BigDecimal], numDigits: TExpr[BigDecimal]): TExpr[BigDecimal] =
    Call(FunctionSpecs.trunc, (number, Some(numDigits)))

  /**
   * SIGN: sign of a number (1, -1, or 0).
   *
   * Example: TExpr.sign(TExpr.Lit(-5))
   */
  def sign(value: TExpr[BigDecimal]): TExpr[BigDecimal] =
    Call(FunctionSpecs.sign, value)

  /**
   * INT: round down to nearest integer (floor).
   *
   * Example: TExpr.int_(TExpr.Lit(8.9))
   */
  def int_(value: TExpr[BigDecimal]): TExpr[BigDecimal] =
    Call(FunctionSpecs.int, value)

  /**
   * PI mathematical constant.
   *
   * Example: TExpr.pi()
   */
  def pi(): TExpr[BigDecimal] = Call(FunctionSpecs.pi, EmptyTuple)
