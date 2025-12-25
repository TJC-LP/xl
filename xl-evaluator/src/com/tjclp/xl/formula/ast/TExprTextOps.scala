package com.tjclp.xl.formula.ast

import com.tjclp.xl.formula.functions.FunctionSpecs
import com.tjclp.xl.formula.eval.EvalError
import com.tjclp.xl.formula.functions.EvalContext

import TExpr.*

trait TExprTextOps:
  /**
   * CONCATENATE text values.
   *
   * Example: TExpr.concatenate(List(TExpr.Lit("Hello"), TExpr.Lit(" "), TExpr.Lit("World")))
   */
  def concatenate(xs: List[TExpr[String]]): TExpr[String] =
    Call(FunctionSpecs.concatenate, xs)

  /**
   * LEFT substring extraction.
   *
   * Example: TExpr.left(TExpr.Lit("Hello"), TExpr.Lit(3))
   */
  def left(text: TExpr[String], n: TExpr[Int]): TExpr[String] =
    Call(FunctionSpecs.left, (text, n))

  /**
   * RIGHT substring extraction.
   *
   * Example: TExpr.right(TExpr.Lit("Hello"), TExpr.Lit(3))
   */
  def right(text: TExpr[String], n: TExpr[Int]): TExpr[String] =
    Call(FunctionSpecs.right, (text, n))

  /**
   * LEN text length.
   *
   * Returns BigDecimal to match Excel semantics.
   *
   * Example: TExpr.len(TExpr.Lit("Hello"))
   */
  def len(text: TExpr[String]): TExpr[BigDecimal] =
    Call(FunctionSpecs.len, text)

  /**
   * UPPER convert to uppercase.
   *
   * Example: TExpr.upper(TExpr.Lit("hello"))
   */
  def upper(text: TExpr[String]): TExpr[String] =
    Call(FunctionSpecs.upper, text)

  /**
   * LOWER convert to lowercase.
   *
   * Example: TExpr.lower(TExpr.Lit("HELLO"))
   */
  def lower(text: TExpr[String]): TExpr[String] =
    Call(FunctionSpecs.lower, text)
