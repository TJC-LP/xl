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

  /** TRIM whitespace per Excel rules (collapses ASCII space runs; preserves nbsp/tab). */
  def trim(text: TExpr[String]): TExpr[String] =
    Call(FunctionSpecs.trim, text)

  /** MID(text, start, len) — 1-indexed substring extraction. */
  def mid(text: TExpr[String], start: TExpr[Int], length: TExpr[Int]): TExpr[String] =
    Call(FunctionSpecs.mid, (text, start, length))

  /** FIND(find_text, within_text, [start_num]) — case-sensitive 1-indexed search. */
  def find(
    findText: TExpr[String],
    withinText: TExpr[String],
    start: Option[TExpr[Int]] = None
  ): TExpr[BigDecimal] =
    Call(FunctionSpecs.find, (findText, withinText, start))

  /** SUBSTITUTE(text, old, new, [instance]) — match-by-content text replacement. */
  def substitute(
    text: TExpr[String],
    oldText: TExpr[String],
    newText: TExpr[String],
    instance: Option[TExpr[Int]] = None
  ): TExpr[String] =
    Call(FunctionSpecs.substitute, (text, oldText, newText, instance))

  /** VALUE(text) — parse text as numeric (handles currency, percent, whitespace). */
  def value(text: TExpr[String]): TExpr[BigDecimal] =
    Call(FunctionSpecs.value, text)

  /** TEXT(value, format) — format numeric/date/text using Excel format codes. */
  def text(v: TExpr[Any], format: TExpr[String]): TExpr[String] =
    Call(FunctionSpecs.text, (v, format))
