package com.tjclp.xl.formula.ast

import com.tjclp.xl.formula.functions.FunctionSpecs
import com.tjclp.xl.formula.eval.EvalError
import com.tjclp.xl.formula.functions.EvalContext

import com.tjclp.xl.CellRange
import TExpr.*

trait TExprAggregateOps:
  /**
   * SUM aggregation: sum all numeric values in range.
   *
   * Example: TExpr.sum(CellRange("A1:A10"))
   */
  def sum(range: CellRange): TExpr[BigDecimal] =
    Call(FunctionSpecs.sum, List(Left(RangeLocation.Local(range))))

  /**
   * SUM aggregation: sum variadic numeric values.
   *
   * Example: TExpr.sum(TExpr.Lit(1), TExpr.Lit(2), TExpr.Lit(3))
   */
  def sum(values: TExpr[BigDecimal]*): TExpr[BigDecimal] =
    Call(FunctionSpecs.sum, values.map(Right(_)).toList)

  /**
   * COUNT aggregation: count numeric cells in range.
   *
   * Example: TExpr.count(CellRange("A1:A10"))
   */
  def count(range: CellRange): TExpr[BigDecimal] =
    Call(FunctionSpecs.count, List(Left(RangeLocation.Local(range))))

  /**
   * AVERAGE aggregation: average of numeric values in range.
   *
   * Example: TExpr.average(CellRange("A1:A10"))
   */
  def average(range: CellRange): TExpr[BigDecimal] =
    Call(FunctionSpecs.average, List(Left(RangeLocation.Local(range))))

  /**
   * AVERAGE aggregation: average variadic numeric values.
   *
   * Example: TExpr.average(TExpr.Lit(1), TExpr.Lit(2), TExpr.Lit(3))
   */
  def average(values: TExpr[BigDecimal]*): TExpr[BigDecimal] =
    Call(FunctionSpecs.average, values.map(Right(_)).toList)

  /**
   * MIN aggregation: minimum numeric value in range.
   *
   * Example: TExpr.min(CellRange("A1:A10"))
   */
  def min(range: CellRange): TExpr[BigDecimal] =
    Call(FunctionSpecs.min, List(Left(RangeLocation.Local(range))))

  /**
   * MIN aggregation: minimum of variadic numeric values.
   *
   * Example: TExpr.min(TExpr.Lit(1), TExpr.Lit(2), TExpr.Lit(3))
   */
  def min(values: TExpr[BigDecimal]*): TExpr[BigDecimal] =
    Call(FunctionSpecs.min, values.map(Right(_)).toList)

  /**
   * MAX aggregation: maximum numeric value in range.
   *
   * Example: TExpr.max(CellRange("A1:A10"))
   */
  def max(range: CellRange): TExpr[BigDecimal] =
    Call(FunctionSpecs.max, List(Left(RangeLocation.Local(range))))

  /**
   * MAX aggregation: maximum of variadic numeric values.
   *
   * Example: TExpr.max(TExpr.Lit(1), TExpr.Lit(2), TExpr.Lit(3))
   */
  def max(values: TExpr[BigDecimal]*): TExpr[BigDecimal] =
    Call(FunctionSpecs.max, values.map(Right(_)).toList)

  // Conditional aggregation function smart constructors

  /**
   * SUMIF: sum cells where criteria matches.
   *
   * Example: TExpr.sumIf(CellRange("A1:A10"), TExpr.Lit("Apple"), Some(CellRange("B1:B10")))
   */
  def sumIf(
    range: CellRange,
    criteria: TExpr[?],
    sumRange: Option[CellRange] = None
  ): TExpr[BigDecimal] =
    Call(
      FunctionSpecs.sumif,
      (range, criteria.asInstanceOf[TExpr[Any]], sumRange)
    )

  /**
   * COUNTIF: count cells where criteria matches.
   *
   * Example: TExpr.countIf(CellRange("A1:A10"), TExpr.Lit(">100"))
   */
  def countIf(range: CellRange, criteria: TExpr[?]): TExpr[BigDecimal] =
    Call(FunctionSpecs.countif, (range, criteria.asInstanceOf[TExpr[Any]]))

  /**
   * SUMIFS: sum with multiple criteria (AND logic).
   *
   * Example: TExpr.sumIfs(CellRange("C1:C10"), List((CellRange("A1:A10"), TExpr.Lit("Apple"))))
   */
  def sumIfs(
    sumRange: CellRange,
    conditions: List[(CellRange, TExpr[?])]
  ): TExpr[BigDecimal] =
    Call(
      FunctionSpecs.sumifs,
      (sumRange, conditions.map { case (r, c) => (r, c.asInstanceOf[TExpr[Any]]) })
    )

  /**
   * COUNTIFS: count with multiple criteria (AND logic).
   *
   * Example: TExpr.countIfs(List((CellRange("A1:A10"), TExpr.Lit("Apple"))))
   */
  def countIfs(conditions: List[(CellRange, TExpr[?])]): TExpr[BigDecimal] =
    Call(
      FunctionSpecs.countifs,
      conditions.map { case (r, c) => (r, c.asInstanceOf[TExpr[Any]]) }
    )

  /**
   * AVERAGEIF: average cells where criteria matches.
   *
   * Example: TExpr.averageIf(CellRange("A1:A10"), TExpr.Lit("Apple"), Some(CellRange("B1:B10")))
   */
  def averageIf(
    range: CellRange,
    criteria: TExpr[?],
    averageRange: Option[CellRange] = None
  ): TExpr[BigDecimal] =
    Call(
      FunctionSpecs.averageif,
      (range, criteria.asInstanceOf[TExpr[Any]], averageRange)
    )

  /**
   * AVERAGEIFS: average with multiple criteria (AND logic).
   *
   * Example: TExpr.averageIfs(CellRange("C1:C10"), List((CellRange("A1:A10"), TExpr.Lit("Apple"))))
   */
  def averageIfs(
    averageRange: CellRange,
    conditions: List[(CellRange, TExpr[?])]
  ): TExpr[BigDecimal] =
    Call(
      FunctionSpecs.averageifs,
      (averageRange, conditions.map { case (r, c) => (r, c.asInstanceOf[TExpr[Any]]) })
    )

  // Array and advanced lookup function smart constructors

  /**
   * SUMPRODUCT: multiply corresponding elements and sum.
   *
   * Example: TExpr.sumProduct(List(CellRange.parse("A1:A3").toOption.get,
   * CellRange.parse("B1:B3").toOption.get))
   */
  def sumProduct(arrays: List[CellRange]): TExpr[BigDecimal] =
    Call(FunctionSpecs.sumproduct, arrays)
