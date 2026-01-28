package com.tjclp.xl.formula.ast

import com.tjclp.xl.formula.functions.FunctionSpecs
import com.tjclp.xl.formula.eval.EvalError
import com.tjclp.xl.formula.functions.EvalContext

import TExpr.*

@SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
trait TExprReferenceOps:
  /**
   * Create ROW expression with a cell reference argument.
   *
   * Example: TExpr.row(TExpr.PolyRef(ref"A5", Anchor.Relative))
   */
  def row(ref: TExpr[?]): TExpr[BigDecimal] =
    Call(FunctionSpecs.row, Some(ref.asInstanceOf[TExpr[Any]]))

  /**
   * Create ROW expression with no arguments (returns row of current cell).
   *
   * Example: TExpr.row() in cell B5 returns 5
   */
  def row(): TExpr[BigDecimal] =
    Call(FunctionSpecs.row, None)

  /**
   * Create COLUMN expression with a cell reference argument.
   *
   * Example: TExpr.column(TExpr.PolyRef(ref"C1", Anchor.Relative))
   */
  def column(ref: TExpr[?]): TExpr[BigDecimal] =
    Call(FunctionSpecs.column, Some(ref.asInstanceOf[TExpr[Any]]))

  /**
   * Create COLUMN expression with no arguments (returns column of current cell).
   *
   * Example: TExpr.column() in cell C5 returns 3
   */
  def column(): TExpr[BigDecimal] =
    Call(FunctionSpecs.column, None)

  /**
   * Create ROWS expression.
   *
   * Example: TExpr.rows(TExpr.RangeRef(range))
   */
  def rows(range: TExpr[?]): TExpr[BigDecimal] =
    Call(FunctionSpecs.rows, range.asInstanceOf[TExpr[Any]])

  /**
   * Create COLUMNS expression.
   *
   * Example: TExpr.columns(TExpr.RangeRef(range))
   */
  def columns(range: TExpr[?]): TExpr[BigDecimal] =
    Call(FunctionSpecs.columns, range.asInstanceOf[TExpr[Any]])

  /**
   * Create ADDRESS expression.
   *
   * Example: TExpr.address(TExpr.Lit(1), TExpr.Lit(1), TExpr.Lit(1), TExpr.Lit(true), None)
   */
  def address(
    row: TExpr[BigDecimal],
    col: TExpr[BigDecimal],
    absNum: TExpr[BigDecimal] = Lit(BigDecimal(1)),
    a1Style: TExpr[Boolean] = Lit(true),
    sheetName: Option[TExpr[String]] = None
  ): TExpr[String] =
    Call(FunctionSpecs.address, (row, col, Some(absNum), Some(a1Style), sheetName))
