package com.tjclp.xl.formula

import com.tjclp.xl.formula.TExpr.*

trait TExprReferenceOps:
  /**
   * Create ROW expression.
   *
   * Example: TExpr.row(TExpr.PolyRef(ref"A5", Anchor.Relative))
   */
  def row(ref: TExpr[?]): TExpr[BigDecimal] =
    Call(FunctionSpecs.row, ref.asInstanceOf[TExpr[Any]])

  /**
   * Create COLUMN expression.
   *
   * Example: TExpr.column(TExpr.PolyRef(ref"C1", Anchor.Relative))
   */
  def column(ref: TExpr[?]): TExpr[BigDecimal] =
    Call(FunctionSpecs.column, ref.asInstanceOf[TExpr[Any]])

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
