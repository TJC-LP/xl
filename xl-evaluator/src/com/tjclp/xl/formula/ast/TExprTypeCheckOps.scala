package com.tjclp.xl.formula.ast

import com.tjclp.xl.formula.functions.FunctionSpecs
import com.tjclp.xl.formula.eval.EvalError
import com.tjclp.xl.formula.functions.EvalContext

import com.tjclp.xl.cells.CellValue
import TExpr.*

trait TExprTypeCheckOps:
  /**
   * IFERROR: return value_if_error if value results in error.
   *
   * Example: TExpr.iferror(TExpr.Div(...), TExpr.Lit(CellValue.Number(0)))
   */
  def iferror(value: TExpr[CellValue], valueIfError: TExpr[CellValue]): TExpr[CellValue] =
    Call(FunctionSpecs.iferror, (value, valueIfError))

  /**
   * ISERROR: check if expression results in error.
   *
   * Example: TExpr.iserror(TExpr.Div(...))
   */
  def iserror(value: TExpr[CellValue]): TExpr[Boolean] =
    Call(FunctionSpecs.iserror, value)

  /**
   * ISERR: check if expression results in error (excluding #N/A).
   *
   * Example: TExpr.iserr(TExpr.Div(...))
   */
  def iserr(value: TExpr[CellValue]): TExpr[Boolean] =
    Call(FunctionSpecs.iserr, value)

  /**
   * ISNUMBER: check if value is numeric.
   *
   * Example: TExpr.isnumber(TExpr.ref(ARef("A1")))
   */
  def isnumber(value: TExpr[CellValue]): TExpr[Boolean] =
    Call(FunctionSpecs.isnumber, value)

  /**
   * ISTEXT: check if value is text.
   *
   * Example: TExpr.istext(TExpr.ref(ARef("A1")))
   */
  def istext(value: TExpr[CellValue]): TExpr[Boolean] =
    Call(FunctionSpecs.istext, value)

  /**
   * ISBLANK: check if cell is empty.
   *
   * Example: TExpr.isblank(TExpr.ref(ARef("A1")))
   */
  def isblank(value: TExpr[CellValue]): TExpr[Boolean] =
    Call(FunctionSpecs.isblank, value)
