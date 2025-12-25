package com.tjclp.xl.formula.ast

import com.tjclp.xl.formula.functions.FunctionSpecs
import com.tjclp.xl.formula.eval.EvalError
import com.tjclp.xl.formula.functions.EvalContext

import com.tjclp.xl.CellRange
import TExpr.*

trait TExprFinancialOps:
  /**
   * Smart constructor for NPV over a range of cash flows.
   *
   * Example: TExpr.npv(TExpr.Lit(BigDecimal("0.1")), CellRange("A2:A6"))
   */
  def npv(rate: TExpr[BigDecimal], values: CellRange): TExpr[BigDecimal] =
    Call(FunctionSpecs.npv, (rate, values))

  /**
   * Smart constructor for IRR with optional guess.
   *
   * Example: TExpr.irr(CellRange("A1:A6"), Some(TExpr.Lit(BigDecimal("0.15"))))
   */
  def irr(values: CellRange, guess: Option[TExpr[BigDecimal]] = None): TExpr[BigDecimal] =
    Call(FunctionSpecs.irr, (values, guess))

  /**
   * Smart constructor for XNPV with irregular dates.
   *
   * @param rate
   *   Discount rate
   * @param values
   *   Cash flow values range
   * @param dates
   *   Corresponding dates range
   *
   * Example: TExpr.xnpv(TExpr.Lit(0.1), valuesRange, datesRange)
   */
  def xnpv(
    rate: TExpr[BigDecimal],
    values: CellRange,
    dates: CellRange
  ): TExpr[BigDecimal] =
    Call(FunctionSpecs.xnpv, (rate, values, dates))

  /**
   * Smart constructor for XIRR with irregular dates.
   *
   * @param values
   *   Cash flow values range (must have positive and negative)
   * @param dates
   *   Corresponding dates range
   * @param guess
   *   Optional starting rate for Newton-Raphson (default 0.1)
   *
   * Example: TExpr.xirr(valuesRange, datesRange, Some(TExpr.Lit(0.15)))
   */
  def xirr(
    values: CellRange,
    dates: CellRange,
    guess: Option[TExpr[BigDecimal]] = None
  ): TExpr[BigDecimal] =
    Call(FunctionSpecs.xirr, (values, dates, guess))

  // ===== TVM Smart Constructors =====

  /**
   * PMT: calculate payment per period.
   *
   * Example: TExpr.pmt(TExpr.Lit(0.05/12), TExpr.Lit(24), TExpr.Lit(10000))
   */
  def pmt(
    rate: TExpr[BigDecimal],
    nper: TExpr[BigDecimal],
    pv: TExpr[BigDecimal],
    fv: Option[TExpr[BigDecimal]] = None,
    pmtType: Option[TExpr[BigDecimal]] = None
  ): TExpr[BigDecimal] =
    Call(FunctionSpecs.pmt, (rate, nper, pv, fv, pmtType))

  /**
   * FV: calculate future value.
   *
   * Example: TExpr.fv(TExpr.Lit(0.05/12), TExpr.Lit(24), TExpr.Lit(-100))
   */
  def fv(
    rate: TExpr[BigDecimal],
    nper: TExpr[BigDecimal],
    pmt: TExpr[BigDecimal],
    pv: Option[TExpr[BigDecimal]] = None,
    pmtType: Option[TExpr[BigDecimal]] = None
  ): TExpr[BigDecimal] =
    Call(FunctionSpecs.fv, (rate, nper, pmt, pv, pmtType))

  /**
   * PV: calculate present value.
   *
   * Example: TExpr.pv(TExpr.Lit(0.05/12), TExpr.Lit(24), TExpr.Lit(-500))
   */
  def pv(
    rate: TExpr[BigDecimal],
    nper: TExpr[BigDecimal],
    pmt: TExpr[BigDecimal],
    fv: Option[TExpr[BigDecimal]] = None,
    pmtType: Option[TExpr[BigDecimal]] = None
  ): TExpr[BigDecimal] =
    Call(FunctionSpecs.pv, (rate, nper, pmt, fv, pmtType))

  /**
   * NPER: calculate number of periods.
   *
   * Example: TExpr.nper(TExpr.Lit(0.05/12), TExpr.Lit(-500), TExpr.Lit(10000))
   */
  def nper(
    rate: TExpr[BigDecimal],
    pmt: TExpr[BigDecimal],
    pv: TExpr[BigDecimal],
    fv: Option[TExpr[BigDecimal]] = None,
    pmtType: Option[TExpr[BigDecimal]] = None
  ): TExpr[BigDecimal] =
    Call(FunctionSpecs.nper, (rate, pmt, pv, fv, pmtType))

  /**
   * RATE: calculate interest rate per period.
   *
   * Example: TExpr.rate(TExpr.Lit(24), TExpr.Lit(-500), TExpr.Lit(10000))
   */
  def rate(
    nper: TExpr[BigDecimal],
    pmt: TExpr[BigDecimal],
    pv: TExpr[BigDecimal],
    fv: Option[TExpr[BigDecimal]] = None,
    pmtType: Option[TExpr[BigDecimal]] = None,
    guess: Option[TExpr[BigDecimal]] = None
  ): TExpr[BigDecimal] =
    Call(FunctionSpecs.rate, (nper, pmt, pv, fv, pmtType, guess))
