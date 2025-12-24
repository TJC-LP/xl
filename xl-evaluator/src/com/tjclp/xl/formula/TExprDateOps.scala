package com.tjclp.xl.formula

import com.tjclp.xl.CellRange
import com.tjclp.xl.formula.TExpr.*

trait TExprDateOps:
  /**
   * TODAY current date.
   *
   * Example: TExpr.today()
   */
  def today(): TExpr[java.time.LocalDate] = Call(FunctionSpecs.today, EmptyTuple)

  /**
   * NOW current date and time.
   *
   * Example: TExpr.now()
   */
  def now(): TExpr[java.time.LocalDateTime] = Call(FunctionSpecs.now, EmptyTuple)

  /**
   * DATE construct from year, month, day.
   *
   * Example: TExpr.date(TExpr.Lit(2025), TExpr.Lit(11), TExpr.Lit(21))
   */
  def date(year: TExpr[Int], month: TExpr[Int], day: TExpr[Int]): TExpr[java.time.LocalDate] =
    Call(FunctionSpecs.date, (year, month, day))

  /**
   * YEAR extract year from date.
   *
   * Returns BigDecimal to match Excel semantics and enable arithmetic composition.
   *
   * Example: TExpr.year(TExpr.date(TExpr.Lit(2025), TExpr.Lit(11), TExpr.Lit(21)))
   */
  def year(date: TExpr[java.time.LocalDate]): TExpr[BigDecimal] = Call(FunctionSpecs.year, date)

  /**
   * MONTH extract month from date.
   *
   * Returns BigDecimal to match Excel semantics and enable arithmetic composition.
   *
   * Example: TExpr.month(TExpr.date(TExpr.Lit(2025), TExpr.Lit(11), TExpr.Lit(21)))
   */
  def month(date: TExpr[java.time.LocalDate]): TExpr[BigDecimal] = Call(FunctionSpecs.month, date)

  /**
   * DAY extract day from date.
   *
   * Returns BigDecimal to match Excel semantics and enable arithmetic composition.
   *
   * Example: TExpr.day(TExpr.date(TExpr.Lit(2025), TExpr.Lit(11), TExpr.Lit(21)))
   */
  def day(date: TExpr[java.time.LocalDate]): TExpr[BigDecimal] = Call(FunctionSpecs.day, date)

  /**
   * EOMONTH end of month N months from start.
   *
   * @param startDate
   *   The starting date
   * @param months
   *   Number of months to add (can be negative)
   *
   * Example: TExpr.eomonth(dateExpr, TExpr.Lit(1))
   */
  def eomonth(
    startDate: TExpr[java.time.LocalDate],
    months: TExpr[Int]
  ): TExpr[java.time.LocalDate] =
    Call(FunctionSpecs.eomonth, (startDate, months))

  /**
   * EDATE add months to date.
   *
   * @param startDate
   *   The starting date
   * @param months
   *   Number of months to add (can be negative)
   *
   * Example: TExpr.edate(dateExpr, TExpr.Lit(3))
   */
  def edate(startDate: TExpr[java.time.LocalDate], months: TExpr[Int]): TExpr[java.time.LocalDate] =
    Call(FunctionSpecs.edate, (startDate, months))

  /**
   * DATEDIF difference between dates.
   *
   * @param startDate
   *   The starting date
   * @param endDate
   *   The ending date
   * @param unit
   *   Unit: "Y", "M", "D", "MD", "YM", "YD"
   *
   * Example: TExpr.datedif(start, end, TExpr.Lit("Y"))
   */
  def datedif(
    startDate: TExpr[java.time.LocalDate],
    endDate: TExpr[java.time.LocalDate],
    unit: TExpr[String]
  ): TExpr[BigDecimal] =
    Call(FunctionSpecs.datedif, (startDate, endDate, unit))

  /**
   * NETWORKDAYS count working days between dates.
   *
   * @param startDate
   *   The starting date (inclusive)
   * @param endDate
   *   The ending date (inclusive)
   * @param holidays
   *   Optional range of holiday dates to exclude
   *
   * Example: TExpr.networkdays(start, end, Some(holidayRange))
   */
  def networkdays(
    startDate: TExpr[java.time.LocalDate],
    endDate: TExpr[java.time.LocalDate],
    holidays: Option[CellRange] = None
  ): TExpr[BigDecimal] =
    Call(FunctionSpecs.networkdays, (startDate, endDate, holidays))

  /**
   * WORKDAY add working days to date.
   *
   * @param startDate
   *   The starting date
   * @param days
   *   Number of working days to add (can be negative)
   * @param holidays
   *   Optional range of holiday dates to exclude
   *
   * Example: TExpr.workday(start, TExpr.Lit(5), Some(holidayRange))
   */
  def workday(
    startDate: TExpr[java.time.LocalDate],
    days: TExpr[Int],
    holidays: Option[CellRange] = None
  ): TExpr[java.time.LocalDate] =
    Call(FunctionSpecs.workday, (startDate, days, holidays))

  /**
   * YEARFRAC year fraction between dates.
   *
   * @param startDate
   *   The starting date
   * @param endDate
   *   The ending date
   * @param basis
   *   Day count basis: 0=US 30/360, 1=Actual/actual, 2=Actual/360, 3=Actual/365, 4=EU 30/360
   *
   * Example: TExpr.yearfrac(start, end, TExpr.Lit(1))
   */
  def yearfrac(
    startDate: TExpr[java.time.LocalDate],
    endDate: TExpr[java.time.LocalDate],
    basis: TExpr[Int] = Lit(0)
  ): TExpr[BigDecimal] =
    Call(FunctionSpecs.yearfrac, (startDate, endDate, Some(basis)))
