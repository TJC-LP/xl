package com.tjclp.xl.formula

import com.tjclp.xl.{ARef, Anchor, CellRange, Column, Row, SheetName}
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.codec.CodecError

import scala.math.BigDecimal

/**
 * Typed expression GADT for Excel formulas.
 *
 * TExpr[A] represents a formula expression that evaluates to type A. This is a purely functional,
 * total representation - all operations return Either for error handling, no exceptions thrown.
 *
 * Type parameter A captures the result type:
 *   - TExpr[BigDecimal] for numeric expressions
 *   - TExpr[Boolean] for boolean/logical expressions
 *   - TExpr[String] for text expressions
 *
 * Laws satisfied:
 *   1. If-fusion: If(c, Lit(x), Lit(y)) ≡ Lit(if ⟦c⟧ then x else y)
 *   2. Ring laws: Add/Mul form commutative semiring over BigDecimal
 *   3. Short-circuit: And/Or respect left-to-right semantics
 *   4. Round-trip: parse(print(expr)) == Right(expr)
 */
enum TExpr[A] derives CanEqual:
  /**
   * Literal value - compile-time constant.
   *
   * Example: Lit(42) represents the number 42
   */
  case Lit[A](value: A) extends TExpr[A]

  /**
   * Cell reference - reads value from a cell at runtime.
   *
   * @param at
   *   The cell address to read from
   * @param anchor
   *   Anchoring mode for formula dragging (default: Relative)
   * @param decode
   *   Function to decode the cell's value to type A
   *
   * Example: Ref(ARef("A1"), Anchor.Relative, decodeNumber) reads numeric value from A1
   */
  case Ref[A](at: ARef, anchor: Anchor, decode: Cell => Either[CodecError, A]) extends TExpr[A]

  /**
   * Polymorphic cell reference - defers type commitment until function context.
   *
   * This is used during parsing when the expected type of a cell reference is not yet known.
   * Function parsers convert PolyRef to Ref with appropriate type-coercing decoder based on
   * context.
   *
   * Matches Excel's type coercion semantics: A1=42 in =LEFT(A1,1) converts 42→"4".
   *
   * @param at
   *   The cell address to read from
   * @param anchor
   *   Anchoring mode for formula dragging (default: Relative)
   *
   * Example: PolyRef(ARef("A1"), Anchor.Relative) - type determined by enclosing function
   */
  case PolyRef(at: ARef, anchor: Anchor = Anchor.Relative) extends TExpr[Nothing]

  /**
   * Sheet-qualified cell reference - reads value from a cell in another sheet.
   *
   * @param sheet
   *   The target sheet name
   * @param at
   *   The cell address within the target sheet
   * @param anchor
   *   Anchoring mode for formula dragging
   * @param decode
   *   Function to decode the cell's value to type A
   *
   * Example: SheetRef("Sales", ARef("F4"), Anchor.Relative, decodeNumber) reads from Sales!F4
   */
  case SheetRef[A](
    sheet: SheetName,
    at: ARef,
    anchor: Anchor,
    decode: Cell => Either[CodecError, A]
  ) extends TExpr[A]

  /**
   * Polymorphic sheet-qualified cell reference - defers type commitment until function context.
   *
   * @param sheet
   *   The target sheet name
   * @param at
   *   The cell address within the target sheet
   * @param anchor
   *   Anchoring mode for formula dragging
   *
   * Example: SheetPolyRef("Sales", ARef("F4"), Anchor.Relative) - type determined by context
   */
  case SheetPolyRef(sheet: SheetName, at: ARef, anchor: Anchor = Anchor.Relative)
      extends TExpr[Nothing]

  /**
   * Sheet-qualified range reference - references a range in another sheet.
   *
   * @param sheet
   *   The target sheet name
   * @param range
   *   The cell range within the target sheet
   *
   * Example: SheetRange("Sales", CellRange(A1, A10)) represents Sales!A1:A10
   */
  case SheetRange(sheet: SheetName, range: CellRange) extends TExpr[Nothing]

  /**
   * Fold over a sheet-qualified range - aggregates values from another sheet.
   *
   * Similar to FoldRange but operates on a range in a different sheet.
   *
   * @param sheet
   *   The target sheet name
   * @param range
   *   The cell range within the target sheet
   * @param z
   *   Initial accumulator value
   * @param step
   *   Fold function combining accumulator with cell value
   * @param decode
   *   Function to decode each cell's value to type A
   *
   * Example: SUM(Sales!A1:A10) = SheetFoldRange(Sales, A1:A10, BigDecimal(0), _ + _, decodeNumber)
   */
  case SheetFoldRange[A, B](
    sheet: SheetName,
    range: CellRange,
    z: B,
    step: (B, A) => B,
    decode: Cell => Either[CodecError, A]
  ) extends TExpr[B]

  /**
   * Conditional expression - if/then/else.
   *
   * @param cond
   *   Boolean condition to evaluate
   * @param ifTrue
   *   Expression to evaluate if condition is true
   * @param ifFalse
   *   Expression to evaluate if condition is false
   *
   * Example: If(Ref(A1) > 0, Lit("Positive"), Lit("Non-positive"))
   */
  case If[A](cond: TExpr[Boolean], ifTrue: TExpr[A], ifFalse: TExpr[A]) extends TExpr[A]

  // Arithmetic operators (form commutative semiring over BigDecimal)

  /**
   * Addition: x + y
   *
   * Laws:
   *   - Commutative: Add(x, y) = Add(y, x)
   *   - Associative: Add(Add(x, y), z) = Add(x, Add(y, z))
   *   - Identity: Add(x, Lit(0)) = x
   */
  case Add(x: TExpr[BigDecimal], y: TExpr[BigDecimal]) extends TExpr[BigDecimal]

  /**
   * Subtraction: x - y
   */
  case Sub(x: TExpr[BigDecimal], y: TExpr[BigDecimal]) extends TExpr[BigDecimal]

  /**
   * Multiplication: x * y
   *
   * Laws:
   *   - Commutative: Mul(x, y) = Mul(y, x)
   *   - Associative: Mul(Mul(x, y), z) = Mul(x, Mul(y, z))
   *   - Identity: Mul(x, Lit(1)) = x
   *   - Distributive: Mul(x, Add(y, z)) = Add(Mul(x, y), Mul(x, z))
   */
  case Mul(x: TExpr[BigDecimal], y: TExpr[BigDecimal]) extends TExpr[BigDecimal]

  /**
   * Division: x / y
   *
   * Note: Division by zero returns evaluation error (not parse error)
   */
  case Div(x: TExpr[BigDecimal], y: TExpr[BigDecimal]) extends TExpr[BigDecimal]

  // Boolean/logical operators

  /**
   * Logical AND: x && y
   *
   * Laws:
   *   - Commutative: And(x, y) = And(y, x)
   *   - Associative: And(And(x, y), z) = And(x, And(y, z))
   *   - Short-circuit: If x evaluates to false, y is not evaluated
   */
  case And(x: TExpr[Boolean], y: TExpr[Boolean]) extends TExpr[Boolean]

  /**
   * Logical OR: x || y
   *
   * Laws:
   *   - Commutative: Or(x, y) = Or(y, x)
   *   - Associative: Or(Or(x, y), z) = Or(x, Or(y, z))
   *   - Short-circuit: If x evaluates to true, y is not evaluated
   */
  case Or(x: TExpr[Boolean], y: TExpr[Boolean]) extends TExpr[Boolean]

  /**
   * Logical NOT: !x
   *
   * Laws:
   *   - Double negation: Not(Not(x)) = x
   *   - De Morgan: Not(And(x, y)) = Or(Not(x), Not(y))
   */
  case Not(x: TExpr[Boolean]) extends TExpr[Boolean]

  // Range aggregation

  /**
   * Fold over a cell range - generalized aggregation.
   *
   * This is the foundation for SUM, COUNT, AVERAGE, etc.
   *
   * @param range
   *   The range of cells to aggregate
   * @param z
   *   Initial accumulator value
   * @param step
   *   Aggregation function: (accumulator, cell_value) => new_accumulator
   * @param decode
   *   Function to decode each cell's value to type A
   *
   * Example: SUM(A1:A10) = FoldRange(A1:A10, BigDecimal(0), _ + _, decodeNumber) Example:
   * COUNT(A1:A10) = FoldRange(A1:A10, 0, (acc, _) => acc + 1, decodeAny)
   */
  case FoldRange[A, B](
    range: CellRange,
    z: B,
    step: (B, A) => B,
    decode: Cell => Either[CodecError, A]
  ) extends TExpr[B]

  // Comparison operators (return Boolean)

  /**
   * Equality: x == y
   */
  case Eq[A](x: TExpr[A], y: TExpr[A]) extends TExpr[Boolean]

  /**
   * Inequality: x != y
   */
  case Neq[A](x: TExpr[A], y: TExpr[A]) extends TExpr[Boolean]

  /**
   * Less than: x < y (numeric comparison)
   */
  case Lt(x: TExpr[BigDecimal], y: TExpr[BigDecimal]) extends TExpr[Boolean]

  /**
   * Less than or equal: x <= y
   */
  case Lte(x: TExpr[BigDecimal], y: TExpr[BigDecimal]) extends TExpr[Boolean]

  /**
   * Greater than: x > y
   */
  case Gt(x: TExpr[BigDecimal], y: TExpr[BigDecimal]) extends TExpr[Boolean]

  /**
   * Greater than or equal: x >= y
   */
  case Gte(x: TExpr[BigDecimal], y: TExpr[BigDecimal]) extends TExpr[Boolean]

  // Type conversions

  /**
   * Convert BigDecimal to Int.
   *
   * Required for DATE(YEAR(A1), MONTH(A1), DAY(A1)) pattern where date extraction functions return
   * BigDecimal but DATE expects Int arguments.
   *
   * Validates at evaluation time that value is a valid integer.
   *
   * Example: ToInt(YEAR(A1)) where YEAR returns BigDecimal(2025) → Int(2025)
   */
  case ToInt(expr: TExpr[BigDecimal]) extends TExpr[Int]

  /**
   * Convert LocalDate to Excel serial number.
   *
   * Excel represents dates as the number of days since December 30, 1899. This enables date
   * arithmetic like TODAY()+30.
   *
   * Example: DateToSerial(TODAY()) where TODAY() = 2025-12-20 → BigDecimal(46011)
   */
  case DateToSerial(expr: TExpr[java.time.LocalDate]) extends TExpr[BigDecimal]

  /**
   * Convert LocalDateTime to Excel serial number.
   *
   * Excel represents dates/times as days since epoch + fractional day for time. This enables
   * datetime arithmetic like NOW()+0.5.
   *
   * Example: DateTimeToSerial(NOW()) where NOW() = 2025-12-20 12:00 → BigDecimal(46011.5)
   */
  case DateTimeToSerial(expr: TExpr[java.time.LocalDateTime]) extends TExpr[BigDecimal]

  // Text functions

  /**
   * Concatenate multiple text values: CONCATENATE(text1, text2, ...)
   *
   * Example: CONCATENATE("Hello", " ", "World") = "Hello World"
   */
  case Concatenate(xs: List[TExpr[String]]) extends TExpr[String]

  /**
   * Extract left substring: LEFT(text, n)
   *
   * @param text
   *   The text to extract from
   * @param n
   *   Number of characters to extract
   *
   * Example: LEFT("Hello", 3) = "Hel"
   */
  case Left(text: TExpr[String], n: TExpr[Int]) extends TExpr[String]

  /**
   * Extract right substring: RIGHT(text, n)
   *
   * @param text
   *   The text to extract from
   * @param n
   *   Number of characters to extract
   *
   * Example: RIGHT("Hello", 3) = "llo"
   */
  case Right(text: TExpr[String], n: TExpr[Int]) extends TExpr[String]

  /**
   * Text length: LEN(text)
   *
   * Returns BigDecimal to match Excel semantics and enable composition with arithmetic.
   *
   * Example: LEN("Hello") = 5
   */
  case Len(text: TExpr[String]) extends TExpr[BigDecimal]

  /**
   * Convert to uppercase: UPPER(text)
   *
   * Example: UPPER("hello") = "HELLO"
   */
  case Upper(text: TExpr[String]) extends TExpr[String]

  /**
   * Convert to lowercase: LOWER(text)
   *
   * Example: LOWER("HELLO") = "hello"
   */
  case Lower(text: TExpr[String]) extends TExpr[String]

  // Date/Time functions

  /**
   * Current date: TODAY()
   *
   * Returns the current date without time component. Requires Clock parameter in evaluator for
   * purity.
   *
   * Example: TODAY() = LocalDate.of(2025, 11, 21)
   */
  case Today() extends TExpr[java.time.LocalDate]

  /**
   * Current date and time: NOW()
   *
   * Returns the current date and time. Requires Clock parameter in evaluator for purity.
   *
   * Example: NOW() = LocalDateTime.of(2025, 11, 21, 18, 30, 0)
   */
  case Now() extends TExpr[java.time.LocalDateTime]

  /**
   * Mathematical constant PI: PI()
   *
   * Returns the mathematical constant pi (3.14159265358979...).
   *
   * Example: PI() = 3.141592653589793
   */
  case Pi() extends TExpr[BigDecimal]

  /**
   * Construct date from components: DATE(year, month, day)
   *
   * @param year
   *   The year (e.g., 2025)
   * @param month
   *   The month (1-12)
   * @param day
   *   The day (1-31)
   *
   * Example: DATE(2025, 11, 21) = LocalDate.of(2025, 11, 21)
   */
  case Date(year: TExpr[Int], month: TExpr[Int], day: TExpr[Int]) extends TExpr[java.time.LocalDate]

  /**
   * Extract year from date: YEAR(date)
   *
   * Returns BigDecimal to match Excel semantics (all numbers are doubles) and enable composition
   * with arithmetic operators (e.g., YEAR(A1) + 1).
   *
   * Example: YEAR(DATE(2025, 11, 21)) = 2025
   */
  case Year(date: TExpr[java.time.LocalDate]) extends TExpr[BigDecimal]

  /**
   * Extract month from date: MONTH(date)
   *
   * Returns BigDecimal to match Excel semantics and enable composition.
   *
   * Example: MONTH(DATE(2025, 11, 21)) = 11
   */
  case Month(date: TExpr[java.time.LocalDate]) extends TExpr[BigDecimal]

  /**
   * Extract day from date: DAY(date)
   *
   * Returns BigDecimal to match Excel semantics and enable composition.
   *
   * Example: DAY(DATE(2025, 11, 21)) = 21
   */
  case Day(date: TExpr[java.time.LocalDate]) extends TExpr[BigDecimal]

  /**
   * End of month: EOMONTH(start_date, months)
   *
   * Returns the last day of the month N months from start_date.
   *
   * @param startDate
   *   The starting date
   * @param months
   *   Number of months to add (can be negative)
   *
   * Example: EOMONTH(DATE(2025, 1, 15), 1) = DATE(2025, 2, 28)
   */
  case Eomonth(startDate: TExpr[java.time.LocalDate], months: TExpr[Int])
      extends TExpr[java.time.LocalDate]

  /**
   * Add months to date: EDATE(start_date, months)
   *
   * Returns the same day N months later (clamped to end of month if needed).
   *
   * @param startDate
   *   The starting date
   * @param months
   *   Number of months to add (can be negative)
   *
   * Example: EDATE(DATE(2025, 1, 31), 1) = DATE(2025, 2, 28)
   */
  case Edate(startDate: TExpr[java.time.LocalDate], months: TExpr[Int])
      extends TExpr[java.time.LocalDate]

  /**
   * Difference between dates: DATEDIF(start, end, unit)
   *
   * Returns the difference between two dates in the specified unit.
   *
   * @param startDate
   *   The starting date
   * @param endDate
   *   The ending date (must be >= startDate)
   * @param unit
   *   Unit of measurement: "Y" (years), "M" (months), "D" (days), "MD" (days ignoring
   *   months/years), "YM" (months ignoring years), "YD" (days ignoring years)
   *
   * Example: DATEDIF(DATE(2020, 1, 1), DATE(2025, 6, 15), "Y") = 5
   */
  case Datedif(
    startDate: TExpr[java.time.LocalDate],
    endDate: TExpr[java.time.LocalDate],
    unit: TExpr[String]
  ) extends TExpr[BigDecimal]

  /**
   * Count working days: NETWORKDAYS(start, end, [holidays])
   *
   * Returns the number of working days (Mon-Fri) between two dates, excluding holidays.
   *
   * @param startDate
   *   The starting date (inclusive)
   * @param endDate
   *   The ending date (inclusive)
   * @param holidays
   *   Optional range of dates to exclude
   *
   * Example: NETWORKDAYS(DATE(2025, 1, 1), DATE(2025, 1, 10)) = 8
   */
  case Networkdays(
    startDate: TExpr[java.time.LocalDate],
    endDate: TExpr[java.time.LocalDate],
    holidays: Option[CellRange]
  ) extends TExpr[BigDecimal]

  /**
   * Add working days: WORKDAY(start, days, [holidays])
   *
   * Returns the date after adding N working days (Mon-Fri), skipping holidays.
   *
   * @param startDate
   *   The starting date
   * @param days
   *   Number of working days to add (can be negative)
   * @param holidays
   *   Optional range of dates to exclude
   *
   * Example: WORKDAY(DATE(2025, 1, 1), 5) = DATE(2025, 1, 8)
   */
  case Workday(
    startDate: TExpr[java.time.LocalDate],
    days: TExpr[Int],
    holidays: Option[CellRange]
  ) extends TExpr[java.time.LocalDate]

  /**
   * Year fraction: YEARFRAC(start, end, [basis])
   *
   * Returns the fraction of a year between two dates based on the day count basis.
   *
   * @param startDate
   *   The starting date
   * @param endDate
   *   The ending date
   * @param basis
   *   Day count basis: 0=US 30/360 (default), 1=Actual/actual, 2=Actual/360, 3=Actual/365, 4=EU
   *   30/360
   *
   * Example: YEARFRAC(DATE(2025, 1, 1), DATE(2025, 7, 1), 1) ≈ 0.4959
   */
  case Yearfrac(
    startDate: TExpr[java.time.LocalDate],
    endDate: TExpr[java.time.LocalDate],
    basis: TExpr[Int]
  ) extends TExpr[BigDecimal]

  // Arithmetic range functions (SUM, COUNT, MIN, MAX, AVERAGE)

  /**
   * Sum of values in range: SUM(range) or SUM(Sheet!range)
   *
   * Sums all numeric values in range. Non-numeric cells are skipped (Excel-style).
   *
   * Example: SUM(A1:A10) = sum of numeric values in range
   */
  case Sum(range: TExpr.RangeLocation) extends TExpr[BigDecimal]

  /**
   * Count of numeric values in range: COUNT(range) or COUNT(Sheet!range)
   *
   * Counts cells containing numeric values. Non-numeric cells are skipped (Excel-style).
   *
   * Example: COUNT(A1:A10) = count of numeric values in range
   */
  case Count(range: TExpr.RangeLocation) extends TExpr[Int]

  /**
   * Minimum value in range: MIN(range) or MIN(Sheet!range)
   *
   * Example: MIN(A1:A10) = smallest numeric value in range Example: MIN(Sales!A1:A10) = smallest
   * value in Sales sheet
   */
  case Min(range: TExpr.RangeLocation) extends TExpr[BigDecimal]

  /**
   * Maximum value in range: MAX(range) or MAX(Sheet!range)
   *
   * Example: MAX(A1:A10) = largest numeric value in range Example: MAX(Sales!A1:A10) = largest
   * value in Sales sheet
   */
  case Max(range: TExpr.RangeLocation) extends TExpr[BigDecimal]

  /**
   * Average value in range: AVERAGE(range) or AVERAGE(Sheet!range)
   *
   * Computes sum/count of numeric values in range. Non-numeric cells are skipped (Excel-style).
   *
   * Example: AVERAGE(A1:A10) = mean of numeric values in range Example: AVERAGE(Sales!A1:A10) =
   * mean of values in Sales sheet
   */
  case Average(range: TExpr.RangeLocation) extends TExpr[BigDecimal]

  // Cross-sheet aggregate functions

  /**
   * Cross-sheet sum: SUM(Sheet!range)
   *
   * Sums numeric values in a range in another sheet.
   *
   * Example: SUM(Sales!A1:A10) = SheetSum(Sales, A1:A10)
   */
  case SheetSum(sheet: SheetName, range: CellRange) extends TExpr[BigDecimal]

  /**
   * Cross-sheet minimum value: MIN(Sheet!range)
   *
   * Evaluates MIN over a range in another sheet.
   *
   * Example: MIN(Sales!A1:A10) = SheetMin(Sales, A1:A10)
   */
  case SheetMin(sheet: SheetName, range: CellRange) extends TExpr[BigDecimal]

  /**
   * Cross-sheet maximum value: MAX(Sheet!range)
   *
   * Evaluates MAX over a range in another sheet.
   *
   * Example: MAX(Sales!A1:A10) = SheetMax(Sales, A1:A10)
   */
  case SheetMax(sheet: SheetName, range: CellRange) extends TExpr[BigDecimal]

  /**
   * Cross-sheet average value: AVERAGE(Sheet!range)
   *
   * Evaluates AVERAGE over a range in another sheet.
   *
   * Example: AVERAGE(Sales!A1:A10) = SheetAverage(Sales, A1:A10)
   */
  case SheetAverage(sheet: SheetName, range: CellRange) extends TExpr[BigDecimal]

  /**
   * Cross-sheet count: COUNT(Sheet!range)
   *
   * Counts numeric values in a range in another sheet.
   *
   * Example: COUNT(Data!B1:B10) = SheetCount(Data, B1:B10)
   */
  case SheetCount(sheet: SheetName, range: CellRange) extends TExpr[Int]

  /**
   * Generic aggregate function over a range using the Aggregator typeclass.
   *
   * This case unifies all simple aggregate functions (SUM, COUNT, MIN, MAX, AVERAGE) into a single
   * representation. The aggregatorId determines which Aggregator instance to use for evaluation.
   *
   * Benefits:
   *   - Single pattern match case in Evaluator, FormulaPrinter, FormulaShifter, DependencyGraph
   *   - Adding new aggregates only requires a new Aggregator given instance
   *   - Automatically supports full row/column references via RangeLocation
   *
   * @param aggregatorId
   *   The name of the aggregator (SUM, COUNT, MIN, MAX, AVERAGE)
   * @param location
   *   The range to aggregate (local or cross-sheet)
   *
   * Example: Aggregate("SUM", RangeLocation.Local(A1:A10))
   */
  case Aggregate(aggregatorId: String, location: TExpr.RangeLocation) extends TExpr[BigDecimal]

  // Financial functions

  /**
   * Net Present Value: NPV(rate, range)
   *
   * Semantics (v1):
   *   - `rate` is a numeric expression (BigDecimal)
   *   - `range` is a cell range of future cash flows at regular intervals
   *   - Non-numeric cells in the range are ignored (Excel-style)
   *   - First numeric value is treated as period 1 (t = 1), consistent with Excel NPV
   */
  case Npv(rate: TExpr[BigDecimal], values: TExpr.RangeLocation) extends TExpr[BigDecimal]

  /**
   * Internal Rate of Return: IRR(values, [guess])
   *
   * Semantics (v1):
   *   - `values` is a range of cash flows including the initial investment (t0)
   *   - Non-numeric cells ignored
   *   - Requires at least one positive and one negative flow
   *   - Optional `guess` is a numeric expression; default is 0.1 (10%)
   */
  case Irr(values: TExpr.RangeLocation, guess: Option[TExpr[BigDecimal]]) extends TExpr[BigDecimal]

  /**
   * Extended Net Present Value with irregular dates: XNPV(rate, values, dates)
   *
   * Semantics:
   *   - `rate` is the discount rate (BigDecimal)
   *   - `values` is a range of cash flows
   *   - `dates` is a range of dates corresponding to each cash flow
   *   - Formula: sum(value_i / (1 + rate)^((date_i - date_0) / 365))
   *   - dates and values must have same length
   *   - First date is the base date (date_0)
   *
   * Example: XNPV(0.1, A1:A5, B1:B5) with irregular payment dates
   */
  case Xnpv(
    rate: TExpr[BigDecimal],
    values: TExpr.RangeLocation,
    dates: TExpr.RangeLocation
  ) extends TExpr[BigDecimal]

  /**
   * Extended Internal Rate of Return with irregular dates: XIRR(values, dates, [guess])
   *
   * Semantics:
   *   - `values` is a range of cash flows (must have positive and negative)
   *   - `dates` is a range of dates corresponding to each cash flow
   *   - `guess` is optional starting point for Newton-Raphson (default 0.1)
   *   - Finds rate where XNPV = 0
   *   - dates and values must have same length
   *
   * Example: XIRR(A1:A5, B1:B5, 0.1) for irregular cash flow schedule
   */
  case Xirr(
    values: TExpr.RangeLocation,
    dates: TExpr.RangeLocation,
    guess: Option[TExpr[BigDecimal]]
  ) extends TExpr[BigDecimal]

  // ===== Time Value of Money (TVM) Functions =====

  /**
   * Payment per period: PMT(rate, nper, pv, [fv], [type])
   *
   * Semantics:
   *   - `rate` is the interest rate per period
   *   - `nper` is the total number of payment periods
   *   - `pv` is the present value (loan amount)
   *   - `fv` is the future value (default 0)
   *   - `type` is 0 for end of period (default), 1 for beginning
   *   - Negative result = outflow (payment made)
   *
   * Example: PMT(0.05/12, 24, 10000) for monthly payment on $10k loan at 5% APR for 2 years
   */
  case Pmt(
    rate: TExpr[BigDecimal],
    nper: TExpr[BigDecimal],
    pv: TExpr[BigDecimal],
    fv: Option[TExpr[BigDecimal]],
    pmtType: Option[TExpr[BigDecimal]]
  ) extends TExpr[BigDecimal]

  /**
   * Future value: FV(rate, nper, pmt, [pv], [type])
   *
   * Semantics:
   *   - `rate` is the interest rate per period
   *   - `nper` is the total number of payment periods
   *   - `pmt` is the payment per period (negative = outflow)
   *   - `pv` is the present value (default 0)
   *   - `type` is 0 for end of period (default), 1 for beginning
   *
   * Example: FV(0.05/12, 24, -100, 0) for future value of $100/month savings at 5% APR
   */
  case Fv(
    rate: TExpr[BigDecimal],
    nper: TExpr[BigDecimal],
    pmt: TExpr[BigDecimal],
    pv: Option[TExpr[BigDecimal]],
    pmtType: Option[TExpr[BigDecimal]]
  ) extends TExpr[BigDecimal]

  /**
   * Present value: PV(rate, nper, pmt, [fv], [type])
   *
   * Semantics:
   *   - `rate` is the interest rate per period
   *   - `nper` is the total number of payment periods
   *   - `pmt` is the payment per period (negative = outflow)
   *   - `fv` is the future value (default 0)
   *   - `type` is 0 for end of period (default), 1 for beginning
   *
   * Example: PV(0.05/12, 24, -500) for present value of $500/month payments at 5% APR
   */
  case Pv(
    rate: TExpr[BigDecimal],
    nper: TExpr[BigDecimal],
    pmt: TExpr[BigDecimal],
    fv: Option[TExpr[BigDecimal]],
    pmtType: Option[TExpr[BigDecimal]]
  ) extends TExpr[BigDecimal]

  /**
   * Number of periods: NPER(rate, pmt, pv, [fv], [type])
   *
   * Semantics:
   *   - `rate` is the interest rate per period
   *   - `pmt` is the payment per period (negative = outflow)
   *   - `pv` is the present value (loan amount)
   *   - `fv` is the future value (default 0)
   *   - `type` is 0 for end of period (default), 1 for beginning
   *
   * Example: NPER(0.05/12, -500, 10000) for months to pay off $10k loan at $500/month
   */
  case Nper(
    rate: TExpr[BigDecimal],
    pmt: TExpr[BigDecimal],
    pv: TExpr[BigDecimal],
    fv: Option[TExpr[BigDecimal]],
    pmtType: Option[TExpr[BigDecimal]]
  ) extends TExpr[BigDecimal]

  /**
   * Interest rate: RATE(nper, pmt, pv, [fv], [type], [guess])
   *
   * Semantics:
   *   - `nper` is the total number of payment periods
   *   - `pmt` is the payment per period (negative = outflow)
   *   - `pv` is the present value (loan amount)
   *   - `fv` is the future value (default 0)
   *   - `type` is 0 for end of period (default), 1 for beginning
   *   - `guess` is the starting guess for iteration (default 0.1)
   *   - Uses Newton-Raphson iteration (like IRR)
   *
   * Example: RATE(24, -500, 10000) for interest rate on $10k loan with $500/month payments
   */
  case Rate(
    nper: TExpr[BigDecimal],
    pmt: TExpr[BigDecimal],
    pv: TExpr[BigDecimal],
    fv: Option[TExpr[BigDecimal]],
    pmtType: Option[TExpr[BigDecimal]],
    guess: Option[TExpr[BigDecimal]]
  ) extends TExpr[BigDecimal]

  /**
   * Vertical lookup: VLOOKUP(lookup, table, colIndex, [rangeLookup])
   *
   * Semantics:
   *   - `lookup` is any value (text or numeric) - supports both text and number lookups
   *   - `table` is a rectangular CellRange; first column is the key
   *   - `colIndex` is 1-based column index into the table
   *   - `rangeLookup` = TRUE → approximate match (largest key <= lookup, numeric only)
   *   - `rangeLookup` = FALSE → exact match (case-insensitive for text)
   *   - Result is the CellValue at the matched row/column (preserves type)
   *
   * Excel-compatible behavior:
   *   - Text comparisons are case-insensitive
   *   - Numeric approximate match requires sorted ascending keys
   *   - #N/A error if no match found
   */
  case VLookup(
    lookup: TExpr[?],
    table: TExpr.RangeLocation,
    colIndex: TExpr[Int],
    rangeLookup: TExpr[Boolean]
  ) extends TExpr[CellValue]

  // Conditional aggregation functions (SUMIF/COUNTIF family)

  /**
   * Sum cells where criteria matches: SUMIF(range, criteria, [sum_range])
   *
   * Semantics:
   *   - `range` is the range to test against criteria
   *   - `criteria` is evaluated at runtime to determine matching rule
   *   - `sumRange` is the range to sum (defaults to `range` if None)
   *   - Non-numeric cells in sumRange are skipped (Excel behavior)
   *   - Criteria supports: exact match, wildcards (*,?), comparisons (>,>=,<,<=,<>)
   *
   * Example: SUMIF(A1:A10, "Apple", B1:B10) sums B values where A equals "Apple"
   */
  case SumIf(
    range: TExpr.RangeLocation,
    criteria: TExpr[?],
    sumRange: Option[TExpr.RangeLocation]
  ) extends TExpr[BigDecimal]

  /**
   * Count cells where criteria matches: COUNTIF(range, criteria)
   *
   * Semantics:
   *   - `range` is the range to test against criteria
   *   - `criteria` is evaluated at runtime to determine matching rule
   *   - Returns count as BigDecimal (Excel convention)
   *
   * Example: COUNTIF(A1:A10, ">100") counts cells greater than 100
   */
  case CountIf(
    range: TExpr.RangeLocation,
    criteria: TExpr[?]
  ) extends TExpr[BigDecimal]

  /**
   * Sum with multiple criteria (AND logic): SUMIFS(sum_range, criteria_range1, criteria1, ...)
   *
   * Semantics:
   *   - `sumRange` is the range to sum
   *   - `conditions` is a list of (range, criteria) pairs - ALL must match
   *   - All ranges must have same dimensions
   *   - Non-numeric cells in sumRange are skipped
   *
   * Example: SUMIFS(C1:C10, A1:A10, "Apple", B1:B10, ">100") sums C where A="Apple" AND B>100
   */
  case SumIfs(
    sumRange: TExpr.RangeLocation,
    conditions: List[(TExpr.RangeLocation, TExpr[?])]
  ) extends TExpr[BigDecimal]

  /**
   * Count with multiple criteria (AND logic): COUNTIFS(criteria_range1, criteria1, ...)
   *
   * Semantics:
   *   - `conditions` is a list of (range, criteria) pairs - ALL must match
   *   - All ranges must have same dimensions
   *   - Returns count as BigDecimal (Excel convention)
   *
   * Example: COUNTIFS(A1:A10, "Apple", B1:B10, ">100") counts where A="Apple" AND B>100
   */
  case CountIfs(
    conditions: List[(TExpr.RangeLocation, TExpr[?])]
  ) extends TExpr[BigDecimal]

  /**
   * Average cells where criteria matches: AVERAGEIF(range, criteria, [average_range])
   *
   * Semantics:
   *   - `range` is the range to test against criteria
   *   - `criteria` is evaluated at runtime to determine matching rule
   *   - `averageRange` is the range to average (defaults to `range` if None)
   *   - Non-numeric cells in averageRange are skipped (Excel behavior)
   *   - Returns #DIV/0! if no cells match or all matching cells are non-numeric
   *   - Criteria supports: exact match, wildcards (*,?), comparisons (>,>=,<,<=,<>)
   *
   * Example: AVERAGEIF(A1:A10, "Apple", B1:B10) averages B values where A equals "Apple"
   */
  case AverageIf(
    range: TExpr.RangeLocation,
    criteria: TExpr[?],
    averageRange: Option[TExpr.RangeLocation]
  ) extends TExpr[BigDecimal]

  /**
   * Average with multiple criteria (AND logic): AVERAGEIFS(avg_range, criteria_range1, criteria1,
   * ...)
   *
   * Semantics:
   *   - `averageRange` is the range to average
   *   - `conditions` is a list of (range, criteria) pairs - ALL must match
   *   - All ranges must have same dimensions
   *   - Non-numeric cells in averageRange are skipped
   *   - Returns #DIV/0! if no cells match or all matching cells are non-numeric
   *
   * Example: AVERAGEIFS(C1:C10, A1:A10, "Apple", B1:B10, ">100") averages C where A="Apple" AND
   * B>100
   */
  case AverageIfs(
    averageRange: TExpr.RangeLocation,
    conditions: List[(TExpr.RangeLocation, TExpr[?])]
  ) extends TExpr[BigDecimal]

  // Error handling functions

  /**
   * Error handler: IFERROR(value, value_if_error)
   *
   * Returns value if no error, otherwise returns value_if_error. Catches both evaluation errors and
   * CellValue.Error values.
   *
   * Example: IFERROR(A1/B1, 0) returns 0 if B1 is 0 (division by zero)
   */
  case Iferror(value: TExpr[CellValue], valueIfError: TExpr[CellValue]) extends TExpr[CellValue]

  /**
   * Error check: ISERROR(value)
   *
   * Returns TRUE if value results in any error, FALSE otherwise.
   *
   * Example: ISERROR(A1/B1) returns TRUE if B1 is 0
   */
  case Iserror(value: TExpr[CellValue]) extends TExpr[Boolean]

  /**
   * Error check (excluding #N/A): ISERR(value)
   *
   * Returns TRUE if value results in any error EXCEPT #N/A, FALSE otherwise. Use ISERROR to check
   * for all errors including #N/A.
   *
   * Example: ISERR(1/0) returns TRUE, ISERR(VLOOKUP("missing", A:A, 1, FALSE)) returns FALSE
   */
  case Iserr(value: TExpr[CellValue]) extends TExpr[Boolean]

  /**
   * Type check for numbers: ISNUMBER(value)
   *
   * Returns TRUE if value is numeric, FALSE otherwise.
   *
   * Example: ISNUMBER(42) returns TRUE, ISNUMBER("hello") returns FALSE
   */
  case Isnumber(value: TExpr[CellValue]) extends TExpr[Boolean]

  /**
   * Type check for text: ISTEXT(value)
   *
   * Returns TRUE if value is a text string, FALSE otherwise.
   *
   * Example: ISTEXT("hello") returns TRUE, ISTEXT(42) returns FALSE
   */
  case Istext(value: TExpr[CellValue]) extends TExpr[Boolean]

  /**
   * Type check for blank cells: ISBLANK(ref)
   *
   * Returns TRUE if the referenced cell is empty, FALSE otherwise. Note: cells containing empty
   * strings ("") are NOT considered blank.
   *
   * Example: ISBLANK(A1) returns TRUE if A1 is empty
   */
  case Isblank(value: TExpr[CellValue]) extends TExpr[Boolean]

  // Rounding and math functions

  /**
   * Round to specified digits: ROUND(number, num_digits)
   *
   * Rounds using HALF_UP mode (standard rounding). Negative num_digits rounds to left of decimal
   * point.
   *
   * Example: ROUND(2.5, 0) = 3, ROUND(1234, -2) = 1200
   */
  case Round(value: TExpr[BigDecimal], numDigits: TExpr[BigDecimal]) extends TExpr[BigDecimal]

  /**
   * Round up (away from zero): ROUNDUP(number, num_digits)
   *
   * Always rounds away from zero.
   *
   * Example: ROUNDUP(2.1, 0) = 3, ROUNDUP(-2.1, 0) = -3
   */
  case RoundUp(value: TExpr[BigDecimal], numDigits: TExpr[BigDecimal]) extends TExpr[BigDecimal]

  /**
   * Round down (toward zero): ROUNDDOWN(number, num_digits)
   *
   * Always rounds toward zero (truncation).
   *
   * Example: ROUNDDOWN(2.9, 0) = 2, ROUNDDOWN(-2.9, 0) = -2
   */
  case RoundDown(value: TExpr[BigDecimal], numDigits: TExpr[BigDecimal]) extends TExpr[BigDecimal]

  /**
   * Absolute value: ABS(number)
   *
   * Returns the absolute value of a number.
   *
   * Example: ABS(-5) = 5, ABS(5) = 5
   */
  case Abs(value: TExpr[BigDecimal]) extends TExpr[BigDecimal]

  /**
   * Square root: SQRT(number)
   *
   * Returns the square root of a number. Returns #NUM! for negative numbers.
   *
   * Example: SQRT(16) = 4, SQRT(2) ≈ 1.414
   */
  case Sqrt(value: TExpr[BigDecimal]) extends TExpr[BigDecimal]

  /**
   * Modulo: MOD(number, divisor)
   *
   * Returns the remainder after division. Result has the same sign as divisor (Excel semantics).
   *
   * Example: MOD(5, 3) = 2, MOD(-5, 3) = 1, MOD(5, -3) = -1
   */
  case Mod(number: TExpr[BigDecimal], divisor: TExpr[BigDecimal]) extends TExpr[BigDecimal]

  /**
   * Power: POWER(number, power)
   *
   * Returns number raised to a power.
   *
   * Example: POWER(2, 3) = 8, POWER(4, 0.5) = 2
   */
  case Power(number: TExpr[BigDecimal], power: TExpr[BigDecimal]) extends TExpr[BigDecimal]

  /**
   * Logarithm: LOG(number, [base])
   *
   * Returns the logarithm of a number to a specified base. Default base is 10.
   *
   * Example: LOG(100) = 2, LOG(8, 2) = 3
   */
  case Log(number: TExpr[BigDecimal], base: TExpr[BigDecimal]) extends TExpr[BigDecimal]

  /**
   * Natural logarithm: LN(number)
   *
   * Returns the natural logarithm (base e) of a number.
   *
   * Example: LN(E()) ≈ 1, LN(2.718281828) ≈ 1
   */
  case Ln(value: TExpr[BigDecimal]) extends TExpr[BigDecimal]

  /**
   * Exponential: EXP(number)
   *
   * Returns e raised to the power of number.
   *
   * Example: EXP(1) ≈ 2.718, EXP(0) = 1
   */
  case Exp(value: TExpr[BigDecimal]) extends TExpr[BigDecimal]

  /**
   * Floor: FLOOR(number, significance)
   *
   * Rounds number down toward zero to the nearest multiple of significance.
   *
   * Example: FLOOR(2.5, 1) = 2, FLOOR(-2.5, -1) = -2, FLOOR(1.5, 0.5) = 1.5
   */
  case Floor(number: TExpr[BigDecimal], significance: TExpr[BigDecimal]) extends TExpr[BigDecimal]

  /**
   * Ceiling: CEILING(number, significance)
   *
   * Rounds number up away from zero to the nearest multiple of significance.
   *
   * Example: CEILING(2.5, 1) = 3, CEILING(-2.5, -1) = -3, CEILING(1.2, 0.5) = 1.5
   */
  case Ceiling(number: TExpr[BigDecimal], significance: TExpr[BigDecimal]) extends TExpr[BigDecimal]

  /**
   * Truncate: TRUNC(number, [num_digits])
   *
   * Truncates number to specified number of decimal places (removes fractional part). Default
   * num_digits is 0.
   *
   * Example: TRUNC(8.9) = 8, TRUNC(-8.9) = -8, TRUNC(3.14159, 2) = 3.14
   */
  case Trunc(number: TExpr[BigDecimal], numDigits: TExpr[BigDecimal]) extends TExpr[BigDecimal]

  /**
   * Sign: SIGN(number)
   *
   * Returns the sign of a number: 1 for positive, -1 for negative, 0 for zero.
   *
   * Example: SIGN(5) = 1, SIGN(-5) = -1, SIGN(0) = 0
   */
  case Sign(value: TExpr[BigDecimal]) extends TExpr[BigDecimal]

  /**
   * Integer part: INT(number)
   *
   * Rounds number down to the nearest integer (floor function). Unlike TRUNC, INT always rounds
   * toward negative infinity.
   *
   * Example: INT(8.9) = 8, INT(-8.9) = -9
   */
  case Int_(value: TExpr[BigDecimal]) extends TExpr[BigDecimal]

  // Reference information functions

  /**
   * Row number: ROW(reference)
   *
   * Returns the 1-based row number of a cell reference. For ranges, returns the row of the top-left
   * cell.
   *
   * Example: ROW(A5) = 5, ROW(B1:C10) = 1
   *
   * @note
   *   Named Row_ to avoid conflict with com.tjclp.xl.Row opaque type
   */
  case Row_(ref: TExpr[?]) extends TExpr[BigDecimal]

  /**
   * Column number: COLUMN(reference)
   *
   * Returns the 1-based column number of a cell reference. For ranges, returns the column of the
   * top-left cell.
   *
   * Example: COLUMN(C1) = 3, COLUMN(B1:D10) = 2
   */
  case Column_(ref: TExpr[?]) extends TExpr[BigDecimal]

  /**
   * Row count: ROWS(range)
   *
   * Returns the number of rows in a range.
   *
   * Example: ROWS(A1:A10) = 10, ROWS(B2:D5) = 4
   */
  case Rows(range: TExpr[?]) extends TExpr[BigDecimal]

  /**
   * Column count: COLUMNS(range)
   *
   * Returns the number of columns in a range.
   *
   * Example: COLUMNS(A1:D1) = 4, COLUMNS(B2:E5) = 4
   */
  case Columns(range: TExpr[?]) extends TExpr[BigDecimal]

  /**
   * Create cell address: ADDRESS(row, column, [abs_num], [a1], [sheet])
   *
   * Returns a cell reference as text string.
   *
   * @param row
   *   1-based row number
   * @param col
   *   1-based column number
   * @param absNum
   *   Anchor style: 1=$A$1, 2=A$1, 3=$A1, 4=A1 (default 1)
   * @param a1Style
   *   TRUE for A1 notation (default), FALSE for R1C1
   * @param sheetName
   *   Optional sheet name to prepend
   *
   * Example: ADDRESS(1, 1) = "$A$1", ADDRESS(1, 1, 4) = "A1"
   */
  case Address(
    row: TExpr[BigDecimal],
    col: TExpr[BigDecimal],
    absNum: TExpr[BigDecimal],
    a1Style: TExpr[Boolean],
    sheetName: Option[TExpr[String]]
  ) extends TExpr[String]

  // Array and advanced lookup functions

  /**
   * Multiply corresponding elements across arrays and sum: SUMPRODUCT(array1, [array2], ...)
   *
   * Semantics:
   *   - All arrays must have same dimensions (width and height)
   *   - Non-numeric cells coerced: TRUE→1, FALSE→0, text/empty→0
   *   - Returns sum of element-wise products
   *
   * Example: SUMPRODUCT(A1:A3, B1:B3) = A1*B1 + A2*B2 + A3*B3
   */
  case SumProduct(arrays: List[TExpr.RangeLocation]) extends TExpr[BigDecimal]

  /**
   * Advanced lookup: XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found],
   * [match_mode], [search_mode])
   *
   * Semantics:
   *   - lookup_array and return_array must have same dimensions
   *   - match_mode: 0=exact (default), -1=next smaller, 1=next larger, 2=wildcard
   *   - search_mode: 1=first-to-last (default), -1=last-to-first, 2=binary asc, -2=binary desc
   *   - if_not_found: expression to return if no match (default #N/A)
   *
   * Example: XLOOKUP("Apple", A1:A10, B1:B10) returns corresponding B value
   */
  case XLookup(
    lookupValue: TExpr[?],
    lookupArray: TExpr.RangeLocation,
    returnArray: TExpr.RangeLocation,
    ifNotFound: Option[TExpr[?]],
    matchMode: TExpr[Int],
    searchMode: TExpr[Int]
  ) extends TExpr[CellValue]

  /**
   * Index into array: INDEX(array, row_num, [column_num])
   *
   * Returns the value at a specific row/column position in a range.
   *
   * Semantics:
   *   - row_num: 1-based row position within the array
   *   - column_num: 1-based column position (optional, defaults to 1 for single-column ranges)
   *   - Returns #REF! if indices are out of bounds
   *
   * Example: INDEX(A1:C3, 2, 3) returns value at row 2, column 3 of the range
   */
  case Index(
    array: TExpr.RangeLocation,
    rowNum: TExpr[BigDecimal],
    colNum: Option[TExpr[BigDecimal]]
  ) extends TExpr[CellValue]

  /**
   * Find position: MATCH(lookup_value, lookup_array, [match_type])
   *
   * Returns the relative position of a value in a range (1-based).
   *
   * Semantics:
   *   - match_type=1 (default): largest value <= lookup_value (array must be sorted ascending)
   *   - match_type=0: exact match (array need not be sorted)
   *   - match_type=-1: smallest value >= lookup_value (array must be sorted descending)
   *   - Returns #N/A if no match found
   *
   * Example: MATCH("B", {"A","B","C"}, 0) returns 2
   */
  case Match(
    lookupValue: TExpr[?],
    lookupArray: TExpr.RangeLocation,
    matchType: TExpr[BigDecimal]
  ) extends TExpr[BigDecimal]

object TExpr:

  // ===== Range Location Abstraction =====

  /**
   * Where a range is located - same sheet or cross-sheet.
   *
   * This enum unifies local and cross-sheet range references, eliminating the need for paired TExpr
   * cases (e.g., Min + SheetMin). Used by TExpr.Aggregate for unified aggregation.
   */
  enum RangeLocation derives CanEqual:
    case Local(range: CellRange)
    case CrossSheet(sheet: SheetName, range: CellRange)

  object RangeLocation:
    extension (loc: RangeLocation)
      /** Extract the CellRange regardless of location */
      def range: CellRange = loc match
        case Local(r) => r
        case CrossSheet(_, r) => r

      /** Get sheet name for cross-sheet, None for local */
      def sheetName: Option[SheetName] = loc match
        case CrossSheet(s, _) => Some(s)
        case _ => None

      /** Get cells for local ranges only (for intra-sheet dependency graphs) */
      def localCells: Set[ARef] = loc match
        case Local(r) => r.cells.toSet
        case CrossSheet(_, _) => Set.empty

      /**
       * Get cells for local ranges, bounded by the sheet's used range.
       *
       * Preferred for dependency graph construction to avoid materializing 1M+ cells for full
       * column/row references like A:A or 1:1.
       *
       * @param bounds
       *   Optional bounding range (typically sheet.usedRange)
       * @return
       *   Set of cell references in the intersection of this range and bounds
       */
      def localCellsBounded(bounds: Option[CellRange]): Set[ARef] = loc match
        case Local(r) =>
          bounds match
            case Some(b) => r.intersect(b).map(_.cells.toSet).getOrElse(Set.empty)
            case None => r.cells.toSet
        case CrossSheet(_, _) => Set.empty

      /** Check if this is a cross-sheet reference */
      def isCrossSheet: Boolean = loc match
        case CrossSheet(_, _) => true
        case _ => false

      /** Get all cells in the range (delegates to underlying CellRange) */
      def cells: Iterator[ARef] = loc.range.cells

      /** Get width of the range */
      def width: Int = loc.range.width

      /** Get height of the range */
      def height: Int = loc.range.height

      /** Get A1 string representation */
      def toA1: String = loc match
        case Local(r) => r.toA1
        case CrossSheet(s, r) => s"${s.value}!${r.toA1}"

      /** Get starting column of the range */
      def colStart: Column = loc.range.colStart

      /** Get ending column of the range */
      def colEnd: Column = loc.range.colEnd

      /** Get starting row of the range */
      def rowStart: Row = loc.range.rowStart

      /** Get ending row of the range */
      def rowEnd: Row = loc.range.rowEnd

      /** Get start cell reference of the range */
      def start: ARef = loc.range.start

      /** Get end cell reference of the range */
      def end: ARef = loc.range.end

  // ===== Smart Constructors =====

  /**
   * Smart constructor for literals.
   *
   * Example: TExpr.lit(42)
   */
  def lit[A](value: A): TExpr[A] = Lit(value)

  /**
   * Smart constructor for cell references.
   *
   * Example: TExpr.ref(ARef("A1"), Anchor.Relative, codec)
   */
  def ref[A](at: ARef, anchor: Anchor, decode: Cell => Either[CodecError, A]): TExpr[A] =
    Ref(at, anchor, decode)

  /**
   * Smart constructor for cell references with default Relative anchor.
   *
   * Example: TExpr.ref(ARef("A1"), codec)
   */
  def ref[A](at: ARef, decode: Cell => Either[CodecError, A]): TExpr[A] =
    Ref(at, Anchor.Relative, decode)

  /**
   * Smart constructor for conditionals.
   *
   * Example: TExpr.cond(test, ifTrue, ifFalse)
   */
  def cond[A](test: TExpr[Boolean], ifTrue: TExpr[A], ifFalse: TExpr[A]): TExpr[A] =
    If(test, ifTrue, ifFalse)

  // Convenience constructors for common operations

  /**
   * SUM aggregation: sum all numeric values in range.
   *
   * Example: TExpr.sum(CellRange("A1:A10"))
   */
  def sum(range: CellRange): TExpr[BigDecimal] =
    FoldRange(
      range,
      BigDecimal(0),
      (acc: BigDecimal, value: BigDecimal) => acc + value,
      decodeNumeric
    )

  /**
   * COUNT aggregation: count non-empty cells in range.
   *
   * Example: TExpr.count(CellRange("A1:A10"))
   */
  def count(range: CellRange): TExpr[Int] =
    FoldRange(
      range,
      0,
      (acc: Int, _: Option[Any]) => acc + 1,
      decodeAny
    )

  /**
   * AVERAGE aggregation: average of numeric values in range.
   *
   * Example: TExpr.average(CellRange("A1:A10"))
   */
  def average(range: CellRange): TExpr[BigDecimal] = Average(RangeLocation.Local(range))

  /**
   * MIN aggregation: minimum numeric value in range.
   *
   * Example: TExpr.min(CellRange("A1:A10"))
   */
  def min(range: CellRange): TExpr[BigDecimal] = Min(RangeLocation.Local(range))

  /**
   * MAX aggregation: maximum numeric value in range.
   *
   * Example: TExpr.max(CellRange("A1:A10"))
   */
  def max(range: CellRange): TExpr[BigDecimal] = Max(RangeLocation.Local(range))

  // Financial function smart constructors

  /**
   * Smart constructor for NPV over a range of cash flows.
   *
   * Example: TExpr.npv(TExpr.Lit(BigDecimal("0.1")), CellRange("A2:A6"))
   */
  def npv(rate: TExpr[BigDecimal], values: CellRange): TExpr[BigDecimal] =
    Npv(rate, RangeLocation.Local(values))

  /**
   * Smart constructor for IRR with optional guess.
   *
   * Example: TExpr.irr(CellRange("A1:A6"), Some(TExpr.Lit(BigDecimal("0.15"))))
   */
  def irr(values: CellRange, guess: Option[TExpr[BigDecimal]] = None): TExpr[BigDecimal] =
    Irr(RangeLocation.Local(values), guess)

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
    Xnpv(rate, RangeLocation.Local(values), RangeLocation.Local(dates))

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
    Xirr(RangeLocation.Local(values), RangeLocation.Local(dates), guess)

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
    Pmt(rate, nper, pv, fv, pmtType)

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
    Fv(rate, nper, pmt, pv, pmtType)

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
    Pv(rate, nper, pmt, fv, pmtType)

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
    Nper(rate, pmt, pv, fv, pmtType)

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
    Rate(nper, pmt, pv, fv, pmtType, guess)

  /**
   * Smart constructor for VLOOKUP (supports text and numeric lookups).
   *
   * Example: TExpr.vlookup(TExpr.Lit("Widget A"), CellRange("A1:D10"), TExpr.Lit(4),
   * TExpr.Lit(false))
   */
  def vlookup(
    lookup: TExpr[?],
    table: CellRange,
    colIndex: TExpr[Int],
    rangeLookup: TExpr[Boolean] = Lit(true)
  ): TExpr[CellValue] =
    VLookup(lookup, RangeLocation.Local(table), colIndex, rangeLookup)

  /**
   * Smart constructor for VLOOKUP with explicit RangeLocation (supports cross-sheet lookups).
   *
   * Example: TExpr.vlookupWithLocation(TExpr.Lit("Widget A"),
   * RangeLocation.CrossSheet(SheetName("Lookup"), CellRange("A1:D10")), TExpr.Lit(4),
   * TExpr.Lit(false))
   */
  def vlookupWithLocation(
    lookup: TExpr[?],
    table: RangeLocation,
    colIndex: TExpr[Int],
    rangeLookup: TExpr[Boolean] = Lit(true)
  ): TExpr[CellValue] =
    VLookup(lookup, table, colIndex, rangeLookup)

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
    SumIf(RangeLocation.Local(range), criteria, sumRange.map(RangeLocation.Local(_)))

  /**
   * COUNTIF: count cells where criteria matches.
   *
   * Example: TExpr.countIf(CellRange("A1:A10"), TExpr.Lit(">100"))
   */
  def countIf(range: CellRange, criteria: TExpr[?]): TExpr[BigDecimal] =
    CountIf(RangeLocation.Local(range), criteria)

  /**
   * SUMIFS: sum with multiple criteria (AND logic).
   *
   * Example: TExpr.sumIfs(CellRange("C1:C10"), List((CellRange("A1:A10"), TExpr.Lit("Apple"))))
   */
  def sumIfs(
    sumRange: CellRange,
    conditions: List[(CellRange, TExpr[?])]
  ): TExpr[BigDecimal] =
    SumIfs(
      RangeLocation.Local(sumRange),
      conditions.map { case (r, c) => (RangeLocation.Local(r), c) }
    )

  /**
   * COUNTIFS: count with multiple criteria (AND logic).
   *
   * Example: TExpr.countIfs(List((CellRange("A1:A10"), TExpr.Lit("Apple"))))
   */
  def countIfs(conditions: List[(CellRange, TExpr[?])]): TExpr[BigDecimal] =
    CountIfs(conditions.map { case (r, c) => (RangeLocation.Local(r), c) })

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
    AverageIf(RangeLocation.Local(range), criteria, averageRange.map(RangeLocation.Local(_)))

  /**
   * AVERAGEIFS: average with multiple criteria (AND logic).
   *
   * Example: TExpr.averageIfs(CellRange("C1:C10"), List((CellRange("A1:A10"), TExpr.Lit("Apple"))))
   */
  def averageIfs(
    averageRange: CellRange,
    conditions: List[(CellRange, TExpr[?])]
  ): TExpr[BigDecimal] =
    AverageIfs(
      RangeLocation.Local(averageRange),
      conditions.map { case (r, c) => (RangeLocation.Local(r), c) }
    )

  // Error handling function smart constructors

  /**
   * IFERROR: return value_if_error if value results in error.
   *
   * Example: TExpr.iferror(TExpr.Div(...), TExpr.Lit(CellValue.Number(0)))
   */
  def iferror(value: TExpr[CellValue], valueIfError: TExpr[CellValue]): TExpr[CellValue] =
    Iferror(value, valueIfError)

  /**
   * ISERROR: check if expression results in error.
   *
   * Example: TExpr.iserror(TExpr.Div(...))
   */
  def iserror(value: TExpr[CellValue]): TExpr[Boolean] =
    Iserror(value)

  /**
   * ISERR: check if expression results in error (excluding #N/A).
   *
   * Example: TExpr.iserr(TExpr.Div(...))
   */
  def iserr(value: TExpr[CellValue]): TExpr[Boolean] =
    Iserr(value)

  /**
   * ISNUMBER: check if value is numeric.
   *
   * Example: TExpr.isnumber(TExpr.ref(ARef("A1")))
   */
  def isnumber(value: TExpr[CellValue]): TExpr[Boolean] =
    Isnumber(value)

  /**
   * ISTEXT: check if value is text.
   *
   * Example: TExpr.istext(TExpr.ref(ARef("A1")))
   */
  def istext(value: TExpr[CellValue]): TExpr[Boolean] =
    Istext(value)

  /**
   * ISBLANK: check if cell is empty.
   *
   * Example: TExpr.isblank(TExpr.ref(ARef("A1")))
   */
  def isblank(value: TExpr[CellValue]): TExpr[Boolean] =
    Isblank(value)

  // Rounding and math function smart constructors

  /**
   * ROUND: round number to specified digits.
   *
   * Example: TExpr.round(TExpr.Lit(2.5), TExpr.Lit(0))
   */
  def round(value: TExpr[BigDecimal], numDigits: TExpr[BigDecimal]): TExpr[BigDecimal] =
    Round(value, numDigits)

  /**
   * ROUNDUP: round away from zero.
   *
   * Example: TExpr.roundUp(TExpr.Lit(2.1), TExpr.Lit(0))
   */
  def roundUp(value: TExpr[BigDecimal], numDigits: TExpr[BigDecimal]): TExpr[BigDecimal] =
    RoundUp(value, numDigits)

  /**
   * ROUNDDOWN: round toward zero (truncate).
   *
   * Example: TExpr.roundDown(TExpr.Lit(2.9), TExpr.Lit(0))
   */
  def roundDown(value: TExpr[BigDecimal], numDigits: TExpr[BigDecimal]): TExpr[BigDecimal] =
    RoundDown(value, numDigits)

  /**
   * ABS: absolute value.
   *
   * Example: TExpr.abs(TExpr.Lit(-5))
   */
  def abs(value: TExpr[BigDecimal]): TExpr[BigDecimal] =
    Abs(value)

  /**
   * SQRT: square root.
   *
   * Example: TExpr.sqrt(TExpr.Lit(16))
   */
  def sqrt(value: TExpr[BigDecimal]): TExpr[BigDecimal] =
    Sqrt(value)

  /**
   * MOD: modulo (remainder after division).
   *
   * Example: TExpr.mod(TExpr.Lit(5), TExpr.Lit(3))
   */
  def mod(number: TExpr[BigDecimal], divisor: TExpr[BigDecimal]): TExpr[BigDecimal] =
    Mod(number, divisor)

  /**
   * POWER: number raised to a power.
   *
   * Example: TExpr.power(TExpr.Lit(2), TExpr.Lit(3))
   */
  def power(number: TExpr[BigDecimal], power: TExpr[BigDecimal]): TExpr[BigDecimal] =
    Power(number, power)

  /**
   * LOG: logarithm to specified base.
   *
   * Example: TExpr.log(TExpr.Lit(100), TExpr.Lit(10))
   */
  def log(number: TExpr[BigDecimal], base: TExpr[BigDecimal]): TExpr[BigDecimal] =
    Log(number, base)

  /**
   * LN: natural logarithm (base e).
   *
   * Example: TExpr.ln(TExpr.Lit(2.718281828))
   */
  def ln(value: TExpr[BigDecimal]): TExpr[BigDecimal] =
    Ln(value)

  /**
   * EXP: e raised to a power.
   *
   * Example: TExpr.exp(TExpr.Lit(1))
   */
  def exp(value: TExpr[BigDecimal]): TExpr[BigDecimal] =
    Exp(value)

  /**
   * FLOOR: round down to nearest multiple of significance.
   *
   * Example: TExpr.floor(TExpr.Lit(2.5), TExpr.Lit(1))
   */
  def floor(number: TExpr[BigDecimal], significance: TExpr[BigDecimal]): TExpr[BigDecimal] =
    Floor(number, significance)

  /**
   * CEILING: round up to nearest multiple of significance.
   *
   * Example: TExpr.ceiling(TExpr.Lit(2.5), TExpr.Lit(1))
   */
  def ceiling(number: TExpr[BigDecimal], significance: TExpr[BigDecimal]): TExpr[BigDecimal] =
    Ceiling(number, significance)

  /**
   * TRUNC: truncate to specified number of decimal places.
   *
   * Example: TExpr.trunc(TExpr.Lit(8.9), TExpr.Lit(0))
   */
  def trunc(number: TExpr[BigDecimal], numDigits: TExpr[BigDecimal]): TExpr[BigDecimal] =
    Trunc(number, numDigits)

  /**
   * SIGN: sign of a number (1, -1, or 0).
   *
   * Example: TExpr.sign(TExpr.Lit(-5))
   */
  def sign(value: TExpr[BigDecimal]): TExpr[BigDecimal] =
    Sign(value)

  /**
   * INT: round down to nearest integer (floor).
   *
   * Example: TExpr.int_(TExpr.Lit(8.9))
   */
  def int_(value: TExpr[BigDecimal]): TExpr[BigDecimal] =
    Int_(value)

  // Reference information function smart constructors

  /**
   * Create ROW expression.
   *
   * Example: TExpr.row(TExpr.PolyRef(ref"A5", Anchor.Relative))
   */
  def row(ref: TExpr[?]): TExpr[BigDecimal] =
    Row_(ref)

  /**
   * Create COLUMN expression.
   *
   * Example: TExpr.column(TExpr.PolyRef(ref"C1", Anchor.Relative))
   */
  def column(ref: TExpr[?]): TExpr[BigDecimal] =
    Column_(ref)

  /**
   * Create ROWS expression.
   *
   * Example: TExpr.rows(TExpr.FoldRange(range, ...))
   */
  def rows(range: TExpr[?]): TExpr[BigDecimal] =
    Rows(range)

  /**
   * Create COLUMNS expression.
   *
   * Example: TExpr.columns(TExpr.FoldRange(range, ...))
   */
  def columns(range: TExpr[?]): TExpr[BigDecimal] =
    Columns(range)

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
    Address(row, col, absNum, a1Style, sheetName)

  // Array and advanced lookup function smart constructors

  /**
   * SUMPRODUCT: multiply corresponding elements and sum.
   *
   * Example: TExpr.sumProduct(List(CellRange.parse("A1:A3").toOption.get,
   * CellRange.parse("B1:B3").toOption.get))
   */
  def sumProduct(arrays: List[CellRange]): TExpr[BigDecimal] =
    SumProduct(arrays.map(RangeLocation.Local(_)))

  /**
   * XLOOKUP: advanced lookup with flexible matching.
   *
   * @param lookupValue
   *   The value to search for
   * @param lookupArray
   *   The range to search in
   * @param returnArray
   *   The range to return values from (same dimensions as lookupArray)
   * @param ifNotFound
   *   Optional value to return if no match (default: #N/A error)
   * @param matchMode
   *   0=exact (default), -1=next smaller, 1=next larger, 2=wildcard
   * @param searchMode
   *   1=first-to-last (default), -1=last-to-first, 2=binary asc, -2=binary desc
   *
   * Example: TExpr.xlookup(TExpr.Lit("Apple"), lookupRange, returnRange)
   */
  def xlookup(
    lookupValue: TExpr[?],
    lookupArray: CellRange,
    returnArray: CellRange,
    ifNotFound: Option[TExpr[?]] = None,
    matchMode: TExpr[Int] = Lit(0),
    searchMode: TExpr[Int] = Lit(1)
  ): TExpr[CellValue] =
    XLookup(
      lookupValue,
      RangeLocation.Local(lookupArray),
      RangeLocation.Local(returnArray),
      ifNotFound,
      matchMode,
      searchMode
    )

  /**
   * INDEX: get value at position in array.
   *
   * @param array
   *   The range to index into
   * @param rowNum
   *   1-based row position
   * @param colNum
   *   Optional 1-based column position (defaults to 1 for single-column ranges)
   *
   * Example: TExpr.index(range, TExpr.Lit(2), Some(TExpr.Lit(3)))
   */
  def index(
    array: CellRange,
    rowNum: TExpr[BigDecimal],
    colNum: Option[TExpr[BigDecimal]] = None
  ): TExpr[CellValue] =
    Index(RangeLocation.Local(array), rowNum, colNum)

  /**
   * MATCH: find position of value in array.
   *
   * @param lookupValue
   *   The value to search for
   * @param lookupArray
   *   The range to search in
   * @param matchType
   *   1=largest <= (default), 0=exact, -1=smallest >=
   *
   * Example: TExpr.matchExpr(TExpr.Lit("B"), range, TExpr.Lit(0))
   */
  def matchExpr(
    lookupValue: TExpr[?],
    lookupArray: CellRange,
    matchType: TExpr[BigDecimal] = Lit(BigDecimal(1))
  ): TExpr[BigDecimal] =
    Match(lookupValue, RangeLocation.Local(lookupArray), matchType)

  // Text function smart constructors

  /**
   * CONCATENATE text values.
   *
   * Example: TExpr.concatenate(List(TExpr.Lit("Hello"), TExpr.Lit(" "), TExpr.Lit("World")))
   */
  def concatenate(xs: List[TExpr[String]]): TExpr[String] = Concatenate(xs)

  /**
   * LEFT substring extraction.
   *
   * Example: TExpr.left(TExpr.Lit("Hello"), TExpr.Lit(3))
   */
  def left(text: TExpr[String], n: TExpr[Int]): TExpr[String] = Left(text, n)

  /**
   * RIGHT substring extraction.
   *
   * Example: TExpr.right(TExpr.Lit("Hello"), TExpr.Lit(3))
   */
  def right(text: TExpr[String], n: TExpr[Int]): TExpr[String] = Right(text, n)

  /**
   * LEN text length.
   *
   * Returns BigDecimal to match Excel semantics.
   *
   * Example: TExpr.len(TExpr.Lit("Hello"))
   */
  def len(text: TExpr[String]): TExpr[BigDecimal] = Len(text)

  /**
   * UPPER convert to uppercase.
   *
   * Example: TExpr.upper(TExpr.Lit("hello"))
   */
  def upper(text: TExpr[String]): TExpr[String] = Upper(text)

  /**
   * LOWER convert to lowercase.
   *
   * Example: TExpr.lower(TExpr.Lit("HELLO"))
   */
  def lower(text: TExpr[String]): TExpr[String] = Lower(text)

  // Date/Time function smart constructors

  /**
   * TODAY current date.
   *
   * Example: TExpr.today()
   */
  def today(): TExpr[java.time.LocalDate] = Today()

  /**
   * NOW current date and time.
   *
   * Example: TExpr.now()
   */
  def now(): TExpr[java.time.LocalDateTime] = Now()

  /**
   * PI mathematical constant.
   *
   * Example: TExpr.pi()
   */
  def pi(): TExpr[BigDecimal] = Pi()

  /**
   * DATE construct from year, month, day.
   *
   * Example: TExpr.date(TExpr.Lit(2025), TExpr.Lit(11), TExpr.Lit(21))
   */
  def date(year: TExpr[Int], month: TExpr[Int], day: TExpr[Int]): TExpr[java.time.LocalDate] =
    Date(year, month, day)

  /**
   * YEAR extract year from date.
   *
   * Returns BigDecimal to match Excel semantics and enable arithmetic composition.
   *
   * Example: TExpr.year(TExpr.date(TExpr.Lit(2025), TExpr.Lit(11), TExpr.Lit(21)))
   */
  def year(date: TExpr[java.time.LocalDate]): TExpr[BigDecimal] = Year(date)

  /**
   * MONTH extract month from date.
   *
   * Returns BigDecimal to match Excel semantics and enable arithmetic composition.
   *
   * Example: TExpr.month(TExpr.date(TExpr.Lit(2025), TExpr.Lit(11), TExpr.Lit(21)))
   */
  def month(date: TExpr[java.time.LocalDate]): TExpr[BigDecimal] = Month(date)

  /**
   * DAY extract day from date.
   *
   * Returns BigDecimal to match Excel semantics and enable arithmetic composition.
   *
   * Example: TExpr.day(TExpr.date(TExpr.Lit(2025), TExpr.Lit(11), TExpr.Lit(21)))
   */
  def day(date: TExpr[java.time.LocalDate]): TExpr[BigDecimal] = Day(date)

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
    Eomonth(startDate, months)

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
    Edate(startDate, months)

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
    Datedif(startDate, endDate, unit)

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
    Networkdays(startDate, endDate, holidays)

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
    Workday(startDate, days, holidays)

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
    Yearfrac(startDate, endDate, basis)

  // Decoder functions for FoldRange

  /**
   * Decode cell as numeric value (Double or BigDecimal).
   *
   * Handles Formula cells by extracting the cached numeric value when available. This enables
   * nested formula evaluation where a cell reference points to another formula cell with a cached
   * result.
   */
  def decodeNumeric(cell: Cell): Either[CodecError, BigDecimal] =
    import com.tjclp.xl.cells.CellValue
    cell.value match
      case CellValue.Number(value) => scala.util.Right(value)
      case CellValue.Formula(_, Some(CellValue.Number(cached))) =>
        // Extract cached numeric value from formula cell
        scala.util.Right(cached)
      case other =>
        scala.util.Left(
          CodecError.TypeMismatch(
            expected = "Numeric",
            actual = other
          )
        )

  /**
   * Decode cell as any value (for COUNT, etc).
   */
  def decodeAny(cell: Cell): Either[CodecError, Option[Any]] =
    import com.tjclp.xl.cells.CellValue
    cell.value match
      case CellValue.Empty => scala.util.Right(None)
      case other => scala.util.Right(Some(other))

  /**
   * Decode cell as text/string value.
   *
   * Handles Formula cells by extracting the cached text value when available.
   */
  def decodeString(cell: Cell): Either[CodecError, String] =
    import com.tjclp.xl.cells.CellValue
    cell.value match
      case CellValue.Text(value) => scala.util.Right(value)
      case CellValue.Formula(_, Some(CellValue.Text(cached))) =>
        scala.util.Right(cached)
      case other =>
        scala.util.Left(
          CodecError.TypeMismatch(
            expected = "Text",
            actual = other
          )
        )

  /**
   * Decode cell as integer value.
   *
   * Handles Formula cells by extracting the cached integer value when available.
   */
  def decodeInt(cell: Cell): Either[CodecError, Int] =
    import com.tjclp.xl.cells.CellValue
    def tryInt(value: BigDecimal, orig: CellValue): Either[CodecError, Int] =
      if value.isValidInt then scala.util.Right(value.toInt)
      else
        scala.util.Left(
          CodecError.TypeMismatch(
            expected = "Int",
            actual = orig
          )
        )
    cell.value match
      case CellValue.Number(value) => tryInt(value, CellValue.Number(value))
      case CellValue.Formula(_, Some(CellValue.Number(cached))) =>
        tryInt(cached, CellValue.Number(cached))
      case other =>
        scala.util.Left(
          CodecError.TypeMismatch(
            expected = "Int",
            actual = other
          )
        )

  /**
   * Decode cell as LocalDate value (extracts date from DateTime).
   *
   * Handles Formula cells by extracting the cached DateTime value when available.
   */
  def decodeDate(cell: Cell): Either[CodecError, java.time.LocalDate] =
    import com.tjclp.xl.cells.CellValue
    cell.value match
      case CellValue.DateTime(value) => scala.util.Right(value.toLocalDate)
      case CellValue.Formula(_, Some(CellValue.DateTime(cached))) =>
        scala.util.Right(cached.toLocalDate)
      case other =>
        scala.util.Left(
          CodecError.TypeMismatch(
            expected = "Date",
            actual = other
          )
        )

  /**
   * Decode cell as LocalDateTime value.
   *
   * Handles Formula cells by extracting the cached DateTime value when available.
   */
  def decodeDateTime(cell: Cell): Either[CodecError, java.time.LocalDateTime] =
    import com.tjclp.xl.cells.CellValue
    cell.value match
      case CellValue.DateTime(value) => scala.util.Right(value)
      case CellValue.Formula(_, Some(CellValue.DateTime(cached))) =>
        scala.util.Right(cached)
      case other =>
        scala.util.Left(
          CodecError.TypeMismatch(
            expected = "DateTime",
            actual = other
          )
        )

  /**
   * Decode cell as Boolean value.
   *
   * Handles Formula cells by extracting the cached Boolean value when available.
   */
  def decodeBool(cell: Cell): Either[CodecError, Boolean] =
    import com.tjclp.xl.cells.CellValue
    cell.value match
      case CellValue.Bool(value) => scala.util.Right(value)
      case CellValue.Formula(_, Some(CellValue.Bool(cached))) =>
        scala.util.Right(cached)
      case other =>
        scala.util.Left(
          CodecError.TypeMismatch(
            expected = "Boolean",
            actual = other
          )
        )

  /**
   * Decode cell as CellValue (always succeeds).
   *
   * Used for IFERROR/ISERROR which need to preserve the raw cell value.
   */
  def decodeCellValue(cell: Cell): Either[CodecError, CellValue] =
    scala.util.Right(cell.value)

  /**
   * Decode cell as resolved CellValue (extracts cached values, converts empty to 0).
   *
   * Used for standalone cell references (e.g., =A1, =Sheet1!B2) where the formula returns the
   * cell's "effective" value:
   *   - Number, Text, Bool, DateTime, RichText → returned as-is
   *   - Formula → returns cached value if present, or Number(0) if no cache
   *   - Empty → returns Number(0) (Excel treats empty as 0 in numeric contexts)
   *   - Error → returns the error
   *
   * This matches Excel semantics for standalone cell references.
   */
  def decodeResolvedValue(cell: Cell): Either[CodecError, CellValue] =
    import com.tjclp.xl.cells.CellValue
    val resolved = cell.value match
      case CellValue.Number(n) => CellValue.Number(n)
      case CellValue.Text(s) => CellValue.Text(s)
      case CellValue.Bool(b) => CellValue.Bool(b)
      case CellValue.DateTime(dt) => CellValue.DateTime(dt)
      case CellValue.RichText(rt) => CellValue.Text(rt.toPlainText)
      case CellValue.Formula(_, cached) =>
        cached match
          case Some(CellValue.Number(n)) => CellValue.Number(n)
          case Some(CellValue.Text(s)) => CellValue.Text(s)
          case Some(CellValue.Bool(b)) => CellValue.Bool(b)
          case Some(CellValue.DateTime(dt)) => CellValue.DateTime(dt)
          case Some(CellValue.RichText(rt)) => CellValue.Text(rt.toPlainText)
          case _ => CellValue.Number(BigDecimal(0))
      case CellValue.Error(err) => CellValue.Error(err)
      case CellValue.Empty => CellValue.Number(BigDecimal(0))
    scala.util.Right(resolved)

  // ===== Type-Coercing Decoders (Excel-compatible automatic conversion) =====

  /**
   * Decode cell as String with automatic type coercion.
   *
   * Matches Excel semantics:
   *   - Text → as-is
   *   - Number → toString (42 → "42")
   *   - Boolean → toString (true → "TRUE", false → "FALSE")
   *   - DateTime → ISO format
   *   - Formula → text representation
   *   - Empty → empty string
   */
  def decodeAsString(cell: Cell): Either[CodecError, String] =
    import com.tjclp.xl.cells.CellValue
    cell.value match
      case CellValue.Empty => scala.util.Right("")
      case CellValue.Text(s) => scala.util.Right(s)
      case CellValue.Number(n) => scala.util.Right(n.toString)
      case CellValue.Bool(b) => scala.util.Right(if b then "TRUE" else "FALSE")
      case CellValue.DateTime(dt) => scala.util.Right(dt.toString)
      case CellValue.Formula(text, _) => scala.util.Right(text)
      case CellValue.RichText(rt) => scala.util.Right(rt.toPlainText)
      case other => scala.util.Left(CodecError.TypeMismatch("String", other))

  /**
   * Decode cell as LocalDate with automatic type coercion.
   *
   * Matches Excel semantics:
   *   - DateTime → extract date component
   *   - Number → interpret as Excel serial number (not yet implemented)
   *   - Text → parse as ISO date (not yet implemented)
   *   - Other → error
   */
  def decodeAsDate(cell: Cell): Either[CodecError, java.time.LocalDate] =
    import com.tjclp.xl.cells.CellValue
    cell.value match
      case CellValue.DateTime(dt) => scala.util.Right(dt.toLocalDate)
      case other =>
        // Future: could add Excel serial number → date conversion for Number
        // Future: could add ISO date string parsing for Text
        scala.util.Left(
          CodecError.TypeMismatch(
            expected = "Date",
            actual = other
          )
        )

  /**
   * Decode cell as Int with automatic type coercion.
   *
   * Matches Excel semantics:
   *   - Number → toInt if valid
   *   - Boolean → 1 for TRUE, 0 for FALSE
   *   - Text → parse as number (not yet implemented)
   *   - Other → error
   */
  def decodeAsInt(cell: Cell): Either[CodecError, Int] =
    import com.tjclp.xl.cells.CellValue
    cell.value match
      case CellValue.Number(n) if n.isValidInt => scala.util.Right(n.toInt)
      case CellValue.Bool(b) => scala.util.Right(if b then 1 else 0)
      case CellValue.Number(n) =>
        scala.util.Left(
          CodecError.TypeMismatch(
            expected = "Int",
            actual = CellValue.Number(n)
          )
        )
      case other =>
        scala.util.Left(
          CodecError.TypeMismatch(
            expected = "Int",
            actual = other
          )
        )

  // ===== PolyRef Conversion Helpers =====

  /**
   * Convert any TExpr to String type with coercion.
   *
   * Used by text functions (LEFT, RIGHT, UPPER, etc.) to handle PolyRef arguments.
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  def asStringExpr(expr: TExpr[?]): TExpr[String] = expr match
    case PolyRef(at, anchor) => Ref(at, anchor, decodeAsString)
    case SheetPolyRef(sheet, at, anchor) => SheetRef(sheet, at, anchor, decodeAsString)
    case other => other.asInstanceOf[TExpr[String]] // Safe: non-PolyRef already has correct type

  /**
   * Convert any TExpr to LocalDate type with coercion.
   *
   * Used by date functions (YEAR, MONTH, DAY) to handle PolyRef arguments.
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  def asDateExpr(expr: TExpr[?]): TExpr[java.time.LocalDate] = expr match
    case PolyRef(at, anchor) => Ref(at, anchor, decodeAsDate)
    case SheetPolyRef(sheet, at, anchor) => SheetRef(sheet, at, anchor, decodeAsDate)
    case other =>
      other.asInstanceOf[TExpr[java.time.LocalDate]] // Safe: non-PolyRef already has correct type

  /**
   * Convert any TExpr to Int type with coercion.
   *
   * Used by functions requiring integer arguments (LEFT, RIGHT, DATE). Automatically converts
   * BigDecimal expressions (like YEAR/MONTH/DAY) to Int.
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  def asIntExpr(expr: TExpr[?]): TExpr[Int] = expr match
    case PolyRef(at, anchor) => Ref(at, anchor, decodeAsInt)
    case SheetPolyRef(sheet, at, anchor) => SheetRef(sheet, at, anchor, decodeAsInt)
    case TExpr.Lit(bd: BigDecimal) if bd.isValidInt => TExpr.Lit(bd.toInt)
    // Convert BigDecimal expressions to Int (YEAR/MONTH/DAY/LEN return BigDecimal)
    case year: TExpr.Year => ToInt(year)
    case month: TExpr.Month => ToInt(month)
    case day: TExpr.Day => ToInt(day)
    case len: TExpr.Len => ToInt(len)
    case other => other.asInstanceOf[TExpr[Int]] // Safe: non-PolyRef already has correct type

  /**
   * Convert any TExpr to BigDecimal type (numeric).
   *
   * Used by arithmetic functions to handle PolyRef arguments.
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  def asNumericExpr(expr: TExpr[?]): TExpr[BigDecimal] = expr match
    case PolyRef(at, anchor) => Ref(at, anchor, decodeNumeric)
    case SheetPolyRef(sheet, at, anchor) => SheetRef(sheet, at, anchor, decodeNumeric)
    // Date functions return LocalDate/LocalDateTime - convert to Excel serial number
    case today: Today => DateToSerial(today)
    case date: Date => DateToSerial(date)
    case edate: Edate => DateToSerial(edate)
    case eomonth: Eomonth => DateToSerial(eomonth)
    case workday: Workday => DateToSerial(workday)
    case now: Now => DateTimeToSerial(now)
    case other =>
      other.asInstanceOf[TExpr[BigDecimal]] // Safe: non-PolyRef already has correct type

  /**
   * Convert any TExpr to Boolean type.
   *
   * Used by logical functions (AND, OR, NOT, IF) to handle PolyRef arguments.
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  def asBooleanExpr(expr: TExpr[?]): TExpr[Boolean] = expr match
    case PolyRef(at, anchor) => Ref(at, anchor, decodeBool)
    case SheetPolyRef(sheet, at, anchor) => SheetRef(sheet, at, anchor, decodeBool)
    case other => other.asInstanceOf[TExpr[Boolean]] // Safe: non-PolyRef already has correct type

  /**
   * Convert any TExpr to CellValue type.
   *
   * Used by error handling functions (IFERROR, ISERROR) to preserve raw cell values.
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  def asCellValueExpr(expr: TExpr[?]): TExpr[CellValue] = expr match
    case PolyRef(at, anchor) => Ref(at, anchor, decodeCellValue)
    case SheetPolyRef(sheet, at, anchor) => SheetRef(sheet, at, anchor, decodeCellValue)
    case other =>
      other.asInstanceOf[TExpr[CellValue]] // Safe: non-PolyRef already has correct type

  /**
   * Convert any TExpr to resolved CellValue type.
   *
   * Used for standalone cell references (e.g., =A1, =Sheet1!B2) where we need the cell's
   * "effective" value with cached formula results extracted and empty cells converted to 0.
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  def asResolvedValueExpr(expr: TExpr[?]): TExpr[CellValue] = expr match
    case PolyRef(at, anchor) => Ref(at, anchor, decodeResolvedValue)
    case SheetPolyRef(sheet, at, anchor) => SheetRef(sheet, at, anchor, decodeResolvedValue)
    case other =>
      other.asInstanceOf[TExpr[CellValue]] // Safe: non-PolyRef already has correct type

  // ===== Date Function Detection =====
  // Used by CLI to auto-apply date formatting when writing formulas

  /**
   * Check if expression contains any date-returning functions.
   *
   * Used to determine if a formula result should be formatted as a date in Excel. Recursively
   * checks compound expressions (arithmetic, conditionals, etc.).
   */
  def containsDateFunction(expr: TExpr[?]): Boolean = expr match
    // Date-returning functions
    case _: Today | _: Now | _: Date | _: Eomonth | _: Edate | _: Workday => true
    // Date-to-serial wrappers (for arithmetic)
    case DateToSerial(_) | DateTimeToSerial(_) => true
    // Arithmetic - recursively check operands
    case Add(l, r) => containsDateFunction(l) || containsDateFunction(r)
    case Sub(l, r) => containsDateFunction(l) || containsDateFunction(r)
    case Mul(l, r) => containsDateFunction(l) || containsDateFunction(r)
    case Div(l, r) => containsDateFunction(l) || containsDateFunction(r)
    // Conditionals
    case If(c, t, e) =>
      containsDateFunction(c) || containsDateFunction(t) || containsDateFunction(e)
    // Boolean operators
    case And(l, r) => containsDateFunction(l) || containsDateFunction(r)
    case Or(l, r) => containsDateFunction(l) || containsDateFunction(r)
    case Not(x) => containsDateFunction(x)
    // Comparisons
    case Eq(l, r) => containsDateFunction(l) || containsDateFunction(r)
    case Neq(l, r) => containsDateFunction(l) || containsDateFunction(r)
    case Lt(l, r) => containsDateFunction(l) || containsDateFunction(r)
    case Lte(l, r) => containsDateFunction(l) || containsDateFunction(r)
    case Gt(l, r) => containsDateFunction(l) || containsDateFunction(r)
    case Gte(l, r) => containsDateFunction(l) || containsDateFunction(r)
    // Error handling
    case Iferror(v, e) => containsDateFunction(v) || containsDateFunction(e)
    case Iserror(v) => containsDateFunction(v)
    case Iserr(v) => containsDateFunction(v)
    case Isnumber(v) => containsDateFunction(v)
    case Istext(v) => containsDateFunction(v)
    case Isblank(v) => containsDateFunction(v)
    // Rounding functions
    case Round(v, _) => containsDateFunction(v)
    case RoundUp(v, _) => containsDateFunction(v)
    case RoundDown(v, _) => containsDateFunction(v)
    case Abs(v) => containsDateFunction(v)
    // Math functions
    case Sqrt(v) => containsDateFunction(v)
    case Mod(n, _) => containsDateFunction(n)
    case Power(n, _) => containsDateFunction(n)
    case Log(n, _) => containsDateFunction(n)
    case Ln(v) => containsDateFunction(v)
    case Exp(v) => containsDateFunction(v)
    case Floor(n, _) => containsDateFunction(n)
    case Ceiling(n, _) => containsDateFunction(n)
    case Trunc(n, _) => containsDateFunction(n)
    case Sign(v) => containsDateFunction(v)
    case Int_(v) => containsDateFunction(v)
    // Type conversion
    case ToInt(e) => containsDateFunction(e)
    // Default: no date function
    case _ => false

  /**
   * Check if expression contains time-returning functions (NOW).
   *
   * Used to distinguish between Date format (m/d/yy) and DateTime format (m/d/yy h:mm). If true,
   * use DateTime format; otherwise use Date format.
   */
  def containsTimeFunction(expr: TExpr[?]): Boolean = expr match
    // Time-returning functions
    case _: Now => true
    case DateTimeToSerial(_) => true
    // Arithmetic - recursively check operands
    case Add(l, r) => containsTimeFunction(l) || containsTimeFunction(r)
    case Sub(l, r) => containsTimeFunction(l) || containsTimeFunction(r)
    case Mul(l, r) => containsTimeFunction(l) || containsTimeFunction(r)
    case Div(l, r) => containsTimeFunction(l) || containsTimeFunction(r)
    // Conditionals
    case If(c, t, e) =>
      containsTimeFunction(c) || containsTimeFunction(t) || containsTimeFunction(e)
    // Boolean operators
    case And(l, r) => containsTimeFunction(l) || containsTimeFunction(r)
    case Or(l, r) => containsTimeFunction(l) || containsTimeFunction(r)
    case Not(x) => containsTimeFunction(x)
    // Comparisons
    case Eq(l, r) => containsTimeFunction(l) || containsTimeFunction(r)
    case Neq(l, r) => containsTimeFunction(l) || containsTimeFunction(r)
    case Lt(l, r) => containsTimeFunction(l) || containsTimeFunction(r)
    case Lte(l, r) => containsTimeFunction(l) || containsTimeFunction(r)
    case Gt(l, r) => containsTimeFunction(l) || containsTimeFunction(r)
    case Gte(l, r) => containsTimeFunction(l) || containsTimeFunction(r)
    // Error handling
    case Iferror(v, e) => containsTimeFunction(v) || containsTimeFunction(e)
    case Iserror(v) => containsTimeFunction(v)
    case Iserr(v) => containsTimeFunction(v)
    case Isnumber(v) => containsTimeFunction(v)
    case Istext(v) => containsTimeFunction(v)
    case Isblank(v) => containsTimeFunction(v)
    // Rounding functions
    case Round(v, _) => containsTimeFunction(v)
    case RoundUp(v, _) => containsTimeFunction(v)
    case RoundDown(v, _) => containsTimeFunction(v)
    case Abs(v) => containsTimeFunction(v)
    // Math functions
    case Sqrt(v) => containsTimeFunction(v)
    case Mod(n, _) => containsTimeFunction(n)
    case Power(n, _) => containsTimeFunction(n)
    case Log(n, _) => containsTimeFunction(n)
    case Ln(v) => containsTimeFunction(v)
    case Exp(v) => containsTimeFunction(v)
    case Floor(n, _) => containsTimeFunction(n)
    case Ceiling(n, _) => containsTimeFunction(n)
    case Trunc(n, _) => containsTimeFunction(n)
    case Sign(v) => containsTimeFunction(v)
    case Int_(v) => containsTimeFunction(v)
    // Type conversion
    case ToInt(e) => containsTimeFunction(e)
    // Default: no time function
    case _ => false

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
    def &&(y: TExpr[Boolean]): TExpr[Boolean] = And(x, y)
    def ||(y: TExpr[Boolean]): TExpr[Boolean] = Or(x, y)
    def unary_! : TExpr[Boolean] = Not(x)

  extension [A](x: TExpr[A])
    /**
     * Equality/inequality operators.
     *
     * Example: expr1 === expr2
     */
    def ===(y: TExpr[A]): TExpr[Boolean] = Eq(x, y)
    def !==(y: TExpr[A]): TExpr[Boolean] = Neq(x, y)
