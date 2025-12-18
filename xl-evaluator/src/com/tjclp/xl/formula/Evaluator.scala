package com.tjclp.xl.formula

import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.{Cell, CellError, CellValue}
import com.tjclp.xl.workbooks.Workbook
import com.tjclp.xl.SheetName
import com.tjclp.xl.syntax.* // Extension methods for Sheet.get, CellRange.cells, ARef.toA1
import scala.math.BigDecimal
import scala.annotation.tailrec
import java.time.LocalDate
import java.time.temporal.ChronoUnit

/**
 * Pure functional formula evaluator.
 *
 * Evaluates TExpr AST against a Sheet, returning either an error or the computed value. All
 * evaluation is total - no exceptions thrown, no side effects.
 *
 * Laws satisfied:
 *   1. Literal identity: eval(Lit(x)) == Right(x)
 *   2. Arithmetic laws: eval(Add(Lit(a), Lit(b))) == Right(a + b)
 *   3. Short-circuit: And(Lit(false), error) == Right(false) (no error raised)
 *   4. Totality: eval always returns Either[EvalError, A] (never throws)
 *
 * Example:
 * {{{
 * val expr = TExpr.Add(TExpr.Lit(BigDecimal(10)), TExpr.Ref(ref"A1", TExpr.decodeNumeric))
 * val evaluator = Evaluator.instance
 * evaluator.eval(expr, sheet) match
 *   case Right(result) => println(s"Result: $$result")
 *   case Left(error) => println(s"Error: $$error")
 * }}}
 */
trait Evaluator:
  /**
   * Evaluate expression against sheet.
   *
   * @param expr
   *   The expression to evaluate
   * @param sheet
   *   The sheet providing cell values
   * @param clock
   *   Clock for date/time functions (defaults to system clock)
   * @param workbook
   *   Optional workbook for cross-sheet references (defaults to None)
   * @return
   *   Either evaluation error or computed value
   */
  def eval[A](
    expr: TExpr[A],
    sheet: Sheet,
    clock: Clock = Clock.system,
    workbook: Option[Workbook] = None
  ): Either[EvalError, A]

object Evaluator:
  /**
   * Default evaluator instance.
   *
   * Pure functional implementation with short-circuit evaluation for And/Or.
   */
  def instance: Evaluator = new EvaluatorImpl

  /**
   * Convenience method for direct evaluation (forwards to instance.eval).
   */
  def eval[A](
    expr: TExpr[A],
    sheet: Sheet,
    clock: Clock = Clock.system,
    workbook: Option[Workbook] = None
  ): Either[EvalError, A] =
    instance.eval(expr, sheet, clock, workbook)

  // Helper methods for consistent cross-sheet error messages
  private[formula] def missingWorkbookError(refStr: String, isRange: Boolean = false): EvalError =
    val refType = if isRange then "range" else "reference"
    EvalError.EvalFailed(
      s"Cross-sheet $refType $refStr requires workbook context, but none was provided.",
      None
    )

  private[formula] def sheetNotFoundError(
    sheetName: SheetName,
    err: com.tjclp.xl.error.XLError
  ): EvalError =
    EvalError.EvalFailed(
      s"Sheet '${sheetName.value}' not found in workbook: ${err.message}",
      None
    )

/**
 * Private implementation of Evaluator.
 *
 * Implements all TExpr cases with proper error handling and short-circuit semantics.
 */
private class EvaluatorImpl extends Evaluator:
  // Suppress asInstanceOf warning for FoldRange GADT type handling (required for type parameter erasure)
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  def eval[A](
    expr: TExpr[A],
    sheet: Sheet,
    clock: Clock = Clock.system,
    workbook: Option[Workbook] = None
  ): Either[EvalError, A] =
    // @unchecked: GADT exhaustivity - PolyRef should be resolved before evaluation
    (expr: @unchecked) match
      // ===== Defensive PolyRef Handling =====
      // PolyRef should be converted to typed Ref during parsing, but if it reaches
      // evaluation unconverted, provide a clear error instead of MatchError
      case TExpr.PolyRef(at, _) =>
        // Cast to ARef to workaround Scala 3 inline method issue with pattern-matched values
        val refStr = (at: ARef).toA1
        Left(
          EvalError.EvalFailed(
            s"Unresolved cell reference $refStr. This indicates a parser bug - PolyRef should be resolved before evaluation.",
            None
          )
        )

      // ===== Sheet-Qualified References (Cross-Sheet) =====
      //
      // SheetPolyRef: Cross-sheet polymorphic reference
      //
      // Unlike same-sheet PolyRef (which errors defensively because it should be
      // resolved to typed Ref during parsing), SheetPolyRef is evaluated directly
      // because cross-sheet references are typically used at the top level of
      // formulas (e.g., =Sales!A1) where type context is "return the raw value".
      //
      // The asInstanceOf[Either[EvalError, A]] cast at line ~150 is safe because:
      // 1. The calling code (SheetEvaluator.evaluateFormula) passes result to
      //    toCellValue(result: Any) which handles all runtime types
      // 2. When used in typed contexts (e.g., =Sales!A1 + 10), the parser would
      //    create SheetRef with proper decoder instead of SheetPolyRef
      //
      // This matches Excel behavior: cross-sheet references return the cell's
      // raw value, with type coercion happening at the operator level.
      //
      case TExpr.SheetPolyRef(sheetName, at, _) =>
        workbook match
          case None =>
            val refStr = s"${sheetName.value}!${(at: ARef).toA1}"
            Left(Evaluator.missingWorkbookError(refStr))
          case Some(wb) =>
            wb(sheetName) match
              case Left(err) =>
                Left(Evaluator.sheetNotFoundError(sheetName, err))
              case Right(targetSheet) =>
                val cell = targetSheet(at)
                // Return the cell value directly - extract raw value from CellValue
                // We cast the whole Either to avoid type issues with TExpr[Nothing]
                val result: Either[EvalError, Any] = cell.value match
                  case CellValue.Number(n) => Right(n)
                  case CellValue.Text(s) => Right(s)
                  case CellValue.Bool(b) => Right(b)
                  case CellValue.DateTime(dt) => Right(dt)
                  case CellValue.Formula(_, cached) =>
                    // Use cached value if available
                    cached match
                      case Some(CellValue.Number(n)) => Right(n)
                      case Some(CellValue.Text(s)) => Right(s)
                      case Some(CellValue.Bool(b)) => Right(b)
                      case Some(CellValue.DateTime(dt)) => Right(dt)
                      case Some(CellValue.RichText(rt)) => Right(rt.toPlainText)
                      case _ => Right(BigDecimal(0)) // Empty formula = 0
                  case CellValue.Error(err) =>
                    Left(EvalError.EvalFailed(s"Cell contains error: $err", None))
                  case CellValue.Empty => Right(BigDecimal(0)) // Empty = 0
                  case CellValue.RichText(rt) => Right(rt.toPlainText)
                result.asInstanceOf[Either[EvalError, A]]

      case TExpr.SheetRef(sheetName, at, _, decode) =>
        // SheetRef: resolve cell from target sheet in workbook
        workbook match
          case None =>
            val refStr = s"${sheetName.value}!${(at: ARef).toA1}"
            Left(Evaluator.missingWorkbookError(refStr))
          case Some(wb) =>
            wb(sheetName) match
              case Left(err) =>
                Left(Evaluator.sheetNotFoundError(sheetName, err))
              case Right(targetSheet) =>
                val cell = targetSheet(at)
                decode(cell).left.map(codecErr => EvalError.CodecFailed(at, codecErr))

      case TExpr.SheetRange(sheetName, range) =>
        // SheetRange should be wrapped in a function (SUM, COUNT, etc.) before evaluation
        val refStr = s"${sheetName.value}!${range.toA1}"
        Left(
          EvalError.EvalFailed(
            s"Cross-sheet range $refStr must be used within a function like SUM or COUNT.",
            None
          )
        )

      // ===== Literals =====
      case TExpr.Lit(value) =>
        // Literal: return value directly (identity law)
        Right(value)

      // ===== Cell References =====
      case TExpr.Ref(at, _, decode) =>
        // Ref: resolve cell, decode value with codec
        // Note: sheet(at) returns empty cell if not present, decode handles empty cells
        val cell = sheet(at)
        decode(cell).left.map(codecErr => EvalError.CodecFailed(at, codecErr))

      // ===== Conditional =====
      case TExpr.If(cond, ifTrue, ifFalse) =>
        // If: evaluate condition, then branch based on result
        eval(cond, sheet, clock, workbook).flatMap { condValue =>
          if condValue then eval(ifTrue, sheet, clock, workbook)
          else eval(ifFalse, sheet, clock, workbook)
        }

      // ===== Arithmetic Operators =====
      case TExpr.Add(x, y) =>
        // Add: evaluate both operands, sum results
        for
          xv <- eval(x, sheet, clock, workbook)
          yv <- eval(y, sheet, clock, workbook)
        yield xv + yv

      case TExpr.Sub(x, y) =>
        // Subtract: evaluate both operands, subtract second from first
        for
          xv <- eval(x, sheet, clock, workbook)
          yv <- eval(y, sheet, clock, workbook)
        yield xv - yv

      case TExpr.Mul(x, y) =>
        // Multiply: evaluate both operands, multiply results
        for
          xv <- eval(x, sheet, clock, workbook)
          yv <- eval(y, sheet, clock, workbook)
        yield xv * yv

      case TExpr.Div(x, y) =>
        // Divide: evaluate both operands, check for division by zero
        for
          xv <- eval(x, sheet, clock, workbook)
          yv <- eval(y, sheet, clock, workbook)
          result <-
            if yv == BigDecimal(0) then
              // Division by zero: provide helpful error message with expressions
              Left(
                EvalError.DivByZero(
                  FormulaPrinter.print(x, includeEquals = false),
                  FormulaPrinter.print(y, includeEquals = false)
                )
              )
            else Right(xv / yv)
        yield result

      // ===== Logical Operators =====
      case TExpr.And(x, y) =>
        // And: short-circuit evaluation (if x is false, don't evaluate y)
        eval(x, sheet, clock, workbook).flatMap {
          case false =>
            // Short-circuit: x is false, result is false (don't evaluate y)
            Right(false)
          case true =>
            // x is true, evaluate y to determine final result
            eval(y, sheet, clock, workbook)
        }

      case TExpr.Or(x, y) =>
        // Or: short-circuit evaluation (if x is true, don't evaluate y)
        eval(x, sheet, clock, workbook).flatMap {
          case true =>
            // Short-circuit: x is true, result is true (don't evaluate y)
            Right(true)
          case false =>
            // x is false, evaluate y to determine final result
            eval(y, sheet, clock, workbook)
        }

      case TExpr.Not(x) =>
        // Not: logical negation
        eval(x, sheet, clock, workbook).map(xv => !xv)

      // ===== Comparison Operators =====
      case TExpr.Lt(x, y) =>
        // Less than: numeric comparison
        for
          xv <- eval(x, sheet, clock, workbook)
          yv <- eval(y, sheet, clock, workbook)
        yield xv < yv

      case TExpr.Lte(x, y) =>
        // Less than or equal: numeric comparison
        for
          xv <- eval(x, sheet, clock, workbook)
          yv <- eval(y, sheet, clock, workbook)
        yield xv <= yv

      case TExpr.Gt(x, y) =>
        // Greater than: numeric comparison
        for
          xv <- eval(x, sheet, clock, workbook)
          yv <- eval(y, sheet, clock, workbook)
        yield xv > yv

      case TExpr.Gte(x, y) =>
        // Greater than or equal: numeric comparison
        for
          xv <- eval(x, sheet, clock, workbook)
          yv <- eval(y, sheet, clock, workbook)
        yield xv >= yv

      case TExpr.Eq(x, y) =>
        // Equality: polymorphic comparison
        for
          xv <- eval(x, sheet, clock, workbook)
          yv <- eval(y, sheet, clock, workbook)
        yield xv == yv

      case TExpr.Neq(x, y) =>
        // Inequality: polymorphic comparison
        for
          xv <- eval(x, sheet, clock, workbook)
          yv <- eval(y, sheet, clock, workbook)
        yield xv != yv

      // ===== Type Conversions =====
      case TExpr.ToInt(expr) =>
        // ToInt: Convert BigDecimal to Int (validates integer range)
        eval(expr, sheet, clock, workbook).flatMap { bd =>
          if bd.isValidInt then Right(bd.toInt)
          else
            Left(
              EvalError.TypeMismatch(
                "ToInt",
                "valid integer",
                s"$bd (out of Int range)"
              )
            )
        }

      // ===== Text Functions =====
      case TExpr.Concatenate(xs) =>
        // Concatenate: evaluate all expressions, concat results
        xs.foldLeft[Either[EvalError, String]](Right("")) { (accEither, expr) =>
          for
            acc <- accEither
            value <- eval(expr, sheet, clock, workbook)
          yield acc + value
        }

      case TExpr.Left(text, n) =>
        // Left: extract left n characters
        for
          textValue <- eval(text, sheet, clock, workbook)
          nValue <- eval(n, sheet, clock, workbook)
          result <-
            if nValue < 0 then
              Left(EvalError.EvalFailed(s"LEFT: n must be non-negative, got $nValue"))
            else if nValue >= textValue.length then Right(textValue)
            else Right(textValue.take(nValue))
        yield result

      case TExpr.Right(text, n) =>
        // Right: extract right n characters
        for
          textValue <- eval(text, sheet, clock, workbook)
          nValue <- eval(n, sheet, clock, workbook)
          result <-
            if nValue < 0 then
              Left(EvalError.EvalFailed(s"RIGHT: n must be non-negative, got $nValue"))
            else if nValue >= textValue.length then Right(textValue)
            else Right(textValue.takeRight(nValue))
        yield result

      case TExpr.Len(text) =>
        // Len: text length (returns BigDecimal to match Excel and enable arithmetic)
        eval(text, sheet, clock, workbook).map(s => BigDecimal(s.length))

      case TExpr.Upper(text) =>
        // Upper: convert to uppercase
        eval(text, sheet, clock, workbook).map(_.toUpperCase)

      case TExpr.Lower(text) =>
        // Lower: convert to lowercase
        eval(text, sheet, clock, workbook).map(_.toLowerCase)

      // ===== Date/Time Functions =====
      case TExpr.Today() =>
        // Today: get current date from clock
        Right(clock.today())

      case TExpr.Now() =>
        // Now: get current date and time from clock
        Right(clock.now())

      case TExpr.Date(year, month, day) =>
        // Date: construct date from components
        for
          y <- eval(year, sheet, clock, workbook)
          m <- eval(month, sheet, clock, workbook)
          d <- eval(day, sheet, clock, workbook)
          result <-
            scala.util.Try(java.time.LocalDate.of(y, m, d)).toEither.left.map { ex =>
              EvalError.EvalFailed(
                s"DATE: invalid date components (year=$y, month=$m, day=$d): ${ex.getMessage}"
              )
            }
        yield result

      case TExpr.Year(date) =>
        // Year: extract year from date (returns BigDecimal to match Excel)
        eval(date, sheet, clock, workbook).map(d => BigDecimal(d.getYear))

      case TExpr.Month(date) =>
        // Month: extract month from date (returns BigDecimal to match Excel)
        eval(date, sheet, clock, workbook).map(d => BigDecimal(d.getMonthValue))

      case TExpr.Day(date) =>
        // Day: extract day from date (returns BigDecimal to match Excel)
        eval(date, sheet, clock, workbook).map(d => BigDecimal(d.getDayOfMonth))

      // ===== Date Calculation Functions =====

      case TExpr.Eomonth(startDate, months) =>
        // EOMONTH: end of month N months from start
        // Note: months may come as Int or BigDecimal depending on parsing path
        for
          date <- eval(startDate, sheet, clock, workbook)
          monthsRaw <- evalAny(months, sheet, clock, workbook)
        yield
          val monthsValue = toInt(monthsRaw)
          val targetMonth = date.plusMonths(monthsValue.toLong)
          targetMonth.withDayOfMonth(targetMonth.lengthOfMonth)

      case TExpr.Edate(startDate, months) =>
        // EDATE: same day N months later (clamped to end of month if needed)
        // Note: months may come as Int or BigDecimal depending on parsing path
        for
          date <- eval(startDate, sheet, clock, workbook)
          monthsRaw <- evalAny(months, sheet, clock, workbook)
        yield date.plusMonths(toInt(monthsRaw).toLong)

      case TExpr.Datedif(startDate, endDate, unit) =>
        // DATEDIF: difference between dates in specified unit
        for
          start <- eval(startDate, sheet, clock, workbook)
          end <- eval(endDate, sheet, clock, workbook)
          unitStr <- eval(unit, sheet, clock, workbook)
          result <- unitStr.toUpperCase match
            case "Y" =>
              // Years between dates
              Right(BigDecimal(ChronoUnit.YEARS.between(start, end)))
            case "M" =>
              // Months between dates
              Right(BigDecimal(ChronoUnit.MONTHS.between(start, end)))
            case "D" =>
              // Days between dates
              Right(BigDecimal(ChronoUnit.DAYS.between(start, end)))
            case "MD" =>
              // Days ignoring months and years
              // Excel's MD: days between as if both dates were in the same month
              val daysDiff = end.getDayOfMonth - start.getDayOfMonth
              val adjustedDays =
                if daysDiff >= 0 then daysDiff
                else
                  // Need to "borrow" from previous month
                  val prevMonthLength = end.minusMonths(1).lengthOfMonth
                  // If start day exceeds prev month length, treat as end of prev month
                  val effectiveStartDay = math.min(start.getDayOfMonth, prevMonthLength)
                  prevMonthLength - effectiveStartDay + end.getDayOfMonth
              Right(BigDecimal(adjustedDays))
            case "YM" =>
              // Months ignoring years
              val monthsDiff = end.getMonthValue - start.getMonthValue
              val adjustedMonths = if monthsDiff < 0 then 12 + monthsDiff else monthsDiff
              val finalMonths =
                if end.getDayOfMonth < start.getDayOfMonth then (adjustedMonths - 1 + 12) % 12
                else adjustedMonths
              Right(BigDecimal(finalMonths))
            case "YD" =>
              // Days ignoring years
              val startAdjusted = start.withYear(end.getYear)
              val days =
                if startAdjusted.isAfter(end) then
                  ChronoUnit.DAYS.between(startAdjusted.minusYears(1), end)
                else ChronoUnit.DAYS.between(startAdjusted, end)
              Right(BigDecimal(days))
            case other =>
              Left(
                EvalError.EvalFailed(
                  s"DATEDIF: invalid unit '$other'. Valid units: Y, M, D, MD, YM, YD",
                  Some("DATEDIF(start, end, unit)")
                )
              )
        yield result

      case TExpr.Networkdays(startDate, endDate, holidaysOpt) =>
        // NETWORKDAYS: count working days (Mon-Fri) between dates, excluding holidays
        for
          start <- eval(startDate, sheet, clock, workbook)
          end <- eval(endDate, sheet, clock, workbook)
        yield
          // Collect holiday dates from range if provided
          val holidays: Set[LocalDate] = holidaysOpt
            .map { range =>
              range.cells
                .flatMap(ref => TExpr.decodeDate(sheet(ref)).toOption)
                .toSet
            }
            .getOrElse(Set.empty)

          // Count working days (Mon-Fri, excluding holidays)
          val (earlier, later) = if start.isBefore(end) then (start, end) else (end, start)
          val count = countWorkingDays(earlier, later, holidays)
          // If dates were reversed, negate the count
          BigDecimal(if start.isBefore(end) || start.isEqual(end) then count else -count)

      case TExpr.Workday(startDate, days, holidaysOpt) =>
        // WORKDAY: add N working days to start date, skipping weekends and holidays
        // Note: days may come as Int or BigDecimal depending on parsing path
        for
          start <- eval(startDate, sheet, clock, workbook)
          daysRaw <- evalAny(days, sheet, clock, workbook)
        yield
          val daysValue = toInt(daysRaw)
          // Collect holiday dates from range if provided
          val holidays: Set[LocalDate] = holidaysOpt
            .map { range =>
              range.cells
                .flatMap(ref => TExpr.decodeDate(sheet(ref)).toOption)
                .toSet
            }
            .getOrElse(Set.empty)

          // Add working days (Mon-Fri, skipping holidays)
          addWorkingDays(start, daysValue, holidays)

      case TExpr.Yearfrac(startDate, endDate, basis) =>
        // YEARFRAC: year fraction between dates based on day count basis
        // Note: basis may come as Int or BigDecimal depending on parsing path
        for
          start <- eval(startDate, sheet, clock, workbook)
          end <- eval(endDate, sheet, clock, workbook)
          basisRaw <- evalAny(basis, sheet, clock, workbook)
          basisValue = toInt(basisRaw)
          result <- basisValue match
            case 0 =>
              // US 30/360
              val d1 = math.min(30, start.getDayOfMonth)
              val d2 = if start.getDayOfMonth == 31 then 30 else end.getDayOfMonth
              val m1 = start.getMonthValue
              val m2 = end.getMonthValue
              val y1 = start.getYear
              val y2 = end.getYear
              val days = ((y2 - y1) * 360) + ((m2 - m1) * 30) + (d2 - d1)
              Right(BigDecimal(days) / BigDecimal(360))
            case 1 =>
              // Actual/actual
              val daysBetween = ChronoUnit.DAYS.between(start, end)
              val year = start.getYear
              val daysInYear = if java.time.Year.isLeap(year) then 366 else 365
              Right(BigDecimal(daysBetween) / BigDecimal(daysInYear))
            case 2 =>
              // Actual/360
              val daysBetween = ChronoUnit.DAYS.between(start, end)
              Right(BigDecimal(daysBetween) / BigDecimal(360))
            case 3 =>
              // Actual/365
              val daysBetween = ChronoUnit.DAYS.between(start, end)
              Right(BigDecimal(daysBetween) / BigDecimal(365))
            case 4 =>
              // European 30/360
              val d1 = math.min(30, start.getDayOfMonth)
              val d2 = math.min(30, end.getDayOfMonth)
              val m1 = start.getMonthValue
              val m2 = end.getMonthValue
              val y1 = start.getYear
              val y2 = end.getYear
              val days = ((y2 - y1) * 360) + ((m2 - m1) * 30) + (d2 - d1)
              Right(BigDecimal(days) / BigDecimal(360))
            case other =>
              Left(
                EvalError.EvalFailed(
                  s"YEARFRAC: invalid basis $other. Valid values: 0-4",
                  Some("YEARFRAC(start, end, [basis])")
                )
              )
        yield result

      // ===== Financial Functions =====
      case TExpr.Npv(rateExpr, range) =>
        // Evaluate discount rate
        eval(rateExpr, sheet, clock, workbook).flatMap { rate =>
          val one = BigDecimal(1)
          val onePlusR = one + rate

          if onePlusR == BigDecimal(0) then
            Left(
              EvalError.EvalFailed(
                "NPV: rate = -1 would require division by zero",
                Some("NPV(rate, values)")
              )
            )
          else
            // Collect numeric cash flows from the range (skip non-numeric, Excel-style)
            val cashFlows: List[BigDecimal] =
              range.cells
                .map(ref => sheet(ref))
                .flatMap(cell => TExpr.decodeNumeric(cell).toOption)
                .toList

            // Excel's NPV treats the first value as period 1
            val npv =
              cashFlows.zipWithIndex.foldLeft(BigDecimal(0)) { case (acc, (cf, idx)) =>
                val period = idx + 1 // t = 1 for first cash flow
                val discount = onePlusR.pow(period)
                acc + (cf / discount)
              }

            Right(npv)
        }

      case TExpr.Irr(range, guessExprOpt) =>
        // Collect numeric cash flows from range (including t0)
        val cashFlows: List[BigDecimal] =
          range.cells
            .map(ref => sheet(ref))
            .flatMap(cell => TExpr.decodeNumeric(cell).toOption)
            .toList

        // Need at least one positive and one negative for IRR to make sense
        if cashFlows.isEmpty || !cashFlows.exists(_ < 0) || !cashFlows.exists(_ > 0) then
          Left(
            EvalError.EvalFailed(
              "IRR requires at least one positive and one negative cash flow",
              Some("IRR(values[, guess])")
            )
          )
        else
          val guessEither: Either[EvalError, BigDecimal] =
            guessExprOpt match
              case Some(guessExpr) => eval(guessExpr, sheet, clock, workbook)
              case None => Right(BigDecimal("0.1"))

          guessEither.flatMap { guess0 =>
            val maxIter = 50
            val tolerance = BigDecimal("1e-7")
            val one = BigDecimal(1)

            def npvAt(rate: BigDecimal): BigDecimal =
              val onePlusR = one + rate
              cashFlows.zipWithIndex.foldLeft(BigDecimal(0)) { case (acc, (cf, idx)) =>
                if idx == 0 then acc + cf
                else acc + cf / onePlusR.pow(idx) // t = idx for remaining flows
              }

            def dNpvAt(rate: BigDecimal): BigDecimal =
              val onePlusR = one + rate
              cashFlows.zipWithIndex.foldLeft(BigDecimal(0)) { case (acc, (cf, idx)) =>
                if idx == 0 then acc
                else acc - (idx * cf) / onePlusR.pow(idx + 1)
              }

            @tailrec
            def loop(iter: Int, r: BigDecimal): Either[EvalError, BigDecimal] =
              if iter >= maxIter then
                Left(
                  EvalError.EvalFailed(
                    s"IRR did not converge after $maxIter iterations",
                    Some("IRR(values[, guess])")
                  )
                )
              else
                val f = npvAt(r)
                val df = dNpvAt(r)
                if df == BigDecimal(0) then
                  Left(
                    EvalError.EvalFailed(
                      "IRR derivative is zero; cannot continue iteration",
                      Some("IRR(values[, guess])")
                    )
                  )
                else
                  val next = r - f / df
                  if (next - r).abs <= tolerance then Right(next)
                  else loop(iter + 1, next)

            loop(0, guess0)
          }

      case TExpr.Xnpv(rateExpr, valuesRange, datesRange) =>
        // XNPV: sum(value_i / (1 + rate)^((date_i - date_0) / 365))
        for
          rate <- eval(rateExpr, sheet, clock, workbook)
          result <- {
            // Collect values and dates from ranges
            val values: List[BigDecimal] =
              valuesRange.cells
                .map(ref => sheet(ref))
                .flatMap(cell => TExpr.decodeNumeric(cell).toOption)
                .toList

            val dates: List[LocalDate] =
              datesRange.cells
                .map(ref => sheet(ref))
                .flatMap(cell => TExpr.decodeDate(cell).toOption)
                .toList

            if values.isEmpty || dates.isEmpty then
              Left(
                EvalError.EvalFailed(
                  "XNPV requires non-empty values and dates ranges",
                  Some("XNPV(rate, values, dates)")
                )
              )
            else if values.length != dates.length then
              Left(
                EvalError.EvalFailed(
                  s"XNPV: values (${values.length}) and dates (${dates.length}) must have same length",
                  Some("XNPV(rate, values, dates)")
                )
              )
            else
              // Pattern match to extract first date (safe: non-empty check above)
              dates match
                case date0 :: _ =>
                  val onePlusR = BigDecimal(1) + rate
                  val npv = values.zip(dates).foldLeft(BigDecimal(0)) { case (acc, (value, date)) =>
                    val daysDiff = ChronoUnit.DAYS.between(date0, date)
                    val yearFraction = BigDecimal(daysDiff) / BigDecimal(365)
                    // Note: Using Double for fractional exponents matches Excel's precision behavior.
                    // BigDecimal.pow only supports integer exponents. For financial calculations,
                    // Double precision (~15 decimal digits) is sufficient and consistent with Excel.
                    val discountFactor = math.pow(onePlusR.toDouble, yearFraction.toDouble)
                    acc + value / BigDecimal(discountFactor)
                  }
                  Right(npv)
                case Nil =>
                  // Unreachable: dates verified non-empty at line 430
                  Left(EvalError.EvalFailed("XNPV: dates cannot be empty", None))
          }
        yield result

      case TExpr.Xirr(valuesRange, datesRange, guessExprOpt) =>
        // XIRR: Find rate where XNPV = 0 using Newton-Raphson
        // Collect values and dates from ranges
        val values: List[BigDecimal] =
          valuesRange.cells
            .map(ref => sheet(ref))
            .flatMap(cell => TExpr.decodeNumeric(cell).toOption)
            .toList

        val dates: List[LocalDate] =
          datesRange.cells
            .map(ref => sheet(ref))
            .flatMap(cell => TExpr.decodeDate(cell).toOption)
            .toList

        if values.isEmpty || dates.isEmpty then
          Left(
            EvalError.EvalFailed(
              "XIRR requires non-empty values and dates ranges",
              Some("XIRR(values, dates[, guess])")
            )
          )
        else if values.length != dates.length then
          Left(
            EvalError.EvalFailed(
              s"XIRR: values (${values.length}) and dates (${dates.length}) must have same length",
              Some("XIRR(values, dates[, guess])")
            )
          )
        else if !values.exists(_ < 0) || !values.exists(_ > 0) then
          Left(
            EvalError.EvalFailed(
              "XIRR requires at least one positive and one negative cash flow",
              Some("XIRR(values, dates[, guess])")
            )
          )
        else
          val guessEither: Either[EvalError, BigDecimal] =
            guessExprOpt match
              case Some(guessExpr) => eval(guessExpr, sheet, clock, workbook)
              case None => Right(BigDecimal("0.1"))

          guessEither.flatMap { guess0 =>
            val maxIter = 100
            val tolerance = BigDecimal("1e-7")
            // Pattern match to extract first date (safe: non-empty check at line 480)
            dates match
              case date0 :: _ =>
                // Calculate year fractions for each date
                val yearFractions: List[BigDecimal] = dates.map { date =>
                  val daysDiff = ChronoUnit.DAYS.between(date0, date)
                  BigDecimal(daysDiff) / BigDecimal(365)
                }

                // XNPV at given rate: sum(value_i / (1 + rate)^yearFraction_i)
                // Note: Uses Double for fractional exponents (matches Excel precision, see XNPV comment)
                def xnpvAt(rate: BigDecimal): BigDecimal =
                  val onePlusR = BigDecimal(1) + rate
                  values.zip(yearFractions).foldLeft(BigDecimal(0)) { case (acc, (cf, yf)) =>
                    val discountFactor = math.pow(onePlusR.toDouble, yf.toDouble)
                    acc + cf / BigDecimal(discountFactor)
                  }

                // Derivative of XNPV: sum(-yearFraction_i * value_i / (1 + rate)^(yearFraction_i + 1))
                def dXnpvAt(rate: BigDecimal): BigDecimal =
                  val onePlusR = BigDecimal(1) + rate
                  values.zip(yearFractions).foldLeft(BigDecimal(0)) { case (acc, (cf, yf)) =>
                    val discountFactor = math.pow(onePlusR.toDouble, (yf + 1).toDouble)
                    acc - (yf * cf) / BigDecimal(discountFactor)
                  }

                @tailrec
                def loop(iter: Int, r: BigDecimal): Either[EvalError, BigDecimal] =
                  if iter >= maxIter then
                    Left(
                      EvalError.EvalFailed(
                        s"XIRR did not converge after $maxIter iterations",
                        Some("XIRR(values, dates[, guess])")
                      )
                    )
                  else
                    val f = xnpvAt(r)
                    val df = dXnpvAt(r)
                    if df.abs < BigDecimal("1e-10") then
                      Left(
                        EvalError.EvalFailed(
                          "XIRR derivative is near zero; cannot continue iteration",
                          Some("XIRR(values, dates[, guess])")
                        )
                      )
                    else
                      val next = r - f / df
                      if (next - r).abs <= tolerance then Right(next)
                      else loop(iter + 1, next)

                loop(0, guess0)
              case Nil =>
                // Unreachable: dates verified non-empty at line 480
                Left(EvalError.EvalFailed("XIRR: dates cannot be empty", None))
          }

      case TExpr.VLookup(lookupExpr, table, colIndexExpr, rangeLookupExpr) =>
        for
          lookupValue <- eval(lookupExpr, sheet, clock, workbook)
          colIndex <- eval(colIndexExpr, sheet, clock, workbook)
          rangeMatch <- eval(rangeLookupExpr, sheet, clock, workbook)
          result <-
            if colIndex < 1 || colIndex > table.width then
              Left(
                EvalError.EvalFailed(
                  s"VLOOKUP: col_index_num $colIndex is outside 1..${table.width}",
                  Some(s"VLOOKUP(…, ${table.toA1})")
                )
              )
            else
              val rowIndices = 0 until table.height
              val keyCol0 = table.colStart.index0
              val rowStart0 = table.rowStart.index0
              val resultCol0 = keyCol0 + (colIndex - 1)

              // Helper to extract text from a cell value for comparison
              def extractTextForMatch(cv: CellValue): Option[String] = cv match
                case CellValue.Text(s) => Some(s)
                case CellValue.Number(n) => Some(n.bigDecimal.stripTrailingZeros().toPlainString)
                case CellValue.Bool(b) => Some(if b then "TRUE" else "FALSE")
                case CellValue.Formula(_, Some(cached)) => extractTextForMatch(cached)
                case _ => None

              // Helper to extract numeric from a cell value
              def extractNumericForMatch(cv: CellValue): Option[BigDecimal] = cv match
                case CellValue.Number(n) => Some(n)
                case CellValue.Text(s) => scala.util.Try(BigDecimal(s.trim)).toOption
                case CellValue.Bool(b) => Some(if b then BigDecimal(1) else BigDecimal(0))
                case CellValue.Formula(_, Some(cached)) => extractNumericForMatch(cached)
                case _ => None

              // Determine if lookup is text or numeric
              val isTextLookup = lookupValue match
                case _: String => true
                case _: BigDecimal => false
                case _: Int => false
                case _: Boolean => false
                case _ => true // Default to text for other types

              val chosenRowOpt: Option[Int] =
                if rangeMatch then
                  // Approximate match: numeric only, find largest key <= lookup
                  val numericLookup: Option[BigDecimal] = lookupValue match
                    case n: BigDecimal => Some(n)
                    case i: Int => Some(BigDecimal(i))
                    case s: String => scala.util.Try(BigDecimal(s.trim)).toOption
                    case _ => None

                  numericLookup.flatMap { lookup =>
                    val keyedRows: List[(Int, BigDecimal)] =
                      rowIndices.toList.flatMap { i =>
                        val keyRef = ARef.from0(keyCol0, rowStart0 + i)
                        extractNumericForMatch(sheet(keyRef).value).map(k => (i, k))
                      }
                    keyedRows
                      .filter(_._2 <= lookup)
                      .sortBy(_._2)
                      .lastOption
                      .map(_._1)
                  }
                else
                  // Exact match: supports both text (case-insensitive) and numeric
                  if isTextLookup then
                    val lookupText = lookupValue.toString.toLowerCase
                    rowIndices.find { i =>
                      val keyRef = ARef.from0(keyCol0, rowStart0 + i)
                      extractTextForMatch(sheet(keyRef).value)
                        .exists(_.toLowerCase == lookupText)
                    }
                  else
                    // Numeric exact match
                    val numericLookup: Option[BigDecimal] = lookupValue match
                      case n: BigDecimal => Some(n)
                      case i: Int => Some(BigDecimal(i))
                      case _ => None
                    numericLookup.flatMap { lookup =>
                      rowIndices.find { i =>
                        val keyRef = ARef.from0(keyCol0, rowStart0 + i)
                        extractNumericForMatch(sheet(keyRef).value).contains(lookup)
                      }
                    }

              chosenRowOpt match
                case Some(rowIndex) =>
                  val resultRef = ARef.from0(resultCol0, rowStart0 + rowIndex)
                  val resultCell = sheet(resultRef)
                  // Return the CellValue directly (preserves type)
                  Right(resultCell.value)
                case None =>
                  Left(
                    EvalError.EvalFailed(
                      if rangeMatch then "VLOOKUP approximate match not found"
                      else "VLOOKUP exact match not found",
                      Some(s"VLOOKUP($lookupValue, ${table.toA1}, $colIndex, $rangeMatch)")
                    )
                  )
        yield result

      // ===== Arithmetic Range Functions =====
      case TExpr.Min(range) =>
        // Min: find minimum value in range (single-pass)
        val cells = range.cells.map(cellRef => sheet(cellRef))
        val values = cells.flatMap { cell =>
          TExpr.decodeNumeric(cell).toOption
        }
        val result = values.foldLeft(Option.empty[BigDecimal]) { (acc, v) =>
          Some(acc.fold(v)(_ min v))
        }
        // Excel behavior: MIN of empty range returns 0
        Right(result.getOrElse(BigDecimal(0)))

      case TExpr.Max(range) =>
        // Max: find maximum value in range (single-pass)
        val cells = range.cells.map(cellRef => sheet(cellRef))
        val values = cells.flatMap { cell =>
          TExpr.decodeNumeric(cell).toOption
        }
        val result = values.foldLeft(Option.empty[BigDecimal]) { (acc, v) =>
          Some(acc.fold(v)(_ max v))
        }
        // Excel behavior: MAX of empty range returns 0
        Right(result.getOrElse(BigDecimal(0)))

      case TExpr.Average(range) =>
        // Average: compute sum/count of numeric values in range
        val cells = range.cells.map(cellRef => sheet(cellRef))
        val values = cells.flatMap { cell =>
          TExpr.decodeNumeric(cell).toOption
        }
        // Single-pass sum and count (avoids .toList materialization)
        val (sum, count) = values.foldLeft((BigDecimal(0), 0)) { case ((s, c), v) =>
          (s + v, c + 1)
        }
        if count == 0 then
          // Excel behavior: AVERAGE of empty range returns #DIV/0!
          Left(EvalError.DivByZero("AVERAGE(empty range)", "count=0"))
        else Right(sum / count)

      // ===== Range Aggregation =====
      case foldExpr: TExpr.FoldRange[a, b] =>
        // FoldRange: iterate cells in range, apply step function with accumulator
        // Note: Use pattern match with type parameters to preserve types
        // Excel behavior: SUM/COUNT/AVERAGE skip cells that can't be decoded (empty cells, text in numeric context)
        val cells = foldExpr.range.cells.map(cellRef => sheet(cellRef))
        val result: Either[EvalError, b] = cells.foldLeft[Either[EvalError, b]](
          Right(foldExpr.z)
        ) { (accEither, cellInstance) =>
          accEither.flatMap { acc =>
            foldExpr.decode(cellInstance) match
              case Right(value) =>
                // Successfully decoded, apply step function
                Right(foldExpr.step(acc, value))
              case Left(_codecErr) =>
                // Failed to decode: skip this cell (Excel behavior for SUM/COUNT/AVERAGE)
                // Empty cells and type mismatches are silently ignored
                Right(acc)
          }
        }
        result.asInstanceOf[Either[EvalError, A]]

      // ===== Cross-Sheet Range Aggregation =====
      case sheetFold: TExpr.SheetFoldRange[a, b] =>
        // SheetFoldRange: like FoldRange but operates on a range in another sheet
        workbook match
          case None =>
            val refStr = s"${sheetFold.sheet.value}!${sheetFold.range.toA1}"
            Left(Evaluator.missingWorkbookError(refStr, isRange = true))
          case Some(wb) =>
            wb(sheetFold.sheet) match
              case Left(err) =>
                Left(Evaluator.sheetNotFoundError(sheetFold.sheet, err))
              case Right(targetSheet) =>
                val cells = sheetFold.range.cells.map(cellRef => targetSheet(cellRef))
                val result: Either[EvalError, b] = cells.foldLeft[Either[EvalError, b]](
                  Right(sheetFold.z)
                ) { (accEither, cellInstance) =>
                  accEither.flatMap { acc =>
                    sheetFold.decode(cellInstance) match
                      case Right(value) =>
                        Right(sheetFold.step(acc, value))
                      case Left(_codecErr) =>
                        Right(acc) // Skip cells that can't be decoded
                  }
                }
                result.asInstanceOf[Either[EvalError, A]]

      // ===== Cross-Sheet Aggregate Functions =====

      case TExpr.SheetMin(sheetName, range) =>
        workbook match
          case None =>
            val refStr = s"${sheetName.value}!${range.toA1}"
            Left(Evaluator.missingWorkbookError(refStr, isRange = true))
          case Some(wb) =>
            wb(sheetName) match
              case Left(err) =>
                Left(Evaluator.sheetNotFoundError(sheetName, err))
              case Right(targetSheet) =>
                val cells = range.cells.map(cellRef => targetSheet(cellRef))
                val values = cells.flatMap { cell =>
                  TExpr.decodeNumeric(cell).toOption
                }
                val result = values.foldLeft(Option.empty[BigDecimal]) { (acc, v) =>
                  Some(acc.fold(v)(_ min v))
                }
                // Excel behavior: MIN of empty range returns 0
                Right(result.getOrElse(BigDecimal(0)))

      case TExpr.SheetMax(sheetName, range) =>
        workbook match
          case None =>
            val refStr = s"${sheetName.value}!${range.toA1}"
            Left(Evaluator.missingWorkbookError(refStr, isRange = true))
          case Some(wb) =>
            wb(sheetName) match
              case Left(err) =>
                Left(Evaluator.sheetNotFoundError(sheetName, err))
              case Right(targetSheet) =>
                val cells = range.cells.map(cellRef => targetSheet(cellRef))
                val values = cells.flatMap { cell =>
                  TExpr.decodeNumeric(cell).toOption
                }
                val result = values.foldLeft(Option.empty[BigDecimal]) { (acc, v) =>
                  Some(acc.fold(v)(_ max v))
                }
                // Excel behavior: MAX of empty range returns 0
                Right(result.getOrElse(BigDecimal(0)))

      case TExpr.SheetAverage(sheetName, range) =>
        workbook match
          case None =>
            val refStr = s"${sheetName.value}!${range.toA1}"
            Left(Evaluator.missingWorkbookError(refStr, isRange = true))
          case Some(wb) =>
            wb(sheetName) match
              case Left(err) =>
                Left(Evaluator.sheetNotFoundError(sheetName, err))
              case Right(targetSheet) =>
                val cells = range.cells.map(cellRef => targetSheet(cellRef))
                val values = cells.flatMap { cell =>
                  TExpr.decodeNumeric(cell).toOption
                }
                // Single-pass sum and count
                val (sum, count) = values.foldLeft((BigDecimal(0), 0)) { case ((s, c), v) =>
                  (s + v, c + 1)
                }
                if count == 0 then
                  // Excel behavior: AVERAGE of empty range returns #DIV/0!
                  Left(EvalError.DivByZero("AVERAGE(empty range)", "count=0"))
                else Right(sum / count)

      case TExpr.SheetCount(sheetName, range) =>
        workbook match
          case None =>
            val refStr = s"${sheetName.value}!${range.toA1}"
            Left(Evaluator.missingWorkbookError(refStr, isRange = true))
          case Some(wb) =>
            wb(sheetName) match
              case Left(err) =>
                Left(Evaluator.sheetNotFoundError(sheetName, err))
              case Right(targetSheet) =>
                val cells = range.cells.map(cellRef => targetSheet(cellRef))
                val count = cells.count { cell =>
                  TExpr.decodeNumeric(cell).isRight
                }
                Right(count)

      // ===== Unified Aggregation (new type-class based) =====

      case TExpr.Aggregate(location, aggregator) =>
        // Get cells based on location (local or cross-sheet)
        val cellsResult: Either[EvalError, Iterator[Cell]] = location match
          case TExpr.RangeLocation.Local(range) =>
            Right(range.cells.iterator.map(ref => sheet(ref)))
          case TExpr.RangeLocation.CrossSheet(sheetName, range) =>
            workbook match
              case None =>
                val refStr = s"${sheetName.value}!${range.toA1}"
                Left(Evaluator.missingWorkbookError(refStr, isRange = true))
              case Some(wb) =>
                wb(sheetName) match
                  case Left(err) =>
                    Left(Evaluator.sheetNotFoundError(sheetName, err))
                  case Right(targetSheet) =>
                    Right(range.cells.iterator.map(ref => targetSheet(ref)))

        // Evaluate using the aggregator tag
        cellsResult.flatMap { cells =>
          aggregator.evaluate(cells).asInstanceOf[Either[EvalError, A]]
        }

      // ===== Conditional Aggregation Functions =====

      case TExpr.SumIf(range, criteriaExpr, sumRangeOpt) =>
        // SUMIF: Sum cells in sumRange where corresponding cells in range match criteria
        eval(criteriaExpr, sheet, clock, workbook).flatMap { criteriaValue =>
          val criterion = CriteriaMatcher.parse(criteriaValue)
          val effectiveRange = sumRangeOpt.getOrElse(range)
          val rangeRefsList = range.cells.toList
          val sumRefsList = effectiveRange.cells.toList

          // Validate dimensions match (both width and height, not just total count)
          if range.width != effectiveRange.width || range.height != effectiveRange.height then
            Left(
              EvalError.EvalFailed(
                s"SUMIF: range and sum_range must have same dimensions (${range.height}×${range.width} vs ${effectiveRange.height}×${effectiveRange.width})",
                Some(s"SUMIF(${range.toA1}, ..., ${effectiveRange.toA1})")
              )
            )
          else
            val pairs = rangeRefsList.zip(sumRefsList)
            val sum = pairs.foldLeft(BigDecimal(0)) { case (acc, (testRef, sumRef)) =>
              val testCell = sheet(testRef)
              if CriteriaMatcher.matches(testCell.value, criterion) then
                sheet(sumRef).value match
                  case CellValue.Number(n) => acc + n
                  case _ => acc // Skip non-numeric (Excel behavior)
              else acc
            }
            Right(sum)
        }

      case TExpr.CountIf(range, criteriaExpr) =>
        // COUNTIF: Count cells in range matching criteria
        eval(criteriaExpr, sheet, clock, workbook).map { criteriaValue =>
          val criterion = CriteriaMatcher.parse(criteriaValue)
          val count = range.cells.count { ref =>
            CriteriaMatcher.matches(sheet(ref).value, criterion)
          }
          BigDecimal(count)
        }

      case TExpr.SumIfs(sumRange, conditions) =>
        // SUMIFS: Sum cells where ALL criteria match (AND logic)
        // First evaluate all criteria expressions
        val criteriaEithers = conditions.map { case (_, criteriaExpr) =>
          eval(criteriaExpr, sheet, clock, workbook)
        }

        // Collect all results or return first error (use :: prepend for O(1), reverse at end)
        criteriaEithers
          .foldLeft[Either[EvalError, List[Any]]](Right(List.empty)) { (acc, either) =>
            acc.flatMap(list => either.map(v => v :: list))
          }
          .map(_.reverse)
          .flatMap { criteriaValues =>
            val parsedConditions = conditions
              .zip(criteriaValues)
              .map { case ((range, _), criteriaValue) =>
                (range, CriteriaMatcher.parse(criteriaValue))
              }

            // Validate all ranges have same dimensions as sumRange (both width and height)
            val sumRefsList = sumRange.cells.toList
            val dimensionError = parsedConditions.collectFirst {
              case (range, _) if range.width != sumRange.width || range.height != sumRange.height =>
                EvalError.EvalFailed(
                  s"SUMIFS: all ranges must have same dimensions (sum_range is ${sumRange.height}×${sumRange.width}, criteria_range is ${range.height}×${range.width})",
                  Some(s"SUMIFS(${sumRange.toA1}, ${range.toA1}, ...)")
                )
            }

            dimensionError match
              case Some(err) => Left(err)
              case None =>
                val sum = sumRefsList.indices.foldLeft(BigDecimal(0)) { (acc, idx) =>
                  // Check if ALL criteria match for this row
                  val allMatch = parsedConditions.forall { case (criteriaRange, criterion) =>
                    val testRef = criteriaRange.cells.toList(idx)
                    CriteriaMatcher.matches(sheet(testRef).value, criterion)
                  }
                  if allMatch then
                    sheet(sumRefsList(idx)).value match
                      case CellValue.Number(n) => acc + n
                      case _ => acc
                  else acc
                }
                Right(sum)
          }

      case TExpr.CountIfs(conditions) =>
        // COUNTIFS: Count cells where ALL criteria match (AND logic)
        // First evaluate all criteria expressions
        val criteriaEithers = conditions.map { case (_, criteriaExpr) =>
          eval(criteriaExpr, sheet, clock, workbook)
        }

        // Collect all results or return first error (use :: prepend for O(1), reverse at end)
        criteriaEithers
          .foldLeft[Either[EvalError, List[Any]]](Right(List.empty)) { (acc, either) =>
            acc.flatMap(list => either.map(v => v :: list))
          }
          .map(_.reverse)
          .flatMap { criteriaValues =>
            val parsedConditions = conditions
              .zip(criteriaValues)
              .map { case ((range, _), criteriaValue) =>
                (range, CriteriaMatcher.parse(criteriaValue))
              }

            // Pattern match to safely extract first range (non-empty per parser validation)
            parsedConditions match
              case (firstRange, _) :: _ =>
                val refCount = firstRange.cells.toList.length

                // Validate all ranges have same dimensions (both width and height)
                val dimensionError = parsedConditions.collectFirst {
                  case (range, _)
                      if range.width != firstRange.width || range.height != firstRange.height =>
                    EvalError.EvalFailed(
                      s"COUNTIFS: all ranges must have same dimensions (first is ${firstRange.height}×${firstRange.width}, this is ${range.height}×${range.width})",
                      Some(s"COUNTIFS(${firstRange.toA1}, ..., ${range.toA1}, ...)")
                    )
                }

                dimensionError match
                  case Some(err) => Left(err)
                  case None =>
                    val count = (0 until refCount).count { idx =>
                      // Check if ALL criteria match for this row
                      parsedConditions.forall { case (criteriaRange, criterion) =>
                        val testRef = criteriaRange.cells.toList(idx)
                        CriteriaMatcher.matches(sheet(testRef).value, criterion)
                      }
                    }
                    Right(BigDecimal(count))
              case Nil =>
                // Should never happen per parser validation, but handle gracefully
                Right(BigDecimal(0))
          }

      // ===== Array and Advanced Lookup Functions =====

      case TExpr.SumProduct(arrays) =>
        // SUMPRODUCT: multiply corresponding elements across arrays and sum
        arrays match
          case Nil =>
            // Empty arrays case - return 0 (should not happen per parser validation)
            Right(BigDecimal(0))
          case first :: rest =>
            val firstWidth = first.width
            val firstHeight = first.height

            // Validate all arrays have same dimensions
            val dimensionError = rest.collectFirst {
              case range if range.width != firstWidth || range.height != firstHeight =>
                EvalError.EvalFailed(
                  s"SUMPRODUCT: all arrays must have same dimensions (first is ${firstHeight}×${firstWidth}, got ${range.height}×${range.width})",
                  Some(s"SUMPRODUCT(${first.toA1}, ${range.toA1}, ...)")
                )
            }

            dimensionError match
              case Some(err) => Left(err)
              case None =>
                // Get all cell references as parallel lists
                val cellLists = arrays.map(_.cells.toList)
                val cellCount = cellLists.headOption.map(_.length).getOrElse(0)

                // For each position, multiply corresponding values across all arrays
                val sum = (0 until cellCount).foldLeft(BigDecimal(0)) { (acc, idx) =>
                  // Get values at position idx from each array, coerce to numeric
                  val values = cellLists.map { cells =>
                    val ref = cells(idx)
                    coerceToNumeric(sheet(ref).value)
                  }
                  // Product of all values at this position
                  val product = values.foldLeft(BigDecimal(1))(_ * _)
                  acc + product
                }

                Right(sum)

      case TExpr.XLookup(
            lookupValueExpr,
            lookupArray,
            returnArray,
            ifNotFoundOpt,
            matchModeExpr,
            searchModeExpr
          ) =>
        // XLOOKUP: advanced lookup with flexible matching
        // Validate dimensions first
        if lookupArray.width != returnArray.width || lookupArray.height != returnArray.height then
          Left(
            EvalError.EvalFailed(
              s"XLOOKUP: lookup_array and return_array must have same dimensions (${lookupArray.height}×${lookupArray.width} vs ${returnArray.height}×${returnArray.width})",
              Some(s"XLOOKUP(..., ${lookupArray.toA1}, ${returnArray.toA1}, ...)")
            )
          )
        else
          // Resolve sheets for cross-sheet lookup (or use current sheet for local)
          val lookupSheet = lookupArray.sheetName
            .flatMap(name => workbook.flatMap(_.apply(name).toOption))
            .getOrElse(sheet)
          val returnSheet = returnArray.sheetName
            .flatMap(name => workbook.flatMap(_.apply(name).toOption))
            .getOrElse(sheet)
          for
            lookupValue <- evalAny(lookupValueExpr, sheet, clock, workbook)
            matchModeRaw <- evalAny(matchModeExpr, sheet, clock, workbook)
            searchModeRaw <- evalAny(searchModeExpr, sheet, clock, workbook)
            matchMode = toInt(matchModeRaw)
            searchMode = toInt(searchModeRaw)
            result <- performXLookup(
              lookupValue,
              lookupArray.range,
              returnArray.range,
              ifNotFoundOpt,
              matchMode,
              searchMode,
              lookupSheet, // Use resolved sheet for lookup
              clock,
              workbook
            )
          yield result

      // ===== Error Handling Functions =====

      case TExpr.Iferror(valueExpr, valueIfErrorExpr) =>
        // IFERROR: return valueIfError if value results in any error
        evalAny(valueExpr, sheet, clock, workbook) match
          case Left(_) =>
            // Evaluation error occurred - return fallback
            evalAny(valueIfErrorExpr, sheet, clock, workbook).map(convertToCellValue)
          case Right(cv: CellValue) =>
            cv match
              case CellValue.Error(_) =>
                // Cell contains error value - return fallback
                evalAny(valueIfErrorExpr, sheet, clock, workbook).map(convertToCellValue)
              case _ =>
                // No error - return original value
                Right(cv)
          case Right(other) =>
            // Non-CellValue result - wrap it
            Right(convertToCellValue(other))

      case TExpr.Iserror(valueExpr) =>
        // ISERROR: return TRUE if value results in any error
        evalAny(valueExpr, sheet, clock, workbook) match
          case Left(_) => Right(true) // Evaluation error
          case Right(cv: CellValue) =>
            cv match
              case CellValue.Error(_) => Right(true) // Cell error value
              case _ => Right(false) // No error
          case Right(_) => Right(false) // Non-CellValue result, no error

      // ===== Rounding and Math Functions =====

      case TExpr.Round(valueExpr, numDigitsExpr) =>
        // ROUND: round to specified digits using HALF_UP
        for
          value <- eval(valueExpr, sheet, clock, workbook)
          numDigits <- eval(numDigitsExpr, sheet, clock, workbook)
        yield roundToDigits(value, numDigits.toInt, BigDecimal.RoundingMode.HALF_UP)

      case TExpr.RoundUp(valueExpr, numDigitsExpr) =>
        // ROUNDUP: round away from zero (Excel semantics)
        // Examples: ROUNDUP(2.1, 0) = 3, ROUNDUP(-2.1, 0) = -3
        for
          value <- eval(valueExpr, sheet, clock, workbook)
          numDigits <- eval(numDigitsExpr, sheet, clock, workbook)
        yield
          // Excel ROUNDUP rounds away from zero, which requires different RoundingMode by sign:
          // - Positive numbers: CEILING (rounds toward +∞, away from 0)
          // - Negative numbers: FLOOR (rounds toward -∞, away from 0)
          val mode =
            if value >= 0 then BigDecimal.RoundingMode.CEILING
            else BigDecimal.RoundingMode.FLOOR
          roundToDigits(value, numDigits.toInt, mode)

      case TExpr.RoundDown(valueExpr, numDigitsExpr) =>
        // ROUNDDOWN: round toward zero / truncate (Excel semantics)
        // Examples: ROUNDDOWN(2.9, 0) = 2, ROUNDDOWN(-2.9, 0) = -2
        for
          value <- eval(valueExpr, sheet, clock, workbook)
          numDigits <- eval(numDigitsExpr, sheet, clock, workbook)
        yield
          // Excel ROUNDDOWN rounds toward zero, which requires different RoundingMode by sign:
          // - Positive numbers: FLOOR (rounds toward -∞, toward 0)
          // - Negative numbers: CEILING (rounds toward +∞, toward 0)
          val mode =
            if value >= 0 then BigDecimal.RoundingMode.FLOOR
            else BigDecimal.RoundingMode.CEILING
          roundToDigits(value, numDigits.toInt, mode)

      case TExpr.Abs(valueExpr) =>
        // ABS: absolute value
        eval(valueExpr, sheet, clock, workbook).map(_.abs)

      // ===== Lookup Functions =====

      case TExpr.Index(array, rowNumExpr, colNumExpr) =>
        // INDEX: return value at position in array
        // Returns #REF! error with descriptive message if position is out of bounds
        for
          rowNum <- eval(rowNumExpr, sheet, clock, workbook)
          colNum <- colNumExpr match
            case Some(expr) => eval(expr, sheet, clock, workbook).map(Some(_))
            case None => Right(None)
          result <- {
            val rowIdx = rowNum.toInt - 1 // 1-based to 0-based
            val colIdx = colNum.map(_.toInt - 1).getOrElse(0) // Default to first column

            // Get dimensions of the range
            val startCol = array.colStart.index0
            val startRow = array.rowStart.index0
            val numCols = array.colEnd.index0 - startCol + 1
            val numRows = array.rowEnd.index0 - startRow + 1

            // Bounds check with descriptive error
            if rowIdx < 0 || rowIdx >= numRows then
              Left(
                EvalError.EvalFailed(
                  s"INDEX: row_num ${rowNum.toInt} is out of bounds (array has $numRows rows, valid range: 1-$numRows) (#REF!)",
                  Some(s"INDEX(${array.toA1}, $rowNum${colNum.map(c => s", $c").getOrElse("")})")
                )
              )
            else if colIdx < 0 || colIdx >= numCols then
              Left(
                EvalError.EvalFailed(
                  s"INDEX: col_num ${colNum.map(_.toInt).getOrElse(1)} is out of bounds (array has $numCols columns, valid range: 1-$numCols) (#REF!)",
                  Some(s"INDEX(${array.toA1}, $rowNum${colNum.map(c => s", $c").getOrElse("")})")
                )
              )
            else
              val targetRef = ARef.from0(startCol + colIdx, startRow + rowIdx)
              Right(sheet(targetRef).value)
          }
        yield result

      case TExpr.Match(lookupValueExpr, lookupArray, matchTypeExpr) =>
        // MATCH: find position of value in array (1-based)
        // Returns #N/A error if no match found (Excel behavior)
        for
          lookupValue <- evalAny(lookupValueExpr, sheet, clock, workbook)
          matchType <- eval(matchTypeExpr, sheet, clock, workbook)
          result <- {
            val matchTypeInt = matchType.toInt

            // Get cells from lookup array (must be 1D - single row or column)
            val cells: List[(Int, CellValue)] =
              lookupArray.cells.toList.zipWithIndex.map { case (ref, idx) =>
                (idx + 1, sheet(ref).value) // 1-based position
              }

            val positionOpt: Option[Int] = matchTypeInt match
              case 0 =>
                // Exact match
                cells
                  .find { case (_, cv) =>
                    compareCellValues(cv, lookupValue) == 0
                  }
                  .map(_._1)

              case 1 =>
                // Largest value <= lookup (array should be sorted ascending)
                val numericLookup = coerceToBigDecimal(lookupValue)
                val candidates = cells.flatMap { case (pos, cv) =>
                  val numericCv = coerceToNumeric(cv)
                  if numericCv <= numericLookup then Some((pos, numericCv))
                  else None
                }
                candidates.maxByOption(_._2).map(_._1)

              case -1 =>
                // Smallest value >= lookup (array should be sorted descending)
                val numericLookup = coerceToBigDecimal(lookupValue)
                val candidates = cells.flatMap { case (pos, cv) =>
                  val numericCv = coerceToNumeric(cv)
                  if numericCv >= numericLookup then Some((pos, numericCv))
                  else None
                }
                candidates.minByOption(_._2).map(_._1)

              case _ =>
                None // Unknown match type

            positionOpt match
              case Some(pos) => Right(BigDecimal(pos))
              case None =>
                Left(
                  EvalError.EvalFailed(
                    "MATCH: no match found for lookup value (#N/A)",
                    Some("MATCH(lookup_value, lookup_array, [match_type])")
                  )
                )
          }
        yield result

  // ===== Helper Methods =====

  /**
   * Coerce CellValue to BigDecimal for SUMPRODUCT.
   *
   * Excel semantics:
   *   - Number → as-is
   *   - TRUE → 1
   *   - FALSE → 0
   *   - Text/Empty/Error → 0
   *   - Formula with cached value → coerce cached value
   */
  private def coerceToNumeric(value: CellValue): BigDecimal =
    value match
      case CellValue.Number(n) => n
      case CellValue.Bool(true) => BigDecimal(1)
      case CellValue.Bool(false) => BigDecimal(0)
      case CellValue.Formula(_, Some(cached)) => coerceToNumeric(cached)
      case _ => BigDecimal(0)

  /**
   * Evaluate any TExpr to its runtime value (type-erased).
   *
   * Used for polymorphic lookup values in XLOOKUP.
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  private def evalAny(
    expr: TExpr[?],
    sheet: Sheet,
    clock: Clock,
    workbook: Option[Workbook]
  ): Either[EvalError, Any] =
    eval(expr.asInstanceOf[TExpr[Any]], sheet, clock, workbook)

  /**
   * Convert runtime value to Int for match_mode/search_mode parameters.
   *
   * Parser produces BigDecimal literals, but these parameters need Int.
   */
  private def toInt(value: Any): Int =
    value match
      case i: Int => i
      case bd: BigDecimal => bd.toInt
      case n: Number => n.intValue()
      case _ => 0

  /**
   * Round BigDecimal to specified number of digits.
   *
   * Handles negative numDigits by rounding to left of decimal point. Excel behavior: ROUND(1234,
   * -2) = 1200
   */
  private def roundToDigits(
    value: BigDecimal,
    numDigits: Int,
    mode: BigDecimal.RoundingMode.Value
  ): BigDecimal =
    if numDigits >= 0 then value.setScale(numDigits, mode)
    else
      // Negative numDigits: round to left of decimal
      // ROUND(1234, -2) means round to nearest 100 = 1200
      val scale = math.pow(10, -numDigits).toLong
      val divided = value / scale
      val rounded = divided.setScale(0, mode)
      rounded * scale

  /**
   * Convert any value to CellValue for IFERROR result.
   */
  private def convertToCellValue(value: Any): CellValue =
    value match
      case cv: CellValue => cv
      case s: String => CellValue.Text(s)
      case n: BigDecimal => CellValue.Number(n)
      case b: Boolean => CellValue.Bool(b)
      case n: Int => CellValue.Number(BigDecimal(n))
      case n: Long => CellValue.Number(BigDecimal(n))
      case n: Double => CellValue.Number(BigDecimal(n))
      case d: java.time.LocalDate =>
        CellValue.DateTime(d.atStartOfDay())
      case dt: java.time.LocalDateTime =>
        CellValue.DateTime(dt)
      case other => CellValue.Text(other.toString)

  /**
   * Compare two cell values for MATCH function. Returns: 0 if equal, negative if cv < value,
   * positive if cv > value
   */
  private def compareCellValues(cv: CellValue, value: Any): Int =
    (cv, value) match
      case (CellValue.Number(n1), n2: BigDecimal) => n1.compare(n2)
      case (CellValue.Number(n1), n2: Int) => n1.compare(BigDecimal(n2))
      case (CellValue.Number(n1), n2: Long) => n1.compare(BigDecimal(n2))
      case (CellValue.Number(n1), n2: Double) => n1.compare(BigDecimal(n2))
      case (CellValue.Text(s1), s2: String) => s1.compareToIgnoreCase(s2)
      case (CellValue.Bool(b1), b2: Boolean) => b1.compare(b2)
      case (CellValue.Number(n), CellValue.Number(n2)) => n.compare(n2)
      case (CellValue.Text(s), CellValue.Text(s2)) => s.compareToIgnoreCase(s2)
      case (CellValue.Bool(b), CellValue.Bool(b2)) => b.compare(b2)
      case _ => -2 // Different types, no match

  /**
   * Coerce Any value to BigDecimal for MATCH comparisons.
   */
  private def coerceToBigDecimal(value: Any): BigDecimal =
    value match
      case n: BigDecimal => n
      case n: Int => BigDecimal(n)
      case n: Long => BigDecimal(n)
      case n: Double => BigDecimal(n)
      case CellValue.Number(n) => n
      case CellValue.Bool(true) => BigDecimal(1)
      case CellValue.Bool(false) => BigDecimal(0)
      case _ => BigDecimal(0)

  /**
   * Perform XLOOKUP search with specified match and search modes.
   */
  private def performXLookup(
    lookupValue: Any,
    lookupArray: CellRange,
    returnArray: CellRange,
    ifNotFoundOpt: Option[TExpr[?]],
    matchMode: Int,
    searchMode: Int,
    sheet: Sheet,
    clock: Clock,
    workbook: Option[Workbook]
  ): Either[EvalError, CellValue] =
    val lookupCells = lookupArray.cells.toList
    val returnCells = returnArray.cells.toList

    // Determine search order based on searchMode
    val indices: List[Int] = searchMode match
      case -1 => lookupCells.indices.reverse.toList // last-to-first
      case _ => lookupCells.indices.toList // first-to-last (default, including binary modes)

    // Find matching index based on matchMode
    val matchingIndexOpt: Option[Int] = matchMode match
      case 0 => // Exact match
        indices.find { idx =>
          val cellValue = sheet(lookupCells(idx)).value
          matchesExactForXLookup(cellValue, lookupValue)
        }

      case 2 => // Wildcard match (only for string lookups)
        lookupValue match
          case pattern: String =>
            val criterion = CriteriaMatcher.parse(pattern)
            indices.find { idx =>
              CriteriaMatcher.matches(sheet(lookupCells(idx)).value, criterion)
            }
          case _ => None

      case -1 => // Next smaller (exact or next smaller)
        findNextSmaller(lookupValue, lookupCells, sheet, indices)

      case 1 => // Next larger (exact or next larger)
        findNextLarger(lookupValue, lookupCells, sheet, indices)

      case _ => None // Invalid match mode

    matchingIndexOpt match
      case Some(idx) =>
        Right(sheet(returnCells(idx)).value)
      case None =>
        // No match found - return if_not_found or #N/A
        ifNotFoundOpt match
          case Some(expr) =>
            evalAny(expr, sheet, clock, workbook).map(convertToXLookupResult)
          case None =>
            Right(CellValue.Error(com.tjclp.xl.cells.CellError.NA))

  /**
   * Check if cell value matches lookup value for exact matching.
   */
  private def matchesExactForXLookup(cellValue: CellValue, lookupValue: Any): Boolean =
    // Reuse CriteriaMatcher's Exact matching logic
    CriteriaMatcher.matches(cellValue, CriteriaMatcher.Exact(lookupValue))

  /**
   * Find the largest value ≤ lookup value (next smaller match mode).
   */
  private def findNextSmaller(
    lookupValue: Any,
    lookupCells: List[ARef],
    sheet: Sheet,
    indices: List[Int]
  ): Option[Int] =
    lookupValue match
      case targetNum: BigDecimal =>
        val candidates = indices
          .flatMap { idx =>
            extractNumericValue(sheet(lookupCells(idx)).value).map(n => (idx, n))
          }
          .filter(_._2 <= targetNum)
        candidates.sortBy(-_._2).headOption.map(_._1) // Largest value <= target
      case _ => None

  /**
   * Find the smallest value ≥ lookup value (next larger match mode).
   */
  private def findNextLarger(
    lookupValue: Any,
    lookupCells: List[ARef],
    sheet: Sheet,
    indices: List[Int]
  ): Option[Int] =
    lookupValue match
      case targetNum: BigDecimal =>
        val candidates = indices
          .flatMap { idx =>
            extractNumericValue(sheet(lookupCells(idx)).value).map(n => (idx, n))
          }
          .filter(_._2 >= targetNum)
        candidates.sortBy(_._2).headOption.map(_._1) // Smallest value >= target
      case _ => None

  /**
   * Extract numeric value from CellValue.
   */
  private def extractNumericValue(value: CellValue): Option[BigDecimal] =
    value match
      case CellValue.Number(n) => Some(n)
      case CellValue.Formula(_, Some(CellValue.Number(n))) => Some(n)
      case _ => None

  /**
   * Convert any evaluated value to CellValue for XLOOKUP result.
   */
  private def convertToXLookupResult(value: Any): CellValue =
    value match
      case cv: CellValue => cv
      case s: String => CellValue.Text(s)
      case n: BigDecimal => CellValue.Number(n)
      case b: Boolean => CellValue.Bool(b)
      case n: Int => CellValue.Number(BigDecimal(n))
      case n: Long => CellValue.Number(BigDecimal(n))
      case n: Double => CellValue.Number(BigDecimal(n))
      case other => CellValue.Text(other.toString)

  /**
   * Count working days (Mon-Fri) between two dates, excluding holidays.
   *
   * Uses imperative loop for efficiency - iterating day-by-day is the most straightforward approach
   * for this date range calculation with weekend/holiday exclusions.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def countWorkingDays(
    earlier: LocalDate,
    later: LocalDate,
    holidays: Set[LocalDate]
  ): Int =
    var count = 0
    var current = earlier
    while !current.isAfter(later) do
      val dayOfWeek = current.getDayOfWeek
      if dayOfWeek != java.time.DayOfWeek.SATURDAY &&
        dayOfWeek != java.time.DayOfWeek.SUNDAY &&
        !holidays.contains(current)
      then count += 1
      current = current.plusDays(1)
    count

  /**
   * Add N working days to a start date, skipping weekends and holidays.
   *
   * Uses imperative loop for efficiency - stepping day-by-day while counting only business days is
   * the most straightforward approach for WORKDAY.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def addWorkingDays(
    start: LocalDate,
    daysValue: Int,
    holidays: Set[LocalDate]
  ): LocalDate =
    var remaining = daysValue
    var current = start
    val direction = if remaining >= 0 then 1L else -1L

    while remaining != 0 do
      current = current.plusDays(direction)
      val dayOfWeek = current.getDayOfWeek
      if dayOfWeek != java.time.DayOfWeek.SATURDAY &&
        dayOfWeek != java.time.DayOfWeek.SUNDAY &&
        !holidays.contains(current)
      then remaining -= direction.toInt
    current
