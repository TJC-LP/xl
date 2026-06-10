package com.tjclp.xl.formula.functions

import com.tjclp.xl.formula.ast.{TExpr, ExprValue}
import com.tjclp.xl.formula.eval.{EvalError, Evaluator}
import com.tjclp.xl.formula.parser.ParseError
import com.tjclp.xl.formula.{Clock, Arity}

import com.tjclp.xl.addressing.CellRange
import java.time.LocalDate
import java.time.temporal.ChronoUnit

trait FunctionSpecsDateTime extends FunctionSpecsBase:
  private def holidaySet(rangeOpt: Option[CellRange], ctx: EvalContext): Set[LocalDate] =
    rangeOpt
      .map { range =>
        range.cells
          .flatMap(ref => TExpr.decodeDate(ctx.sheet(ref)).toOption)
          .toSet
      }
      .getOrElse(Set.empty)

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

  val today: FunctionSpec[LocalDate] { type Args = NoArgs } =
    FunctionSpec.simple[LocalDate, NoArgs](
      "TODAY",
      Arity.none,
      flags = FunctionFlags(returnsDate = true)
    ) { (_, ctx) =>
      Right(ctx.clock.today())
    }

  val now: FunctionSpec[java.time.LocalDateTime] { type Args = NoArgs } =
    FunctionSpec.simple[java.time.LocalDateTime, NoArgs](
      "NOW",
      Arity.none,
      flags = FunctionFlags(returnsTime = true)
    ) { (_, ctx) =>
      Right(ctx.clock.now())
    }

  val date: FunctionSpec[LocalDate] { type Args = DateTripleInt } =
    FunctionSpec.simple[LocalDate, DateTripleInt](
      "DATE",
      Arity.three,
      flags = FunctionFlags(returnsDate = true)
    ) { (args, ctx) =>
      val (yearExpr, monthExpr, dayExpr) = args
      for
        y <- ctx.evalExpr(yearExpr)
        m <- ctx.evalExpr(monthExpr)
        d <- ctx.evalExpr(dayExpr)
        result <- scala.util.Try(LocalDate.of(y, m, d)).toEither.left.map { ex =>
          EvalError.EvalFailed(
            s"DATE: invalid date components (year=$y, month=$m, day=$d): ${ex.getMessage}"
          )
        }
      yield result
    }

  val year: FunctionSpec[BigDecimal] { type Args = UnaryDate } =
    FunctionSpec.simple[BigDecimal, UnaryDate](
      "YEAR",
      Arity.one,
      flags = FunctionFlags(returnsNumeric = true)
    ) { (expr, ctx) =>
      ctx.evalExpr(expr).map(date => BigDecimal(date.getYear))
    }

  val month: FunctionSpec[BigDecimal] { type Args = UnaryDate } =
    FunctionSpec.simple[BigDecimal, UnaryDate](
      "MONTH",
      Arity.one,
      flags = FunctionFlags(returnsNumeric = true)
    ) { (expr, ctx) =>
      ctx.evalExpr(expr).map(date => BigDecimal(date.getMonthValue))
    }

  val day: FunctionSpec[BigDecimal] { type Args = UnaryDate } =
    FunctionSpec.simple[BigDecimal, UnaryDate](
      "DAY",
      Arity.one,
      flags = FunctionFlags(returnsNumeric = true)
    ) { (expr, ctx) =>
      ctx.evalExpr(expr).map(date => BigDecimal(date.getDayOfMonth))
    }

  val eomonth: FunctionSpec[LocalDate] { type Args = DateInt } =
    FunctionSpec.simple[LocalDate, DateInt](
      "EOMONTH",
      Arity.two,
      flags = FunctionFlags(returnsDate = true)
    ) { (args, ctx) =>
      val (startDateExpr, monthsExpr) = args
      for
        date <- ctx.evalExpr(startDateExpr)
        monthsRaw <- evalValue(ctx, monthsExpr)
      yield
        val monthsValue = toInt(monthsRaw)
        val targetMonth = date.plusMonths(monthsValue.toLong)
        targetMonth.withDayOfMonth(targetMonth.lengthOfMonth)
    }

  val edate: FunctionSpec[LocalDate] { type Args = DateInt } =
    FunctionSpec.simple[LocalDate, DateInt](
      "EDATE",
      Arity.two,
      flags = FunctionFlags(returnsDate = true)
    ) { (args, ctx) =>
      val (startDateExpr, monthsExpr) = args
      for
        date <- ctx.evalExpr(startDateExpr)
        monthsRaw <- evalValue(ctx, monthsExpr)
      yield date.plusMonths(toInt(monthsRaw).toLong)
    }

  val datedif: FunctionSpec[BigDecimal] { type Args = DatePairUnit } =
    FunctionSpec.simple[BigDecimal, DatePairUnit](
      "DATEDIF",
      Arity.three,
      flags = FunctionFlags(returnsNumeric = true)
    ) { (args, ctx) =>
      val (startDateExpr, endDateExpr, unitExpr) = args
      for
        start <- ctx.evalExpr(startDateExpr)
        end <- ctx.evalExpr(endDateExpr)
        unitStr <- ctx.evalExpr(unitExpr)
        result <- unitStr.toUpperCase match
          case "Y" =>
            Right(BigDecimal(ChronoUnit.YEARS.between(start, end)))
          case "M" =>
            Right(BigDecimal(ChronoUnit.MONTHS.between(start, end)))
          case "D" =>
            Right(BigDecimal(ChronoUnit.DAYS.between(start, end)))
          case "MD" =>
            val daysDiff = end.getDayOfMonth - start.getDayOfMonth
            val adjustedDays =
              if daysDiff >= 0 then daysDiff
              else
                val prevMonthLength = end.minusMonths(1).lengthOfMonth
                val effectiveStartDay = math.min(start.getDayOfMonth, prevMonthLength)
                prevMonthLength - effectiveStartDay + end.getDayOfMonth
            Right(BigDecimal(adjustedDays))
          case "YM" =>
            val monthsDiff = end.getMonthValue - start.getMonthValue
            val adjustedMonths = if monthsDiff < 0 then 12 + monthsDiff else monthsDiff
            val finalMonths =
              if end.getDayOfMonth < start.getDayOfMonth then (adjustedMonths - 1 + 12) % 12
              else adjustedMonths
            Right(BigDecimal(finalMonths))
          case "YD" =>
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
    }

  val networkdays: FunctionSpec[BigDecimal] { type Args = DatePairOptRange } =
    FunctionSpec.simple[BigDecimal, DatePairOptRange](
      "NETWORKDAYS",
      Arity.Range(2, 3),
      flags = FunctionFlags(returnsNumeric = true)
    ) { (args, ctx) =>
      val (startDateExpr, endDateExpr, holidaysOpt) = args
      for
        start <- ctx.evalExpr(startDateExpr)
        end <- ctx.evalExpr(endDateExpr)
      yield
        val holidays = holidaySet(holidaysOpt, ctx)

        val (earlier, later) = if start.isBefore(end) then (start, end) else (end, start)
        val count = countWorkingDays(earlier, later, holidays)
        BigDecimal(if start.isBefore(end) || start.isEqual(end) then count else -count)
    }

  val workday: FunctionSpec[LocalDate] { type Args = DateIntOptRange } =
    FunctionSpec.simple[LocalDate, DateIntOptRange](
      "WORKDAY",
      Arity.Range(2, 3),
      flags = FunctionFlags(returnsDate = true)
    ) { (args, ctx) =>
      val (startDateExpr, daysExpr, holidaysOpt) = args
      for
        start <- ctx.evalExpr(startDateExpr)
        daysRaw <- evalValue(ctx, daysExpr)
      yield
        val daysValue = toInt(daysRaw)
        val holidays = holidaySet(holidaysOpt, ctx)
        addWorkingDays(start, daysValue, holidays)
    }

  val yearfrac: FunctionSpec[BigDecimal] { type Args = DatePairOptBasis } =
    FunctionSpec.simple[BigDecimal, DatePairOptBasis](
      "YEARFRAC",
      Arity.Range(2, 3),
      renderFn = Some { (args, printer) =>
        val (startDateExpr, endDateExpr, basisOpt) = args
        val rendered = basisOpt match
          case None =>
            List(printer.expr(startDateExpr), printer.expr(endDateExpr))
          case Some(TExpr.Lit(0)) =>
            List(printer.expr(startDateExpr), printer.expr(endDateExpr))
          case Some(basisExpr) =>
            List(
              printer.expr(startDateExpr),
              printer.expr(endDateExpr),
              printer.expr(basisExpr)
            )
        s"YEARFRAC(${rendered.mkString(", ")})"
      },
      flags = FunctionFlags(returnsNumeric = true)
    ) { (args, ctx) =>
      val (startDateExpr, endDateExpr, basisOpt) = args
      val basisValueEither = basisOpt match
        case Some(expr) => evalValue(ctx, expr).map(toInt)
        case None => Right(0)
      for
        start <- ctx.evalExpr(startDateExpr)
        end <- ctx.evalExpr(endDateExpr)
        basisValue <- basisValueEither
        result <- yearFraction(start, end, basisValue)
      yield result
    }

  /** Last day of February (Feb 28, or Feb 29 in a leap year). */
  private def isLastDayOfFebruary(date: LocalDate): Boolean =
    date.getMonthValue == 2 && date.getDayOfMonth == date.lengthOfMonth

  /** 30/360 day count: 360 days/year, 30 days/month, with pre-adjusted day-of-month values. */
  private def days360(start: LocalDate, d1: Int, end: LocalDate, d2: Int): Int =
    (end.getYear - start.getYear) * 360 + (end.getMonthValue - start.getMonthValue) * 30 +
      (d2 - d1)

  /**
   * Excel-compatible YEARFRAC for bases 0-4.
   *
   * Semantics follow Excel as reverse-engineered by David A. Wheeler
   * (https://dwheeler.com/yearfrac/) and the ODF OpenFormula v1.2 date-basis definitions (§4.11.7);
   * LibreOffice ships the same algorithm for Excel interop (GetYearFrac in
   * sc/source/core/tool/interpr2.cxx). Excel normalizes argument order first, so the result is
   * always non-negative: YEARFRAC(b, a) = YEARFRAC(a, b).
   */
  private def yearFraction(
    startDate: LocalDate,
    endDate: LocalDate,
    basis: Int
  ): Either[EvalError, BigDecimal] =
    val (start, end) =
      if startDate.isAfter(endDate) then (endDate, startDate) else (startDate, endDate)
    basis match
      case 0 =>
        // US (NASD) 30/360. Rule order matters: February end-of-month adjustments
        // apply BEFORE the 31st-day adjustments (ODFF v1.2 §4.11.7 basis 0):
        //   1. d1 and d2 both last day of February -> d2 = 30
        //   2. d1 last day of February            -> d1 = 30
        //   3. d2 = 31 and adjusted d1 >= 30      -> d2 = 30
        //   4. d1 = 31                            -> d1 = 30
        val startLastFeb = isLastDayOfFebruary(start)
        val d1 = if startLastFeb || start.getDayOfMonth == 31 then 30 else start.getDayOfMonth
        val d2 =
          if startLastFeb && isLastDayOfFebruary(end) then 30
          else if end.getDayOfMonth == 31 && d1 >= 30 then 30
          else end.getDayOfMonth
        Right(BigDecimal(days360(start, d1, end, d2)) / BigDecimal(360))
      case 1 =>
        // Excel "Actual/actual" — deliberately NOT ISDA Act/Act (no per-year proration).
        // Per Wheeler's reverse-engineering, Excel picks a single denominator:
        //  * span at most one calendar year (y1 == y2, or y2 == y1+1 with
        //    (m1, d1) on/after (m2, d2)): 366 when the single year is leap, when the
        //    span crosses Feb 29, or when the end date IS Feb 29; otherwise 365.
        //    Hence every exact one-year span yields exactly 1.0.
        //  * span over one year: the average length of all calendar years touched,
        //    days(Jan1(y1)..Jan1(y2+1)) / (y2 - y1 + 1).
        val days = ChronoUnit.DAYS.between(start, end)
        val y1 = start.getYear
        val y2 = end.getYear
        val withinOneYear =
          y1 == y2 || (y2 == y1 + 1 &&
            (start.getMonthValue > end.getMonthValue ||
              (start.getMonthValue == end.getMonthValue &&
                start.getDayOfMonth >= end.getDayOfMonth)))
        if withinOneYear then
          val leapDenominator =
            if y1 == y2 then java.time.Year.isLeap(y1)
            else
              (java.time.Year.isLeap(y1) && start.getMonthValue <= 2) ||
              (java.time.Year.isLeap(y2) &&
                (end.getMonthValue > 2 ||
                  (end.getMonthValue == 2 && end.getDayOfMonth == 29)))
          Right(BigDecimal(days) / BigDecimal(if leapDenominator then 366 else 365))
        else
          val yearCount = y2 - y1 + 1
          // Total days in calendar years y1..y2. Computed as Jan1(y1)..Jan1(y2) plus the
          // length of y2 so no date outside the input years is ever constructed (totality).
          val totalDays =
            ChronoUnit.DAYS.between(LocalDate.of(y1, 1, 1), LocalDate.of(y2, 1, 1)) +
              end.lengthOfYear
          // days / (totalDays / yearCount), as a single exact division
          Right(BigDecimal(days * yearCount) / BigDecimal(totalDays))
      case 2 =>
        Right(BigDecimal(ChronoUnit.DAYS.between(start, end)) / BigDecimal(360))
      case 3 =>
        Right(BigDecimal(ChronoUnit.DAYS.between(start, end)) / BigDecimal(365))
      case 4 =>
        // European 30E/360: only the 31st is pulled back to 30; no February rules.
        val d1 = math.min(30, start.getDayOfMonth)
        val d2 = math.min(30, end.getDayOfMonth)
        Right(BigDecimal(days360(start, d1, end, d2)) / BigDecimal(360))
      case other =>
        Left(
          EvalError.EvalFailed(
            s"YEARFRAC: invalid basis $other. Valid values: 0-4",
            Some("YEARFRAC(start, end, [basis])")
          )
        )
