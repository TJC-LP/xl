package com.tjclp.xl.formula

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
    FunctionSpec.simple[BigDecimal, UnaryDate]("YEAR", Arity.one) { (expr, ctx) =>
      ctx.evalExpr(expr).map(date => BigDecimal(date.getYear))
    }

  val month: FunctionSpec[BigDecimal] { type Args = UnaryDate } =
    FunctionSpec.simple[BigDecimal, UnaryDate]("MONTH", Arity.one) { (expr, ctx) =>
      ctx.evalExpr(expr).map(date => BigDecimal(date.getMonthValue))
    }

  val day: FunctionSpec[BigDecimal] { type Args = UnaryDate } =
    FunctionSpec.simple[BigDecimal, UnaryDate]("DAY", Arity.one) { (expr, ctx) =>
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
    FunctionSpec.simple[BigDecimal, DatePairUnit]("DATEDIF", Arity.three) { (args, ctx) =>
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
    FunctionSpec.simple[BigDecimal, DatePairOptRange]("NETWORKDAYS", Arity.Range(2, 3)) {
      (args, ctx) =>
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
      }
    ) { (args, ctx) =>
      val (startDateExpr, endDateExpr, basisOpt) = args
      val basisValueEither = basisOpt match
        case Some(expr) => evalValue(ctx, expr).map(toInt)
        case None => Right(0)
      for
        start <- ctx.evalExpr(startDateExpr)
        end <- ctx.evalExpr(endDateExpr)
        basisValue <- basisValueEither
        result <- basisValue match
          case 0 =>
            val d1 = math.min(30, start.getDayOfMonth)
            val d2 = if start.getDayOfMonth == 31 then 30 else end.getDayOfMonth
            val m1 = start.getMonthValue
            val m2 = end.getMonthValue
            val y1 = start.getYear
            val y2 = end.getYear
            val days = ((y2 - y1) * 360) + ((m2 - m1) * 30) + (d2 - d1)
            Right(BigDecimal(days) / BigDecimal(360))
          case 1 =>
            val daysBetween = ChronoUnit.DAYS.between(start, end)
            val year = start.getYear
            val daysInYear = if java.time.Year.isLeap(year) then 366 else 365
            Right(BigDecimal(daysBetween) / BigDecimal(daysInYear))
          case 2 =>
            val daysBetween = ChronoUnit.DAYS.between(start, end)
            Right(BigDecimal(daysBetween) / BigDecimal(360))
          case 3 =>
            val daysBetween = ChronoUnit.DAYS.between(start, end)
            Right(BigDecimal(daysBetween) / BigDecimal(365))
          case 4 =>
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
    }
