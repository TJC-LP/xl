package com.tjclp.xl.formula

import java.time.LocalDate
import java.time.temporal.ChronoUnit

trait FunctionSpecsFinancial extends FunctionSpecsBase:
  val npv: FunctionSpec[BigDecimal] { type Args = NpvArgs } =
    FunctionSpec.simple[BigDecimal, NpvArgs]("NPV", Arity.two) { (args, ctx) =>
      val (rateExpr, range) = args
      ctx.evalExpr(rateExpr).flatMap { rate =>
        val onePlusR = BigDecimal(1) + rate
        if onePlusR == BigDecimal(0) then
          Left(
            EvalError.EvalFailed(
              "NPV: rate = -1 would require division by zero",
              Some("NPV(rate, values)")
            )
          )
        else
          val cashFlows: List[BigDecimal] =
            range.cells
              .map(ref => ctx.sheet(ref))
              .flatMap(cell => TExpr.decodeNumeric(cell).toOption)
              .toList

          val npv =
            cashFlows.zipWithIndex.foldLeft(BigDecimal(0)) { case (acc, (cf, idx)) =>
              val period = idx + 1
              acc + cf / onePlusR.pow(period)
            }
          Right(npv)
      }
    }

  val irr: FunctionSpec[BigDecimal] { type Args = IrrArgs } =
    FunctionSpec.simple[BigDecimal, IrrArgs]("IRR", Arity.Range(1, 2)) { (args, ctx) =>
      val (range, guessOpt) = args
      val cashFlows: List[BigDecimal] =
        range.cells
          .map(ref => ctx.sheet(ref))
          .flatMap(cell => TExpr.decodeNumeric(cell).toOption)
          .toList

      if cashFlows.isEmpty || !cashFlows.exists(_ < 0) || !cashFlows.exists(_ > 0) then
        Left(
          EvalError.EvalFailed(
            "IRR requires at least one positive and one negative cash flow",
            Some("IRR(values[, guess])")
          )
        )
      else
        val guessEither: Either[EvalError, BigDecimal] =
          guessOpt match
            case Some(guessExpr) => ctx.evalExpr(guessExpr)
            case None => Right(BigDecimal("0.1"))

        guessEither.flatMap { guess0 =>
          val maxIter = 50
          val tolerance = BigDecimal("1e-7")
          val one = BigDecimal(1)

          def npvAt(rate: BigDecimal): BigDecimal =
            val onePlusR = one + rate
            cashFlows.zipWithIndex.foldLeft(BigDecimal(0)) { case (acc, (cf, idx)) =>
              if idx == 0 then acc + cf
              else acc + cf / onePlusR.pow(idx)
            }

          def dNpvAt(rate: BigDecimal): BigDecimal =
            val onePlusR = one + rate
            cashFlows.zipWithIndex.foldLeft(BigDecimal(0)) { case (acc, (cf, idx)) =>
              if idx == 0 then acc
              else acc - (idx * cf) / onePlusR.pow(idx + 1)
            }

          @annotation.tailrec
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
    }

  val xnpv: FunctionSpec[BigDecimal] { type Args = XnpvArgs } =
    FunctionSpec.simple[BigDecimal, XnpvArgs]("XNPV", Arity.three) { (args, ctx) =>
      val (rateExpr, valuesRange, datesRange) = args
      for
        rate <- ctx.evalExpr(rateExpr)
        result <- {
          val values: List[BigDecimal] =
            valuesRange.cells
              .map(ref => ctx.sheet(ref))
              .flatMap(cell => TExpr.decodeNumeric(cell).toOption)
              .toList

          val dates: List[LocalDate] =
            datesRange.cells
              .map(ref => ctx.sheet(ref))
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
            dates match
              case date0 :: _ =>
                val onePlusR = BigDecimal(1) + rate
                val npv = values.zip(dates).foldLeft(BigDecimal(0)) { case (acc, (value, date)) =>
                  val daysDiff = ChronoUnit.DAYS.between(date0, date)
                  val yearFraction = BigDecimal(daysDiff) / BigDecimal(365)
                  val discountFactor = math.pow(onePlusR.toDouble, yearFraction.toDouble)
                  acc + value / BigDecimal(discountFactor)
                }
                Right(npv)
              case Nil =>
                Left(EvalError.EvalFailed("XNPV: dates cannot be empty", None))
        }
      yield result
    }

  val xirr: FunctionSpec[BigDecimal] { type Args = XirrArgs } =
    FunctionSpec.simple[BigDecimal, XirrArgs]("XIRR", Arity.Range(2, 3)) { (args, ctx) =>
      val (valuesRange, datesRange, guessOpt) = args
      val values: List[BigDecimal] =
        valuesRange.cells
          .map(ref => ctx.sheet(ref))
          .flatMap(cell => TExpr.decodeNumeric(cell).toOption)
          .toList

      val dates: List[LocalDate] =
        datesRange.cells
          .map(ref => ctx.sheet(ref))
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
          guessOpt match
            case Some(guessExpr) => ctx.evalExpr(guessExpr)
            case None => Right(BigDecimal("0.1"))

        guessEither.flatMap { guess0 =>
          val maxIter = 100
          val tolerance = BigDecimal("1e-7")
          dates match
            case date0 :: _ =>
              val yearFractions: List[BigDecimal] = dates.map { date =>
                val daysDiff = ChronoUnit.DAYS.between(date0, date)
                BigDecimal(daysDiff) / BigDecimal(365)
              }

              def xnpvAt(rate: BigDecimal): BigDecimal =
                val onePlusR = BigDecimal(1) + rate
                values.zip(yearFractions).foldLeft(BigDecimal(0)) { case (acc, (cf, yf)) =>
                  val discountFactor = math.pow(onePlusR.toDouble, yf.toDouble)
                  acc + cf / BigDecimal(discountFactor)
                }

              def dXnpvAt(rate: BigDecimal): BigDecimal =
                val onePlusR = BigDecimal(1) + rate
                values.zip(yearFractions).foldLeft(BigDecimal(0)) { case (acc, (cf, yf)) =>
                  val discountFactor = math.pow(onePlusR.toDouble, (yf + 1).toDouble)
                  acc - (yf * cf) / BigDecimal(discountFactor)
                }

              @annotation.tailrec
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
              Left(EvalError.EvalFailed("XIRR: dates cannot be empty", None))
        }
    }

  val pmt: FunctionSpec[BigDecimal] { type Args = TvmArgs } =
    FunctionSpec.simple[BigDecimal, TvmArgs]("PMT", Arity.Range(3, 5)) { (args, ctx) =>
      val (rateExpr, nperExpr, pvExpr, fvOpt, typeOpt) = args
      for
        rate <- ctx.evalExpr(rateExpr).map(_.toDouble)
        nper <- ctx.evalExpr(nperExpr).map(_.toDouble)
        pv <- ctx.evalExpr(pvExpr).map(_.toDouble)
        fv <- fvOpt match
          case Some(expr) => ctx.evalExpr(expr).map(_.toDouble)
          case None => Right(0.0)
        pmtType <- typeOpt match
          case Some(expr) => ctx.evalExpr(expr).map(v => if v.toInt != 0 then 1 else 0)
          case None => Right(0)
      yield
        if math.abs(rate) < 1e-10 then
          if nper == 0.0 then BigDecimal(Double.NaN)
          else BigDecimal(-(pv + fv) / nper)
        else
          val pvif = math.pow(1.0 + rate, nper)
          BigDecimal(-rate * (pv * pvif + fv) / ((1.0 + rate * pmtType) * (pvif - 1.0)))
    }

  val fv: FunctionSpec[BigDecimal] { type Args = TvmArgs } =
    FunctionSpec.simple[BigDecimal, TvmArgs]("FV", Arity.Range(3, 5)) { (args, ctx) =>
      val (rateExpr, nperExpr, pmtExpr, pvOpt, typeOpt) = args
      for
        rate <- ctx.evalExpr(rateExpr).map(_.toDouble)
        nper <- ctx.evalExpr(nperExpr).map(_.toDouble)
        pmt <- ctx.evalExpr(pmtExpr).map(_.toDouble)
        pv <- pvOpt match
          case Some(expr) => ctx.evalExpr(expr).map(_.toDouble)
          case None => Right(0.0)
        pmtType <- typeOpt match
          case Some(expr) => ctx.evalExpr(expr).map(v => if v.toInt != 0 then 1 else 0)
          case None => Right(0)
      yield
        if math.abs(rate) < 1e-10 then BigDecimal(-pv - pmt * nper)
        else
          val pvif = math.pow(1.0 + rate, nper)
          val fvifa = (pvif - 1.0) / rate
          BigDecimal(-pv * pvif - pmt * (1.0 + rate * pmtType) * fvifa)
    }

  val pv: FunctionSpec[BigDecimal] { type Args = TvmArgs } =
    FunctionSpec.simple[BigDecimal, TvmArgs]("PV", Arity.Range(3, 5)) { (args, ctx) =>
      val (rateExpr, nperExpr, pmtExpr, fvOpt, typeOpt) = args
      for
        rate <- ctx.evalExpr(rateExpr).map(_.toDouble)
        nper <- ctx.evalExpr(nperExpr).map(_.toDouble)
        pmt <- ctx.evalExpr(pmtExpr).map(_.toDouble)
        fv <- fvOpt match
          case Some(expr) => ctx.evalExpr(expr).map(_.toDouble)
          case None => Right(0.0)
        pmtType <- typeOpt match
          case Some(expr) => ctx.evalExpr(expr).map(v => if v.toInt != 0 then 1 else 0)
          case None => Right(0)
      yield
        if math.abs(rate) < 1e-10 then BigDecimal(-fv - pmt * nper)
        else
          val pvif = math.pow(1.0 + rate, nper)
          val fvifa = (pvif - 1.0) / rate
          BigDecimal((-fv - pmt * (1.0 + rate * pmtType) * fvifa) / pvif)
    }

  val nper: FunctionSpec[BigDecimal] { type Args = TvmArgs } =
    FunctionSpec.simple[BigDecimal, TvmArgs]("NPER", Arity.Range(3, 5)) { (args, ctx) =>
      val (rateExpr, pmtExpr, pvExpr, fvOpt, typeOpt) = args
      for
        rate <- ctx.evalExpr(rateExpr).map(_.toDouble)
        pmt <- ctx.evalExpr(pmtExpr).map(_.toDouble)
        pv <- ctx.evalExpr(pvExpr).map(_.toDouble)
        fv <- fvOpt match
          case Some(expr) => ctx.evalExpr(expr).map(_.toDouble)
          case None => Right(0.0)
        pmtType <- typeOpt match
          case Some(expr) => ctx.evalExpr(expr).map(v => if v.toInt != 0 then 1 else 0)
          case None => Right(0)
      yield
        if math.abs(rate) < 1e-10 then
          if pmt == 0.0 then BigDecimal(Double.NaN)
          else BigDecimal(-(pv + fv) / pmt)
        else
          val ratep1 = 1.0 + rate
          val numerator = -fv * rate + pmt * (1.0 + rate * pmtType)
          val denominator = pv * rate + pmt * (1.0 + rate * pmtType)
          BigDecimal(math.log(numerator / denominator) / math.log(ratep1))
    }

  val rate: FunctionSpec[BigDecimal] { type Args = RateArgs } =
    FunctionSpec.simple[BigDecimal, RateArgs]("RATE", Arity.Range(3, 6)) { (args, ctx) =>
      val (nperExpr, pmtExpr, pvExpr, fvOpt, typeOpt, guessOpt) = args
      for
        nper <- ctx.evalExpr(nperExpr).map(_.toDouble)
        pmt <- ctx.evalExpr(pmtExpr).map(_.toDouble)
        pv <- ctx.evalExpr(pvExpr).map(_.toDouble)
        fv <- fvOpt match
          case Some(expr) => ctx.evalExpr(expr).map(_.toDouble)
          case None => Right(0.0)
        pmtType <- typeOpt match
          case Some(expr) => ctx.evalExpr(expr).map(v => if v.toInt != 0 then 1 else 0)
          case None => Right(0)
        guess <- guessOpt match
          case Some(expr) => ctx.evalExpr(expr).map(_.toDouble)
          case None => Right(0.1)
        result <- {
          val maxIter = 100
          val tolerance = 1e-7

          def f(rate: Double): Double =
            if math.abs(rate) < 1e-10 then pv + pmt * nper + fv
            else
              val pvif = math.pow(1.0 + rate, nper)
              val fvifa = (pvif - 1.0) / rate
              pv * pvif + pmt * (1.0 + rate * pmtType) * fvifa + fv

          def df(rate: Double): Double =
            if math.abs(rate) < 1e-10 then pv * nper + pmt * nper * (nper - 1.0) / 2.0
            else
              val pvif = math.pow(1.0 + rate, nper)
              val dpvif = nper * math.pow(1.0 + rate, nper - 1.0)
              val fvifa = (pvif - 1.0) / rate
              val dfvifa = (dpvif * rate - (pvif - 1.0)) / (rate * rate)
              pv * dpvif + pmt * pmtType * fvifa + pmt * (1.0 + rate * pmtType) * dfvifa

          @annotation.tailrec
          def loop(iter: Int, r: Double): Either[EvalError, BigDecimal] =
            if iter >= maxIter then
              Left(
                EvalError.EvalFailed(
                  s"RATE did not converge after $maxIter iterations",
                  Some("RATE(nper, pmt, pv, [fv], [type], [guess])")
                )
              )
            else
              val fVal = f(r)
              val dfVal = df(r)
              if math.abs(dfVal) < 1e-14 then
                Left(
                  EvalError.EvalFailed(
                    "RATE derivative is zero; cannot continue iteration",
                    Some("RATE(nper, pmt, pv, [fv], [type], [guess])")
                  )
                )
              else
                val next = r - fVal / dfVal
                if math.abs(next - r) <= tolerance then Right(BigDecimal(next))
                else loop(iter + 1, next)

          loop(0, guess)
        }
      yield result
    }
