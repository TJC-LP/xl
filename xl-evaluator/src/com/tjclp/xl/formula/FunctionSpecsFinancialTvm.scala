package com.tjclp.xl.formula

trait FunctionSpecsFinancialTvm extends FunctionSpecsBase:
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
