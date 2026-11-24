package com.tjclp.xl.formula.functions

import com.tjclp.xl.formula.ast.{TExpr, ExprValue}
import com.tjclp.xl.formula.eval.EvalError
import com.tjclp.xl.formula.parser.ParseError
import com.tjclp.xl.formula.{Clock, Arity}

import com.tjclp.xl.addressing.CellRange
import com.tjclp.xl.cells.{Cell, CellError, CellValue}
import java.time.LocalDate
import java.time.temporal.ChronoUnit
import scala.util.control.NonFatal

/**
 * GH-388: containment for numeric blowups in financial math.
 *
 * Newton–Raphson divergence can push BigDecimal scales past `Int.MaxValue`
 * (`java.lang.ArithmeticException` from `checkScale`), and NaN/Infinity doubles explode in
 * `BigDecimal.apply` (`NumberFormatException`). Totality requires those to surface as a contained
 * per-cell result — never as an exception unwinding `recalculate()`. Same precedent as the
 * `ArrayArithmetic.pow` guard. #670(h): the result is Excel's `#NUM!` (what it caches for an
 * overflowing or undefined financial result), like the non-convergence path.
 */
private[functions] object NumericGuard:
  def contained[A](fn: String, usage: String)(
    body: => Either[EvalError, A]
  ): Either[EvalError, A] =
    try body
    catch
      case NonFatal(_) =>
        Left(EvalError.ErrorValue(CellError.Num, Some(s"$fn diverged (numeric overflow): $usage")))

/**
 * #954: the fallback IRR and XIRR use when Newton from the guess fails (no convergence, zero slope,
 * numeric overflow). Excel still finds a rate for flows such as -100, 50 (-50%), where a Newton
 * step from 10% overshoots past -100% and runs away.
 *
 * Two stages, both bounded:
 *   1. [[bracket]] samples the NPV on a grid of `ln(1 + rate)` around the guess, outward one step
 *      at a time on each side, and returns the first interval whose ends differ in sign: the root
 *      nearest the guess. Rates stay above -100%, where the NPV is continuous, so a sign change
 *      always holds a root. Sampling uses a `Double` NPV; non-finite samples are skipped.
 *   2. [[refine]] narrows that bracket with safeguarded Newton (`rtsafe`): a Newton step is taken
 *      only when it lands inside the bracket, otherwise the bracket is bisected. It stops at the
 *      caller's tolerance, the same one Newton uses, within [[MaxRefineIterations]] steps.
 *
 * Flows whose NPV never changes sign have no rate, and the caller keeps its `#NUM!`.
 */
private[functions] object RateSolver:
  // ±20 in ln(1 + rate) around the guess: from -99.9999998% up to ~e^20, 401 samples at most
  private val Step = 0.1
  private val MaxSteps = 200
  private val MaxRefineIterations = 100

  /** Root of `f` nearest `guess`, or None when no sign change is found or the guess is ≤ -100%. */
  def fallback(
    guess: BigDecimal,
    fDouble: Double => Double,
    f: BigDecimal => BigDecimal,
    df: BigDecimal => BigDecimal,
    tolerance: BigDecimal
  ): Option[BigDecimal] =
    bracket(guess.toDouble, fDouble).flatMap { case (lo, hi) =>
      refine(BigDecimal(lo), BigDecimal(hi), f, df, tolerance)
    }

  /** The sign-change interval `(lo, hi)` (as rates, `lo < hi`) nearest `guess`. */
  def bracket(guess: Double, f: Double => Double): Option[(Double, Double)] =
    val center = math.log1p(guess)
    if guess <= -1.0 || !center.isFinite then None
    else
      def rateAt(j: Int): Double = math.expm1(center + j * Step)
      def crosses(fa: Double, fb: Double): Boolean =
        fa.isFinite && fb.isFinite && (fa == 0.0 || fb == 0.0 || (fa < 0.0) != (fb < 0.0))

      // Outward from the guess, one step above then one below, each point evaluated once: a
      // root near the guess costs a handful of NPV evaluations, not the whole grid.
      @annotation.tailrec
      def scan(k: Int, fAbove: Double, fBelow: Double): Option[(Double, Double)] =
        if k > MaxSteps then None
        else
          val fUp = f(rateAt(k))
          if crosses(fAbove, fUp) then Some((rateAt(k - 1), rateAt(k)))
          else
            val fDown = f(rateAt(-k))
            if crosses(fDown, fBelow) then Some((rateAt(-k), rateAt(-(k - 1))))
            else scan(k + 1, fUp, fDown)

      val fCenter = f(rateAt(0))
      scan(1, fCenter, fCenter)

  /** Safeguarded Newton inside `[lo, hi]`, whose ends must bracket a root of `f`. */
  def refine(
    lo: BigDecimal,
    hi: BigDecimal,
    f: BigDecimal => BigDecimal,
    df: BigDecimal => BigDecimal,
    tolerance: BigDecimal
  ): Option[BigDecimal] =
    val zero = BigDecimal(0)
    val two = BigDecimal(2)
    val fLo = f(lo)
    val fHi = f(hi)
    if fLo == zero then Some(lo)
    else if fHi == zero then Some(hi)
    else if fLo.signum == fHi.signum then None
    else
      // Orient so that f(neg) < 0 < f(pos); neg may lie above pos.
      val (neg0, pos0) = if fLo < zero then (lo, hi) else (hi, lo)

      @annotation.tailrec
      def loop(iter: Int, x: BigDecimal, neg: BigDecimal, pos: BigDecimal): Option[BigDecimal] =
        if iter >= MaxRefineIterations then None
        else
          val fx = f(x)
          if fx == zero then Some(x)
          else
            val (nNeg, nPos) = if fx < zero then (x, pos) else (neg, x)
            val low = nNeg.min(nPos)
            val high = nNeg.max(nPos)
            val dfx = df(x)
            val newton = if dfx == zero then None else Some(x - fx / dfx)
            val next = newton.filter(n => n > low && n < high).getOrElse((low + high) / two)
            if (next - x).abs <= tolerance || (high - low) <= tolerance then Some(next)
            else loop(iter + 1, next, nNeg, nPos)

      loop(0, (lo + hi) / two, neg0, pos0)

/**
 * #670(h): a financial function's argument failure as Excel's `#NUM!` — XIRR/XNPV over values and
 * dates of different counts, IRR/XIRR without a sign change, a Newton step with no slope — so the
 * cell caches the error value (visible to IFERROR and ERROR.TYPE) instead of reading as uncached.
 */
private[functions] def numError(detail: String, usage: String): EvalError =
  EvalError.ErrorValue(CellError.Num, Some(s"$detail: $usage"))

trait FunctionSpecsFinancialCashflow extends FunctionSpecsBase:
  private def rangeValues(
    location: TExpr.RangeLocation,
    ctx: EvalContext
  ): Either[EvalError, Vector[Cell]] =
    ctx.resolveRange(location).flatMap { case (targetSheet, range) =>
      val readCell = rangeCellReader(targetSheet, ctx)
      range.cellsRowMajor.foldLeft[Either[EvalError, Vector[Cell]]](Right(Vector.empty)) {
        (result, at) =>
          result.flatMap(cells => readCell(at).map(value => cells :+ Cell(at, value)))
      }
    }

  /**
   * Resolve before applying the function's skip policy: an unevaluated formula is never a blank.
   * NPV's documented range policy also skips Excel error VALUES; host failures stay Left.
   */
  private def numericValues(
    function: String,
    location: TExpr.RangeLocation,
    ctx: EvalContext
  ): Either[EvalError, List[BigDecimal]] =
    rangeValues(location, ctx).flatMap { cells =>
      cells
        .foldLeft[Either[EvalError, List[BigDecimal]]](Right(Nil)) { (result, cell) =>
          result.flatMap { values =>
            cell.value match
              case CellValue.Empty | CellValue.Text(_) | CellValue.RichText(_) |
                  CellValue.Bool(_) =>
                Right(values)
              case CellValue.Error(_) if function == "NPV" => Right(values)
              case CellValue.Error(error) => Left(EvalError.ErrorValue(error, Some(function)))
              case _ =>
                TExpr
                  .decodeNumeric(cell)
                  .left
                  .map(error => EvalError.CodecFailed(cell.ref, error))
                  .map(_ :: values)
          }
        }
        .map(_.reverse)
    }

  /** GH-499: validate each original pair before filtering aligned blank rows. */
  private def datedValues(
    function: String,
    valuesLocation: TExpr.RangeLocation,
    datesLocation: TExpr.RangeLocation,
    ctx: EvalContext
  ): Either[EvalError, (List[BigDecimal], List[LocalDate])] =
    // Excel: differing counts are #NUM!, a value or date it cannot read #VALUE!
    def mismatched(message: String): EvalError =
      EvalError.ErrorValue(CellError.Num, Some(s"$function: $message"))
    def invalid(message: String): EvalError =
      EvalError.ErrorValue(CellError.Value, Some(s"$function: $message"))

    for
      values <- rangeValues(valuesLocation, ctx)
      dates <- rangeValues(datesLocation, ctx)
      pairs <-
        if values.size != dates.size then
          Left(mismatched("values and dates must have the same length"))
        else
          values
            .zip(dates)
            .foldLeft[Either[EvalError, List[(BigDecimal, LocalDate)]]](Right(Nil)) {
              case (result, (valueCell, dateCell)) =>
                result.flatMap { acc =>
                  if valueCell.value == CellValue.Empty && dateCell.value == CellValue.Empty then
                    Right(acc)
                  else if valueCell.value == CellValue.Empty || dateCell.value == CellValue.Empty
                  then
                    Left(
                      mismatched(
                        "values and dates must have the same length and aligned positions " +
                          s"(${valueCell.ref.toA1}, ${dateCell.ref.toA1})"
                      )
                    )
                  else
                    for
                      value <- valueCell.value match
                        case CellValue.Error(error) =>
                          Left(EvalError.ErrorValue(error, Some(function)))
                        case _ =>
                          TExpr
                            .decodeNumeric(valueCell)
                            .left
                            .map(_ => invalid(s"invalid cash flow at ${valueCell.ref.toA1}"))
                      date <- dateCell.value match
                        case CellValue.Error(error) =>
                          Left(EvalError.ErrorValue(error, Some(function)))
                        case _ =>
                          TExpr
                            .decodeAsDate(dateCell)
                            .left
                            .map(_ => invalid(s"invalid date at ${dateCell.ref.toA1}"))
                    yield (value, date) :: acc
                }
            }
            .map(_.reverse)
    yield pairs.unzip

  val npv: FunctionSpec[BigDecimal] { type Args = NpvArgs } =
    FunctionSpec.simple[BigDecimal, NpvArgs](
      "NPV",
      Arity.two,
      flags = FunctionFlags(returnsNumeric = true, lift = ArrayLift.all)
    ) { (args, ctx) =>
      val (rateExpr, range) = args
      ctx.evalExpr(rateExpr).flatMap { rate =>
        val onePlusR = BigDecimal(1) + rate
        if onePlusR == BigDecimal(0) then
          // #670(h): #NUM!, a cached error value (LibreOffice: NPV(-1,…) is #NUM!)
          Left(numError("NPV: rate = -1 would require division by zero", "NPV(rate, values)"))
        else
          NumericGuard.contained("NPV", "NPV(rate, values)") {
            numericValues("NPV", range, ctx).map { cashFlows =>
              cashFlows.zipWithIndex.foldLeft(BigDecimal(0)) { case (acc, (cf, idx)) =>
                val period = idx + 1
                acc + cf / onePlusR.pow(period)
              }
            }
          }
      }
    }

  val irr: FunctionSpec[BigDecimal] { type Args = IrrArgs } =
    FunctionSpec.simple[BigDecimal, IrrArgs](
      "IRR",
      Arity.Range(1, 2),
      flags = FunctionFlags(returnsNumeric = true)
    ) { (args, ctx) =>
      val (range, guessOpt) = args
      numericValues("IRR", range, ctx).flatMap { cashFlows =>
        if cashFlows.isEmpty || !cashFlows.exists(_ < 0) || !cashFlows.exists(_ > 0) then
          Left(
            numError(
              "IRR requires at least one positive and one negative cash flow",
              "IRR(values[, guess])"
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
                // GH-344: Excel #NUM! for non-convergence (op-level classification)
                Left(
                  EvalError.ErrorValue(
                    CellError.Num,
                    Some(s"IRR did not converge after $maxIter iterations")
                  )
                )
              else
                val f = npvAt(r)
                val df = dNpvAt(r)
                if df == BigDecimal(0) then
                  Left(
                    numError(
                      "IRR derivative is zero; cannot continue iteration",
                      "IRR(values[, guess])"
                    )
                  )
                else
                  val next = r - f / df
                  if (next - r).abs <= tolerance then Right(next)
                  else loop(iter + 1, next)

            // Horner in x = 1/(1+rate), for the fallback's bracket scan. The flows are converted
            // once, and only if the fallback runs: BigDecimal.toDouble per sample dominated the scan.
            lazy val flowsLastFirst: Array[Double] =
              cashFlows.reverseIterator.map(_.toDouble).toArray
            def npvDouble(rate: Double): Double =
              val x = 1.0 / (1.0 + rate)
              flowsLastFirst.foldLeft(0.0)((acc, cf) => cf + x * acc)

            val usage = "IRR(values[, guess])"
            // #954: Newton first, so every rate it already finds is unchanged
            NumericGuard.contained("IRR", usage)(loop(0, guess0)).left.flatMap { newtonError =>
              NumericGuard.contained("IRR", usage)(
                RateSolver
                  .fallback(guess0, npvDouble, npvAt, dNpvAt, tolerance)
                  .toRight(newtonError)
              )
            }
          }
      }
    }

  val xnpv: FunctionSpec[BigDecimal] { type Args = XnpvArgs } =
    FunctionSpec.simple[BigDecimal, XnpvArgs](
      "XNPV",
      Arity.three,
      flags = FunctionFlags(returnsNumeric = true)
    ) { (args, ctx) =>
      val (rateExpr, valuesRange, datesRange) = args
      for
        rate <- ctx.evalExpr(rateExpr)
        flows <- datedValues("XNPV", valuesRange, datesRange, ctx)
        (values, dates) = flows
        result <- {
          if values.isEmpty || dates.isEmpty then
            Left(
              numError(
                "XNPV requires non-empty values and dates ranges",
                "XNPV(rate, values, dates)"
              )
            )
          else if values.length != dates.length then
            Left(
              numError(
                s"XNPV: values (${values.length}) and dates (${dates.length}) must have same length",
                "XNPV(rate, values, dates)"
              )
            )
          else
            dates match
              case date0 :: _ =>
                NumericGuard.contained("XNPV", "XNPV(rate, values, dates)") {
                  val onePlusR = BigDecimal(1) + rate
                  val npv =
                    values.zip(dates).foldLeft(BigDecimal(0)) { case (acc, (value, date)) =>
                      val daysDiff = ChronoUnit.DAYS.between(date0, date)
                      val yearFraction = BigDecimal(daysDiff) / BigDecimal(365)
                      val discountFactor = math.pow(onePlusR.toDouble, yearFraction.toDouble)
                      acc + value / BigDecimal(discountFactor)
                    }
                  Right(npv)
                }
              case Nil =>
                Left(numError("XNPV: dates cannot be empty", "XNPV(rate, values, dates)"))
        }
      yield result
    }

  val xirr: FunctionSpec[BigDecimal] { type Args = XirrArgs } =
    FunctionSpec.simple[BigDecimal, XirrArgs](
      "XIRR",
      Arity.Range(2, 3),
      flags = FunctionFlags(returnsNumeric = true)
    ) { (args, ctx) =>
      val (valuesRange, datesRange, guessOpt) = args
      for
        flows <- datedValues("XIRR", valuesRange, datesRange, ctx)
        (values, dates) = flows
        result <- {
          if values.isEmpty || dates.isEmpty then
            Left(
              numError(
                "XIRR requires non-empty values and dates ranges",
                "XIRR(values, dates[, guess])"
              )
            )
          else if values.length != dates.length then
            Left(
              numError(
                s"XIRR: values (${values.length}) and dates (${dates.length}) must have same length",
                "XIRR(values, dates[, guess])"
              )
            )
          else if !values.exists(_ < 0) || !values.exists(_ > 0) then
            Left(
              numError(
                "XIRR requires at least one positive and one negative cash flow",
                "XIRR(values, dates[, guess])"
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
                      // GH-344: Excel #NUM! for non-convergence (op-level classification)
                      Left(
                        EvalError.ErrorValue(
                          CellError.Num,
                          Some(s"XIRR did not converge after $maxIter iterations")
                        )
                      )
                    else
                      val f = xnpvAt(r)
                      val df = dXnpvAt(r)
                      if df.abs < BigDecimal("1e-10") then
                        Left(
                          numError(
                            "XIRR derivative is near zero; cannot continue iteration",
                            "XIRR(values, dates[, guess])"
                          )
                        )
                      else
                        val next = r - f / df
                        if (next - r).abs <= tolerance then Right(next)
                        else loop(iter + 1, next)

                  // Converted once, and only if the fallback runs
                  lazy val cfDouble: Array[Double] = values.map(_.toDouble).toArray
                  lazy val yfDouble: Array[Double] = yearFractions.map(_.toDouble).toArray
                  def xnpvDouble(rate: Double): Double =
                    cfDouble.indices.foldLeft(0.0) { (acc, i) =>
                      acc + cfDouble(i) / math.pow(1.0 + rate, yfDouble(i))
                    }

                  val usage = "XIRR(values, dates[, guess])"
                  // #954: Newton first, so every rate it already finds is unchanged
                  NumericGuard.contained("XIRR", usage)(loop(0, guess0)).left.flatMap {
                    newtonError =>
                      NumericGuard.contained("XIRR", usage)(
                        RateSolver
                          .fallback(guess0, xnpvDouble, xnpvAt, dXnpvAt, tolerance)
                          .toRight(newtonError)
                      )
                  }
                case Nil =>
                  Left(numError("XIRR: dates cannot be empty", "XIRR(values, dates[, guess])"))
            }
        }
      yield result
    }
