package com.tjclp.xl.formula.functions

import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.eval.EvalError
import com.tjclp.xl.formula.Arity

/**
 * GH-115: random number functions.
 *
 * Randomness flows through the explicit [[com.tjclp.xl.formula.Rng]] capability on EvalContext
 * (Clock pattern): production draws from Rng.system, tests and reproducible pipelines from
 * Rng.seeded. Like Excel's volatile functions, every (re)evaluation draws a fresh value; xl
 * re-evaluates volatile formulas on every recalculate.
 */
trait FunctionSpecsRandom extends FunctionSpecsBase:

  /** RAND() — uniformly distributed number in [0, 1). */
  val rand: FunctionSpec[BigDecimal] { type Args = NoArgs } =
    FunctionSpec.simple[BigDecimal, NoArgs](
      "RAND",
      Arity.none,
      flags = FunctionFlags(returnsNumeric = true)
    ) { (_, ctx) =>
      Right(BigDecimal(ctx.rng.nextDouble()))
    }

  /**
   * RANDBETWEEN(bottom, top) — random integer in [bottom, top] inclusive.
   *
   * Excel semantics: bottom must be <= top (else #NUM!); fractional bounds tighten inward (bottom
   * rounds up, top rounds down) since only integers can be returned; negative ranges are fine.
   */
  val randbetween: FunctionSpec[BigDecimal] { type Args = BinaryNumeric } =
    FunctionSpec.simple[BigDecimal, BinaryNumeric](
      "RANDBETWEEN",
      Arity.two,
      flags = FunctionFlags(returnsNumeric = true)
    ) { (args, ctx) =>
      val (bottomExpr, topExpr) = args
      for
        bottom <- ctx.evalExpr(bottomExpr)
        top <- ctx.evalExpr(topExpr)
        result <-
          if bottom > top then
            Left(
              EvalError.EvalFailed(
                s"RANDBETWEEN: bottom must be <= top (#NUM!)",
                Some(s"RANDBETWEEN($bottom, $top)")
              )
            )
          else
            val lo = bottom.setScale(0, BigDecimal.RoundingMode.CEILING)
            val hi = top.setScale(0, BigDecimal.RoundingMode.FLOOR)
            if lo > hi then
              Left(
                EvalError.EvalFailed(
                  s"RANDBETWEEN: no integer between $bottom and $top (#NUM!)",
                  Some(s"RANDBETWEEN($bottom, $top)")
                )
              )
            else
              // draw in [0, span): floor(d * span) <= span - 1 since d < 1, so the result
              // stays within [lo, hi] inclusive
              val span = hi - lo + 1
              val offset =
                (BigDecimal(ctx.rng.nextDouble()) * span).setScale(0, BigDecimal.RoundingMode.FLOOR)
              Right(lo + offset)
      yield result
    }
