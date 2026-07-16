package com.tjclp.xl.formula.functions

import com.tjclp.xl.formula.ast.{TExpr, ExprValue}
import com.tjclp.xl.formula.eval.{EvalError, Evaluator}
import com.tjclp.xl.formula.parser.ParseError
import com.tjclp.xl.formula.{Clock, Arity}

import com.tjclp.xl.cells.CellError

trait FunctionSpecsMath extends FunctionSpecsBase:
  private def roundToDigits(
    value: BigDecimal,
    numDigits: Int,
    mode: BigDecimal.RoundingMode.Value
  ): BigDecimal =
    if numDigits >= 0 then value.setScale(numDigits, mode)
    else
      val scale = math.pow(10, -numDigits).toLong
      val divided = value / scale
      val rounded = divided.setScale(0, mode)
      rounded * scale

  val abs: FunctionSpec[BigDecimal] { type Args = UnaryNumeric } =
    FunctionSpec.simple[BigDecimal, UnaryNumeric](
      "ABS",
      Arity.one,
      flags = FunctionFlags(returnsNumeric = true)
    ) { (expr, ctx) =>
      ctx.evalExpr(expr).map(_.abs)
    }

  val sqrt: FunctionSpec[BigDecimal] { type Args = UnaryNumeric } =
    FunctionSpec.simple[BigDecimal, UnaryNumeric](
      "SQRT",
      Arity.one,
      flags = FunctionFlags(returnsNumeric = true)
    ) { (expr, ctx) =>
      ctx.evalExpr(expr).flatMap { value =>
        if value < 0 then
          // GH-344: Excel #NUM! (op-level classification)
          Left(
            EvalError.ErrorValue(
              CellError.Num,
              Some(s"SQRT: cannot take square root of negative number ($value)")
            )
          )
        else Right(BigDecimal(Math.sqrt(value.toDouble)))
      }
    }

  val round: FunctionSpec[BigDecimal] { type Args = BinaryNumeric } =
    FunctionSpec.simple[BigDecimal, BinaryNumeric](
      "ROUND",
      Arity.two,
      flags = FunctionFlags(returnsNumeric = true)
    ) { (args, ctx) =>
      val (valueExpr, numDigitsExpr) = args
      for
        value <- ctx.evalExpr(valueExpr)
        numDigits <- ctx.evalExpr(numDigitsExpr)
      yield roundToDigits(value, numDigits.toInt, BigDecimal.RoundingMode.HALF_UP)
    }

  val roundUp: FunctionSpec[BigDecimal] { type Args = BinaryNumeric } =
    FunctionSpec.simple[BigDecimal, BinaryNumeric](
      "ROUNDUP",
      Arity.two,
      flags = FunctionFlags(returnsNumeric = true)
    ) { (args, ctx) =>
      val (valueExpr, numDigitsExpr) = args
      for
        value <- ctx.evalExpr(valueExpr)
        numDigits <- ctx.evalExpr(numDigitsExpr)
      yield
        val mode =
          if value >= 0 then BigDecimal.RoundingMode.CEILING
          else BigDecimal.RoundingMode.FLOOR
        roundToDigits(value, numDigits.toInt, mode)
    }

  val roundDown: FunctionSpec[BigDecimal] { type Args = BinaryNumeric } =
    FunctionSpec.simple[BigDecimal, BinaryNumeric](
      "ROUNDDOWN",
      Arity.two,
      flags = FunctionFlags(returnsNumeric = true)
    ) { (args, ctx) =>
      val (valueExpr, numDigitsExpr) = args
      for
        value <- ctx.evalExpr(valueExpr)
        numDigits <- ctx.evalExpr(numDigitsExpr)
      yield
        val mode =
          if value >= 0 then BigDecimal.RoundingMode.FLOOR
          else BigDecimal.RoundingMode.CEILING
        roundToDigits(value, numDigits.toInt, mode)
    }

  val mod: FunctionSpec[BigDecimal] { type Args = BinaryNumeric } =
    FunctionSpec.simple[BigDecimal, BinaryNumeric](
      "MOD",
      Arity.two,
      flags = FunctionFlags(returnsNumeric = true)
    ) { (args, ctx) =>
      val (numberExpr, divisorExpr) = args
      for
        number <- ctx.evalExpr(numberExpr)
        divisor <- ctx.evalExpr(divisorExpr)
        result <-
          if divisor == 0 then
            // GH-344: Excel #DIV/0! (op-level classification)
            Left(
              EvalError
                .ErrorValue(CellError.Div0, Some(s"MOD: division by zero ($number, $divisor)"))
            )
          else
            val quotient = (number / divisor).setScale(0, BigDecimal.RoundingMode.FLOOR)
            Right(number - divisor * quotient)
      yield result
    }

  val power: FunctionSpec[BigDecimal] { type Args = BinaryNumeric } =
    FunctionSpec.simple[BigDecimal, BinaryNumeric](
      "POWER",
      Arity.two,
      flags = FunctionFlags(returnsNumeric = true)
    ) { (args, ctx) =>
      val (numberExpr, powerExpr) = args
      for
        number <- ctx.evalExpr(numberExpr)
        power <- ctx.evalExpr(powerExpr)
      yield BigDecimal(Math.pow(number.toDouble, power.toDouble))
    }

  val log: FunctionSpec[BigDecimal] { type Args = BinaryNumericOpt } =
    FunctionSpec.simple[BigDecimal, BinaryNumericOpt](
      "LOG",
      Arity.Range(1, 2),
      flags = FunctionFlags(returnsNumeric = true)
    ) { (args, ctx) =>
      val (numberExpr, baseExprOpt) = args
      val baseExpr = baseExprOpt.getOrElse(TExpr.Lit(BigDecimal(10)))
      for
        number <- ctx.evalExpr(numberExpr)
        base <- ctx.evalExpr(baseExpr)
        result <-
          // GH-344: Excel codes at source — non-positive domain is #NUM!, base 1 is #DIV/0!
          if number <= 0 then
            Left(
              EvalError.ErrorValue(CellError.Num, Some(s"LOG: argument must be positive ($number)"))
            )
          else if base <= 0 then
            Left(EvalError.ErrorValue(CellError.Num, Some(s"LOG: base must be positive ($base)")))
          else if base == 1 then
            Left(EvalError.ErrorValue(CellError.Div0, Some("LOG: base cannot be 1")))
          else Right(BigDecimal(Math.log(number.toDouble) / Math.log(base.toDouble)))
      yield result
    }

  val ln: FunctionSpec[BigDecimal] { type Args = UnaryNumeric } =
    FunctionSpec.simple[BigDecimal, UnaryNumeric](
      "LN",
      Arity.one,
      flags = FunctionFlags(returnsNumeric = true)
    ) { (expr, ctx) =>
      ctx.evalExpr(expr).flatMap { value =>
        if value <= 0 then
          // GH-344: Excel #NUM! (op-level classification)
          Left(EvalError.ErrorValue(CellError.Num, Some(s"LN: argument must be positive ($value)")))
        else Right(BigDecimal(Math.log(value.toDouble)))
      }
    }

  val exp: FunctionSpec[BigDecimal] { type Args = UnaryNumeric } =
    FunctionSpec.simple[BigDecimal, UnaryNumeric](
      "EXP",
      Arity.one,
      flags = FunctionFlags(returnsNumeric = true)
    ) { (expr, ctx) =>
      ctx.evalExpr(expr).map { value =>
        BigDecimal(Math.exp(value.toDouble))
      }
    }

  val floor: FunctionSpec[BigDecimal] { type Args = BinaryNumeric } =
    FunctionSpec.simple[BigDecimal, BinaryNumeric](
      "FLOOR",
      Arity.two,
      flags = FunctionFlags(returnsNumeric = true)
    ) { (args, ctx) =>
      val (numberExpr, significanceExpr) = args
      for
        number <- ctx.evalExpr(numberExpr)
        significance <- ctx.evalExpr(significanceExpr)
        result <-
          if significance == 0 then
            Left(
              EvalError.EvalFailed(
                "FLOOR: significance cannot be zero",
                Some(s"FLOOR($number, $significance)")
              )
            )
          else if (number > 0 && significance < 0) || (number < 0 && significance > 0) then
            Left(
              EvalError.EvalFailed(
                "FLOOR: number and significance must have same sign",
                Some(s"FLOOR($number, $significance)")
              )
            )
          else
            val quotient = (number / significance).setScale(0, BigDecimal.RoundingMode.FLOOR)
            Right(quotient * significance)
      yield result
    }

  val ceiling: FunctionSpec[BigDecimal] { type Args = BinaryNumeric } =
    FunctionSpec.simple[BigDecimal, BinaryNumeric](
      "CEILING",
      Arity.two,
      flags = FunctionFlags(returnsNumeric = true)
    ) { (args, ctx) =>
      val (numberExpr, significanceExpr) = args
      for
        number <- ctx.evalExpr(numberExpr)
        significance <- ctx.evalExpr(significanceExpr)
        result <-
          if significance == 0 then
            Left(
              EvalError.EvalFailed(
                "CEILING: significance cannot be zero",
                Some(s"CEILING($number, $significance)")
              )
            )
          else if (number > 0 && significance < 0) || (number < 0 && significance > 0) then
            Left(
              EvalError.EvalFailed(
                "CEILING: number and significance must have same sign",
                Some(s"CEILING($number, $significance)")
              )
            )
          else
            val quotient =
              (number / significance).setScale(0, BigDecimal.RoundingMode.CEILING)
            Right(quotient * significance)
      yield result
    }

  val mround: FunctionSpec[BigDecimal] { type Args = BinaryNumeric } =
    FunctionSpec.simple[BigDecimal, BinaryNumeric](
      "MROUND",
      Arity.two,
      flags = FunctionFlags(returnsNumeric = true)
    ) { (args, ctx) =>
      val (numberExpr, multipleExpr) = args
      for
        number <- ctx.evalExpr(numberExpr)
        multiple <- ctx.evalExpr(multipleExpr)
        result <-
          // Excel: MROUND(x, 0) = 0 — unlike CEILING/FLOOR, a zero multiple is NOT an error
          if multiple == 0 then Right(BigDecimal(0))
          else if (number > 0 && multiple < 0) || (number < 0 && multiple > 0) then
            Left(
              EvalError.EvalFailed(
                "MROUND: number and multiple must have same sign",
                Some(s"MROUND($number, $multiple)")
              )
            )
          else
            // multiple * ROUND(number/multiple, 0); HALF_UP matches Excel's half-away-from-zero
            // on the sign-equal domain (7.5/5 = 1.5 -> 2 -> 10; -7.5/-5 = 1.5 -> 2 -> -10)
            val quotient = (number / multiple).setScale(0, BigDecimal.RoundingMode.HALF_UP)
            Right(quotient * multiple)
      yield result
    }

  val trunc: FunctionSpec[BigDecimal] { type Args = BinaryNumericOpt } =
    FunctionSpec.simple[BigDecimal, BinaryNumericOpt](
      "TRUNC",
      Arity.Range(1, 2),
      flags = FunctionFlags(returnsNumeric = true)
    ) { (args, ctx) =>
      val (numberExpr, numDigitsExprOpt) = args
      val numDigitsExpr = numDigitsExprOpt.getOrElse(TExpr.Lit(BigDecimal(0)))
      for
        number <- ctx.evalExpr(numberExpr)
        numDigits <- ctx.evalExpr(numDigitsExpr)
      yield roundToDigits(number, numDigits.toInt, BigDecimal.RoundingMode.DOWN)
    }

  val sign: FunctionSpec[BigDecimal] { type Args = UnaryNumeric } =
    FunctionSpec.simple[BigDecimal, UnaryNumeric](
      "SIGN",
      Arity.one,
      flags = FunctionFlags(returnsNumeric = true)
    ) { (expr, ctx) =>
      ctx.evalExpr(expr).map { value =>
        if value > 0 then BigDecimal(1)
        else if value < 0 then BigDecimal(-1)
        else BigDecimal(0)
      }
    }

  val int: FunctionSpec[BigDecimal] { type Args = UnaryNumeric } =
    FunctionSpec.simple[BigDecimal, UnaryNumeric](
      "INT",
      Arity.one,
      flags = FunctionFlags(returnsNumeric = true)
    ) { (expr, ctx) =>
      ctx.evalExpr(expr).map(_.setScale(0, BigDecimal.RoundingMode.FLOOR))
    }

  val pi: FunctionSpec[BigDecimal] { type Args = NoArgs } =
    FunctionSpec.simple[BigDecimal, NoArgs](
      "PI",
      Arity.none,
      flags = FunctionFlags(returnsNumeric = true)
    ) { (_, _) =>
      Right(BigDecimal(Math.PI))
    }
