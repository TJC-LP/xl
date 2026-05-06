package com.tjclp.xl.formula.functions

import com.tjclp.xl.formula.ast.{TExpr, ExprValue}
import com.tjclp.xl.formula.eval.{EvalError, Evaluator}
import com.tjclp.xl.formula.parser.ParseError
import com.tjclp.xl.formula.{Clock, Arity}

trait FunctionSpecsText extends FunctionSpecsBase:
  val concatenate: FunctionSpec[String] { type Args = TextList } =
    FunctionSpec.simple[String, TextList]("CONCATENATE", Arity.atLeastOne) { (args, ctx) =>
      args.foldLeft[Either[EvalError, String]](Right("")) { (accEither, expr) =>
        for
          acc <- accEither
          value <- ctx.evalExpr(expr)
        yield acc + value
      }
    }

  val left: FunctionSpec[String] { type Args = BinaryTextInt } =
    FunctionSpec.simple[String, BinaryTextInt]("LEFT", Arity.two) { (args, ctx) =>
      val (textExpr, nExpr) = args
      for
        text <- ctx.evalExpr(textExpr)
        nValue <- ctx.evalExpr(nExpr)
        result <-
          if nValue < 0 then
            Left(EvalError.EvalFailed(s"LEFT: n must be non-negative, got $nValue"))
          else if nValue >= text.length then Right(text)
          else Right(text.take(nValue))
      yield result
    }

  val right: FunctionSpec[String] { type Args = BinaryTextInt } =
    FunctionSpec.simple[String, BinaryTextInt]("RIGHT", Arity.two) { (args, ctx) =>
      val (textExpr, nExpr) = args
      for
        text <- ctx.evalExpr(textExpr)
        nValue <- ctx.evalExpr(nExpr)
        result <-
          if nValue < 0 then
            Left(EvalError.EvalFailed(s"RIGHT: n must be non-negative, got $nValue"))
          else if nValue >= text.length then Right(text)
          else Right(text.takeRight(nValue))
      yield result
    }

  val len: FunctionSpec[BigDecimal] { type Args = UnaryText } =
    FunctionSpec.simple[BigDecimal, UnaryText]("LEN", Arity.one) { (expr, ctx) =>
      ctx.evalExpr(expr).map(text => BigDecimal(text.length))
    }

  val upper: FunctionSpec[String] { type Args = UnaryText } =
    FunctionSpec.simple[String, UnaryText]("UPPER", Arity.one) { (expr, ctx) =>
      ctx.evalExpr(expr).map(_.toUpperCase)
    }

  val lower: FunctionSpec[String] { type Args = UnaryText } =
    FunctionSpec.simple[String, UnaryText]("LOWER", Arity.one) { (expr, ctx) =>
      ctx.evalExpr(expr).map(_.toLowerCase)
    }

  val trim: FunctionSpec[String] { type Args = UnaryText } =
    FunctionSpec.simple[String, UnaryText]("TRIM", Arity.one) { (textExpr, ctx) =>
      ctx.evalExpr(textExpr).map(trimAsciiSpaces)
    }

  /**
   * Excel TRIM rule: collapse runs of ASCII space (0x20) to a single space and strip leading /
   * trailing 0x20s. All other whitespace (tab, newline, nbsp, BOM, ZWSP) is preserved verbatim.
   */
  private def trimAsciiSpaces(s: String): String =
    s.split(' ').iterator.filter(_.nonEmpty).mkString(" ")

  val mid: FunctionSpec[String] { type Args = TextIntInt } =
    FunctionSpec.simple[String, TextIntInt]("MID", Arity.three) { (args, ctx) =>
      val (textExpr, startExpr, lengthExpr) = args
      for
        text <- ctx.evalExpr(textExpr)
        start <- ctx.evalExpr(startExpr)
        length <- ctx.evalExpr(lengthExpr)
        result <-
          if start < 1 then
            Left(EvalError.EvalFailed(s"MID: start must be >= 1, got $start"))
          else if length < 0 then
            Left(EvalError.EvalFailed(s"MID: length must be >= 0, got $length"))
          else if start > text.length then Right("")
          else
            val from = start - 1
            val to = math.min(from.toLong + length.toLong, text.length.toLong).toInt
            Right(text.substring(from, to))
      yield result
    }

  val find: FunctionSpec[BigDecimal] { type Args = FindArgs } =
    FunctionSpec.simple[BigDecimal, FindArgs]("FIND", Arity.Range(2, 3)) { (_, _) =>
      Left(EvalError.EvalFailed("FIND: not yet implemented"))
    }

  val substitute: FunctionSpec[String] { type Args = SubstituteArgs } =
    FunctionSpec.simple[String, SubstituteArgs]("SUBSTITUTE", Arity.Range(3, 4)) { (_, _) =>
      Left(EvalError.EvalFailed("SUBSTITUTE: not yet implemented"))
    }

  val value: FunctionSpec[BigDecimal] { type Args = UnaryText } =
    FunctionSpec.simple[BigDecimal, UnaryText]("VALUE", Arity.one) { (_, _) =>
      Left(EvalError.EvalFailed("VALUE: not yet implemented"))
    }

  val text: FunctionSpec[String] { type Args = TextArgs } =
    FunctionSpec.simple[String, TextArgs]("TEXT", Arity.two) { (_, _) =>
      Left(EvalError.EvalFailed("TEXT: not yet implemented"))
    }
