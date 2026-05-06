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
    FunctionSpec.simple[BigDecimal, FindArgs]("FIND", Arity.Range(2, 3)) { (args, ctx) =>
      val (findExpr, withinExpr, startOpt) = args
      for
        needle <- ctx.evalExpr(findExpr)
        haystack <- ctx.evalExpr(withinExpr)
        start <- startOpt.fold[Either[EvalError, Int]](Right(1))(e => ctx.evalExpr(e))
        result <-
          if start < 1 then
            Left(EvalError.EvalFailed(s"FIND: start must be >= 1, got $start"))
          else if start > haystack.length + 1 then
            Left(EvalError.EvalFailed(s"FIND: start ($start) is past end of text"))
          else if needle.isEmpty then Right(BigDecimal(start))
          else
            val idx = haystack.indexOf(needle, start - 1)
            if idx < 0 then
              Left(EvalError.EvalFailed(s"FIND: '$needle' not found in '$haystack'"))
            else Right(BigDecimal(idx + 1))
      yield result
    }

  val substitute: FunctionSpec[String] { type Args = SubstituteArgs } =
    FunctionSpec.simple[String, SubstituteArgs]("SUBSTITUTE", Arity.Range(3, 4)) { (args, ctx) =>
      val (textExpr, oldExpr, newExpr, instExpr) = args
      for
        text <- ctx.evalExpr(textExpr)
        oldS <- ctx.evalExpr(oldExpr)
        newS <- ctx.evalExpr(newExpr)
        instOpt <-
          instExpr.fold[Either[EvalError, Option[Int]]](Right(None))(e =>
            ctx.evalExpr(e).map(Some(_))
          )
        result <- substituteImpl(text, oldS, newS, instOpt)
      yield result
    }

  private def substituteImpl(
    text: String,
    oldS: String,
    newS: String,
    instOpt: Option[Int]
  ): Either[EvalError, String] =
    if oldS.isEmpty then Right(text)
    else
      instOpt match
        case Some(n) if n < 1 =>
          Left(EvalError.EvalFailed(s"SUBSTITUTE: instance must be >= 1, got $n"))
        case Some(n) => Right(replaceNthOccurrence(text, oldS, newS, n))
        case None => Right(text.replace(oldS, newS))

  /** Replace only the nth (1-indexed) forward, non-overlapping occurrence. */
  private def replaceNthOccurrence(s: String, oldS: String, newS: String, n: Int): String =
    @annotation.tailrec
    def findNth(idx: Int, count: Int): Int =
      val next = s.indexOf(oldS, idx)
      if next < 0 then -1
      else if count == n then next
      else findNth(next + oldS.length, count + 1)
    val pos = findNth(0, 1)
    if pos < 0 then s
    else s.substring(0, pos) + newS + s.substring(pos + oldS.length)

  val value: FunctionSpec[BigDecimal] { type Args = UnaryText } =
    FunctionSpec.simple[BigDecimal, UnaryText]("VALUE", Arity.one) { (_, _) =>
      Left(EvalError.EvalFailed("VALUE: not yet implemented"))
    }

  val text: FunctionSpec[String] { type Args = TextArgs } =
    FunctionSpec.simple[String, TextArgs]("TEXT", Arity.two) { (_, _) =>
      Left(EvalError.EvalFailed("TEXT: not yet implemented"))
    }
