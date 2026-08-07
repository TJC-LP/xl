package com.tjclp.xl.formula.functions

import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.display.NumFmtFormatter
import com.tjclp.xl.formula.ast.{TExpr, ExprValue}
import com.tjclp.xl.formula.eval.{CriteriaMatcher, EvalError, Evaluator}
import com.tjclp.xl.formula.parser.ParseError
import com.tjclp.xl.formula.{Clock, Arity}
import com.tjclp.xl.styles.numfmt.NumFmt

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
    FunctionSpec.simple[BigDecimal, UnaryText](
      "LEN",
      Arity.one,
      flags = FunctionFlags(returnsNumeric = true)
    ) { (expr, ctx) =>
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
          if start < 1 then Left(EvalError.EvalFailed(s"MID: start must be >= 1, got $start"))
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
    FunctionSpec.simple[BigDecimal, FindArgs](
      "FIND",
      Arity.Range(2, 3),
      flags = FunctionFlags(returnsNumeric = true)
    ) { (args, ctx) =>
      val (findExpr, withinExpr, startOpt) = args
      for
        needle <- ctx.evalExpr(findExpr)
        haystack <- ctx.evalExpr(withinExpr)
        start <- startOpt.fold[Either[EvalError, Int]](Right(1))(e => ctx.evalExpr(e))
        result <-
          if start < 1 then Left(EvalError.EvalFailed(s"FIND: start must be >= 1, got $start"))
          else if start > haystack.length then
            // Excel: "If start_num is greater than the length of within_text,
            // FIND returns the #VALUE! error value." This applies to both empty
            // and non-empty needles — start past length is invalid regardless.
            Left(EvalError.EvalFailed(s"FIND: start ($start) is past end of text"))
          else if needle.isEmpty then Right(BigDecimal(start))
          else
            val idx = haystack.indexOf(needle, start - 1)
            if idx < 0 then Left(EvalError.EvalFailed(s"FIND: '$needle' not found in '$haystack'"))
            else Right(BigDecimal(idx + 1))
      yield result
    }

  /**
   * GH-476: SEARCH(find_text, within_text, [start_num]) — FIND's case-INSENSITIVE, wildcard-aware
   * sibling and the more common banker idiom. Absent from the roster, it aborted whole-book recalc
   * with an UnknownFunction host error.
   *
   * Excel semantics: 1-based result, `*`/`?` wildcards (`~` escapes), an out-of-range start_num or
   * a miss is a cached `#VALUE!` (catchable by IFERROR), never a host failure.
   */
  val search: FunctionSpec[BigDecimal] { type Args = FindArgs } =
    FunctionSpec.simple[BigDecimal, FindArgs](
      "SEARCH",
      Arity.Range(2, 3),
      flags = FunctionFlags(returnsNumeric = true)
    ) { (args, ctx) =>
      val (findExpr, withinExpr, startOpt) = args
      for
        needle <- ctx.evalExpr(findExpr)
        haystack <- ctx.evalExpr(withinExpr)
        start <- startOpt.fold[Either[EvalError, Int]](Right(1))(e => ctx.evalExpr(e))
        result <-
          if start < 1 || start > haystack.length then
            Left(
              EvalError.ErrorValue(
                CellError.Value,
                Some(s"SEARCH: start_num $start is outside 1..${haystack.length}")
              )
            )
          else
            searchIndex(needle, haystack, start).toRight(
              EvalError.ErrorValue(
                CellError.Value,
                Some(s"SEARCH: '$needle' not found in '$haystack'")
              )
            )
      yield BigDecimal(result)
    }

  /** 1-based index of the first case-insensitive (wildcard-aware) hit at or after `start`. */
  private def searchIndex(needle: String, haystack: String, start: Int): Option[Int] =
    if needle.isEmpty then Some(start)
    else if needle.exists(c => c == '*' || c == '?' || c == '~') then
      val matcher = java.util.regex.Pattern
        .compile(
          CriteriaMatcher.wildcardRegexBody(needle),
          java.util.regex.Pattern.CASE_INSENSITIVE
        )
        .matcher(haystack)
      Option.when(matcher.find(start - 1))(matcher.start + 1)
    else
      val idx = haystack.toLowerCase.indexOf(needle.toLowerCase, start - 1)
      Option.when(idx >= 0)(idx + 1)

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
    instOpt match
      case Some(n) if n < 1 =>
        Left(EvalError.EvalFailed(s"SUBSTITUTE: instance must be >= 1, got $n"))
      case _ if oldS.isEmpty => Right(text)
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
    FunctionSpec.simple[BigDecimal, UnaryText](
      "VALUE",
      Arity.one,
      flags = FunctionFlags(returnsNumeric = true)
    ) { (textExpr, ctx) =>
      ctx.evalExpr(textExpr).flatMap(parseExcelNumber)
    }

  /**
   * GH-476: unparseable VALUE input is an Excel error VALUE, not a host failure.
   *
   * `=VALUE("None")` is ordinary in banker books (a text sentinel in a numeric column); Excel
   * caches `#VALUE!` there and recalc carries on. Raising EvalFailed instead aborted cache refresh
   * for the whole cell — and IFERROR could not see a code to swallow.
   */
  private def unparseableValue(input: String): EvalError =
    EvalError.ErrorValue(CellError.Value, Some(s"VALUE: cannot parse '$input'"))

  /**
   * Parse an Excel-style numeric string. Handles currency ($), thousands commas, percent suffix
   * (×1/100), accounting parentheses (negative), scientific notation, sign, and whitespace.
   *
   * Date and time strings are rejected (deferred per TJC-1055 scope decision).
   */
  private def parseExcelNumber(input: String): Either[EvalError, BigDecimal] =
    val trimmed = input.trim
    if trimmed.isEmpty then Right(BigDecimal(0))
    else
      val (negFromParens, afterParens) =
        if trimmed.startsWith("(") && trimmed.endsWith(")") then
          (true, trimmed.substring(1, trimmed.length - 1))
        else (false, trimmed)
      // Reject "(-N)" / "(+N)": accounting parens combined with an inner sign
      // double-negates. Excel returns #VALUE! for this pattern.
      val innerStartsWithSign =
        val inner = afterParens.trim
        val afterCurrency = if inner.startsWith("$") then inner.drop(1).trim else inner
        val innerHasLeadingSign = inner.startsWith("-") || inner.startsWith("+")
        val currencyStrippedHasLeadingSign =
          afterCurrency.startsWith("-") || afterCurrency.startsWith("+")
        innerHasLeadingSign || currencyStrippedHasLeadingSign
      if negFromParens && innerStartsWithSign then Left(unparseableValue(input))
      else
        val (isPercent, afterPercent) =
          if afterParens.endsWith("%") then (true, afterParens.substring(0, afterParens.length - 1))
          else (false, afterParens)
        val cleaned = afterPercent.replace(",", "").replace("$", "").trim
        scala.util.Try(BigDecimal(cleaned)).toEither match
          case Right(n) =>
            val signed = if negFromParens then -n else n
            Right(if isPercent then signed / 100 else signed)
          case Left(_) =>
            Left(unparseableValue(input))

  val text: FunctionSpec[String] { type Args = TextArgs } =
    FunctionSpec.simple[String, TextArgs]("TEXT", Arity.two) { (args, ctx) =>
      val (valueExpr, formatExpr) = args
      for
        formatStr <- ctx.evalExpr(formatExpr)
        exprValue <- evalValue(ctx, valueExpr)
      yield
        if formatStr.isEmpty then ""
        else
          val cv = exprValueForTextFn(exprValue)
          NumFmtFormatter.formatValue(cv, NumFmt.Custom(formatStr))
    }

  /**
   * Coerce ExprValue → CellValue for TEXT. Empty cells are treated as Number(0) per Excel
   * convention; other types pass through the standard toCellValue path.
   */
  private def exprValueForTextFn(ev: ExprValue): CellValue =
    ev match
      case ExprValue.Cell(CellValue.Empty) => CellValue.Number(BigDecimal(0))
      case other => toCellValue(other)
