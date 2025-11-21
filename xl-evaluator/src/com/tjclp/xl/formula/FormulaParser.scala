package com.tjclp.xl.formula

import com.tjclp.xl.{ARef, CellRange}
import com.tjclp.xl.cells.Cell
import com.tjclp.xl.codec

import scala.annotation.tailrec

/**
 * Parser for Excel formula strings to typed TExpr AST.
 *
 * Implements a recursive descent parser with operator precedence. Pure functional - no mutation, no
 * exceptions, all errors as Either.
 *
 * Supported syntax:
 *   - Literals: numbers (42, 3.14), booleans (TRUE, FALSE), strings ("text")
 *   - Cell references: A1, $A$1, Sheet1!A1
 *   - Ranges: A1:B10
 *   - Operators: +, -, *, /, =, <>, <, <=, >, >=, &
 *   - Functions: SUM, COUNT, IF, AND, OR, NOT, etc.
 *   - Parentheses for grouping
 *
 * Operator precedence (highest to lowest):
 *   1. Parentheses ()
 *   2. Function calls
 *   3. Unary minus -
 *   4. Exponentiation ^ (future)
 *   5. Multiplication *, Division /
 *   6. Addition +, Subtraction -
 *   7. Concatenation &
 *   8. Comparison =, <>, <, <=, >, >=
 *   9. Logical AND
 *   10. Logical OR
 */
object FormulaParser:
  /**
   * Parse a formula string into a TExpr AST.
   *
   * @param input
   *   The formula string (with or without leading '=')
   * @return
   *   Right(expr) on success, Left(error) on parse failure
   *
   * Example:
   * {{{
   * parse("=SUM(A1:B10)") // Right(TExpr.FoldRange(...))
   * parse("=A1+B2")       // Right(TExpr.Add(...))
   * parse("=IF(A1>0, "Yes", "No")") // Right(TExpr.If(...))
   * }}}
   */
  def parse(input: String): Either[ParseError, TExpr[?]] =
    // Strip leading '=' if present
    val formula = if input.startsWith("=") then input.substring(1) else input

    // Validate non-empty
    if formula.trim.isEmpty then return Left(ParseError.EmptyFormula)

    // Validate length (Excel limit: 8192 chars)
    if formula.length > 8192 then return Left(ParseError.FormulaTooLong(formula.length, 8192))

    // Create parser state and parse
    val state = ParserState(formula, 0)
    parseExpr(state) match
      case Right((expr, finalState)) =>
        // Ensure we consumed all input
        skipWhitespace(finalState) match
          case s if s.pos >= s.input.length => Right(expr)
          case s =>
            Left(
              ParseError.UnexpectedChar(
                s.input(s.pos),
                s.pos,
                "unexpected characters after expression"
              )
            )
      case Left(err) => Left(err)

  /**
   * Parser state - tracks position in input string.
   *
   * @param input
   *   The formula string being parsed
   * @param pos
   *   Current position (0-based offset)
   */
  private case class ParserState(input: String, pos: Int):
    def advance(n: Int = 1): ParserState = copy(pos = pos + n)
    def currentChar: Option[Char] =
      if pos < input.length then Some(input(pos)) else None
    def remaining: String =
      if pos < input.length then input.substring(pos) else ""
    def atEnd: Boolean = pos >= input.length

  /**
   * Parse result - either error or (expression, remaining state).
   */
  private type ParseResult[A] = Either[ParseError, (A, ParserState)]

  /**
   * Skip whitespace characters.
   */
  private def skipWhitespace(state: ParserState): ParserState =
    @tailrec
    def loop(s: ParserState): ParserState =
      s.currentChar match
        case Some(c) if c.isWhitespace => loop(s.advance())
        case _ => s
    loop(state)

  /**
   * Parse top-level expression (handles all operator precedence).
   */
  private def parseExpr(state: ParserState): ParseResult[TExpr[?]] =
    parseLogicalOr(skipWhitespace(state))

  /**
   * Parse logical OR (lowest precedence).
   */
  private def parseLogicalOr(state: ParserState): ParseResult[TExpr[?]] =
    parseLogicalAnd(state).flatMap { case (left, s1) =>
      val s2 = skipWhitespace(s1)
      if s2.remaining.toUpperCase.startsWith("OR") then
        val s3 = skipWhitespace(s2.advance(2))
        parseLogicalOr(s3).map { case (right, s4) =>
          (TExpr.Or(left.asInstanceOf[TExpr[Boolean]], right.asInstanceOf[TExpr[Boolean]]), s4)
        }
      else Right((left, s2))
    }

  /**
   * Parse logical AND.
   */
  private def parseLogicalAnd(state: ParserState): ParseResult[TExpr[?]] =
    parseComparison(state).flatMap { case (left, s1) =>
      val s2 = skipWhitespace(s1)
      if s2.remaining.toUpperCase.startsWith("AND") then
        val s3 = skipWhitespace(s2.advance(3))
        parseLogicalAnd(s3).map { case (right, s4) =>
          (TExpr.And(left.asInstanceOf[TExpr[Boolean]], right.asInstanceOf[TExpr[Boolean]]), s4)
        }
      else Right((left, s2))
    }

  /**
   * Parse comparison operators: =, <>, <, <=, >, >=
   */
  private def parseComparison(state: ParserState): ParseResult[TExpr[?]] =
    parseConcatenation(state).flatMap { case (left, s1) =>
      val s2 = skipWhitespace(s1)
      s2.currentChar match
        case Some('=') =>
          val s3 = skipWhitespace(s2.advance())
          parseComparison(s3).map { case (right, s4) =>
            // Runtime parsing loses type info - use asInstanceOf
            (
              TExpr
                .Eq(left.asInstanceOf[TExpr[Any]], right.asInstanceOf[TExpr[Any]])
                .asInstanceOf[TExpr[Boolean]],
              s4
            )
          }
        case Some('<') =>
          s2.advance().currentChar match
            case Some('>') => // <>
              val s3 = skipWhitespace(s2.advance(2))
              parseComparison(s3).map { case (right, s4) =>
                // Runtime parsing loses type info - use asInstanceOf
                (
                  TExpr
                    .Neq(left.asInstanceOf[TExpr[Any]], right.asInstanceOf[TExpr[Any]])
                    .asInstanceOf[TExpr[Boolean]],
                  s4
                )
              }
            case Some('=') => // <=
              val s3 = skipWhitespace(s2.advance(2))
              parseComparison(s3).map { case (right, s4) =>
                (
                  TExpr.Lte(
                    left.asInstanceOf[TExpr[BigDecimal]],
                    right.asInstanceOf[TExpr[BigDecimal]]
                  ),
                  s4
                )
              }
            case _ => // <
              val s3 = skipWhitespace(s2.advance())
              parseComparison(s3).map { case (right, s4) =>
                (
                  TExpr.Lt(
                    left.asInstanceOf[TExpr[BigDecimal]],
                    right.asInstanceOf[TExpr[BigDecimal]]
                  ),
                  s4
                )
              }
        case Some('>') =>
          s2.advance().currentChar match
            case Some('=') => // >=
              val s3 = skipWhitespace(s2.advance(2))
              parseComparison(s3).map { case (right, s4) =>
                (
                  TExpr.Gte(
                    left.asInstanceOf[TExpr[BigDecimal]],
                    right.asInstanceOf[TExpr[BigDecimal]]
                  ),
                  s4
                )
              }
            case _ => // >
              val s3 = skipWhitespace(s2.advance())
              parseComparison(s3).map { case (right, s4) =>
                (
                  TExpr.Gt(
                    left.asInstanceOf[TExpr[BigDecimal]],
                    right.asInstanceOf[TExpr[BigDecimal]]
                  ),
                  s4
                )
              }
        case _ => Right((left, s2))
    }

  /**
   * Parse concatenation operator: &
   */
  private def parseConcatenation(state: ParserState): ParseResult[TExpr[?]] =
    parseAddSub(state).flatMap { case (left, s1) =>
      val s2 = skipWhitespace(s1)
      s2.currentChar match
        case Some('&') =>
          val s3 = skipWhitespace(s2.advance())
          parseConcatenation(s3).map { case (right, s4) =>
            // Concatenation not yet in TExpr - treat as string literal for now
            // Future: Add TExpr.Concat case
            (left, s4) // Placeholder
          }
        case _ => Right((left, s2))
    }

  /**
   * Parse addition and subtraction (left-associative).
   */
  private def parseAddSub(state: ParserState): ParseResult[TExpr[?]] =
    parseMulDiv(state).flatMap { case (left, s1) =>
      @tailrec
      def loop(acc: TExpr[?], s: ParserState): ParseResult[TExpr[?]] =
        val s2 = skipWhitespace(s)
        s2.currentChar match
          case Some('+') =>
            val s3 = skipWhitespace(s2.advance())
            parseMulDiv(s3) match
              case Right((right, s4)) =>
                loop(
                  TExpr.Add(
                    acc.asInstanceOf[TExpr[BigDecimal]],
                    right.asInstanceOf[TExpr[BigDecimal]]
                  ),
                  s4
                )
              case Left(err) => Left(err)
          case Some('-') if !s2.remaining.startsWith("->") =>
            val s3 = skipWhitespace(s2.advance())
            parseMulDiv(s3) match
              case Right((right, s4)) =>
                loop(
                  TExpr.Sub(
                    acc.asInstanceOf[TExpr[BigDecimal]],
                    right.asInstanceOf[TExpr[BigDecimal]]
                  ),
                  s4
                )
              case Left(err) => Left(err)
          case _ => Right((acc, s2))

      loop(left, s1)
    }

  /**
   * Parse multiplication and division (left-associative).
   */
  private def parseMulDiv(state: ParserState): ParseResult[TExpr[?]] =
    parseUnary(state).flatMap { case (left, s1) =>
      @tailrec
      def loop(acc: TExpr[?], s: ParserState): ParseResult[TExpr[?]] =
        val s2 = skipWhitespace(s)
        s2.currentChar match
          case Some('*') =>
            val s3 = skipWhitespace(s2.advance())
            parseUnary(s3) match
              case Right((right, s4)) =>
                loop(
                  TExpr.Mul(
                    acc.asInstanceOf[TExpr[BigDecimal]],
                    right.asInstanceOf[TExpr[BigDecimal]]
                  ),
                  s4
                )
              case Left(err) => Left(err)
          case Some('/') =>
            val s3 = skipWhitespace(s2.advance())
            parseUnary(s3) match
              case Right((right, s4)) =>
                loop(
                  TExpr.Div(
                    acc.asInstanceOf[TExpr[BigDecimal]],
                    right.asInstanceOf[TExpr[BigDecimal]]
                  ),
                  s4
                )
              case Left(err) => Left(err)
          case _ => Right((acc, s2))

      loop(left, s1)
    }

  /**
   * Parse unary operators: -, NOT
   */
  private def parseUnary(state: ParserState): ParseResult[TExpr[?]] =
    val s = skipWhitespace(state)
    s.currentChar match
      case Some('-') =>
        val s2 = skipWhitespace(s.advance())
        parseUnary(s2).map { case (expr, s3) =>
          // Unary minus: 0 - expr
          (
            TExpr.Sub(TExpr.Lit(BigDecimal(0)), expr.asInstanceOf[TExpr[BigDecimal]]),
            s3
          )
        }
      case Some('N') | Some('n') if s.remaining.toUpperCase.startsWith("NOT") =>
        val s2 = skipWhitespace(s.advance(3))
        parseUnary(s2).map { case (expr, s3) =>
          (TExpr.Not(expr.asInstanceOf[TExpr[Boolean]]), s3)
        }
      case _ => parsePrimary(s)

  /**
   * Parse primary expressions: literals, cell refs, functions, parentheses.
   */
  private def parsePrimary(state: ParserState): ParseResult[TExpr[?]] =
    val s = skipWhitespace(state)
    s.currentChar match
      case None => Left(ParseError.UnexpectedEOF(s.pos, "expected expression"))
      case Some('(') =>
        // Parenthesized expression
        val s2 = skipWhitespace(s.advance())
        parseExpr(s2).flatMap { case (expr, s3) =>
          val s4 = skipWhitespace(s3)
          s4.currentChar match
            case Some(')') => Right((expr, s4.advance()))
            case Some(c) =>
              Left(ParseError.UnbalancedDelimiter(s4.pos, c, "expected ')'"))
            case None => Left(ParseError.UnexpectedEOF(s4.pos, "expected ')'"))
        }
      case Some('"') =>
        // String literal
        parseStringLiteral(s)
      case Some(c) if c.isDigit || c == '.' =>
        // Number literal
        parseNumberLiteral(s)
      case Some(c) if c.isLetter =>
        // Function call, cell reference, or boolean literal
        parseFunctionOrRef(s)
      case Some(c) =>
        Left(ParseError.UnexpectedChar(c, s.pos, "expected expression"))

  /**
   * Parse string literal: "text"
   */
  private def parseStringLiteral(state: ParserState): ParseResult[TExpr[?]] =
    val startPos = state.pos
    @tailrec
    def loop(s: ParserState, acc: StringBuilder): ParseResult[TExpr[String]] =
      s.currentChar match
        case None => Left(ParseError.UnexpectedEOF(s.pos, "unterminated string"))
        case Some('"') =>
          s.advance().currentChar match
            case Some('"') => // Escaped quote
              loop(s.advance(2), acc.append('"'))
            case _ => // End of string
              Right((TExpr.Lit(acc.toString), s.advance()))
        case Some(c) =>
          loop(s.advance(), acc.append(c))

    state.currentChar match
      case Some('"') => loop(state.advance(), new StringBuilder)
      case _ =>
        Left(ParseError.UnexpectedChar(state.input(state.pos), state.pos, "expected '\"'"))

  /**
   * Parse number literal: 42, 3.14, .5, 1.5E10, 2.3E-7
   *
   * Supports scientific notation (E or e) with optional +/- sign.
   */
  private def parseNumberLiteral(state: ParserState): ParseResult[TExpr[BigDecimal]] =
    val startPos = state.pos
    @tailrec
    def loop(s: ParserState, hasDecimal: Boolean, hasExponent: Boolean): ParserState =
      s.currentChar match
        case Some(c) if c.isDigit =>
          loop(s.advance(), hasDecimal, hasExponent)
        case Some('.') if !hasDecimal && !hasExponent =>
          loop(s.advance(), hasDecimal = true, hasExponent)
        case Some('E' | 'e') if !hasExponent =>
          // Start of exponent - may be followed by +/- sign
          val s2 = s.advance()
          s2.currentChar match
            case Some('+' | '-') =>
              // Optional sign after E
              val s3 = s2.advance()
              s3.currentChar match
                case Some(c) if c.isDigit =>
                  loop(s3, hasDecimal, hasExponent = true)
                case _ => s // Invalid: E+ or E- not followed by digit
            case Some(c) if c.isDigit =>
              loop(s2, hasDecimal, hasExponent = true)
            case _ => s // Invalid: E not followed by sign or digit
        case _ => s

    val s2 = loop(state, hasDecimal = state.currentChar.contains('.'), hasExponent = false)
    val numStr = state.input.substring(startPos, s2.pos)

    try
      val value = BigDecimal(numStr)
      Right((TExpr.Lit(value), s2))
    catch
      case _: NumberFormatException =>
        Left(ParseError.InvalidNumber(numStr, startPos, "invalid number format"))

  /**
   * Parse function call, cell reference, range, or boolean literal.
   */
  private def parseFunctionOrRef(state: ParserState): ParseResult[TExpr[?]] =
    val startPos = state.pos

    // Read identifier (letters, digits, underscores)
    @tailrec
    def readIdent(s: ParserState): ParserState =
      s.currentChar match
        case Some(c) if c.isLetterOrDigit || c == '_' => readIdent(s.advance())
        case Some('!') => s // Sheet reference separator
        case _ => s

    val s2 = readIdent(state)
    val ident = state.input.substring(startPos, s2.pos).toUpperCase

    // Check for boolean literals
    ident match
      case "TRUE" => return Right((TExpr.Lit(true), s2))
      case "FALSE" => return Right((TExpr.Lit(false), s2))
      case _ => // Continue parsing

    // Check for function call (identifier followed by '(')
    val s3 = skipWhitespace(s2)
    s3.currentChar match
      case Some('(') =>
        // Function call
        parseFunction(ident, s3, startPos)
      case Some(':') =>
        // Range (e.g., A1:B10)
        parseRange(state.input.substring(startPos, s2.pos), s2, startPos)
      case _ =>
        // Cell reference
        parseCellReference(state.input.substring(startPos, s2.pos), s2, startPos)

  /**
   * Parse function call: FUNC(arg1, arg2, ...)
   */
  private def parseFunction(
    name: String,
    state: ParserState,
    startPos: Int
  ): ParseResult[TExpr[?]] =
    // Skip opening '('
    val s2 = skipWhitespace(state.advance())

    // Parse arguments (comma-separated)
    def parseArgs(s: ParserState, args: List[TExpr[?]]): ParseResult[List[TExpr[?]]] =
      val s1 = skipWhitespace(s)
      s1.currentChar match
        case Some(')') => Right((args.reverse, s1.advance()))
        case _ =>
          parseExpr(s1).flatMap { case (arg, s2) =>
            val s3 = skipWhitespace(s2)
            s3.currentChar match
              case Some(',') => parseArgs(skipWhitespace(s3.advance()), arg :: args)
              case Some(')') => Right(((arg :: args).reverse, s3.advance()))
              case Some(c) =>
                Left(ParseError.UnexpectedChar(c, s3.pos, "expected ',' or ')'"))
              case None => Left(ParseError.UnexpectedEOF(s3.pos, "expected ',' or ')'"))
          }

    parseArgs(s2, Nil).flatMap { case (args, finalState) =>
      // Dispatch to specific function parser
      name match
        case "SUM" => parseSumFunction(args, startPos)
        case "COUNT" => parseCountFunction(args, startPos)
        case "AVERAGE" => parseAverageFunction(args, startPos)
        case "IF" => parseIfFunction(args, startPos)
        case "AND" => parseAndFunction(args, startPos)
        case "OR" => parseOrFunction(args, startPos)
        case "NOT" => parseNotFunction(args, startPos)
        case _ =>
          // Unknown function - provide suggestions
          val suggestions = suggestFunctions(name)
          Left(ParseError.UnknownFunction(name, startPos, suggestions))
      match
        case Right(expr) => Right((expr, finalState))
        case Left(err) => Left(err)
    }

  /**
   * Parse SUM function: SUM(range) or SUM(expr1, expr2, ...)
   */
  private def parseSumFunction(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
    args match
      case Nil =>
        Left(ParseError.InvalidArguments("SUM", pos, "at least 1 argument", "0 arguments"))
      case head :: _ =>
        // For now, only support single range argument
        // Future: support multiple arguments
        head match
          case fold: TExpr.FoldRange[?, ?] => Right(fold)
          case _ =>
            Left(ParseError.InvalidArguments("SUM", pos, "range argument", "non-range"))

  /**
   * Parse COUNT function: COUNT(range)
   */
  private def parseCountFunction(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
    args match
      case List(fold: TExpr.FoldRange[?, ?]) => Right(fold)
      case _ =>
        Left(
          ParseError.InvalidArguments("COUNT", pos, "1 range argument", s"${args.length} arguments")
        )

  /**
   * Parse AVERAGE function: AVERAGE(range)
   */
  private def parseAverageFunction(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
    args match
      case List(fold: TExpr.FoldRange[?, ?]) => Right(fold)
      case _ =>
        Left(
          ParseError.InvalidArguments(
            "AVERAGE",
            pos,
            "1 range argument",
            s"${args.length} arguments"
          )
        )

  /**
   * Parse IF function: IF(condition, ifTrue, ifFalse)
   */
  private def parseIfFunction(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
    args match
      case List(cond, ifTrue, ifFalse) =>
        // Runtime parsing: ifTrue and ifFalse must have same type but we can't verify at parse time
        Right(
          TExpr.If(
            cond.asInstanceOf[TExpr[Boolean]],
            ifTrue.asInstanceOf[TExpr[Any]],
            ifFalse.asInstanceOf[TExpr[Any]]
          )
        )
      case _ =>
        Left(ParseError.InvalidArguments("IF", pos, "3 arguments", s"${args.length} arguments"))

  /**
   * Parse AND function: AND(expr1, expr2, ...)
   */
  private def parseAndFunction(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
    args match
      case Nil =>
        Left(ParseError.InvalidArguments("AND", pos, "at least 1 argument", "0 arguments"))
      case head :: tail =>
        val result = tail.foldLeft(head.asInstanceOf[TExpr[Boolean]]) { (acc, expr) =>
          TExpr.And(acc, expr.asInstanceOf[TExpr[Boolean]])
        }
        Right(result)

  /**
   * Parse OR function: OR(expr1, expr2, ...)
   */
  private def parseOrFunction(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
    args match
      case Nil =>
        Left(ParseError.InvalidArguments("OR", pos, "at least 1 argument", "0 arguments"))
      case head :: tail =>
        val result = tail.foldLeft(head.asInstanceOf[TExpr[Boolean]]) { (acc, expr) =>
          TExpr.Or(acc, expr.asInstanceOf[TExpr[Boolean]])
        }
        Right(result)

  /**
   * Parse NOT function: NOT(expr)
   */
  private def parseNotFunction(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
    args match
      case List(expr) => Right(TExpr.Not(expr.asInstanceOf[TExpr[Boolean]]))
      case _ =>
        Left(ParseError.InvalidArguments("NOT", pos, "1 argument", s"${args.length} arguments"))

  /**
   * Parse cell reference: A1, $A$1, Sheet1!A1
   */
  private def parseCellReference(
    refStr: String,
    state: ParserState,
    startPos: Int
  ): ParseResult[TExpr[?]] =
    ARef.parse(refStr) match
      case Right(aref: ARef) =>
        // Create Ref expression with numeric decoder
        Right((TExpr.Ref(aref, TExpr.decodeNumeric), state))
      case Left(err) =>
        Left(ParseError.InvalidCellRef(refStr, startPos, err))

  /**
   * Parse range: A1:B10
   */
  private def parseRange(
    startRef: String,
    state: ParserState,
    startPos: Int
  ): ParseResult[TExpr[?]] =
    // Skip ':'
    val s2 = state.advance()
    val endPos = s2.pos

    // Read end reference
    @tailrec
    def readRef(s: ParserState): ParserState =
      s.currentChar match
        case Some(c) if c.isLetterOrDigit || c == '$' || c == '!' => readRef(s.advance())
        case _ => s

    val s3 = readRef(s2)
    val endRef = state.input.substring(endPos, s3.pos)
    val rangeStr = s"$startRef:$endRef"

    CellRange.parse(rangeStr) match
      case Right(range) =>
        // Create FoldRange for SUM by default
        Right((TExpr.sum(range), s3))
      case Left(err) =>
        Left(ParseError.InvalidCellRef(rangeStr, startPos, err))

  /**
   * Suggest similar function names for unknown functions.
   */
  private def suggestFunctions(name: String): List[String] =
    val knownFunctions = List(
      "SUM",
      "COUNT",
      "AVERAGE",
      "IF",
      "AND",
      "OR",
      "NOT",
      "MIN",
      "MAX",
      "ROUND",
      "ABS"
    )

    // Simple Levenshtein distance for suggestions
    knownFunctions
      .map(f => (f, levenshteinDistance(name.toUpperCase, f)))
      .filter(_._2 <= 3) // Max distance 3
      .sortBy(_._2)
      .take(3)
      .map(_._1)

  /**
   * Levenshtein distance for function name suggestions.
   */
  private def levenshteinDistance(s1: String, s2: String): Int =
    val len1 = s1.length
    val len2 = s2.length
    val matrix = Array.ofDim[Int](len1 + 1, len2 + 1)

    for i <- 0 to len1 do matrix(i)(0) = i
    for j <- 0 to len2 do matrix(0)(j) = j

    for
      i <- 1 to len1
      j <- 1 to len2
    do
      val cost = if s1(i - 1) == s2(j - 1) then 0 else 1
      matrix(i)(j) = math.min(
        math.min(matrix(i - 1)(j) + 1, matrix(i)(j - 1) + 1),
        matrix(i - 1)(j - 1) + cost
      )

    matrix(len1)(len2)
