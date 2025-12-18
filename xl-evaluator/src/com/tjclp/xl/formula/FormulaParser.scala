package com.tjclp.xl.formula

import com.tjclp.xl.{ARef, Anchor, CellRange, SheetName}
import com.tjclp.xl.addressing.RefParser
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
 *
 * @note
 *   Suppression rationale:
 *   - AsInstanceOf: Runtime parsing loses GADT type information. Type casts safely restore type
 *     parameters that are statically known to be correct based on parser context.
 */
@SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
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
    import scala.util.boundary, boundary.break

    boundary:
      // Strip leading '=' if present
      val formula = if input.startsWith("=") then input.substring(1) else input

      // Validate non-empty (early exit)
      if formula.trim.isEmpty then break(Left(ParseError.EmptyFormula))

      // Validate length (Excel limit: 8192 chars)
      if formula.length > 8192 then break(Left(ParseError.FormulaTooLong(formula.length, 8192)))
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
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
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
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
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
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
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
                    TExpr.asNumericExpr(left), // Convert PolyRef to typed Ref
                    TExpr.asNumericExpr(right)
                  ),
                  s4
                )
              }
            case _ => // <
              val s3 = skipWhitespace(s2.advance())
              parseComparison(s3).map { case (right, s4) =>
                (
                  TExpr.Lt(
                    TExpr.asNumericExpr(left), // Convert PolyRef to typed Ref
                    TExpr.asNumericExpr(right)
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
                    TExpr.asNumericExpr(left), // Convert PolyRef to typed Ref
                    TExpr.asNumericExpr(right)
                  ),
                  s4
                )
              }
            case _ => // >
              val s3 = skipWhitespace(s2.advance())
              parseComparison(s3).map { case (right, s4) =>
                (
                  TExpr.Gt(
                    TExpr.asNumericExpr(left), // Convert PolyRef to typed Ref
                    TExpr.asNumericExpr(right)
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
          // Concatenation operator not yet supported
          Left(
            ParseError.InvalidOperator(
              "&",
              s2.pos,
              "concatenation operator not yet supported (see LIMITATIONS.md)"
            )
          )
        case _ => Right((left, s2))
    }

  /**
   * Parse addition and subtraction (left-associative).
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
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
                    TExpr.asNumericExpr(acc), // Convert PolyRef to typed Ref
                    TExpr.asNumericExpr(right)
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
                    TExpr.asNumericExpr(acc), // Convert PolyRef to typed Ref
                    TExpr.asNumericExpr(right)
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
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
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
                    TExpr.asNumericExpr(acc), // Convert PolyRef to typed Ref
                    TExpr.asNumericExpr(right)
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
                    TExpr.asNumericExpr(acc), // Convert PolyRef to typed Ref
                    TExpr.asNumericExpr(right)
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
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  private def parseUnary(state: ParserState): ParseResult[TExpr[?]] =
    val s = skipWhitespace(state)
    s.currentChar match
      case Some('-') =>
        val s2 = skipWhitespace(s.advance())
        parseUnary(s2).map { case (expr, s3) =>
          // Unary minus: 0 - expr
          (
            TExpr.Sub(TExpr.Lit(BigDecimal(0)), TExpr.asNumericExpr(expr)), // Convert PolyRef
            s3
          )
        }
      case Some('N') | Some('n') if s.remaining.toUpperCase.startsWith("NOT") =>
        val s2 = skipWhitespace(s.advance(3))
        parseUnary(s2).map { case (expr, s3) =>
          (TExpr.Not(TExpr.asBooleanExpr(expr)), s3) // Convert PolyRef
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
      case Some('$') =>
        // Anchored cell reference (e.g., $A$1, $A1)
        parseAnchoredCellRef(s)
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

    // Read identifier (letters, digits, underscores, $ for anchored refs like A$1)
    @tailrec
    def readIdent(s: ParserState): ParserState =
      s.currentChar match
        case Some(c) if c.isLetterOrDigit || c == '_' || c == '$' => readIdent(s.advance())
        case Some('!') => s // Sheet reference separator
        case _ => s

    val s2 = readIdent(state)
    val ident = state.input.substring(startPos, s2.pos).toUpperCase

    // Note: This function is called within the boundary block from parseExpr,
    // so we can use early returns via simple control flow

    // Check for boolean literals
    ident match
      case "TRUE" => Right((TExpr.Lit(true), s2))
      case "FALSE" => Right((TExpr.Lit(false), s2))
      case _ =>
        // Check for sheet-qualified reference (identifier followed by '!')
        s2.currentChar match
          case Some('!') =>
            // Sheet-qualified reference: Sheet1!A1 or Sheet1!A1:B10
            parseSheetQualifiedRef(state.input.substring(startPos, s2.pos), s2.advance(), startPos)
          case _ =>
            // Not a boolean - check for function call (identifier followed by '(')
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
      // Lookup function in type class registry
      FunctionParser.lookup(name) match
        case Some(parser) =>
          // Parse using registered function parser
          parser.parse(args, startPos) match
            case Right(expr) => Right((expr, finalState))
            case Left(err) => Left(err)
        case None =>
          // Unknown function - provide suggestions
          val suggestions = suggestFunctions(name)
          Left(ParseError.UnknownFunction(name, startPos, suggestions))
    }

  /**
   * Parse cell reference with anchor support: A1, $A$1, $A1, A$1, Sheet1!A1
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  private def parseCellReference(
    refStr: String,
    state: ParserState,
    startPos: Int
  ): ParseResult[TExpr[?]] =
    // Parse anchor from refStr: "$A$1" → ("A1", Anchor.Absolute)
    val (cleanRef, anchor) = Anchor.parse(refStr)

    // Safe: ARef.parse returns Either[String, ARef], match is exhaustive
    ARef.parse(cleanRef) match
      case Right(aref) =>
        // Create PolyRef with anchor - type will be determined by function context
        // Cast needed due to opaque type erasure in pattern matching
        Right((TExpr.PolyRef(aref.asInstanceOf[ARef], anchor), state))
      case Left(err) =>
        Left(ParseError.InvalidCellRef(refStr, startPos, err))

  /**
   * Parse sheet-qualified reference: Sheet1!A1 or Sheet1!A1:B10
   *
   * @param sheetStr
   *   The sheet name (already parsed, before the !)
   * @param state
   *   Parser state positioned after the !
   * @param startPos
   *   Start position for error reporting
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  private def parseSheetQualifiedRef(
    sheetStr: String,
    state: ParserState,
    startPos: Int
  ): ParseResult[TExpr[?]] =
    // Validate sheet name - use unsafe since we've validated the string
    SheetName(sheetStr) match
      case Left(err) =>
        Left(ParseError.InvalidCellRef(s"$sheetStr!", startPos, s"invalid sheet name: $err"))
      case Right(_) =>
        // SheetName validated, create using unsafe to preserve opaque type
        val sheetName: SheetName = SheetName.unsafe(sheetStr)

        // Read the cell reference or range after the !
        val refStartPos = state.pos

        @tailrec
        def readRef(s: ParserState): ParserState =
          s.currentChar match
            case Some(c) if c.isLetterOrDigit || c == '$' || c == ':' => readRef(s.advance())
            case _ => s

        val s2 = readRef(state)
        val refPart = state.input.substring(refStartPos, s2.pos)

        if refPart.isEmpty then
          Left(ParseError.InvalidCellRef(s"$sheetStr!", startPos, "missing cell reference after !"))
        else if refPart.contains(':') then
          // Range reference: Sheet1!A1:B10
          CellRange.parse(refPart) match
            case Right(range) =>
              Right((TExpr.SheetRange(sheetName, range), s2))
            case Left(err) =>
              Left(ParseError.InvalidCellRef(s"$sheetStr!$refPart", startPos, err))
        else
          // Single cell reference: Sheet1!A1
          val (cleanRef, anchor) = Anchor.parse(refPart)
          ARef.parse(cleanRef) match
            case Right(aref) =>
              Right((TExpr.SheetPolyRef(sheetName, aref.asInstanceOf[ARef], anchor), s2))
            case Left(err) =>
              Left(ParseError.InvalidCellRef(s"$sheetStr!$refPart", startPos, err))

  /**
   * Parse anchored cell reference starting with $.
   *
   * This handles references like $A$1, $A1 that start with $.
   */
  private def parseAnchoredCellRef(state: ParserState): ParseResult[TExpr[?]] =
    val startPos = state.pos

    // Read the full reference including $ signs
    @tailrec
    def readRef(s: ParserState): ParserState =
      s.currentChar match
        case Some(c) if c.isLetterOrDigit || c == '$' => readRef(s.advance())
        case _ => s

    val s2 = readRef(state)
    val refStr = state.input.substring(startPos, s2.pos)

    // Check if followed by ':' for range
    s2.currentChar match
      case Some(':') =>
        parseRange(refStr, s2, startPos)
      case _ =>
        parseCellReference(refStr, s2, startPos)

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
        // Create FoldRange for SUM by default (function parsers will re-wrap as needed)
        // E.g., COUNT(A1:B10) will extract the range and create TExpr.count(range)
        Right((TExpr.sum(range), s3))
      case Left(err) =>
        Left(ParseError.InvalidCellRef(rangeStr, startPos, err))

  /**
   * Suggest similar function names for unknown functions.
   */
  private def suggestFunctions(name: String): List[String] =
    // Use FunctionParser registry for all available functions
    val knownFunctions = FunctionParser.allFunctions

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
