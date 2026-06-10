package com.tjclp.xl.formula.parser

import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.functions.{FunctionSpec, FunctionSpecs, FunctionRegistry}
import com.tjclp.xl.formula.{Arity}

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
 * Operator precedence (highest to lowest, Excel-compatible):
 *   1. Parentheses ()
 *   2. Function calls
 *   3. Exponentiation ^ (right-associative)
 *   4. Unary minus -
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
   * parse("=SUM(A1:B10)") // Right(TExpr.Call(...))
   * parse("=A1+B2")       // Right(TExpr.Add(...))
   * parse("=IF(A1>0, "Yes", "No")") // Right(TExpr.Call(...))
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
            case s if s.pos >= s.input.length =>
              // Resolve top-level PolyRef to typed Ref (default: numeric)
              // This eliminates unsafe asInstanceOf casts in the evaluator
              Right(resolveTopLevelPolyRef(expr))
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
   * Resolve top-level polymorphic references to typed references.
   *
   * When a formula is just a cell reference (e.g., "=A1" or "=Sheet1!A1"), the parser returns a
   * PolyRef/SheetPolyRef with no static type information. This function converts these to typed
   * Ref/SheetRef with a resolved value decoder that:
   *   - Extracts cached values from Formula cells
   *   - Converts Empty cells to Number(0)
   *   - Returns other cell types as-is
   *
   * This resolution eliminates unsafe asInstanceOf casts in the evaluator by ensuring all cell
   * references are properly typed before evaluation. The resolved value decoder matches Excel's
   * behavior where standalone cell references return the cell's effective value.
   */
  private def resolveTopLevelPolyRef(expr: TExpr[?]): TExpr[?] = expr match
    case _: TExpr.PolyRef | _: TExpr.SheetPolyRef => TExpr.asResolvedValueExpr(expr)
    case other => other

  /**
   * GH-193: an in-scope LET binding visible to the parser.
   *
   * @param name
   *   The declared binding name (original case; lookup is case-insensitive)
   * @param substitution
   *   For range-shaped binding values (RangeRef/SheetRange), the range expression to substitute at
   *   each use site so range-typed argument positions keep working; None for ordinary bindings
   *   which parse to [[TExpr.BindingRef]]
   */
  private final case class LetScopeEntry(name: String, substitution: Option[TExpr[?]])

  /**
   * Parser state - tracks position in input string.
   *
   * @param input
   *   The formula string being parsed
   * @param pos
   *   Current position (0-based offset)
   * @param scope
   *   GH-193: lexically visible LET bindings, innermost first
   */
  private case class ParserState(
    input: String,
    pos: Int,
    depth: Int = 0,
    scope: List[LetScopeEntry] = Nil
  ):
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
   * Maximum formula nesting depth — a stack-overflow guard (GH-56). Well above Excel's 64-level
   * nesting limit and any realistic formula, but far below the depth (~2000) that overflows the
   * evaluator/parser stack. Capping the parser keeps the AST shallow enough that evaluation is also
   * safe, so a single guard covers both the parser and the evaluator.
   */
  private val MaxNestingDepth = 256

  /**
   * Guarded descent into a nested sub-expression: deepen the state, or fail if already too deep.
   */
  private def descend(s: ParserState): Either[ParseError, ParserState] =
    if s.depth >= MaxNestingDepth then Left(ParseError.NestingTooDeep(s.depth, MaxNestingDepth))
    else Right(s.copy(depth = s.depth + 1))

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
    // parseExpr is the re-entry point for every nesting boundary (parens + function args), so a
    // single depth guard here bounds both. Restore the caller's depth on the way out so siblings
    // (e.g. `(1)+(2)+...`) don't accumulate.
    val s0 = skipWhitespace(state)
    descend(s0).flatMap { sd =>
      parseLogicalOr(sd).map { case (expr, s1) => (expr, s1.copy(depth = s0.depth)) }
    }

  /**
   * Match a logical keyword (NOT/AND/OR) at the current position with a word boundary: the keyword
   * must be followed by whitespace, '(', or end of input. Without the boundary, sheet and defined
   * names starting with a keyword were consumed as the operator — `Notes1!A1` parsed as
   * `NOT(es1!A1)` (found by the cross-sheet round-trip property, GH-268).
   */
  private def isKeywordAt(s: ParserState, keyword: String): Boolean =
    val rem = s.remaining
    rem.length >= keyword.length &&
    rem.substring(0, keyword.length).equalsIgnoreCase(keyword) && {
      rem.length == keyword.length || {
        val next = rem.charAt(keyword.length)
        next.isWhitespace || next == '('
      }
    }

  /**
   * Parse logical OR (lowest precedence).
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  private def parseLogicalOr(state: ParserState): ParseResult[TExpr[?]] =
    parseLogicalAnd(state).flatMap { case (left, s1) =>
      val s2 = skipWhitespace(s1)
      if isKeywordAt(s2, "OR") then
        val s3 = skipWhitespace(s2.advance(2))
        descend(s3).flatMap { sd =>
          parseLogicalOr(sd).map { case (right, s4) =>
            (
              TExpr.Call(
                FunctionSpecs.or,
                List(left.asInstanceOf[TExpr[Boolean]], right.asInstanceOf[TExpr[Boolean]])
              ),
              s4.copy(depth = s3.depth)
            )
          }
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
      if isKeywordAt(s2, "AND") then
        val s3 = skipWhitespace(s2.advance(3))
        descend(s3).flatMap { sd =>
          parseLogicalAnd(sd).map { case (right, s4) =>
            (
              TExpr.Call(
                FunctionSpecs.and,
                List(left.asInstanceOf[TExpr[Boolean]], right.asInstanceOf[TExpr[Boolean]])
              ),
              s4.copy(depth = s3.depth)
            )
          }
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
          descend(s3).flatMap { sd =>
            parseComparison(sd).map { case (right, s4) =>
              // GH-233: resolve PolyRef operands polymorphically (like the inequality
              // branches use asNumericExpr) so =A1=B1 / =IF(A1=B1,…) evaluate instead of
              // erroring with "Unresolved PolyRef". asResolvedValueExpr preserves
              // text/number/bool/date equality (unlike numeric-only asNumericExpr).
              (
                TExpr.Eq(TExpr.asResolvedValueExpr(left), TExpr.asResolvedValueExpr(right)),
                s4.copy(depth = s3.depth)
              )
            }
          }
        case Some('<') =>
          s2.advance().currentChar match
            case Some('>') => // <>
              val s3 = skipWhitespace(s2.advance(2))
              descend(s3).flatMap { sd =>
                parseComparison(sd).map { case (right, s4) =>
                  // GH-233: resolve PolyRef operands so =A1<>B1 evaluates (see Eq above).
                  (
                    TExpr.Neq(TExpr.asResolvedValueExpr(left), TExpr.asResolvedValueExpr(right)),
                    s4.copy(depth = s3.depth)
                  )
                }
              }
            case Some('=') => // <=
              val s3 = skipWhitespace(s2.advance(2))
              descend(s3).flatMap { sd =>
                parseComparison(sd).map { case (right, s4) =>
                  (
                    TExpr.Lte(
                      TExpr.asNumericExpr(left), // Convert PolyRef to typed Ref
                      TExpr.asNumericExpr(right)
                    ),
                    s4.copy(depth = s3.depth)
                  )
                }
              }
            case _ => // <
              val s3 = skipWhitespace(s2.advance())
              descend(s3).flatMap { sd =>
                parseComparison(sd).map { case (right, s4) =>
                  (
                    TExpr.Lt(
                      TExpr.asNumericExpr(left), // Convert PolyRef to typed Ref
                      TExpr.asNumericExpr(right)
                    ),
                    s4.copy(depth = s3.depth)
                  )
                }
              }
        case Some('>') =>
          s2.advance().currentChar match
            case Some('=') => // >=
              val s3 = skipWhitespace(s2.advance(2))
              descend(s3).flatMap { sd =>
                parseComparison(sd).map { case (right, s4) =>
                  (
                    TExpr.Gte(
                      TExpr.asNumericExpr(left), // Convert PolyRef to typed Ref
                      TExpr.asNumericExpr(right)
                    ),
                    s4.copy(depth = s3.depth)
                  )
                }
              }
            case _ => // >
              val s3 = skipWhitespace(s2.advance())
              descend(s3).flatMap { sd =>
                parseComparison(sd).map { case (right, s4) =>
                  (
                    TExpr.Gt(
                      TExpr.asNumericExpr(left), // Convert PolyRef to typed Ref
                      TExpr.asNumericExpr(right)
                    ),
                    s4.copy(depth = s3.depth)
                  )
                }
              }
        case _ => Right((left, s2))
    }

  /**
   * Parse concatenation operator: & (left-associative)
   *
   * Excel's & operator joins strings: "Hello" & "World" → "HelloWorld" Operands are coerced to
   * strings via asStringExpr.
   */
  private def parseConcatenation(state: ParserState): ParseResult[TExpr[?]] =
    parseAddSub(state).flatMap { case (left, s1) =>
      val s2 = skipWhitespace(s1)
      s2.currentChar match
        case Some('&') =>
          val s3 = skipWhitespace(s2.advance())
          descend(s3).flatMap { sd =>
            parseConcatenation(sd).map { case (right, s4) =>
              (
                TExpr.Concat(
                  TExpr.asStringExpr(left),
                  TExpr.asStringExpr(right)
                ),
                s4.copy(depth = s3.depth)
              )
            }
          }
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
            // GH-56: each chained term deepens the left-nested AST, so count it against the depth
            // budget — otherwise a flat `1+1+1+…` overflows the evaluator (the parser loops, but
            // eval recurses the spine).
            descend(s2) match
              case Left(err) => Left(err)
              case Right(sd) =>
                val s3 = skipWhitespace(sd.advance())
                parseMulDiv(s3) match
                  case Right((right, s4)) =>
                    loop(
                      TExpr.Add(
                        TExpr.asNumericOrRangeExpr(acc), // Preserve RangeRef for array arithmetic
                        TExpr.asNumericOrRangeExpr(right)
                      ),
                      s4
                    )
                  case Left(err) => Left(err)
          case Some('-') if !s2.remaining.startsWith("->") =>
            descend(s2) match
              case Left(err) => Left(err)
              case Right(sd) =>
                val s3 = skipWhitespace(sd.advance())
                parseMulDiv(s3) match
                  case Right((right, s4)) =>
                    loop(
                      TExpr.Sub(
                        TExpr.asNumericOrRangeExpr(acc), // Preserve RangeRef for array arithmetic
                        TExpr.asNumericOrRangeExpr(right)
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
            // GH-56: count each chained factor against the depth budget (see parseAddSub).
            descend(s2) match
              case Left(err) => Left(err)
              case Right(sd) =>
                val s3 = skipWhitespace(sd.advance())
                parseUnary(s3) match
                  case Right((right, s4)) =>
                    loop(
                      TExpr.Mul(
                        TExpr.asNumericOrRangeExpr(acc), // Preserve RangeRef for array arithmetic
                        TExpr.asNumericOrRangeExpr(right)
                      ),
                      s4
                    )
                  case Left(err) => Left(err)
          case Some('/') =>
            descend(s2) match
              case Left(err) => Left(err)
              case Right(sd) =>
                val s3 = skipWhitespace(sd.advance())
                parseUnary(s3) match
                  case Right((right, s4)) =>
                    loop(
                      TExpr.Div(
                        TExpr.asNumericOrRangeExpr(acc), // Preserve RangeRef for array arithmetic
                        TExpr.asNumericOrRangeExpr(right)
                      ),
                      s4
                    )
                  case Left(err) => Left(err)
          case _ => Right((acc, s2))

      loop(left, s1)
    }

  /**
   * Parse exponentiation (right-associative, highest arithmetic precedence).
   *
   * Right-associativity: 2^3^2 = 2^(3^2) = 512, not (2^3)^2 = 64 Excel precedence: ^ binds tighter
   * than unary minus, so -2^2 = -(2^2) = -4 But the exponent can have unary minus: 2^-1 = 0.5
   */
  private def parsePow(state: ParserState): ParseResult[TExpr[?]] =
    parsePrimary(state).flatMap { case (left, s1) =>
      val s2 = skipWhitespace(s1)
      s2.currentChar match
        case Some('^') =>
          descend(s2).flatMap { sd =>
            val s3 = skipWhitespace(sd.advance())
            // Allow unary minus in the exponent (2^-1 = 0.5)
            parsePowExponent(s3).map {
              case (right, s4) => // Recursive call for right-associativity
                (
                  TExpr.Pow(
                    TExpr.asNumericExpr(left),
                    TExpr.asNumericExpr(right)
                  ),
                  s4.copy(depth = s2.depth)
                )
            }
          }
        case _ => Right((left, s2))
    }

  /**
   * Parse the exponent of a power expression, allowing unary minus and plus. This handles cases
   * like 2^-1 = 0.5 while keeping -2^2 = -(2^2) = -4. Unary plus is identity (GH-271).
   */
  private def parsePowExponent(state: ParserState): ParseResult[TExpr[?]] =
    val s = skipWhitespace(state)
    s.currentChar match
      case Some('-') =>
        descend(s).flatMap { sd =>
          val s2 = skipWhitespace(sd.advance())
          parsePowExponent(s2).map { case (expr, s3) =>
            (
              TExpr.Sub(TExpr.Lit(BigDecimal(0)), TExpr.asNumericExpr(expr)),
              s3.copy(depth = s.depth)
            )
          }
        }
      case Some('+') =>
        descend(s).flatMap { sd =>
          val s2 = skipWhitespace(sd.advance())
          parsePowExponent(s2).map { case (expr, s3) =>
            (expr, s3.copy(depth = s.depth))
          }
        }
      case _ => parsePow(s)

  /**
   * Parse unary operators: -, +, NOT
   *
   * Excel precedence: unary minus has lower precedence than ^, so -2^2 = -(2^2) = -4
   *
   * Unary plus (GH-271: =+A1, the pervasive banker idiom) is identity: it parses to the same AST as
   * the operand alone, so the printer normalizes it away while parse∘print=id holds on the AST.
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  private def parseUnary(state: ParserState): ParseResult[TExpr[?]] =
    val s = skipWhitespace(state)
    s.currentChar match
      case Some('-') =>
        descend(s).flatMap { sd =>
          val s2 = skipWhitespace(sd.advance())
          parseUnary(s2).map { case (expr, s3) =>
            // Unary minus: 0 - expr
            (
              TExpr.Sub(TExpr.Lit(BigDecimal(0)), TExpr.asNumericExpr(expr)), // Convert PolyRef
              s3.copy(depth = s.depth)
            )
          }
        }
      case Some('+') =>
        // descend: each chained '+' counts against the depth budget (GH-56 recursion guard)
        descend(s).flatMap { sd =>
          val s2 = skipWhitespace(sd.advance())
          parseUnary(s2).map { case (expr, s3) =>
            (expr, s3.copy(depth = s.depth))
          }
        }
      case Some('N') | Some('n') if isKeywordAt(s, "NOT") =>
        descend(s).flatMap { sd =>
          val s2 = skipWhitespace(sd.advance(3))
          parseUnary(s2).map { case (expr, s3) =>
            (TExpr.Call(FunctionSpecs.not, TExpr.asBooleanExpr(expr)), s3.copy(depth = s.depth))
          }
        }
      case _ => parsePow(s)

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
        // Check for full row reference (e.g., 1:1, 1:5) before treating as number
        // Row references are digits followed by ':' and more digits
        if c.isDigit then parseNumberOrRowRange(s)
        else
          // Starts with '.' - definitely a number like .5
          parseNumberLiteral(s)
      case Some(c) if c.isLetter || c == '_' =>
        // Function call, cell reference, boolean literal, or LET binding reference
        // ('_' admits underscore-led LET names, GH-193)
        parseFunctionOrRef(s)
      case Some('$') =>
        // Anchored cell reference (e.g., $A$1, $A1)
        parseAnchoredCellRef(s)
      case Some('\'') =>
        // Quoted sheet name reference (e.g., 'Q1 Report'!A1, 'Debt-Schedule'!H29)
        parseQuotedSheetRef(s)
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
   * Parse either a number literal or a full row range reference.
   *
   * Distinguishes between:
   *   - Numbers: 42, 3.14, 1.5E10
   *   - Row ranges: 1:1, 1:5 (digits followed by ':' and more digits)
   */
  private def parseNumberOrRowRange(state: ParserState): ParseResult[TExpr[?]] =
    val startPos = state.pos

    // First, read all leading digits
    @tailrec
    def readDigits(s: ParserState): ParserState =
      s.currentChar match
        case Some(c) if c.isDigit => readDigits(s.advance())
        case _ => s

    val afterDigits = readDigits(state)

    // Check if followed by ':' and more digits (row range pattern)
    afterDigits.currentChar match
      case Some(':') =>
        val afterColon = afterDigits.advance()
        afterColon.currentChar match
          case Some(c) if c.isDigit =>
            // This is a row range like 1:5
            val afterSecondDigits = readDigits(afterColon)
            val rangeStr = state.input.substring(startPos, afterSecondDigits.pos)
            CellRange.parse(rangeStr) match
              case Right(range) =>
                // Create RangeRef for range arguments
                Right((TExpr.RangeRef(range), afterSecondDigits))
              case Left(err) =>
                Left(ParseError.InvalidCellRef(rangeStr, startPos, err))
          case _ =>
            // Just digits followed by ':' but not more digits - treat as number
            parseNumberLiteral(state)
      case _ =>
        // No ':' - definitely a number
        parseNumberLiteral(state)

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
                // GH-193: a bare identifier matching an in-scope LET binding resolves to the
                // binding (case-insensitive, innermost first) — even when the name shadows a
                // function name. Function-call syntax above still wins for `name(...)`.
                val rawIdent = state.input.substring(startPos, s2.pos)
                state.scope.find(_.name.equalsIgnoreCase(rawIdent)) match
                  case Some(entry) =>
                    entry.substitution match
                      case Some(rangeExpr) => Right((rangeExpr, s2))
                      case None => Right((TExpr.BindingRef(entry.name), s2))
                  case None =>
                    // Cell reference
                    parseCellReference(rawIdent, s2, startPos)

  /**
   * Parse function call: FUNC(arg1, arg2, ...)
   */
  private def parseFunction(
    name: String,
    state: ParserState,
    startPos: Int
  ): ParseResult[TExpr[?]] =
    // GH-193: LET is a special form (it introduces lexical bindings), not a FunctionSpec.
    if name == "LET" then parseLet(state, startPos)
    else parseRegularFunction(name, state, startPos)

  private def parseRegularFunction(
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
      FunctionRegistry.lookup(name) match
        case Some(spec) =>
          spec.arity
            .validate(args.length, spec.name, startPos)
            .flatMap(_ => spec.argSpec.parse(args, startPos, spec.name))
            .flatMap {
              case (parsedArgs, Nil) =>
                Right((TExpr.Call(spec, parsedArgs), finalState))
              case _ =>
                Left(
                  ParseError.InvalidArguments(
                    spec.name,
                    startPos,
                    spec.argSpec.describe,
                    s"${args.length} arguments"
                  )
                )
            }
        case None =>
          // Unknown function - provide suggestions
          val suggestions = suggestFunctions(name)
          Left(ParseError.UnknownFunction(name, startPos, suggestions))
    }

  // ===== GH-193: LET special form =====

  /** Names that collide with literals or operator keywords can never be LET binding names. */
  private val ReservedLetNames = Set("TRUE", "FALSE", "AND", "OR", "NOT")

  /**
   * Valid Excel LET binding name: starts with a letter or underscore, continues with
   * letters/digits/underscores, is not cell-ref shaped (A1, XFD100, ...), and is not a reserved
   * literal/operator keyword.
   */
  private def isValidLetName(name: String): Boolean =
    name.nonEmpty
      && (name.charAt(0).isLetter || name.charAt(0) == '_')
      && name.forall(c => c.isLetterOrDigit || c == '_')
      && ARef.parse(name).isLeft
      && !ReservedLetNames.contains(name.toUpperCase)

  private def invalidLetNameError(name: String, pos: Int): ParseError =
    ParseError.InvalidArguments(
      "LET",
      pos,
      "binding name (starts with a letter or '_', not a cell reference)",
      if name.isEmpty then "empty name" else s"'$name'"
    )

  /** Read an identifier-shaped token (letters/digits/underscores) at the current position. */
  @tailrec
  private def readLetName(s: ParserState): ParserState =
    s.currentChar match
      case Some(c) if c.isLetterOrDigit || c == '_' => readLetName(s.advance())
      case _ => s

  /**
   * Parse LET(name1, value1, [name2, value2, ...], calculation).
   *
   * Lexical scope: binding N is visible to bindings N+1.. and the body (let* semantics, matching
   * Excel). Pair-vs-body disambiguation is a one-token lookahead: at a pair position, an identifier
   * directly followed by ',' is a binding name; anything else is the final calculation. The final
   * argument must therefore be the body — a trailing name/value pair without a body is rejected, as
   * is LET without at least one pair.
   *
   * Range-shaped binding values (A1:B10, Sheet2!A1:A10) are substituted at use sites (see
   * [[LetScopeEntry]]); all other bindings parse to [[TExpr.BindingRef]] and are evaluated against
   * the runtime environment.
   *
   * @param state
   *   Parser state positioned at the opening '('
   */
  private def parseLet(state: ParserState, startPos: Int): ParseResult[TExpr[?]] =
    val entryScope = state.scope

    def expectComma(s: ParserState, context: String): Either[ParseError, ParserState] =
      val sw = skipWhitespace(s)
      sw.currentChar match
        case Some(',') => Right(skipWhitespace(sw.advance()))
        case Some(c) => Left(ParseError.UnexpectedChar(c, sw.pos, context))
        case None => Left(ParseError.UnexpectedEOF(sw.pos, context))

    def scopeEntryFor(name: String, value: TExpr[?]): LetScopeEntry =
      value match
        case r: TExpr.RangeRef => LetScopeEntry(name, Some(r))
        case sr: TExpr.SheetRange => LetScopeEntry(name, Some(sr))
        case _ => LetScopeEntry(name, None)

    /**
     * Lookahead for a pair position: Some((name, namePos, stateAfterName)) iff an identifier-shaped
     * token directly followed (modulo whitespace) by ',' starts here.
     */
    def identCommaLookahead(s: ParserState): Option[(String, Int, ParserState)] =
      val sw = skipWhitespace(s)
      sw.currentChar match
        case Some(c) if c.isLetter || c == '_' =>
          val sEnd = readLetName(sw)
          val name = sw.input.substring(sw.pos, sEnd.pos)
          if skipWhitespace(sEnd).currentChar.contains(',') then Some((name, sw.pos, sEnd))
          else None
        case _ => None

    def parsePairsAndBody(
      s: ParserState,
      scope: List[LetScopeEntry],
      acc: List[(String, TExpr[?])]
    ): ParseResult[TExpr[?]] =
      identCommaLookahead(s) match
        case Some((rawName, namePos, afterName)) =>
          if !isValidLetName(rawName) then Left(invalidLetNameError(rawName, namePos))
          else
            for
              afterComma <- expectComma(afterName, "expected ',' after LET binding name")
              // The value expression sees only PRIOR bindings (lexical, let* semantics)
              (value, afterValue) <- parseExpr(afterComma.copy(scope = scope))
              next <- expectComma(
                afterValue,
                "expected ',' and a final calculation after LET binding value"
              )
              result <- parsePairsAndBody(
                next,
                scopeEntryFor(rawName, value) :: scope,
                (rawName, value) :: acc
              )
            yield result
        case None =>
          if acc.isEmpty then
            Left(
              ParseError.InvalidArguments(
                "LET",
                startPos,
                "at least one name/value pair and a calculation",
                "no name/value pair"
              )
            )
          else
            parseExpr(s.copy(scope = scope)).flatMap { case (body, afterBody) =>
              val sw = skipWhitespace(afterBody)
              sw.currentChar match
                case Some(')') =>
                  Right((TExpr.Let(acc.reverse, body), sw.advance().copy(scope = entryScope)))
                case Some(c) =>
                  Left(ParseError.UnexpectedChar(c, sw.pos, "expected ')' to close LET"))
                case None => Left(ParseError.UnexpectedEOF(sw.pos, "expected ')' to close LET"))
            }

    // Skip opening '('
    parsePairsAndBody(skipWhitespace(state.advance()), entryScope, Nil)

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
   * Parse quoted sheet name reference: 'Sheet Name'!A1 or 'Sheet-Name'!A1:B10
   *
   * Handles:
   *   - Sheet names with spaces: 'Q1 Report'!A1
   *   - Sheet names with special characters: 'Sales&Marketing'!A1, 'Jan-Mar'!A1
   *   - Escaped single quotes: 'O''Brien''s Data'!A1 ('' becomes ')
   *   - Sheet names starting with digits: '2024Q1'!A1
   *
   * @param state
   *   Parser state positioned at the opening single quote
   */
  private def parseQuotedSheetRef(state: ParserState): ParseResult[TExpr[?]] =
    val startPos = state.pos

    // Skip opening quote
    val s1 = state.advance()

    // Read sheet name until closing quote, handling escaped quotes ('')
    @tailrec
    def readQuotedName(
      s: ParserState,
      acc: StringBuilder
    ): Either[ParseError, (String, ParserState)] =
      s.currentChar match
        case None =>
          Left(ParseError.UnexpectedEOF(s.pos, "unterminated quoted sheet name"))
        case Some('\'') =>
          // Check for escaped quote ('') vs closing quote
          s.advance().currentChar match
            case Some('\'') =>
              // Escaped quote - add single quote and continue
              readQuotedName(s.advance(2), acc.append('\''))
            case _ =>
              // Closing quote - done reading sheet name
              Right((acc.toString, s.advance()))
        case Some(c) =>
          readQuotedName(s.advance(), acc.append(c))

    readQuotedName(s1, new StringBuilder).flatMap { case (sheetName, s2) =>
      // Expect '!' after closing quote
      s2.currentChar match
        case Some('!') =>
          // Delegate to existing sheet-qualified ref parser
          parseSheetQualifiedRef(sheetName, s2.advance(), startPos)
        case Some(c) =>
          Left(ParseError.UnexpectedChar(c, s2.pos, "expected '!' after quoted sheet name"))
        case None =>
          Left(ParseError.UnexpectedEOF(s2.pos, "expected '!' after quoted sheet name"))
    }

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
        // Create RangeRef for range arguments
        Right((TExpr.RangeRef(range), s3))
      case Left(err) =>
        Left(ParseError.InvalidCellRef(rangeStr, startPos, err))

  /**
   * Suggest similar function names for unknown functions.
   */
  private def suggestFunctions(name: String): List[String] =
    // Use registry list for suggestions
    val knownFunctions = FunctionRegistry.allNames

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
