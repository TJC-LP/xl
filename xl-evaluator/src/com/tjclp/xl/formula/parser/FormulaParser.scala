package com.tjclp.xl.formula.parser

import com.tjclp.xl.formula.ast.{RangeForm, TExpr}
import com.tjclp.xl.formula.eval.ArrayResult
import com.tjclp.xl.formula.functions.{
  FunctionRegistry,
  FunctionSpec,
  FunctionSpecs,
  ReferenceOperators
}
import com.tjclp.xl.formula.{Arity}

import com.tjclp.xl.{ARef, Anchor, CellRange, SheetName}
import com.tjclp.xl.addressing.RefParser
import com.tjclp.xl.cells.{Cell, CellError, CellValue}
import com.tjclp.xl.codec
import com.tjclp.xl.ooxml.FormulaStorage

import scala.annotation.tailrec

/**
 * Parser for Excel formula strings to typed TExpr AST.
 *
 * Implements a recursive descent parser with operator precedence. Pure functional - no mutation, no
 * exceptions, all errors as Either.
 *
 * Supported syntax:
 *   - Literals: numbers (42, 3.14), booleans (TRUE, FALSE), strings ("text")
 *   - Cell references: A1, $A$1, Sheet1!A1, 'Quoted Name'!A1
 *   - Ranges: A1:B10
 *   - External-workbook references (GH-353): [2]Book1!A1, [2]Book1!A1:B2, '[3]Sheet Name'!B2
 *     (unresolvable — carried as ExternalRef/ExternalRange for closed-workbook cache semantics)
 *   - Operators: +, -, *, /, ^, =, <>, <, <=, >, >=, &, postfix % (GH-355)
 *   - Reference operators (GH-669): union `(A1,B2:C3)` — inside parentheses only; LibreOffice's `~`
 *     is read as `,` — and intersection `A1:C3 B2:D4` (one or more spaces between two references)
 *   - Array constants (GH-669): `{1,2;3,4}` — `,` between columns, `;` between rows
 *   - Functions: SUM, COUNT, IF, AND, OR, NOT, etc.; `TRUE()`/`FALSE()` as calls (GH-669)
 *   - Parentheses for grouping
 *
 * Sheet-name leniency (GH-281): any identifier followed by `!` parses as a sheet reference,
 * including cell-ref-shaped names that Excel itself rejects unquoted — `=Q1!A1` parses here as
 * sheet Q1, while Excel requires `='Q1'!A1`. This is intentional (Postel-style): such input has
 * exactly one plausible meaning (Excel forbids defined names that collide with cell-ref shapes, so
 * no ambiguity is possible), and FormulaPrinter always canonicalizes cell-ref-shaped sheet names to
 * the quoted form on print (GH-263), so every formula xl emits is spec-valid and parse/print
 * round-trips are stable. Consequence: xl accepts a strict superset of Excel's sheet-ref grammar;
 * there is no strict-parity mode that rejects the unquoted form.
 *
 * Operator precedence (highest to lowest, Excel-compatible):
 *   1. Parentheses (), function calls, the range colon `:`
 *   2. Intersection ` ` (GH-669, left-associative; parseIntersection)
 *   3. Union `,` (GH-669, only inside parentheses)
 *   4. Unary minus -, unary plus + (GH-578: tighter than ^, -2^2 = 4)
 *   5. Postfix percent % (GH-355: binds tighter than ^, 2^3% = 2^(3%))
 *   6. Exponentiation ^ (left-associative, GH-480)
 *   7. Multiplication *, Division /
 *   8. Addition +, Subtraction -
 *   9. Concatenation &
 *   10. Comparison =, <>, <, <=, >, >=
 *   11. Logical AND (xl's lenient infix keyword)
 *   12. Logical OR (xl's lenient infix keyword)
 *
 * Grammar sketch (one level per line, each folding left unless noted):
 * {{{
 * expr        := or
 * or          := and ("OR" or)?                       -- right-nested keyword form
 * and         := comparison ("AND" and)?
 * comparison  := concat (("=" | "<>" | "<" | "<=" | ">" | ">=") concat)*
 * concat      := addsub ("&" addsub)*
 * addsub      := muldiv (("+" | "-") muldiv)*
 * muldiv      := unary (("*" | "/") unary)*
 * unary       := "NOT" unary | pow                    -- NOT followed by an operand only
 * pow         := signed ("^" signed)*
 * signed      := ("-" | "+") signed | postfix
 * postfix     := intersection ("%" | "#")*
 * intersection:= primary (" "+ primary)*              -- both operands references
 * primary     := number | string | ref | range | name | call | error | "{" array "}"
 *              | "(" expr ("," expr)* ")"             -- two or more: a union of references
 *              | "@" primary
 * }}}
 * Nesting — parentheses, arguments, prefix and postfix operators, intersections — is bounded by
 * MaxNestingDepth; a flat chain of binary operators costs no level and the operator count is
 * bounded instead (GH-680).
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
    // GH-374: unary plus is a transparent wrapper — a bare `=+A1` must resolve exactly like `=A1`
    case TExpr.UnaryPlus(inner) => TExpr.UnaryPlus(resolveTopLevelPolyRef(inner))
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
   * @param depth
   *   Nesting levels entered on the current path (parentheses, function arguments, prefix and
   *   postfix operators) — restored on the way out, so siblings do not accumulate
   * @param operators
   *   GH-680: binary operators consumed so far in the whole formula — monotone, never restored
   * @param scope
   *   GH-193: lexically visible LET bindings, innermost first
   */
  private case class ParserState(
    input: String,
    pos: Int,
    depth: Int = 0,
    scope: List[LetScopeEntry] = Nil,
    operators: Int = 0
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
   * Maximum formula nesting depth — a stack-overflow guard (GH-56): 2x Excel's own 64-level cap on
   * nesting functions and parentheses. Each level costs a run of recursive-descent frames (the
   * precedence ladder), and 256 levels sat at the edge of a default 1MB thread stack under
   * interpreted execution (CI-observed StackOverflowError in the depth-guard tests themselves).
   * GH-680: only real nesting counts — a flat operator chain costs nothing here (see
   * [[parseChain]]); its length is bounded by [[MaxOperators]].
   */
  private val MaxNestingDepth = 128

  /**
   * GH-680: the most binary operators one formula may hold. A left-associative chain builds a
   * left-nested spine one node per operator. The heavy walkers (evaluator, printer, shifter,
   * dependency extraction, analysis) take a spine in one loop
   * ([[com.tjclp.xl.formula.ast.BinarySpine]]), so this bound is not what keeps them on the stack;
   * it is a margin for what still recurses along a spine — case-class `equals`/`hashCode` —
   * measured cold (interpreted) on a 1MB thread stack clear to 2000 operators under 64 nesting
   * levels. 1024 is eight times the 130-term chains generated books carry, and a longer chain is
   * `TooManyOperators`: a total `Left`, never a StackOverflowError.
   */
  private val MaxOperators = 1024

  /**
   * Guarded descent into a nested sub-expression: deepen the state, or fail if already too deep.
   */
  private def descend(s: ParserState): Either[ParseError, ParserState] =
    if s.depth >= MaxNestingDepth then Left(ParseError.NestingTooDeep(s.depth, MaxNestingDepth))
    else Right(s.copy(depth = s.depth + 1))

  /** GH-680: count one more binary operator against [[MaxOperators]]. */
  private def countOperator(s: ParserState): Either[ParseError, ParserState] =
    if s.operators >= MaxOperators then
      Left(ParseError.TooManyOperators(s.operators + 1, MaxOperators, s.pos))
    else Right(s.copy(operators = s.operators + 1))

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
   * GH-653: the paren-less prefix form `NOT x` is taken only when the first non-whitespace
   * character after the word is not '(' (or the input ends). `NOT(` — and `NOT (`, whitespace
   * before the paren, which [[parseFunctionOrRef]] accepts for every function name — is the
   * function CALL, whose closing paren ends it, so a postfix after it binds to the call as in
   * Excel: `NOT(A1)^2` is `(NOT(A1))^2`, `NOT(A1)%` is `(NOT(A1))%` and `NOT(A1)#` is an error.
   * Read as the keyword with a parenthesized operand, the postfix bound INSIDE (`NOT(A1^2)`,
   * `NOT(A1%)`, `NOT(A1#)`) — a different value (found by the grammar generator; the whitespace
   * spelling by the PR #659 review).
   */
  private def isNotKeywordAt(s: ParserState): Boolean =
    isKeywordAt(s, "NOT") && {
      val rem = s.remaining
      val next = rem.indexWhere(!_.isWhitespace, 3)
      // GH-669: with no operand after it — the end of the text, a binary operator, a closer or a
      // separator — the word is a NAME (`=NOT`, `=NOT&"x"`, `IF(NOT,…)`; NOT is a legal defined
      // name Excel resolves), read by parsePrimary like any other identifier
      next >= 0 && rem.charAt(next) != '(' && startsOperand(rem.charAt(next))
    }

  /**
   * Can `c` begin an operand? Letters, digits and the characters that open a reference, a literal
   * or a group — and the signs, which the paren-less `NOT -x` form has always read as the
   * operand's.
   */
  private def startsOperand(c: Char): Boolean =
    c.isLetterOrDigit || "._\\$'[#@\"{+-".contains(c)

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

  /** A binary operator at the cursor: the characters it spells and the node it builds. */
  private type ChainOperator = (Int, (TExpr[?], TExpr[?]) => TExpr[?])

  /**
   * GH-680: parse a left-associative chain `operand (op operand)*` as an iterative left fold.
   *
   * A chain costs nothing against the nesting budget: its right operands never re-enter
   * [[parseExpr]] (only parentheses and function arguments do), so the parser's stack is flat
   * however long `B2+B3+…` runs, and Excel's limit is the nesting of functions and parentheses, not
   * chain length. What a chain does deepen is the AST's left spine (GH-56's reason for once
   * counting every segment as a level); the walkers take a spine in one loop (BinarySpine) and the
   * per-formula operator budget ([[countOperator]]) bounds what is left.
   */
  private def parseChain(
    state: ParserState,
    operand: ParserState => ParseResult[TExpr[?]],
    operatorAt: ParserState => Option[ChainOperator]
  ): ParseResult[TExpr[?]] =
    operand(state).flatMap { case (first, s1) =>
      @tailrec
      def loop(acc: TExpr[?], s: ParserState): ParseResult[TExpr[?]] =
        val s2 = skipWhitespace(s)
        operatorAt(s2) match
          case None => Right((acc, s2))
          case Some((consumed, build)) =>
            countOperator(s2) match
              case Left(err) => Left(err)
              case Right(sc) =>
                operand(skipWhitespace(sc.advance(consumed))) match
                  case Right((right, s3)) => loop(build(acc, right), s3)
                  case Left(err) => Left(err)
      loop(first, s1)
    }

  /**
   * Parse comparison operators: =, <>, <, <=, >, >= (left-associative).
   *
   * GH-455: chained comparisons fold LEFT like Excel's equal-precedence rule (`=1=1=TRUE` is
   * `(1=1)=TRUE`); right-recursion here would build right-nested ASTs that the printer must then
   * parenthesize, destroying byte-fidelity of untouched chained formulas on every rewrite.
   *
   * GH-233: PolyRef operands resolve polymorphically so =A1=B1 / =IF(A1=B1,…) evaluate instead of
   * erroring with "Unresolved PolyRef". GH-335: the comparable decoder preserves emptiness so empty
   * cells equal 0 / "" / FALSE like Excel, compares text lexicographically, and ranks number < text
   * < logical (numeric-only coercion would reject text operands).
   */
  private def parseComparison(state: ParserState): ParseResult[TExpr[?]] =
    def op(
      width: Int,
      node: (TExpr[CellValue], TExpr[CellValue]) => TExpr[Boolean]
    ): Option[ChainOperator] =
      Some((width, (l, r) => node(TExpr.asComparableValueExpr(l), TExpr.asComparableValueExpr(r))))
    parseChain(
      state,
      parseConcatenation,
      s =>
        s.currentChar match
          case Some('=') => op(1, TExpr.Eq.apply)
          case Some('<') =>
            s.advance().currentChar match
              case Some('>') => op(2, TExpr.Neq.apply)
              case Some('=') => op(2, TExpr.Lte.apply)
              case _ => op(1, TExpr.Lt.apply)
          case Some('>') =>
            s.advance().currentChar match
              case Some('=') => op(2, TExpr.Gte.apply)
              case _ => op(1, TExpr.Gt.apply)
          case _ => None
    )

  /**
   * Parse concatenation operator: & (left-associative).
   *
   * Excel's & operator joins strings: "Hello" & "World" → "HelloWorld" Operands are coerced to
   * strings via asStringExpr. GH-455: chained & folds LEFT (Excel's equal-precedence rule), so
   * `=A1&B1&C1` round-trips without the printer inventing grouping parens.
   */
  private def parseConcatenation(state: ParserState): ParseResult[TExpr[?]] =
    parseChain(
      state,
      parseAddSub,
      s =>
        Option.when(s.currentChar.contains('&'))(
          (1, (l, r) => TExpr.Concat(TExpr.asStringExpr(l), TExpr.asStringExpr(r)))
        )
    )

  /** An arithmetic operator whose operands keep RangeRef for array arithmetic. */
  private def arithmetic(node: (TExpr[BigDecimal], TExpr[BigDecimal]) => TExpr[?]): ChainOperator =
    (1, (l, r) => node(TExpr.asNumericOrRangeExpr(l), TExpr.asNumericOrRangeExpr(r)))

  /** Parse addition and subtraction (left-associative). */
  private def parseAddSub(state: ParserState): ParseResult[TExpr[?]] =
    parseChain(
      state,
      parseMulDiv,
      s =>
        s.currentChar match
          case Some('+') => Some(arithmetic(TExpr.Add.apply))
          case Some('-') if !s.remaining.startsWith("->") => Some(arithmetic(TExpr.Sub.apply))
          case _ => None
    )

  /** Parse multiplication and division (left-associative). */
  private def parseMulDiv(state: ParserState): ParseResult[TExpr[?]] =
    parseChain(
      state,
      parseUnary,
      s =>
        s.currentChar match
          case Some('*') => Some(arithmetic(TExpr.Mul.apply))
          case Some('/') => Some(arithmetic(TExpr.Div.apply))
          case _ => None
    )

  /**
   * Parse exponentiation (LEFT-associative, highest binary arithmetic precedence).
   *
   * GH-480: Excel folds chained '^' from the left like every other binary operator: 2^3^2 = (2^3)^2 =
   * 64, not 2^(3^2) = 512 (the mathematical convention this parser used to follow, which silently
   * changed the value of Excel-authored chained-pow formulas). GH-578: Excel's negation binds
   * TIGHTER than '^' (Microsoft's precedence table lists negation above percent and
   * exponentiation), so both the base and the exponent are signed operands: -2^2 = (-2)^2 = 4, 2^-1 =
   * 0.5, and a signed exponent is one operand of the left fold (2^-3^2 = (2^-3)^2). Binary
   * subtraction is unaffected: 0-2^2 = -4.
   */
  private def parsePow(state: ParserState): ParseResult[TExpr[?]] =
    parseChain(
      state,
      parseSigned,
      s =>
        Option.when(s.currentChar.contains('^'))(
          (1, (l, r) => TExpr.Pow(TExpr.asNumericExpr(l), TExpr.asNumericExpr(r)))
        )
    )

  /**
   * GH-355: parse postfix percent — Excel's tightest-binding operator (value ÷ 100).
   *
   * Sits between parsePow and parsePrimary in the precedence ladder so `%` binds tighter than `^`
   * (2^3% ≡ 2^(3%)) and tighter than unary minus (-2% ≡ -(2%)). A loop consumes chained percents
   * (10%% = 0.001), each wrapping in TExpr.Percent and counting against the depth budget (chained
   * postfixes nest the AST like chained '+' terms, GH-56). The operand coerces numerically with
   * ranges preserved so `=A1:A3%` broadcasts through the array machinery.
   */
  /**
   * GH-669: parse the intersection operator — one or more U+0020 spaces between two references
   * (`A1:C3 B2:D4`, `A:A 3:3`, two names `Jan Sales`). It binds tighter than every operator but `:`
   * (Excel's order: `:` > space > `,` > negation > `%` > `^` > the rest), so it sits between the
   * postfix level and the primaries, and folds left.
   *
   * A space is an intersection only when all of these hold — otherwise it is insignificant
   * whitespace and the state before it is returned: the operand so far is a reference; the gap is
   * spaces only (a tab or line feed never is); the next character can start a reference (a letter,
   * `_`, `\`, `$`, `'`, `[`, `(`, a digit for `5:5`, or `#` opening an error literal — a deletion
   * leaves `A1:B2 #REF!`); and it is not xl's lenient infix `AND`/`OR`. Every continuation that
   * triggers was a parse error before, so no formula that parsed changes meaning. Once triggered,
   * the right operand must be a reference too.
   *
   * Each intersection keeps the nesting level it takes, as a chained `%` does: the chain builds a
   * left spine of intersection calls that the walkers recurse along (BinarySpine covers only the
   * binary operators), so the chain shares the nesting budget instead of the operator budget.
   */
  private def parseIntersection(state: ParserState): ParseResult[TExpr[?]] =
    parsePrimary(state).flatMap { case (first, s1) =>
      def triggers(s: ParserState, s2: ParserState): Boolean =
        s2.pos > s.pos && s.input.substring(s.pos, s2.pos).forall(_ == ' ') &&
          s2.currentChar.exists { c =>
            c.isLetterOrDigit || "_\\$'[(".contains(c) ||
            (c == '#' && parseErrorLiteral(s2).isRight)
          } && !isKeywordAt(s2, "AND") && !isKeywordAt(s2, "OR")
      @tailrec
      def loop(acc: TExpr[?], s: ParserState): ParseResult[TExpr[?]] =
        val s2 = skipWhitespace(s)
        if !(ReferenceOperators.isReferenceShape(acc) && triggers(s, s2)) then Right((acc, s))
        else
          descend(s2) match
            case Left(err) => Left(err)
            case Right(sd) =>
              parsePrimary(sd).flatMap { case (right, s3) =>
                ReferenceOperators.mkIntersection(acc, right, s2.pos).map((_, s3))
              } match
                case Right((node, s3)) => loop(node, s3.copy(depth = sd.depth))
                case Left(err) => Left(err)
      loop(first, s1)
    }

  private def parsePostfix(state: ParserState): ParseResult[TExpr[?]] =
    parseIntersection(state).flatMap { case (first, s1) =>
      @tailrec
      def loop(acc: TExpr[?], s: ParserState): ParseResult[TExpr[?]] =
        val s2 = skipWhitespace(s)
        s2.currentChar match
          case Some('%') =>
            descend(s2) match
              case Left(err) => Left(err)
              case Right(sd) =>
                loop(TExpr.Percent(TExpr.asNumericOrRangeExpr(acc)), sd.advance())
          // GH-655: the spill reference `x#` — the formula-bar spelling of `_xlfn.ANCHORARRAY(x)`,
          // postfix on one cell reference or name (never on a range, a value or a call)
          case Some('#') if FunctionSpecs.isSpillAnchorShape(acc) =>
            parseSpillSuffix(acc, s2) match
              case Left(err) => Left(err)
              case Right((spill, s3)) => loop(spill, s3)
          case _ => Right((acc, s2))
      loop(first, s1)
    }

  /**
   * GH-655: wrap `operand` — a cell reference or a name whose next character is `#` — in the
   * ANCHORARRAY call the spill reference denotes, consuming the `#`. The parser's postfix loop and
   * the `@` operand slot share it, so `@A1#` reads as `@(A1#)`.
   */
  private def parseSpillSuffix(operand: TExpr[?], state: ParserState): ParseResult[TExpr[?]] =
    descend(state).flatMap { sd =>
      FunctionSpecs.anchorArray.argSpec
        .parse(List(operand), state.pos, FunctionSpecs.anchorArray.name)
        .map { case (parsed, _) =>
          (TExpr.Call(FunctionSpecs.anchorArray, parsed), sd.advance().copy(depth = state.depth))
        }
    }

  /**
   * The `#`-suffixed form of a primary when one follows an anchor-shaped operand, else the primary.
   */
  private def withSpillSuffix(primary: TExpr[?], state: ParserState): ParseResult[TExpr[?]] =
    if state.currentChar.contains('#') && FunctionSpecs.isSpillAnchorShape(primary) then
      parseSpillSuffix(primary, state)
    else Right((primary, state))

  /**
   * Parse a signed operand of '^' — either side, GH-578 — allowing unary minus and plus: -2^2 is
   * (-2)^2 = 4 like Excel, 2^-1 = 0.5, 2^-3^2 = (2^-3)^2. Unary minus is `0 - x` (the shape the
   * printer renders as `-x`); unary plus wraps in TExpr.UnaryPlus so `=+A1` and `=2^+2` print back
   * byte-identically (GH-271 acceptance, GH-374 preservation). Chained signs nest, each counting
   * against the depth budget (GH-56 recursion guard). Postfix percent binds tighter still (-2% is
   * -(2%), GH-355). GH-480: the operand is a single postfix term — a following '^' belongs to the
   * enclosing left fold.
   */
  private def parseSigned(state: ParserState): ParseResult[TExpr[?]] =
    val s = skipWhitespace(state)
    s.currentChar match
      case Some('-') =>
        descend(s).flatMap { sd =>
          val s2 = skipWhitespace(sd.advance())
          parseSigned(s2).map { case (expr, s3) =>
            // Unary minus: 0 - expr (asNumericExpr converts PolyRef)
            (
              TExpr.Sub(TExpr.Lit(BigDecimal(0)), TExpr.asNumericExpr(expr)),
              s3.copy(depth = s.depth)
            )
          }
        }
      case Some('+') =>
        descend(s).flatMap { sd =>
          val s2 = skipWhitespace(sd.advance())
          parseSigned(s2).map { case (expr, s3) =>
            (TExpr.UnaryPlus(expr), s3.copy(depth = s.depth))
          }
        }
      // GH-654: a sign may precede the lenient paren-less NOT keyword (`=-NOT x` parsed before
      // GH-578 moved the signs down here) — it is a prefix-level operand, not a primary
      case Some('N') | Some('n') if isNotKeywordAt(s) => parseUnary(s)
      case _ => parsePostfix(s)

  /**
   * Parse the prefix-operator level: NOT (xl's lenient keyword form of the NOT function).
   *
   * The arithmetic signs are NOT parsed here: since GH-578 unary minus and plus bind tighter than
   * '^' (Excel's precedence table), so [[parsePow]] parses them on both of its operands via
   * [[parseSigned]]. Only the NOT keyword sits between multiplication and exponentiation.
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  private def parseUnary(state: ParserState): ParseResult[TExpr[?]] =
    val s = skipWhitespace(state)
    s.currentChar match
      case Some('N') | Some('n') if isNotKeywordAt(s) =>
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
        // Parenthesized expression — or, GH-669, a union of references `(A1,B2:C3)`: after the
        // first element a `,` (or LibreOffice's xlsx spelling `~`, accepted here only and never
        // printed) continues the group. Each element descends through parseExpr; the group itself
        // adds no level. Any other character keeps the unbalanced-delimiter diagnosis.
        @tailrec
        def elements(acc: List[TExpr[?]], st: ParserState): ParseResult[List[TExpr[?]]] =
          val sw = skipWhitespace(st)
          sw.currentChar match
            case Some(')') => Right((acc.reverse, sw.advance()))
            case Some(',' | '~') if acc.nonEmpty =>
              parseExpr(sw.advance()) match
                case Right((next, after)) => elements(next :: acc, after)
                case Left(err) => Left(err)
            case Some(c) => Left(ParseError.UnbalancedDelimiter(sw.pos, c, "expected ')'"))
            case None => Left(ParseError.UnexpectedEOF(sw.pos, "expected ')'"))
        val s2 = skipWhitespace(s.advance())
        parseExpr(s2).flatMap { case (first, s3) =>
          val unionAt = skipWhitespace(s3).pos
          elements(List(first), s3).flatMap {
            case (List(single), after) => Right((single, after))
            case (areas, after) => ReferenceOperators.mkUnion(areas, unionAt).map((_, after))
          }
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
      case Some(c) if c.isLetter || c == '_' || c == '\\' =>
        // Function call, cell reference, boolean literal, LET binding reference, or defined
        // name ('_' admits underscore-led LET names, GH-193; '\\' admits backslash-led
        // defined names, GH-394 — legal Excel name characters)
        parseFunctionOrRef(s)
      case Some('$') =>
        // Anchored cell reference (e.g., $A$1, $A1)
        parseAnchoredCellRef(s)
      case Some('\'') =>
        // Quoted sheet name reference (e.g., 'Q1 Report'!A1, 'Debt-Schedule'!H29)
        parseQuotedSheetRef(s)
      case Some('[') =>
        // GH-353: external-workbook reference (e.g., [2]Book1!A1, [2]Consolidation.xlsx!D5:D9)
        parseExternalRef(s)
      case Some('#') =>
        // GH-612: error literal (#REF!, #N/A, #DIV/0!, …)
        parseErrorLiteral(s)
      case Some('{') =>
        // GH-669: array constant ({1,2;3,4})
        parseArrayConstant(s)
      case Some('@') =>
        // GH-604: implicit intersection — `@x` is the formula-bar spelling of `_xlfn.SINGLE(x)`.
        // The operand is one primary (a reference, name, call, or parenthesized expression), so
        // `@A1:A10%` is (@A1:A10)% and `@A1^2` is (@A1)^2, as in Excel.
        descend(s).flatMap { sd =>
          // GH-655: the operand may carry the spill suffix — `@A1#` is `@(A1#)`
          parsePrimary(skipWhitespace(sd.advance()))
            .flatMap { case (primary, s2) => withSpillSuffix(primary, s2) }
            .flatMap { case (operand, s2) =>
              FunctionSpecs.single.argSpec
                .parse(List(operand), s.pos, FunctionSpecs.single.name)
                .map { case (parsed, _) =>
                  (TExpr.Call(FunctionSpecs.single, parsed), s2.copy(depth = s.depth))
                }
            }
        }
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
        // GH-612: the end row may carry its own anchor (3:$10)
        val afterEndAnchor =
          if afterColon.currentChar.contains('$') then afterColon.advance() else afterColon
        afterEndAnchor.currentChar match
          case Some(c) if c.isDigit =>
            // This is a row range like 1:5
            val afterSecondDigits = readDigits(afterEndAnchor)
            val rangeStr = state.input.substring(startPos, afterSecondDigits.pos)
            CellRange.parse(rangeStr) match
              case Right(range) =>
                // GH-612: digits on both sides is Excel's whole-row form
                Right((TExpr.RangeRef(range, RangeForm.Rows), afterSecondDigits))
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

    // GH-653: the loop consumes the decimal point itself, so a leading-dot fraction (`.5`, Excel's
    // own spelling of a fraction below one) reads as `.5` and not as the empty literal — seeding
    // `hasDecimal` from the first character left nothing consumed and reported Invalid number ''
    val s2 = loop(state, hasDecimal = false, hasExponent = false)
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

    // Read identifier (letters, digits, underscores, $ for anchored refs like A$1; '.' and
    // '\\' are Excel name characters, GH-394 — ARef/range parsing rejects them downstream so
    // ref-shaped tokens are unaffected)
    @tailrec
    def readIdent(s: ParserState): ParserState =
      s.currentChar match
        case Some(c) if c.isLetterOrDigit || c == '_' || c == '$' || c == '.' || c == '\\' =>
          readIdent(s.advance())
        case Some('!') => s // Sheet reference separator
        case _ => s

    val s2 = readIdent(state)
    val ident = state.input.substring(startPos, s2.pos).toUpperCase

    // Note: This function is called within the boundary block from parseExpr,
    // so we can use early returns via simple control flow

    // Check for boolean literals
    ident match
      // GH-669: `TRUE()`/`FALSE()` are calls (the registry's zero-argument TRUE and FALSE), so the
      // text prints back as written; the bare word is the literal
      case "TRUE" if !skipWhitespace(s2).currentChar.contains('(') => Right((TExpr.Lit(true), s2))
      case "FALSE" if !skipWhitespace(s2).currentChar.contains('(') =>
        Right((TExpr.Lit(false), s2))
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
              // GH-669: `A1 (B1:C2)` — a cell reference, a space, then a parenthesized reference —
              // is an intersection, not a call to a function named A1: an ARef-shaped identifier
              // that names no function and is not LET, separated from `(` by whitespace, is the
              // reference, and the state before the whitespace goes back for parseIntersection.
              // `LOG10 (x)` and `ATAN2 (…)` stay calls; `A1(` with no space is unchanged.
              case Some('(') if s3.pos > s2.pos && isReferenceNotCall(ident) =>
                parseCellReference(state.input.substring(startPos, s2.pos), s2, startPos)
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
   * GH-669: an identifier that is a cell reference and cannot be a function name (see above) — not
   * a registered function, not LET, and not one of Excel's own cell-shaped function names that xl
   * does not implement (`LOG10 (x)` stays an unknown-function error, never an intersection).
   */
  private def isReferenceNotCall(ident: String): Boolean =
    val bare = FormulaStorage.bareFunctionName(ident)
    ARef.parse(Anchor.parse(ident)._1).isRight && bare != "LET" &&
    !CellShapedExcelFunctions.contains(bare) && FunctionRegistry.lookup(bare).isEmpty

  /** Excel function names that are also valid cell addresses (column letters, row digits). */
  private val CellShapedExcelFunctions = Set("LOG10")

  /**
   * Parse function call: FUNC(arg1, arg2, ...)
   */
  private def parseFunction(
    name: String,
    state: ParserState,
    startPos: Int
  ): ParseResult[TExpr[?]] =
    // GH-556: Excel stores post-2007 functions as _xlfn.NAME (FILTER/SORT as _xlfn._xlws.NAME);
    // inherited formulas may still carry the prefix — drop it before the registry lookup.
    val bareName = FormulaStorage.bareFunctionName(name)
    // GH-193: LET is a special form (it introduces lexical bindings), not a FunctionSpec.
    if bareName == "LET" then parseLet(state, startPos)
    else parseRegularFunction(bareName, state, startPos)

  private def parseRegularFunction(
    name: String,
    state: ParserState,
    startPos: Int
  ): ParseResult[TExpr[?]] =
    // Skip opening '('
    val s2 = skipWhitespace(state.advance())

    // Parse arguments (comma-separated). GH-603: an empty slot — nothing between two commas, or
    // between a comma and the closing paren — is an omitted argument (TExpr.Missing); `F()` alone
    // stays the zero-argument call.
    def parseArgs(
      s: ParserState,
      args: List[TExpr[?]],
      afterComma: Boolean
    ): ParseResult[List[TExpr[?]]] =
      val s1 = skipWhitespace(s)
      s1.currentChar match
        case Some(')') if !afterComma => Right((args.reverse, s1.advance()))
        case Some(')') => Right(((TExpr.Missing :: args).reverse, s1.advance()))
        case Some(',') => parseArgs(skipWhitespace(s1.advance()), TExpr.Missing :: args, true)
        case _ =>
          parseExpr(s1).flatMap { case (arg, s2) =>
            val s3 = skipWhitespace(s2)
            s3.currentChar match
              case Some(',') => parseArgs(skipWhitespace(s3.advance()), arg :: args, true)
              case Some(')') => Right(((arg :: args).reverse, s3.advance()))
              case Some(c) =>
                Left(ParseError.UnexpectedChar(c, s3.pos, "expected ',' or ')'"))
              case None => Left(ParseError.UnexpectedEOF(s3.pos, "expected ',' or ')'"))
          }

    parseArgs(s2, Nil, afterComma = false).flatMap { case (args, finalState) =>
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
   * Valid Excel name identifier (shared by LET binding names and defined-name references): starts
   * with a letter, underscore, or backslash, continues with letters/digits/underscores plus '.' and
   * '\\' (GH-394 — Excel allows both in defined names, e.g. Sales.Total), is not cell-ref shaped
   * (A1, XFD100, ...), and is not a reserved literal/operator keyword. Number literals (3.14) can
   * never reach this predicate: it runs on identifier-shaped tokens, which always start with a
   * letter/underscore/backslash.
   */
  private def isValidLetName(name: String): Boolean =
    name.nonEmpty
      && (name.charAt(0).isLetter || name.charAt(0) == '_' || name.charAt(0) == '\\')
      && name.forall(c => c.isLetterOrDigit || c == '_' || c == '.' || c == '\\')
      && ARef.parse(name).isLeft
      && !ReservedLetNames.contains(name.toUpperCase)

  private def invalidLetNameError(name: String, pos: Int): ParseError =
    ParseError.InvalidArguments(
      "LET",
      pos,
      "binding name (starts with a letter or '_', not a cell reference)",
      if name.isEmpty then "empty name" else s"'$name'"
    )

  /**
   * Read an identifier-shaped token at the current position (letters/digits/underscores plus '.'
   * and '\\' — the Excel name character set, GH-394).
   */
  @tailrec
  private def readLetName(s: ParserState): ParserState =
    s.currentChar match
      case Some(c) if c.isLetterOrDigit || c == '_' || c == '.' || c == '\\' =>
        readLetName(s.advance())
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
        // GH-384: a name-shaped identifier that is not a cell reference is a defined-name
        // reference (=case, =entry_mult), resolved against the workbook at evaluation time.
        // Disambiguation is total: ARef.parse SUCCESS above means cell ref (TAX1 is a cell —
        // OOXML forbids ref-shaped defined names); name-shaped failure means NameRef; anything
        // else (e.g. '$' inside the token) keeps today's InvalidCellRef. Reuses the LET name
        // shape (letter/underscore start, [A-Za-z0-9_]*, not TRUE/FALSE/AND/OR/NOT).
        // GH-669: NOT reaches here only where it is not the prefix operator (isNotKeywordAt) —
        // then it is a defined name, as Excel reads it
        if isValidLetName(refStr) || refStr.equalsIgnoreCase("NOT") then
          Right((TExpr.NameRef(refStr), state))
        else Left(ParseError.InvalidCellRef(refStr, startPos, err))

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
          // GH-353: quoted external-workbook reference — the bracket prefix sits INSIDE the
          // quotes ('[3]Sheet Name'!B2). Split it off and route to the external-ref parser.
          splitExternalPrefix(sheetName) match
            case Some((index, externalName)) =>
              parseExternalQualifiedRef(index, externalName, s2.advance(), startPos)
            case None =>
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

        // '.'/'_'/'\\' are Excel name characters (GH-394): a name-shaped tail becomes a
        // sheet-qualified defined name below; ref shapes are unaffected (ARef/CellRange
        // parsing decides)
        @tailrec
        def readRef(s: ParserState): ParserState =
          s.currentChar match
            case Some(c)
                if c.isLetterOrDigit || c == '$' || c == ':' || c == '_' || c == '.' ||
                  c == '\\' =>
              readRef(s.advance())
            case _ => s

        val s1 = readRef(state)
        val written = state.input.substring(refStartPos, s1.pos)

        // GH-669: `Sheet1!A1:Sheet1!B2` — the range's end repeats the sheet qualifier
        // (LibreOffice and hand-written formulas spell it so): the same range as `Sheet1!A1:B2`.
        // Two different sheets would be a 3-D reference, which xl does not implement — refused,
        // never silently narrowed to one sheet.
        def repeatedEnd: Either[ParseError, (String, ParserState)] =
          val colon = written.lastIndexOf(':')
          val qualifier: Option[(String, ParserState)] =
            if colon < 0 then None
            else if s1.currentChar.contains('!') && colon < written.length - 1 then
              Some((written.substring(colon + 1), s1.advance()))
            else if s1.currentChar.contains('\'') && colon == written.length - 1 then
              quotedQualifier(s1)
            else None
          qualifier match
            case None => Right((written, s1))
            case Some((endSheet, afterBang)) if endSheet.equalsIgnoreCase(sheetStr) =>
              val sEnd = readRef(afterBang)
              val endRef = state.input.substring(afterBang.pos, sEnd.pos)
              Right((written.substring(0, colon + 1) + endRef, sEnd))
            case Some((endSheet, _)) =>
              Left(
                ParseError.InvalidCellRef(
                  state.input.substring(startPos, s1.pos),
                  startPos,
                  s"a range's two ends must be on the same sheet ('$sheetStr' and '$endSheet'); " +
                    "3-D references are not supported"
                )
              )

        repeatedEnd match
          case Left(err) => Left(err)
          case Right((refPart, s2)) =>
            if refPart.isEmpty then
              Left(
                ParseError.InvalidCellRef(s"$sheetStr!", startPos, "missing cell reference after !")
              )
            else if refPart.contains(':') then
              // Range reference: Sheet1!A1:B10
              CellRange.parse(refPart) match
                case Right(range) =>
                  // GH-612: keep the whole-column / whole-row form the text spelled (a corner range
                  // over every row/column canonicalises to it, as Excel does at entry)
                  Right((TExpr.SheetRange(sheetName, range, RangeForm.of(refPart, range)), s2))
                case Left(err) =>
                  Left(ParseError.InvalidCellRef(s"$sheetStr!$refPart", startPos, err))
            else
              // Single cell reference: Sheet1!A1
              val (cleanRef, anchor) = Anchor.parse(refPart)
              ARef.parse(cleanRef) match
                case Right(aref) =>
                  Right((TExpr.SheetPolyRef(sheetName, aref.asInstanceOf[ARef], anchor), s2))
                case Left(err) =>
                  // GH-394: a name-shaped tail is a sheet-qualified defined name (=Model!case —
                  // legal Excel syntax for sheet-scoped names), resolved against the workbook at
                  // evaluation. Total disambiguation like parseCellReference: ARef success means
                  // cell ref, name-shaped failure means SheetNameRef, anything else keeps the
                  // InvalidCellRef.
                  if isValidLetName(refPart) then
                    Right((TExpr.SheetNameRef(sheetName, refPart), s2))
                  else Left(ParseError.InvalidCellRef(s"$sheetStr!$refPart", startPos, err))

  /**
   * GH-669: a quoted sheet qualifier at the cursor — `'My Sheet'!`, `''` escaping a quote — as the
   * sheet name and the state after its `!`; None when the text is not one.
   */
  private def quotedQualifier(state: ParserState): Option[(String, ParserState)] =
    @tailrec
    def loop(s: ParserState, acc: StringBuilder): Option[(String, ParserState)] =
      s.currentChar match
        case None => None
        case Some('\'') if s.advance().currentChar.contains('\'') =>
          loop(s.advance(2), acc.append('\''))
        case Some('\'') =>
          val after = s.advance()
          Option.when(after.currentChar.contains('!'))((acc.toString, after.advance()))
        case Some(c) => loop(s.advance(), acc.append(c))
    if state.currentChar.contains('\'') then loop(state.advance(), new StringBuilder) else None

  // ===== GH-353: external-workbook references =====

  /**
   * Split a `[N]Name` external-workbook prefix off a quoted sheet-name string (the quoted form
   * carries the bracket INSIDE the quotes: `'[3]Sheet Name'!B2`). Returns the 1-based workbook
   * index and the external sheet name, or None when the string is not external-shaped.
   */
  private def splitExternalPrefix(name: String): Option[(Int, String)] =
    if name.startsWith("[") then
      val close = name.indexOf(']')
      if close > 1 then
        val digits = name.substring(1, close)
        val rest = name.substring(close + 1)
        if digits.forall(_.isDigit) && rest.nonEmpty then
          digits.toIntOption.filter(_ > 0).map(index => (index, rest))
        else None
      else None
    else None

  /**
   * Parse an unquoted external-workbook reference: `[2]Book1!A1`, `[2]Book1!A1:B2`,
   * `[2]Consolidation.xlsx!D5:D9`.
   *
   * The workbook index is the positive integer Excel assigns to the external link; the unquoted
   * sheet name accepts letters, digits, underscores, and periods (names needing more get the quoted
   * form, handled in parseQuotedSheetRef). The referenced cells live outside this workbook, so the
   * result is a dedicated [[TExpr.ExternalRef]]/[[TExpr.ExternalRange]] node — never resolved here,
   * only carried for closed-workbook cache semantics (GH-353).
   *
   * @param state
   *   Parser state positioned at the opening '['
   */
  private def parseExternalRef(state: ParserState): ParseResult[TExpr[?]] =
    val startPos = state.pos
    val s1 = state.advance() // skip '['

    @tailrec
    def readDigits(s: ParserState): ParserState =
      s.currentChar match
        case Some(c) if c.isDigit => readDigits(s.advance())
        case _ => s

    val s2 = readDigits(s1)
    val digits = state.input.substring(s1.pos, s2.pos)
    (digits.toIntOption.filter(_ > 0), s2.currentChar) match
      case (Some(index), Some(']')) =>
        val s3 = s2.advance()
        @tailrec
        def readName(s: ParserState): ParserState =
          s.currentChar match
            case Some(c) if c.isLetterOrDigit || c == '_' || c == '.' => readName(s.advance())
            case _ => s
        val s4 = readName(s3)
        val name = state.input.substring(s3.pos, s4.pos)
        s4.currentChar match
          case Some('!') if name.nonEmpty =>
            parseExternalQualifiedRef(index, name, s4.advance(), startPos)
          case _ =>
            Left(
              ParseError.InvalidCellRef(
                state.input.substring(startPos, s4.pos),
                startPos,
                "expected external sheet name followed by '!' after workbook index [n]"
              )
            )
      case _ =>
        Left(
          ParseError.InvalidCellRef(
            state.input.substring(startPos, s2.pos),
            startPos,
            "expected external workbook index [n] (a positive integer in brackets)"
          )
        )

  /**
   * Parse the cell/range part of an external-workbook reference: `[index]name!<here>`.
   *
   * Mirrors parseSheetQualifiedRef but produces [[TExpr.ExternalRef]]/[[TExpr.ExternalRange]].
   *
   * @param state
   *   Parser state positioned after the '!'
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  private def parseExternalQualifiedRef(
    index: Int,
    name: String,
    state: ParserState,
    startPos: Int
  ): ParseResult[TExpr[?]] =
    val refStartPos = state.pos

    @tailrec
    def readRef(s: ParserState): ParserState =
      s.currentChar match
        case Some(c) if c.isLetterOrDigit || c == '$' || c == ':' => readRef(s.advance())
        case _ => s

    val s2 = readRef(state)
    val refPart = state.input.substring(refStartPos, s2.pos)
    val prefix = s"[$index]$name"

    if refPart.isEmpty then
      Left(ParseError.InvalidCellRef(s"$prefix!", startPos, "missing cell reference after !"))
    else if refPart.contains(':') then
      // External range reference: [2]Book1!A1:B2
      CellRange.parse(refPart) match
        case Right(range) =>
          Right((TExpr.ExternalRange(index, name, range, RangeForm.of(refPart, range)), s2))
        case Left(err) =>
          Left(ParseError.InvalidCellRef(s"$prefix!$refPart", startPos, err))
    else
      // External single cell reference: [2]Book1!A1
      val (cleanRef, anchor) = Anchor.parse(refPart)
      ARef.parse(cleanRef) match
        case Right(aref) =>
          Right((TExpr.ExternalRef(index, name, aref.asInstanceOf[ARef], anchor), s2))
        case Left(err) =>
          Left(ParseError.InvalidCellRef(s"$prefix!$refPart", startPos, err))

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
        // GH-612: the form is the syntax consumed — A:C is whole columns, A1:C10 corners — so a
        // whole-column reference prints back as written and drags only along columns; a corner
        // spelling over every row (A1:A1048576) is the whole-column form, as Excel canonicalises it
        Right((TExpr.RangeRef(range, RangeForm.of(rangeStr, range)), s3))
      case Left(err) =>
        Left(ParseError.InvalidCellRef(rangeStr, startPos, err))

  /**
   * The Excel error codes, longest first, so a prefix match never stops short (`#N/A` vs `#NAME?`).
   */
  private val errorCodesLongestFirst: List[(String, CellError)] =
    CellError.values.toList.map(e => (e.toExcel, e)).sortBy(-_._1.length)

  /**
   * GH-612: parse an Excel error literal (`#REF!`, `#N/A`, `#DIV/0!`, `#NAME?`, …) by matching the
   * known codes case-insensitively at the cursor (`#ref!` is `#REF!`, as Excel upper-cases it at
   * entry). Matching the codes rather than scanning a character class keeps `#N/A` — the one code
   * with no `!`/`?` terminator — from swallowing what follows it: `#N/A/2` is `#N/A` divided by 2,
   * as Excel reads it. When no code matches, the whole `#`-token (letters, digits, `/`, `_` and a
   * closing `!`/`?`) is reported, so `#GETTING_DATA` fails as itself — never a silent literal.
   */
  private def parseErrorLiteral(state: ParserState): ParseResult[TExpr[?]] =
    val startPos = state.pos
    val matched = errorCodesLongestFirst.collectFirst {
      case (code, error) if state.input.regionMatches(true, startPos, code, 0, code.length) =>
        (code, error)
    }
    matched match
      case Some((code, error)) => Right((TExpr.ErrorLit(error), state.advance(code.length)))
      case None =>
        @tailrec
        def readBody(s: ParserState): ParserState =
          s.currentChar match
            case Some(c) if c.isLetterOrDigit || c == '/' || c == '_' => readBody(s.advance())
            case _ => s
        val afterBody = readBody(state.advance())
        val afterTerminator = afterBody.currentChar match
          case Some('!' | '?') => afterBody.advance()
          case _ => afterBody
        val text = state.input.substring(startPos, afterTerminator.pos)
        Left(ParseError.UnexpectedChar('#', startPos, s"unknown error literal '$text'"))

  /**
   * GH-669: parse an array constant — `{1,2;3,4}`: `,` separates the columns of a row, `;` the
   * rows. As in Excel, every element is a constant (a number, optionally negative; text; TRUE or
   * FALSE; an error value) and every row has the same width; a reference, an expression or an empty
   * element is refused. The result is a literal array (`TExpr.Lit(ArrayResult)`), which the
   * evaluator already treats as an array value and every reference walker as a leaf; the printer
   * spells it back as written.
   *
   * @param state
   *   Parser state positioned at the opening '{'
   */
  private def parseArrayConstant(state: ParserState): ParseResult[TExpr[?]] =
    val startPos = state.pos
    val expected = "an array element: a number, text, TRUE, FALSE or an error value"

    // the scalar literal parsers' result as an element value
    def constant(parsed: ParseResult[TExpr[?]], negate: Boolean) =
      parsed.flatMap {
        case (TExpr.Lit(n: BigDecimal), next) =>
          Right((CellValue.Number(if negate then -n else n), next))
        case (TExpr.Lit(text: String), next) => Right((CellValue.Text(text), next))
        case (TExpr.ErrorLit(error), next) => Right((CellValue.Error(error), next))
        case (_, next) => Left(ParseError.GenericError(expected, Some(next.pos)))
      }

    def element(s0: ParserState): Either[ParseError, (CellValue, ParserState)] =
      val s = skipWhitespace(s0)
      s.currentChar match
        case Some('"') => constant(parseStringLiteral(s), negate = false)
        case Some(c) if c.isDigit || c == '.' => constant(parseNumberLiteral(s), negate = false)
        case Some('-') if s.advance().currentChar.exists(c => c.isDigit || c == '.') =>
          constant(parseNumberLiteral(s.advance()), negate = true)
        case Some('#') => constant(parseErrorLiteral(s), negate = false)
        case Some(c) if c.isLetter =>
          @tailrec
          def word(w: ParserState): ParserState =
            if w.currentChar.exists(ch => ch.isLetterOrDigit || ch == '_' || ch == '.') then
              word(w.advance())
            else w
          val end = word(s)
          s.input.substring(s.pos, end.pos).toUpperCase match
            case "TRUE" => Right((CellValue.Bool(true), end))
            case "FALSE" => Right((CellValue.Bool(false), end))
            case _ => Left(ParseError.UnexpectedChar(c, s.pos, expected))
        case Some(c) => Left(ParseError.UnexpectedChar(c, s.pos, expected))
        case None => Left(ParseError.UnexpectedEOF(s.pos, expected))

    @tailrec
    def rows(
      s: ParserState,
      row: Vector[CellValue],
      done: Vector[Vector[CellValue]]
    ): Either[ParseError, (Vector[Vector[CellValue]], ParserState)] =
      element(s) match
        case Left(err) => Left(err)
        case Right((value, next)) =>
          val after = skipWhitespace(next)
          val current = row :+ value
          after.currentChar match
            case Some(',') => rows(after.advance(), current, done)
            case Some(';') => rows(after.advance(), Vector.empty, done :+ current)
            case Some('}') => Right((done :+ current, after.advance()))
            case Some(c) =>
              Left(ParseError.UnexpectedChar(c, after.pos, "expected ',', ';' or '}'"))
            case None => Left(ParseError.UnexpectedEOF(after.pos, "expected ',', ';' or '}'"))

    descend(state).flatMap { sd =>
      rows(sd.advance(), Vector.empty, Vector.empty).flatMap { case (grid, next) =>
        val width = grid.headOption.map(_.size).getOrElse(0)
        if grid.forall(_.size == width) then
          Right((TExpr.Lit(ArrayResult(grid)), next.copy(depth = state.depth)))
        else
          Left(
            ParseError.generic(
              "every row of an array constant must have the same number of columns",
              startPos
            )
          )
      }
    }

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
