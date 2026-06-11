package com.tjclp.xl.cli.helpers

import java.util.regex.Pattern

import com.tjclp.xl.cells.CellValue

/**
 * Predicate grammar and evaluator for the filter command (GH-134, phase 1).
 *
 * Grammar (keywords case-insensitive):
 * {{{
 * expr      := orExpr
 * orExpr    := andExpr (OR andExpr)*
 * andExpr   := unary (AND unary)*
 * unary     := NOT unary | '(' expr ')' | predicate
 * predicate := col (= | != | <> | > | >= | < | <=) literal
 *            | col LIKE 'pat%'                  (% wildcard only)
 *            | col BETWEEN literal AND literal  (inclusive, same-type bounds)
 *            | col IN '(' literal (',' literal)* ')'
 *            | col IS [NOT] EMPTY
 * literal   := number | 'string' | "string" | TRUE | FALSE
 * }}}
 *
 * Deliberately NOT the full query algebra from docs/design/query-api.md — no SELECT/GROUP BY/ORDER
 * BY (out of scope for phase 1).
 *
 * Total: `parse` returns Left on malformed input; `evaluate` never throws — a type mismatch (e.g.
 * `B > 100` against a Text cell) means the row doesn't match. String comparisons and LIKE are
 * case-insensitive (Excel semantics).
 */
object FilterPredicate:

  enum Literal derives CanEqual:
    case Num(value: BigDecimal)
    case Str(value: String)
    case Bool(value: Boolean)

  enum CmpOp derives CanEqual:
    case Eq, Ne, Gt, Ge, Lt, Le

  enum Pred derives CanEqual:
    case Cmp(col: String, op: CmpOp, lit: Literal)
    case Like(col: String, pattern: String)
    case Between(col: String, lo: Literal, hi: Literal)
    case In(col: String, values: List[Literal])
    case IsEmpty(col: String, negated: Boolean)
    case And(left: Pred, right: Pred)
    case Or(left: Pred, right: Pred)
    case Not(pred: Pred)

  /** All column identifiers referenced by the predicate (as written, before resolution). */
  def columnRefs(pred: Pred): Set[String] =
    pred match
      case Pred.Cmp(col, _, _) => Set(col)
      case Pred.Like(col, _) => Set(col)
      case Pred.Between(col, _, _) => Set(col)
      case Pred.In(col, _) => Set(col)
      case Pred.IsEmpty(col, _) => Set(col)
      case Pred.And(l, r) => columnRefs(l) ++ columnRefs(r)
      case Pred.Or(l, r) => columnRefs(l) ++ columnRefs(r)
      case Pred.Not(p) => columnRefs(p)

  // ==========================================================================
  // Tokenizer
  // ==========================================================================

  private enum Token derives CanEqual:
    case Ident(name: String)
    case StrLit(value: String)
    case NumLit(value: BigDecimal)
    case Op(sym: String)
    case LParen, RParen, Comma

  private def describe(token: Token): String = token match
    case Token.Ident(name) => s"'$name'"
    case Token.StrLit(value) => s"string '$value'"
    case Token.NumLit(value) => s"number $value"
    case Token.Op(sym) => s"operator '$sym'"
    case Token.LParen => "'('"
    case Token.RParen => "')'"
    case Token.Comma => "','"

  private def tokenize(input: String): Either[String, Vector[Token]] =
    @annotation.tailrec
    def loop(i: Int, acc: Vector[Token]): Either[String, Vector[Token]] =
      if i >= input.length then Right(acc)
      else
        val c = input.charAt(i)
        if c.isWhitespace then loop(i + 1, acc)
        else if c == '(' then loop(i + 1, acc :+ Token.LParen)
        else if c == ')' then loop(i + 1, acc :+ Token.RParen)
        else if c == ',' then loop(i + 1, acc :+ Token.Comma)
        else if c == '\'' || c == '"' then
          readString(i + 1, c) match
            case Right((value, next)) => loop(next, acc :+ Token.StrLit(value))
            case Left(err) => Left(err)
        else if c == '>' || c == '<' then
          if i + 1 < input.length && input.charAt(i + 1) == '=' then
            loop(i + 2, acc :+ Token.Op(s"$c="))
          else if c == '<' && i + 1 < input.length && input.charAt(i + 1) == '>' then
            loop(i + 2, acc :+ Token.Op("!="))
          else loop(i + 1, acc :+ Token.Op(c.toString))
        else if c == '=' then loop(i + 1, acc :+ Token.Op("="))
        else if c == '!' && i + 1 < input.length && input.charAt(i + 1) == '=' then
          loop(i + 2, acc :+ Token.Op("!="))
        else if c.isDigit || (c == '-' && i + 1 < input.length && input.charAt(i + 1).isDigit) then
          val end = numberEnd(i + 1)
          val text = input.substring(i, end)
          scala.util.Try(BigDecimal(text)).toOption match
            case Some(n) => loop(end, acc :+ Token.NumLit(n))
            case scala.None => Left(s"Invalid number '$text' at position $i")
        else if c.isLetter || c == '_' then
          val end = identEnd(i + 1)
          loop(end, acc :+ Token.Ident(input.substring(i, end)))
        else Left(s"Unexpected character '$c' at position $i")

    def readString(start: Int, quote: Char): Either[String, (String, Int)] =
      @annotation.tailrec
      def go(i: Int, sb: String): Either[String, (String, Int)] =
        if i >= input.length then Left(s"Unterminated string (started at position ${start - 1})")
        else if input.charAt(i) == quote then
          // Doubled quote = escaped literal quote
          if i + 1 < input.length && input.charAt(i + 1) == quote then go(i + 2, sb + quote)
          else Right((sb, i + 1))
        else go(i + 1, sb + input.charAt(i))
      go(start, "")

    @annotation.tailrec
    def numberEnd(i: Int): Int =
      if i < input.length && (input.charAt(i).isDigit || input.charAt(i) == '.') then
        numberEnd(i + 1)
      else i

    @annotation.tailrec
    def identEnd(i: Int): Int =
      if i < input.length && (input.charAt(i).isLetterOrDigit || input.charAt(i) == '_') then
        identEnd(i + 1)
      else i

    loop(0, Vector.empty)

  // ==========================================================================
  // Recursive-descent parser (index-passing, total)
  // ==========================================================================

  /** Parse a predicate expression. Left with a description on malformed input. */
  def parse(input: String): Either[String, Pred] =
    tokenize(input).flatMap { tokens =>
      if tokens.isEmpty then Left("Empty predicate")
      else
        parseOr(tokens, 0).flatMap { (pred, next) =>
          if next == tokens.length then Right(pred)
          else Left(s"Unexpected ${describe(tokens(next))} after end of predicate")
        }
    }

  private type ParseResult = Either[String, (Pred, Int)]

  private def isKeyword(tokens: Vector[Token], i: Int, kw: String): Boolean =
    tokens.lift(i) match
      case Some(Token.Ident(name)) => name.equalsIgnoreCase(kw)
      case _ => false

  private def parseOr(tokens: Vector[Token], start: Int): ParseResult =
    parseAnd(tokens, start).flatMap { (left, i) =>
      if isKeyword(tokens, i, "OR") then
        parseOr(tokens, i + 1).map((right, next) => (Pred.Or(left, right), next))
      else Right((left, i))
    }

  private def parseAnd(tokens: Vector[Token], start: Int): ParseResult =
    parseUnary(tokens, start).flatMap { (left, i) =>
      if isKeyword(tokens, i, "AND") then
        parseAnd(tokens, i + 1).map((right, next) => (Pred.And(left, right), next))
      else Right((left, i))
    }

  private def parseUnary(tokens: Vector[Token], start: Int): ParseResult =
    if isKeyword(tokens, start, "NOT") then
      parseUnary(tokens, start + 1).map((p, next) => (Pred.Not(p), next))
    else if tokens.lift(start).contains(Token.LParen) then
      parseOr(tokens, start + 1).flatMap { (p, i) =>
        if tokens.lift(i).contains(Token.RParen) then Right((p, i + 1))
        else Left("Expected ')' to close '('")
      }
    else parsePredicate(tokens, start)

  private def parsePredicate(tokens: Vector[Token], start: Int): ParseResult =
    tokens.lift(start) match
      case Some(Token.Ident(col)) if !isReserved(col) =>
        parsePredicateTail(tokens, start + 1, col)
      case Some(token) => Left(s"Expected a column reference, found ${describe(token)}")
      case scala.None => Left("Expected a column reference, found end of input")

  private val reservedWords =
    Set("AND", "OR", "NOT", "LIKE", "BETWEEN", "IN", "IS", "EMPTY", "TRUE", "FALSE")

  private def isReserved(name: String): Boolean = reservedWords.contains(name.toUpperCase)

  private def parsePredicateTail(tokens: Vector[Token], i: Int, col: String): ParseResult =
    tokens.lift(i) match
      case Some(Token.Op(sym)) =>
        val op = sym match
          case "=" => CmpOp.Eq
          case "!=" => CmpOp.Ne
          case ">" => CmpOp.Gt
          case ">=" => CmpOp.Ge
          case "<" => CmpOp.Lt
          case _ => CmpOp.Le
        parseLiteral(tokens, i + 1).map((lit, next) => (Pred.Cmp(col, op, lit), next))

      case Some(Token.Ident(kw)) if kw.equalsIgnoreCase("LIKE") =>
        tokens.lift(i + 1) match
          case Some(Token.StrLit(pattern)) => Right((Pred.Like(col, pattern), i + 2))
          case _ => Left(s"LIKE requires a quoted string pattern (e.g. $col LIKE 'Widget%')")

      case Some(Token.Ident(kw)) if kw.equalsIgnoreCase("BETWEEN") =>
        for
          (lo, afterLo) <- parseLiteral(tokens, i + 1)
          _ <- Either.cond(
            isKeyword(tokens, afterLo, "AND"),
            (),
            "BETWEEN requires AND between bounds (e.g. B BETWEEN 10 AND 100)"
          )
          (hi, next) <- parseLiteral(tokens, afterLo + 1)
          _ <- Either.cond(
            sameKind(lo, hi),
            (),
            "BETWEEN bounds must have the same type (both numbers or both strings)"
          )
        yield (Pred.Between(col, lo, hi), next)

      case Some(Token.Ident(kw)) if kw.equalsIgnoreCase("IN") =>
        if !tokens.lift(i + 1).contains(Token.LParen) then
          Left(s"IN requires a parenthesized list (e.g. $col IN ('a', 'b'))")
        else
          parseLiteralList(tokens, i + 2, Vector.empty).map { (values, next) =>
            (Pred.In(col, values.toList), next)
          }

      case Some(Token.Ident(kw)) if kw.equalsIgnoreCase("IS") =>
        if isKeyword(tokens, i + 1, "EMPTY") then Right((Pred.IsEmpty(col, negated = false), i + 2))
        else if isKeyword(tokens, i + 1, "NOT") && isKeyword(tokens, i + 2, "EMPTY") then
          Right((Pred.IsEmpty(col, negated = true), i + 3))
        else Left("IS must be followed by EMPTY or NOT EMPTY")

      case Some(token) =>
        Left(s"Expected an operator after '$col', found ${describe(token)}")
      case scala.None =>
        Left(s"Expected an operator after '$col', found end of input")

  private def parseLiteralList(
    tokens: Vector[Token],
    i: Int,
    acc: Vector[Literal]
  ): Either[String, (Vector[Literal], Int)] =
    parseLiteral(tokens, i).flatMap { (lit, next) =>
      tokens.lift(next) match
        case Some(Token.Comma) => parseLiteralList(tokens, next + 1, acc :+ lit)
        case Some(Token.RParen) => Right((acc :+ lit, next + 1))
        case _ => Left("Expected ',' or ')' in IN list")
    }

  private def parseLiteral(tokens: Vector[Token], i: Int): Either[String, (Literal, Int)] =
    tokens.lift(i) match
      case Some(Token.NumLit(n)) => Right((Literal.Num(n), i + 1))
      case Some(Token.StrLit(s)) => Right((Literal.Str(s), i + 1))
      case Some(Token.Ident(w)) if w.equalsIgnoreCase("TRUE") => Right((Literal.Bool(true), i + 1))
      case Some(Token.Ident(w)) if w.equalsIgnoreCase("FALSE") =>
        Right((Literal.Bool(false), i + 1))
      case Some(token) =>
        Left(s"Expected a literal (number, quoted string, TRUE/FALSE), found ${describe(token)}")
      case scala.None =>
        Left("Expected a literal (number, quoted string, TRUE/FALSE), found end of input")

  private def sameKind(a: Literal, b: Literal): Boolean =
    (a, b) match
      case (Literal.Num(_), Literal.Num(_)) => true
      case (Literal.Str(_), Literal.Str(_)) => true
      case (Literal.Bool(_), Literal.Bool(_)) => true
      case _ => false

  // ==========================================================================
  // Evaluation
  // ==========================================================================

  /**
   * Evaluate against one row. `resolve` maps an identifier to a 0-based column index; `cellAt`
   * fetches the row's cell value at that index (None = no cell). Unresolvable identifiers fail the
   * match (callers validate them upfront for a proper error).
   */
  def evaluate(
    pred: Pred,
    resolve: String => Option[Int],
    cellAt: Int => Option[CellValue]
  ): Boolean =
    def valueOf(col: String): Option[CellValue] =
      resolve(col).flatMap(cellAt).map {
        case CellValue.Formula(_, Some(cached)) => cached
        case other => other
      }

    pred match
      case Pred.And(l, r) => evaluate(l, resolve, cellAt) && evaluate(r, resolve, cellAt)
      case Pred.Or(l, r) => evaluate(l, resolve, cellAt) || evaluate(r, resolve, cellAt)
      case Pred.Not(p) => !evaluate(p, resolve, cellAt)
      case Pred.Cmp(col, op, lit) => valueOf(col).exists(compare(_, op, lit))
      case Pred.Like(col, pattern) =>
        valueOf(col).flatMap(textOf).exists(likeMatches(pattern, _))
      case Pred.Between(col, lo, hi) =>
        valueOf(col).exists(v => compare(v, CmpOp.Ge, lo) && compare(v, CmpOp.Le, hi))
      case Pred.In(col, values) =>
        valueOf(col).exists(v => values.exists(compare(v, CmpOp.Eq, _)))
      case Pred.IsEmpty(col, negated) =>
        val empty = resolve(col) match
          case scala.None => true
          case Some(idx) =>
            cellAt(idx) match
              case scala.None => true
              case Some(CellValue.Empty) => true
              case Some(CellValue.Text(s)) => s.trim.isEmpty
              case Some(_) => false
        empty != negated

  /** Typed comparison: Number vs Num, Text vs Str (case-insensitive), Bool vs Bool. */
  private def compare(value: CellValue, op: CmpOp, lit: Literal): Boolean =
    (value, lit) match
      case (CellValue.Number(n), Literal.Num(target)) => applyOp(op, n.compare(target))
      case (v, Literal.Str(target)) =>
        textOf(v).exists(s => applyOp(op, s.compareToIgnoreCase(target)))
      case (CellValue.Bool(b), Literal.Bool(target)) => applyOp(op, b.compareTo(target))
      case _ => false // Type mismatch = no match, never an error

  private def applyOp(op: CmpOp, cmp: Int): Boolean = op match
    case CmpOp.Eq => cmp == 0
    case CmpOp.Ne => cmp != 0
    case CmpOp.Gt => cmp > 0
    case CmpOp.Ge => cmp >= 0
    case CmpOp.Lt => cmp < 0
    case CmpOp.Le => cmp <= 0

  private def textOf(value: CellValue): Option[String] = value match
    case CellValue.Text(s) => Some(s)
    case CellValue.RichText(rt) => Some(rt.toPlainText)
    case _ => scala.None

  /** `%` matches any run of characters; everything else is literal (case-insensitive). */
  private def likeMatches(pattern: String, text: String): Boolean =
    val regex = pattern.split("%", -1).map(Pattern.quote).mkString(".*")
    Pattern.compile(s"(?is)$regex").matcher(text).matches()
