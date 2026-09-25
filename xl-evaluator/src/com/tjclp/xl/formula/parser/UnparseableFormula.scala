package com.tjclp.xl.formula.parser

import scala.annotation.tailrec

/**
 * GH-663: the formula oracle behind the lint's `formula-unparseable` rule — the evaluator's parser,
 * the gate `putf` and batch `putf` apply before writing. Given a cell `<f>`'s stored text (file
 * form, no leading '='), [[check]] is the parser's diagnostic when Excel would show the repair
 * prompt on open, `None` otherwise. Its type is `WorkbookLint.FormulaCheck` (xl-ooxml's oracle
 * slot, which cannot see this parser): the CLI's `lint` and the aggregate module's `Excel.lint`
 * (GH-674) pass this one function, so a script lints exactly what `xl lint` lints.
 *
 * A repair-tier finding must be a text no Excel dialect accepts, and the parser's grammar is
 * narrower than Excel's (LibreOffice writes `TRUE()`, which the parser refuses as an unexpected
 * '('; add-in names such as `BDP(…)` are `#NAME?` on recalculation, not a repair; an argument count
 * the registry's arity model refuses opens intact), so only the classes the parser is certain of
 * are findings: text that ends before the expression does (`SUM(A1:A2`, an unterminated string), a
 * `]` or `}` closing a `(`, and Excel's 8192-character limit. The parser's stack guards are NOT
 * findings: its 128-level nesting budget is twice Excel's 64-level rule and its 1024-operator
 * budget (#680) bounds a flat chain's length, which Excel does not — both bound the parser, not a
 * certain repair. Every other refusal stays `xl audit`'s to list under "Unparseable formulas" —
 * including an extra or wrong closer after a complete expression (`SUM(A1:A2))`, `SUM(A1:A2]`),
 * which surfaces as an unexpected character, and a `,` or space inside parentheses (`SUM((A1,A2))`,
 * `(A1:B2 B1:C2)`: Excel's union and intersection reference operators, which the parser does not
 * implement). The parser reports ANY character but `)` after a parenthesized expression as an
 * `UnbalancedDelimiter`, so that class is a finding only when the character is a `]` or `}`; a `,`,
 * a space or a reference character there is a grammar gap, not a certain repair. Likewise an
 * `UnexpectedEOF` is a finding only when the TEXT shows the truncation — an open `(`, `{` or `[`,
 * an unterminated `"…"` or `'…'`, a trailing operator (`A1+`, `A1:`, `Sheet1!`): the parser also
 * reports it for a complete text its grammar cannot finish, such as the bare word `NOT` (a legal
 * defined name Excel resolves, read here as the prefix operator awaiting its operand; PR #679
 * review) — a grammar gap, not a repair.
 */
object UnparseableFormula:

  /**
   * The oracle: the diagnostic for a text Excel repairs on open, `None` for every other text.
   *
   * Findings-preserving fast path (PR #679 review): each finding class has a cheap necessary
   * precondition — UnexpectedEOF needs [[certainTruncation]], FormulaTooLong needs the length, the
   * `]`/`}` arm needs one of those characters in the text — so a text meeting none is exactly the
   * None the parse would return, and a well-formed book never pays for a parse per `<f>`.
   */
  val check: String => Option[String] = text =>
    if !(certainTruncation(text) || text.length >= ExcelFormulaMaxChars ||
        text.exists(ch => ch == ']' || ch == '}'))
    then None
    else checkSlow(text)

  /**
   * Excel's formula length limit: the parser refuses a text LONGER than this (`FormulaTooLong`);
   * the fast path parses at exactly this length too — the conservative side.
   */
  private val ExcelFormulaMaxChars = 8192

  /** The parse-backed classification [[check]] short-circuits; exposed for the parity pin. */
  private[xl] val checkSlow: String => Option[String] = text =>
    FormulaParser.parse(s"=$text") match
      case Right(_) => None
      case Left(err: ParseError.UnexpectedEOF) if certainTruncation(text) =>
        Some(ParseError.describe(err))
      case Left(err: ParseError.FormulaTooLong) => Some(ParseError.describe(err))
      case Left(err @ ParseError.UnbalancedDelimiter(_, ']' | '}', _)) =>
        Some(ParseError.describe(err))
      case Left(_) => None

  /**
   * Does the formula TEXT itself show that it ends early? True when a `(`, `{` or `[` is still
   * open, a `"…"` string or `'…'` sheet name is unterminated (`""` and `''` are the escapes), or
   * the last non-blank character outside quotes is one Excel never ends a formula with — a binary
   * operator, a range or sheet separator, an argument comma or an open paren. A complete text the
   * parser merely cannot finish (`NOT`, a name spelled like its prefix operator) is not a
   * truncation.
   *
   * A tail-recursive scan over `charAt`: this is the fast path every well-formed `<f>` in a book
   * takes under `lint` and `lint --stream`, so it allocates nothing per character (PR #679 review).
   */
  private[xl] def certainTruncation(text: String): Boolean =
    @tailrec
    def scan(i: Int, quote: Char, justClosed: Boolean, depth: Int, last: Char): Boolean =
      if i >= text.length then quote != NoQuote || depth > 0 || TrailingOperators.contains(last)
      else
        val c = text.charAt(i)
        if quote != NoQuote then
          if c == quote then scan(i + 1, NoQuote, true, depth, last)
          else scan(i + 1, quote, false, depth, last)
        else if justClosed && c == last && (c == '"' || c == '\'') then
          // the closing quote was the first half of an escaped pair: still inside the literal
          scan(i + 1, c, false, depth, last)
        else
          c match
            case '"' | '\'' => scan(i + 1, c, false, depth, c)
            case '(' | '{' | '[' => scan(i + 1, NoQuote, false, depth + 1, c)
            case ')' | '}' | ']' => scan(i + 1, NoQuote, false, depth - 1, c)
            case ' ' | '\t' | '\n' | '\r' => scan(i + 1, NoQuote, false, depth, last)
            case other => scan(i + 1, NoQuote, false, depth, other)
    scan(0, NoQuote, false, 0, ' ')

  /** "No quote open" in [[certainTruncation]]'s scan; NUL never occurs in a formula. */
  private val NoQuote: Char = '\u0000'

  /** A formula cannot end on one of these outside quotes; `(` is covered by the depth count too. */
  private val TrailingOperators: Set[Char] =
    Set('+', '-', '*', '/', '^', '&', '=', '<', '>', ',', ':', '!', '(')
