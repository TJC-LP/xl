package com.tjclp.xl.cli.commands

import scala.annotation.tailrec

import com.tjclp.xl.formula.parser.{FormulaParser, ParseError}
import com.tjclp.xl.ooxml.lint.{Finding, LintSeverity, WorkbookLint}

/**
 * Output rendering for the lint command (GH-397).
 *
 * Pure: the renderers never throw and perform no IO. The caller (Main.runLint) runs the structural
 * lint on the raw zip and maps the findings to the exit-code convention (0 clean or hygiene-only, 1
 * repair findings — or any finding under `--strict` — 2 usage, 3 error).
 */
object LintCommands:

  /**
   * GH-663: the formula oracle behind `formula-unparseable` — the evaluator's parser, the gate
   * `putf` and batch `putf` apply before writing. A repair-tier finding must be a text no Excel
   * dialect accepts, and the parser's grammar is narrower than Excel's (LibreOffice writes
   * `TRUE()`, which the parser refuses as an unexpected '('; add-in names such as `BDP(…)` are
   * `#NAME?` on recalculation, not a repair; an argument count the registry's arity model refuses
   * opens intact), so only the classes the parser is certain of are findings: text that ends before
   * the expression does (`SUM(A1:A2`, an unterminated string), a `]` or `}` closing a `(`, and
   * Excel's 8192-character limit. The parser's 128-level depth budget is NOT a finding: it counts
   * every chained operator segment as a level (GH-56), so a flat 130-term `B2+B3+…` chain that
   * Excel opens intact fails it while a 100-deep `SUM(SUM(…))` nest passes — it bounds the parser,
   * not Excel's 64-level nesting rule (#680). Every other refusal stays `xl audit`'s to list under
   * "Unparseable formulas" — including an extra or wrong closer after a complete expression
   * (`SUM(A1:A2))`, `SUM(A1:A2]`), which surfaces as an unexpected character, and a `,` or space
   * inside parentheses (`SUM((A1,A2))`, `(A1:B2 B1:C2)`: Excel's union and intersection reference
   * operators, which the parser does not implement). The parser reports ANY character but `)` after
   * a parenthesized expression as an `UnbalancedDelimiter`, so that class is a finding only when
   * the character is a `]` or `}`; a `,`, a space or a reference character there is a grammar gap,
   * not a certain repair. Likewise an `UnexpectedEOF` is a finding only when the TEXT shows the
   * truncation — an open `(`, `{` or `[`, an unterminated `"…"` or `'…'`, a trailing operator
   * (`A1+`, `A1:`, `Sheet1!`): the parser also reports it for a complete text its grammar cannot
   * finish, such as the bare word `NOT` (a legal defined name Excel resolves, read here as the
   * prefix operator awaiting its operand; PR #679 review) — a grammar gap, not a repair.
   */
  val formulaCheck: WorkbookLint.FormulaCheck = text =>
    // Findings-preserving fast path (PR #679 review): each finding class has a cheap necessary
    // precondition — UnexpectedEOF needs certainTruncation, FormulaTooLong needs the length, the
    // `]`/`}` arm needs one of those characters in the text — so a text meeting none is exactly
    // the None the parse would return, and a well-formed book never pays for a parse per <f>.
    if !(certainTruncation(text) || text.length >= ExcelFormulaMaxChars ||
        text.exists(ch => ch == ']' || ch == '}'))
    then None
    else formulaCheckSlow(text)

  /**
   * Excel's formula length limit: the parser refuses a text LONGER than this (`FormulaTooLong`);
   * the fast path parses at exactly this length too — the conservative side.
   */
  private val ExcelFormulaMaxChars = 8192

  /**
   * The parse-backed classification [[formulaCheck]] short-circuits; exposed for the parity pin.
   */
  private[cli] val formulaCheckSlow: WorkbookLint.FormulaCheck = text =>
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
  private[cli] def certainTruncation(text: String): Boolean =
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

  /** The findings that fail the gate: every repair, plus the hygiene tier under `--strict`. */
  def gating(findings: Vector[Finding], strict: Boolean): Vector[Finding] =
    if strict then findings else findings.filter(_.severity == LintSeverity.Repair)

  private def countOf(findings: Vector[Finding], severity: LintSeverity): Int =
    findings.count(_.severity == severity)

  /**
   * Human-readable findings list (default format). Hygiene findings are tagged on their line and
   * counted in the header; when they are the only findings and the gate is not strict, a trailer
   * says why the exit is 0.
   */
  def renderText(file: String, findings: Vector[Finding], strict: Boolean = false): String =
    if findings.isEmpty then s"$file: clean (no findings)"
    else
      val repair = countOf(findings, LintSeverity.Repair)
      val hygiene = countOf(findings, LintSeverity.Hygiene)
      val header =
        if hygiene == 0 then s"$file: ${findings.size} finding(s)"
        else s"$file: ${findings.size} finding(s) ($repair repair, $hygiene hygiene)"
      val lines = findings.map { f =>
        val tier = if f.severity == LintSeverity.Hygiene then " (hygiene)" else ""
        s"  [${f.category.slug}]$tier ${f.part}: ${f.message} — ${f.locator}"
      }
      val trailer =
        if repair == 0 && !strict then
          Vector("  hygiene findings only — exit 0 (pass --strict to exit 1 on them)")
        else Vector.empty
      ((header +: lines) ++ trailer).mkString("\n")

  /**
   * Stable machine-readable schema for pipelines:
   * {{{
   * {
   *   "file": "report.xlsx",
   *   "clean": false,
   *   "findings": [{"part": "xl/workbook.xml", "category": "child-order", "severity": "repair",
   *                 "locator": "<externalReferences> (element #9)", "message": "..."}]
   * }
   * }}}
   * `clean` is "no findings at all"; the exit code is the gate (`severity: "repair"` fails it,
   * `"hygiene"` only under `--strict`).
   */
  def renderJson(file: String, findings: Vector[Finding]): String =
    val arr = ujson.Arr.from(findings.map { f =>
      ujson.Obj(
        "part" -> ujson.Str(f.part),
        "category" -> ujson.Str(f.category.slug),
        "severity" -> ujson.Str(f.severity.slug),
        "locator" -> ujson.Str(f.locator),
        "message" -> ujson.Str(f.message)
      )
    })
    val root = ujson.Obj(
      "file" -> ujson.Str(file),
      "clean" -> ujson.Bool(findings.isEmpty),
      "findings" -> arr
    )
    ujson.write(root, indent = 2)
