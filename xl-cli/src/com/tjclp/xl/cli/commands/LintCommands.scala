package com.tjclp.xl.cli.commands

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
   * the expression does (`SUM(A1:A2`, an unterminated string), a `]` or `}` closing a `(`, Excel's
   * 8192-character limit, and the parser's 128-level nesting limit (Excel's is 64). Every other
   * refusal stays `xl audit`'s to list under "Unparseable formulas" — including an extra or wrong
   * closer after a complete expression (`SUM(A1:A2))`, `SUM(A1:A2]`), which surfaces as an
   * unexpected character, and a `,` or space inside parentheses (`SUM((A1,A2))`, `(A1:B2 B1:C2)`:
   * Excel's union and intersection reference operators, which the parser does not implement). The
   * parser reports ANY character but `)` after a parenthesized expression as an
   * `UnbalancedDelimiter`, so that class is a finding only when the character is a `]` or `}`; a
   * `,`, a space or a reference character there is a grammar gap, not a certain repair.
   */
  val formulaCheck: WorkbookLint.FormulaCheck = text =>
    FormulaParser.parse(s"=$text") match
      case Right(_) => None
      case Left(
            err @ (_: ParseError.UnexpectedEOF | _: ParseError.FormulaTooLong |
            _: ParseError.NestingTooDeep)
          ) =>
        Some(ParseError.describe(err))
      case Left(err @ ParseError.UnbalancedDelimiter(_, ']' | '}', _)) =>
        Some(ParseError.describe(err))
      case Left(_) => None

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
