package com.tjclp.xl.io

import java.nio.file.{Path, Paths}

import com.tjclp.xl.error.XLResult
import com.tjclp.xl.formula.parser.UnparseableFormula
import com.tjclp.xl.ooxml.lint.{Finding, WorkbookLint}

/**
 * The structural lint `xl lint` runs, for the sync `Excel` facade (GH-674): every repair and
 * hygiene rule of [[com.tjclp.xl.ooxml.lint.WorkbookLint]] INCLUDING `formula-unparseable`, which
 * the one-argument `WorkbookLint.lint(path)` leaves off because xl-ooxml cannot see the formula
 * parser. The oracle is the CLI's own ([[com.tjclp.xl.formula.parser.UnparseableFormula]]), so a
 * script and `xl lint` report the same findings for the same file.
 *
 * {{{
 * import com.tjclp.xl.scripting.{*, given}
 *
 * val findings = Excel.lint("deliverable.xlsx").unsafe
 * findings.filter(_.severity == LintSeverity.Repair).foreach(f => println(s"$${f.part}: $${f.message}"))
 * }}}
 *
 * Lives in the aggregate `xl` module — the only library module that sees both the parser and the
 * lint. Explicit overloads instead of default parameters: defaulted extension methods do not
 * survive the prelude's wildcard export (see the WorkbookEvaluator note).
 */
object ExcelLint:

  extension (excel: Excel.type)

    /** Lint the XLSX file at `path` (the `xl lint` rule set; findings in part order). */
    def lint(path: String): XLResult[Vector[Finding]] = excel.lint(Paths.get(path))

    /** [[lint]] for a `java.nio.file.Path`. */
    def lint(path: Path): XLResult[Vector[Finding]] =
      WorkbookLint.lint(path, UnparseableFormula.check)

    /**
     * `xl lint --stream`: sheet and table parts are SAX-scanned, so memory stays O(1) in the row
     * count; the findings are [[lint]]'s (pinned by the parity suite).
     */
    def lintStream(path: String): XLResult[Vector[Finding]] = excel.lintStream(Paths.get(path))

    /** [[lintStream]] for a `java.nio.file.Path`. */
    def lintStream(path: Path): XLResult[Vector[Finding]] =
      WorkbookLint.lintStream(path, UnparseableFormula.check)
