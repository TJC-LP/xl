package com.tjclp.xl.cli.commands

import com.tjclp.xl.ooxml.lint.{Finding, LintSeverity}

/**
 * Output rendering for the lint command (GH-397).
 *
 * Pure: the renderers never throw and perform no IO. The caller (Main.runLint) runs the structural
 * lint on the raw zip and maps the findings to the exit-code convention (0 clean or hygiene-only, 1
 * repair findings — or any finding under `--strict` — 2 usage, 3 error).
 */
object LintCommands:

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
