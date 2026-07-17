package com.tjclp.xl.cli.commands

import com.tjclp.xl.ooxml.lint.Finding

/**
 * Output rendering for the lint command (GH-397).
 *
 * Pure: the renderers never throw and perform no IO. The caller (Main.runLint) runs the structural
 * lint on the raw zip and maps the findings to the exit-code convention (0 clean, 1 findings, 2
 * error).
 */
object LintCommands:

  /** Human-readable findings list (default format). */
  def renderText(file: String, findings: Vector[Finding]): String =
    if findings.isEmpty then s"$file: clean (no findings)"
    else
      val lines = findings.map { f =>
        s"  [${f.category.slug}] ${f.part}: ${f.message} — ${f.locator}"
      }
      (s"$file: ${findings.size} finding(s)" +: lines).mkString("\n")

  /**
   * Stable machine-readable schema for pipelines:
   * {{{
   * {
   *   "file": "report.xlsx",
   *   "clean": false,
   *   "findings": [{"part": "xl/workbook.xml", "category": "child-order",
   *                 "locator": "<externalReferences> (element #9)", "message": "..."}]
   * }
   * }}}
   */
  def renderJson(file: String, findings: Vector[Finding]): String =
    val arr = ujson.Arr.from(findings.map { f =>
      ujson.Obj(
        "part" -> ujson.Str(f.part),
        "category" -> ujson.Str(f.category.slug),
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
