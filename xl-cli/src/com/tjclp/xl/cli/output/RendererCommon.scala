package com.tjclp.xl.cli.output

import com.tjclp.xl.cells.FormulaKind
import com.tjclp.xl.cli.read.RecordWindow

/**
 * Shared utilities for the table renderers (Markdown, JSON, CSV): the notices they emit and the
 * formula-bar spelling. Cell text itself is the record's ([[com.tjclp.xl.cli.read.CellRecord]]).
 */
object RendererCommon:

  /**
   * Format evaluation errors as Excel-style error codes.
   *
   * Maps error messages to standard Excel error codes for consistent display across all output
   * formats.
   */
  def formatEvalError(message: String): String =
    if message.toLowerCase.contains("circular") then "#CIRC!"
    else if message.toLowerCase.contains("division") || message.toLowerCase.contains("div") then
      "#DIV/0!"
    else if message.toLowerCase.contains("parse") || message.toLowerCase.contains("unknown") then
      "#NAME?"
    else if message.toLowerCase.contains("ref") then "#REF!"
    else "#ERROR!"

  /**
   * Formula display text: leading `=` always; `{=...}` braces when the cell carries a non-Normal
   * CT_CellFormula record (array/dataTable) — Excel's own formula-bar convention (GH-430).
   */
  def formulaDisplay(expr: String, kind: FormulaKind): String =
    val withEquals = if expr.startsWith("=") then expr else s"=$expr"
    kind match
      case _: FormulaKind.Normal => withEquals
      case _ => s"{$withEquals}"

  /**
   * Human-readable truncation notice emitted when --limit clips output (GH-351).
   *
   * Rendered as a trailer line after markdown tables, or on stderr for machine-parseable formats
   * (CSV) so stdout stays clean.
   */
  def truncationNotice(shown: Int, total: Int, noun: String = "rows"): String =
    s"… showing $shown of $total $noun (use --limit to raise; --limit 0 = no limit)"

  /**
   * The truncation notice of a paged window (`--offset`): which rows of the range are shown.
   * `first`/`last` are 1-based positions within the range.
   */
  def pagingNotice(first: Int, last: Int, total: Int): String =
    s"… showing rows $first–$last of $total (use --offset/--limit to page; --limit 0 = no limit)"

  /**
   * The notice of a `search` whose scan stopped at `--limit` (GH-637): one more match was seen, so
   * more exist, and the exact total is a `--total` (or `--limit 0`) away.
   */
  def searchStoppedNotice(shown: Int): String =
    s"… showing first $shown matches; more exist (use --limit to raise; --limit 0 = no limit; " +
      "--total for the exact count)"

  /** The notice for a window whose columns `--max-cols` clipped. */
  def columnNotice(shown: Int, total: Int): String =
    s"… showing $shown of $total columns (use --max-cols to raise; --max-cols 0 = no limit)"

  /** Cap long index listings so the notice stays one readable line. */
  private def summarizeIndices(label: String, items: Vector[String]): String =
    val shown = items.take(10).mkString(", ")
    if items.size > 10 then s"$label $shown, … (${items.size} total)" else s"$label $shown"

  /**
   * GH-474: the hidden-line marker for a rendered window, or None when it holds no hidden rows or
   * columns.
   *
   * A viewer that drops data from a range the caller named reads as corruption (`search` finds the
   * cell, `cell` reads it, `view` showed nothing). Hidden lines are therefore rendered by default
   * and MARKED; `--skip-hidden` opts back into elision and the marker then names what was dropped.
   *
   * Rendered as a trailer line after markdown tables, on stderr for machine-parseable formats
   * (CSV), and as structured fields in JSON.
   */
  def hiddenNotice(window: RecordWindow, skipHidden: Boolean): Option[String] =
    val rows = window.hiddenRowNumbers.map(_.toString)
    val cols = window.hiddenColLetters
    if rows.isEmpty && cols.isEmpty then None
    else
      val parts = List(
        Option.when(rows.nonEmpty)(summarizeIndices("row(s)", rows)),
        Option.when(cols.nonEmpty)(summarizeIndices("column(s)", cols))
      ).flatten.mkString(" and ")
      Some(
        if skipHidden then
          s"note: omitted hidden $parts from the requested range (drop --skip-hidden to include them)"
        else
          s"note: range includes hidden $parts — shown because the range was explicitly requested (use --skip-hidden to omit)"
      )

  /**
   * GH-474: `--skip-hidden` cannot be honored under `--stream`.
   *
   * The streaming reader yields cell values only — it never parses row/column properties — so it
   * has no way to know which lines are hidden and renders all of them. Accepting the flag and doing
   * nothing is the silent no-op this issue is about, so the streaming view says so on stderr
   * instead.
   */
  val streamingSkipHiddenNotice: String =
    "note: --skip-hidden is ignored with --stream — the streaming reader does not read row/column " +
      "properties, so hidden rows/columns were rendered. Re-run without --stream to omit them."
