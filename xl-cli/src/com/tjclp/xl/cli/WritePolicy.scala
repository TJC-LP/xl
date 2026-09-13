package com.tjclp.xl.cli

import scala.util.control.NoStackTrace

/**
 * Cross-cutting posture for the write verbs, carried by the global `--no-recalc` /
 * `--preserve-caches` and `--strict` flags.
 *
 * @param noRecalc
 *   GH-468: apply the edit and recalculate nothing — the escape hatch for books whose caches come
 *   from an engine other than xl (an external calculator, a LibreOffice arbiter, a replica
 *   computation). How the file says so is per verb class. A NON-STRUCTURAL write (put, putf, batch,
 *   style, ...) leaves every cached formula value in the file verbatim: it moves no cell and
 *   rewrites no formula, so no untouched formula can have changed answer, and the file needs no
 *   marker. A STRUCTURAL write (insert / delete rows or columns) shifts cells, rewrites formula
 *   text and rewrites defined names, and no local test can tell which pre-edit cache still holds —
 *   so (GH-509) it carries every pre-edit cache forward with its formula, rewritten or relocated or
 *   not, and writes `<calcPr fullCalcOnLoad="1"/>`: Excel and LibreOffice recompute the book on
 *   open and never show a stale `<v>`, while a cache-only reader sees the numbers the source file
 *   had. The one exception is the GH-507 blind closure (readers no static graph can resolve, and
 *   their dependents), withdrawn on both paths and counted in the summary. Without the flag a write
 *   still only refreshes its dirty dependency cone, never the whole book.
 * @param strict
 *   GH-496: promote a write's advisory conditions — formula-evaluation errors, iterative
 *   non-convergence, data-table seed warnings — from "printed in the summary, exit 0" to exit 1.
 *   The default stays advisory so interactive use is unchanged.
 */
final case class WritePolicy(noRecalc: Boolean = false, strict: Boolean = false) derives CanEqual

object WritePolicy:
  /** Recalculate the dirty cone, report problems advisorily — the historical CLI behavior. */
  val default: WritePolicy = WritePolicy()

/**
 * GH-496: a write that COMPLETED (the output file is written) but whose summary carries a condition
 * `--strict` promotes to a non-zero exit.
 *
 * Distinct from a plain failure so the runner can print the full summary verbatim — the counts, the
 * failing refs, the convergence verdict — instead of an `Error:`-prefixed one-liner, while still
 * exiting non-zero. Stack-trace free: nothing here is a defect to debug.
 */
final class StrictFailure(val summary: String) extends Exception(summary) with NoStackTrace
