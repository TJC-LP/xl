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
 *   text and rewrites defined names, so (GH-503, GH-509) it keeps only the caches the edit provably
 *   left unchanged — text, record kind and address identical, and outside the edit's dirty cone
 *   (every reader of a moved or removed cell and everything downstream of it across sheets, every
 *   dynamic reference, every reader the static graph cannot resolve, GH-507) — writes every other
 *   formula WITHOUT a `<v>`, counts both, and writes `<calcPr fullCalcOnLoad="1"/>`. Excel
 *   recomputes the whole book on open. LibreOffice does NOT honor the marker at its shipped default
 *   ("never recalculate on load" for xlsx): it displays whatever `<v>` is present and computes only
 *   the cells without one, which is why the dirty cone is withdrawn rather than carried — a blank
 *   it fills in, a stale number it would show. A cache-only reader (openpyxl `data_only`, pandas,
 *   `xl view` without `--eval`) sees a blank there too, never a pre-edit number. Without the flag a
 *   write still only refreshes its dirty dependency cone, never the whole book.
 * @param strict
 *   GH-496: promote a write's advisory conditions — formula-evaluation errors, iterative
 *   non-convergence, data-table seed warnings — from "printed in the summary, exit 0" to exit 1; a
 *   strict failure publishes nothing (#677). The default stays advisory so interactive use is
 *   unchanged.
 */
final case class WritePolicy(noRecalc: Boolean = false, strict: Boolean = false) derives CanEqual

object WritePolicy:
  /** Recalculate the dirty cone, report problems advisorily — the historical CLI behavior. */
  val default: WritePolicy = WritePolicy()

/**
 * GH-496: a write whose output COMPLETED (the staging file is written) but whose summary carries a
 * condition `--strict` promotes to a non-zero exit; the runner withholds the staged file (#677).
 *
 * Distinct from a plain failure so the runner can print the full summary verbatim — the counts, the
 * failing refs, the convergence verdict — instead of an `Error:`-prefixed one-liner, while still
 * exiting non-zero. Stack-trace free: nothing here is a defect to debug.
 */
final class StrictFailure(val summary: String) extends Exception(summary) with NoStackTrace
