package com.tjclp.xl.cli.output

import com.tjclp.xl.formula.eval.SccReport
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef

/**
 * The one rendering of iterative cycles that did not converge, shared by `view --eval` (#917),
 * `recalc` (#933) and `audit` (#934) so the three never word the same book differently. The verdict
 * itself is the evaluator's ([[SccReport.unconvergedVerdict]]).
 */
object Convergence:

  /** How many cycles a one-line message names before `… and N more`. */
  private val MaxNamed = 3

  /**
   * `Iterative calculation exhausted 5 round(s) without converging (last values kept): S!A1, S!B1:
   * exhausted 5 round(s), max |Δ| = 32`, or `None` when every cycle converged.
   */
  def message(unconverged: Vector[SccReport]): Option[String] =
    SccReport.unconvergedVerdict(unconverged).map { verdict =>
      val named = unconverged.take(MaxNamed).map(_.render).mkString("; ")
      val more =
        if unconverged.sizeIs > MaxNamed then s"; … and ${unconverged.size - MaxNamed} more"
        else ""
      s"Iterative calculation $verdict: $named$more"
    }

  /**
   * One object per cycle: `{cells, verdict, rounds, maxChange, stalled}` — `cells` the members as
   * qualified refs, `maxChange` the largest last-round |Δ| (null when none was measured; a
   * magnitude beyond a double's range as its decimal string).
   */
  def json(unconverged: Vector[SccReport]): ujson.Arr =
    ujson.Arr.from(unconverged.map { scc =>
      ujson.Obj(
        "cells" -> ujson.Arr.from(
          scc.members.map((sheet, ref) => ujson.Str(QualifiedRef(sheet, ref).toString))
        ),
        "verdict" -> ujson.Str(scc.verdict),
        "rounds" -> ujson.Num(scc.rounds),
        "maxChange" -> scc.maxDelta.fold[ujson.Value](ujson.Null) { d =>
          val asDouble = d.toDouble
          if asDouble.isInfinite then ujson.Str(d.toString) else ujson.Num(asDouble)
        },
        "stalled" -> ujson.Bool(scc.stalled)
      )
    })
