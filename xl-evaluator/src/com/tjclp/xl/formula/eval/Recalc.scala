package com.tjclp.xl.formula.eval

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.error.XLError
import com.tjclp.xl.workbooks.{CalcPr, Workbook}

/**
 * A formula cell that failed to evaluate during `Workbook.recalculate`, with its location and
 * structured error. The cell is left uncached — Excel recalculates it on open.
 *
 * GH-344: these are COULD-NOT-EVALUATE host failures only (parse errors, missing sheets, unknown
 * functions, cycle participants and blocked cells, contained blowups). A formula that COMPUTES an
 * Excel error value (#DIV/0!, #N/A, ...) is a successful evaluation — it caches, writes as a
 * `t="e"` cell, and reports through [[RecalcResult.excelErrors]] instead.
 */
final case class CellEvalError(
  sheet: SheetName,
  ref: ARef,
  error: XLError
) derives CanEqual:
  /**
   * Human-readable one-liner: `Sales!B2: Formula error in '=A1/0': ...` — GH-280: cell-ref-shaped
   * sheet names quote (`'A1'!B2: ...`) so the location reads unambiguously.
   */
  def render: String = s"${SheetName.quoteForFormula(sheet.value)}!${ref.toA1}: ${error.message}"

/**
 * GH-373: opt-in bounded iterative calculation for circular workbooks.
 *
 * Professional models are routinely circular by design (interest on average debt balances) and ship
 * with `<calcPr iterate="1">`. Passing an `IterativeCalc` to `recalculate` fixpoints cycle members
 * with Jacobi iteration (every member reads PREVIOUS-iteration values, matching Excel): members
 * seed to 0, iterate until every member's |Δ| < `maxChange` or `maxIter` rounds, and
 * non-convergence keeps the last values with NO error — Excel's semantics, and the deliberate
 * inversion of the default path's circular-reference errors.
 *
 * Deliberately NOT auto-derived from `wb.metadata.calcPr` — iteration is opt-in so the default
 * `recalculate()` stays byte-identical. Bridge explicitly when honoring a file's settings:
 * {{{
 * wb.metadata.calcPr
 *   .filter(_.iterativeCalculation)
 *   .map(IterativeCalc.fromCalcPr)
 *   .fold(wb.recalculate())(it => wb.recalculate(it))
 * }}}
 *
 * GH-492: `maxIter`/`maxChange` are PER strongly-connected component, not per workbook. Each cyclic
 * component fixpoints against its own budget while the condensation is walked in dependency-first
 * order, so one permanently-oscillating cycle can no longer burn the budget of (or block the
 * convergence verdict of) an unrelated one. Worst-case total work is unchanged (`maxIter × |cyclic
 * core|` member evaluations) and matches Excel's per-cell budget.
 *
 * @param maxIter
 *   Iteration cap (Excel UI default 100; values < 1 are treated as 1)
 * @param maxChange
 *   Convergence threshold: iteration stops when every cycle member changes by LESS than this
 *   between rounds (Excel UI default 0.001)
 * @param seedFromCaches
 *   GH-469: seed each cycle member from its LOADED cached value when it has one, falling back to 0
 *   otherwise — Excel's semantics (iterative calculation starts from the current cell values). A
 *   book already at a valid fixpoint therefore re-solves to itself in one round instead of being
 *   driven back through the 0-seed transient, which on guarded topologies (mutually
 *   `IF(ISERROR(…))`-guarded pairs) silently replaced valid numeric caches with the guard's text
 *   branch — itself a fixpoint, so the run reported convergence with no error at all. Set false for
 *   a cold start: the supported escape hatch for a book whose caches are known to be poisoned.
 *   Members whose cache the dynamic (INDIRECT/OFFSET) bucket strips seed 0 regardless — their
 *   caches are declared stale. Any cached VALUE seeds, not just numbers; convergence still requires
 *   re-evaluating the formula and reproducing it.
 */
final case class IterativeCalc(
  maxIter: Int,
  maxChange: BigDecimal,
  seedFromCaches: Boolean = true
) derives CanEqual

object IterativeCalc:
  /**
   * Lift a workbook's modeled `<calcPr>` into iteration settings, applying Excel's defaults (100,
   * 0.001) for absent attributes. Callers gate on `calcPr.iterativeCalculation` themselves — see
   * the bridge example on [[IterativeCalc]].
   */
  def fromCalcPr(cp: CalcPr): IterativeCalc =
    IterativeCalc(cp.maxIterations.getOrElse(100), cp.maxChange.getOrElse(BigDecimal("0.001")))

/**
 * GH-492: one cyclic strongly-connected component's fixpoint verdict.
 *
 * Members are the component's cells sorted by (sheet name, A1) — the same order the Jacobi loop
 * iterates them in. Uses `(SheetName, ARef)` rather than the graph package's `QualifiedRef`: the
 * latter does not derive `CanEqual`, and the public result type stays free of a graph dependency
 * (the same shape [[CellEvalError]] already uses).
 *
 * @param converged
 *   true iff every member's |Δ| dropped below `maxChange` within `maxIter` rounds
 * @param rounds
 *   rounds actually run for THIS component — `maxIter` on exhaustion, else the converging round
 * @param maxDelta
 *   the largest |Δ| among numeric members in the final round (None when no member was numeric) —
 *   the residual a caller can size an exhaustion against
 */
final case class SccReport(
  members: Vector[(SheetName, ARef)],
  converged: Boolean,
  rounds: Int,
  maxDelta: Option[BigDecimal]
) derives CanEqual:

  /** e.g. `'Debt'!B7, 'Debt'!B8 (+3 more): exhausted 400 round(s), max |Δ| = 3.9` */
  def render: String =
    val shown = members
      .take(2)
      .map((s, r) => s"${SheetName.quoteForFormula(s.value)}!${r.toA1}")
      .mkString(", ")
    val more = if members.sizeIs > 2 then s" (+${members.size - 2} more)" else ""
    val verdict =
      if converged then s"converged in $rounds round(s)" else s"exhausted $rounds round(s)"
    val delta = maxDelta.fold("")(d => s", max |Δ| = $d")
    s"$shown$more: $verdict$delta"

/**
 * Result of a total, whole-workbook recalculation (`wb.recalculate()`).
 *
 * Evaluation is per-cell: a failing or cyclic formula is collected into `errors` and left uncached,
 * while every other formula — including those on the same sheet — still evaluates and caches.
 * Cross-sheet references are resolved against the workbook automatically.
 *
 * GH-344: Excel error VALUES are RESULTS, not failures. A formula computing #DIV/0! caches
 * `Formula(expr, Some(CellValue.Error(Div0)))`, writes as a `t="e"` cell (round-trip stable), and
 * its dependents read the error value exactly as Excel's do. Such cells appear in [[excelErrors]],
 * never in [[errors]].
 *
 * @param workbook
 *   The workbook with every successfully evaluated formula cached (`Formula(expr, Some(value))`) —
 *   including formulas whose computed value is an Excel error
 * @param evaluated
 *   Computed values per sheet for inspection without re-reading cells
 * @param errors
 *   Per-cell COULD-NOT-EVALUATE host failures only: parse errors, unknown functions, missing
 *   sheets, cycle participants, and cells blocked by a cycle
 * @param converged
 *   GH-454/GH-492: `cycles.forall(_.converged)` — false iff some cyclic component exhausted
 *   `maxIter` rounds without every member's |Δ| dropping below `maxChange`. The last-round values
 *   are still kept (Excel semantics) but callers can gate instead of mistaking exhaustion for
 *   stationarity. Non-iterative runs (default `recalculate()`, or iterative settings on an acyclic
 *   workbook) report true vacuously.
 *
 * '''GH-492 — what `converged = true` now certifies.''' An iterative recalculation walks the SCC
 * condensation of the workbook graph ONCE in dependency-first order: a run of acyclic components
 * evaluates normally, each cyclic component fixpoints against the freshly computed values of its
 * precedents. By induction over that order, after the last component every cell holds its
 * GLOBAL-fixpoint value (within each component's `maxChange`), so `converged = true` means one more
 * whole-workbook pass would change nothing beyond that tolerance — every component would recognise
 * its input as a fixpoint in round 1 (exactly, and bit-for-bit, with
 * `IterativeCalc.seedFromCaches = false`). Before GH-492 the pass was split pre-order / one flat
 * Jacobi / post-order and `converged` certified only that the Jacobi loop had stopped moving —
 * mid-wavefront, against whatever caches the acyclic precedents happened to hold. Two honest
 * caveats survive: per-component tolerances compose across the condensation only up to error
 * amplification through the DAG, and dynamic (INDIRECT/OFFSET) edges are invisible to Tarjan, so a
 * dynamic cycle is neither iterated nor covered by this flag (see the GH-274 note on
 * `recalculate`).
 *
 * '''Seed dependence.''' With `IterativeCalc.seedFromCaches` (the default, Excel's behavior) the
 * fixpoint reached depends on the input workbook's cached values. For a contraction that is
 * immaterial; for a genuinely multi-fixpoint nonlinear cycle it is not, and
 * `recalculate(wb) != recalculate(stripCaches(wb))` is then a real, documented property — the same
 * exposure Excel has.
 *
 * @param iterationsUsed
 *   GH-454/GH-492: `cycles.map(_.rounds).max` — the rounds run by the WORST component (0 when no
 *   iteration happened). Equals `maxIter` when any component exhausted; otherwise in (0, maxIter].
 * @param cycles
 *   GH-492: one [[SccReport]] per cyclic component actually iterated, sorted by the component's
 *   canonical key (its minimum member under (sheet name, A1)). Empty on non-iterative and acyclic
 *   runs. Future diagnostics extend [[SccReport]], not this class.
 */
final case class RecalcResult(
  workbook: Workbook,
  evaluated: Map[SheetName, Map[ARef, CellValue]],
  errors: Vector[CellEvalError],
  converged: Boolean = true,
  iterationsUsed: Int = 0,
  cycles: Vector[SccReport] = Vector.empty
) derives CanEqual:

  /** GH-492: the cyclic components that exhausted their budget — the offenders to name. */
  def unconverged: Vector[SccReport] = cycles.filterNot(_.converged)

  /**
   * GH-492: no cell failed to evaluate AND every cyclic component reached its fixpoint — the single
   * gate a caller can trust to mean "this workbook is at its global fixpoint".
   */
  def certified: Boolean = errors.isEmpty && converged

  /**
   * True when every formula in the workbook COMPUTED a value — possibly an Excel error value
   * (GH-344: check [[excelErrors]] to distinguish clean-with-error-cells from
   * clean-and-error-free).
   */
  def isClean: Boolean = errors.isEmpty

  /**
   * GH-344: formula cells whose computed VALUE is an Excel error (#DIV/0!, #N/A, ...) — data
   * conditions carried by the model, not evaluation failures. Deterministic order: sorted by (sheet
   * name, A1 reference).
   */
  def excelErrors: Vector[(SheetName, ARef, CellError)] =
    (for
      (sheet, cells) <- evaluated.toVector
      (ref, value) <- cells.toVector
      err <- carriedCellError(value).toList
    yield (sheet, ref, err)).sortBy((sheet, ref, _) => (sheet.value, ref.toA1))

  /** The Excel error a computed value carries, if any (cached formula values included). */
  private def carriedCellError(value: CellValue): Option[CellError] = value match
    case CellValue.Error(err) => Some(err)
    case CellValue.Formula(_, Some(cached), _) => carriedCellError(cached)
    case _ => None

  /**
   * Right(workbook) when clean, Left(errors) otherwise — for scripts that must not proceed on
   * partial results. Note (GH-344): error VALUES do not make a result unclean — gate on
   * [[excelErrors]] too when a script must also reject computed error cells.
   */
  def toEither: Either[Vector[CellEvalError], Workbook] =
    if isClean then Right(workbook) else Left(errors)
