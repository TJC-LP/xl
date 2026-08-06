package com.tjclp.xl.formula.eval

import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cells.{CellValue, FormulaKind}
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.formula.Clock
import com.tjclp.xl.formula.ast.{BindingCoercion, TExpr}
import com.tjclp.xl.formula.functions.ArgValue
import com.tjclp.xl.formula.graph.DependencyGraph
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.formula.parser.FormulaParser
import com.tjclp.xl.formula.printer.FormulaPrinter
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook

/**
 * GH-453: per-table diagnostics from a seeding run — see [[DataTableSeedReport]]. Warnings are
 * ordered sheet order then record-ref row-major, matching the seeding order.
 */
enum SeedTableWarning derives CanEqual:
  /**
   * The table's source formula depends on a reference cycle but iterative calculation is not
   * available (neither declared via `<calcPr iterate="1"/>` nor passed explicitly): the table is
   * left entirely untouched — interiors stay uncached, so the data-table-unseeded lint still fires.
   * `cycle` renders the members as `'Sheet'!A1`, sorted, capped at 8.
   */
  case CircularNotIterated(sheet: SheetName, ref: CellRange, cycle: Vector[String])

  /**
   * `combinations` axis combinations exhausted `maxIter` rounds without every cycle member's |Δ|
   * dropping below `maxChange`; their last-round values were seeded (Excel semantics).
   */
  case NotConverged(sheet: SheetName, ref: CellRange, combinations: Int, maxIter: Int)

  /**
   * GH-453 (review follow-up): `cells` interior cells were left UNSEEDED, and `reason` says why —
   * either an axis value or cycle member failed to evaluate on the iterated path, or the whole
   * table was skipped by a budget guard (`MaxIterativeSeedBudget`, `MaxConeSeedBudget`) and `cells`
   * counts the interior. Both would otherwise read as a clean run. A table whose interiors WERE
   * seeded off a stale precedent reports [[ConeUnresolved]] instead — this case always means
   * unseeded cells.
   */
  case Skipped(sheet: SheetName, ref: CellRange, cells: Int, reason: String)

  /**
   * GH-493 (round-2 review): `cells` DISTINCT precedent cells in the what-if cone could not be
   * re-derived under the substitution, so they stayed on their loaded caches — the interiors were
   * still seeded (tolerant doctrine) but the grid may not reflect the substitution at all, which is
   * the very FLAT-grid mechanism GH-493 filed. Distinct from [[Skipped]]: nothing was left unseeded
   * here. `refs` renders the cells as `'Sheet'!A1`, sorted, capped at 8.
   */
  case ConeUnresolved(sheet: SheetName, ref: CellRange, cells: Int, refs: Vector[String])

  /**
   * GH-494: the source formula's own error guard (`IFERROR`/`IFNA`, or `IF(ISERROR(x),…)` and its
   * `ISERR`/`ISNA` siblings) BANKED ITS FALLBACK for `cells` axis combinations — the seeded value
   * is the fallback, not the modelled quantity. A guarded corner turns an evaluation failure into a
   * plausible-looking grid ("NM" text, or worse, a numeric fallback), so a run that seeds fallbacks
   * must never read as clean. `guard` renders the guarded expression (capped at 120 chars).
   * Legitimate too — a genuinely undefined combination — hence a warning, never a Left. Only guards
   * on the EVALUATED path count: a guard whose fallback never became the cell's value (a chained
   * ladder's inner rung, an untaken `IF` branch) is silent, since a clean run that reads as dirty
   * destroys the same signal as a dirty run that reads as clean.
   */
  case ErrorGuardFired(sheet: SheetName, ref: CellRange, cells: Int, guard: String)

/** GH-453: the seeded workbook plus every per-table warning the run produced. */
final case class DataTableSeedReport(
  workbook: Workbook,
  warnings: Vector[SeedTableWarning]
) derives CanEqual

/**
 * Explicit cache seeding for data-table interiors (GH-419).
 *
 * xl never evaluates `TABLE(...)` implicitly — a data-table record's cache is pinned as its only
 * truthful value (GH-430). Under `calcMode="autoNoTable"` (the house dialect) Excel itself never
 * recomputes tables on open either, so an uncached authored table would open BLANK. This seeder is
 * the explicit lever: it replays Excel's what-if substitution — overlay each axis value into the
 * input cell(s), evaluate the source formula, cache the result per interior cell.
 *
 * Semantics (tolerant, deterministic — never all-or-nothing):
 *   - `Right(v)` caches `v`, INCLUDING error VALUES (`#NUM!` interiors are legitimate data);
 *   - a `Left` (unsupported function, unresolvable axis) leaves that cell untouched and continues;
 *   - an empty or non-formula source cell caches `Number(0)` (Excel-verified bytes);
 *   - SHAPE-PRESERVING: DataTable-kind cells get `cachedValue` refreshed KEEPING their kind
 *     (foreign every-interior books keep per-cell faithfulness — only re-authoring normalizes);
 *     plain interior cells get their values replaced; real (non-record) formulas are never
 *     overwritten;
 *   - records with `del1`/`del2`, missing required inputs, no room for the corner/axes, or an
 *     interior larger than [[MaxInteriorArea]] cells (the record ref is file-controlled — a
 *     corrupt/hostile extent is a malformed record, never materialized) are skipped untouched;
 *   - a seed that changes NOTHING never marks a sheet modified, so a pristine book keeps the
 *     writer's verbatim-copy fidelity loop.
 *
 * Cells are computed row-major against the group's base sheet, so results are order-independent
 * within a group; groups seed in row-major record-ref order, sheets in workbook order.
 *
 * GH-493/GH-494 — the PRECEDENT CONE. The what-if lane must evaluate everything the solve lane
 * evaluates, or the substitution never reaches the source formula. Two failure modes, one cause:
 *   - a cached intermediate between the input cell and the source formula returns its BASE value
 *     (the pinned-cache doctrine), so every axis combination collapses to the same number — a FLAT
 *     grid with no warning (GH-493; the CLI `recalc --tables` lane hits it always, its recalc phase
 *     caching every precedent before the seed phase runs);
 *   - an UNCACHED formula inside a range argument is silently DROPPED by the range readers (they
 *     shrink the array rather than erroring), so XIRR-class functions see a length mismatch and a
 *     guarded corner banks its error arm into every interior (GH-494).
 * Both are cured by re-deriving the cone on TEMP sheets instead of trusting caches: the source
 * formula's transitive precedents split into input-INDEPENDENT uncached cells (evaluated once per
 * table) and axis-DEPENDENT cells (re-evaluated per combination), each in topological order with
 * every value written back before its dependents run — exactly `recalculate`'s `evalPass`. A cone
 * cell that fails to evaluate is left as it was and seeding continues (evalPass parity), but it is
 * COUNTED and reported: a stale cone cell is the very mechanism GH-493 filed, so it must never pass
 * silently. Precedents outside the cone keep the landed pinned-cache doctrine: a cached precedent
 * reads via its cache.
 *
 * GH-453 — circular books: a table whose source formula transitively depends on a reference cycle
 * cannot use the pinned-cache substitution (the what-if never propagates through the cycle and
 * every interior would silently seed the base value — FLAT). Such tables gate three ways:
 *   - source precedents disjoint from the cyclic core: the cone path above;
 *   - cycle reached AND iterative calculation available: each axis combination fixpoints the
 *     relevant cycle members (Jacobi, [[WorkbookEvaluator.jacobiFixpoint]]) on TEMP sheets whose
 *     stale downstream caches are stripped, then evaluates the source formula off the converged
 *     values — `maxIter` exhaustion seeds the last-round values (Excel semantics) and reports
 *     [[SeedTableWarning.NotConverged]];
 *   - cycle reached AND no iterative calculation: the table is skipped untouched and reports
 *     [[SeedTableWarning.CircularNotIterated]] — never a Left (tolerant-Right doctrine).
 *
 * Iterative settings auto-derive from the book's own `<calcPr iterate="1"/>` (or pass an
 * [[IterativeCalc]] explicitly). Auto-derivation is safe HERE, unlike `recalculate()`'s opt-in
 * posture (see [[IterativeCalc]]): the acyclic path is identical with or without settings, the
 * cyclic-without-settings path leaves the table untouched, and the only behavior that changes is
 * the previously silently-wrong flat seeding.
 */
object DataTableSeeder:

  /**
   * Interior-area skip cap, in cells. The record ref comes from the file, so a crafted
   * `B2:XFD1048576` interior (~1.7e10 cells) must skip like any other malformed record instead of
   * being iterated. 1,000,000 cells (a 1000x1000 grid) is orders beyond any real sensitivity table.
   */
  private val MaxInteriorArea: Long = 1000000L

  /**
   * GH-453 (review follow-up): per-table cost cap for the ITERATED path, in cycle-member
   * evaluations (`interior cells × maxIter × relevant cycle members`). `maxIter` is file-controlled
   * (`<calcPr iterateCount=…/>`), so a hostile or misconfigured book can declare effectively
   * unbounded work per table. A table whose worst-case cost exceeds this generous bound skips
   * untouched and reports [[SeedTableWarning.Skipped]] naming the budget. 20,000,000 member
   * evaluations is orders beyond any real circular sensitivity table (the 20x20 battery in the spec
   * costs 400 × 100 × 2 = 80,000).
   */
  private val MaxIterativeSeedBudget: Long = 20000000L

  /**
   * GH-493/GH-494: per-table cost cap for CONE re-derivation, in cell evaluations (`interior cells
   * × axis-dependent cone cells`). Both factors are file-controlled, so the same hostile-book
   * reasoning as [[MaxIterativeSeedBudget]] applies and the bound is the same generous 20,000,000:
   * past it the table skips untouched with a [[SeedTableWarning.Skipped]] naming the budget rather
   * than silently seeding a stale (FLAT) grid.
   */
  private val MaxConeSeedBudget: Long = 20000000L

  /** GH-494: guard predicates whose TRUE arm is an error arm (`IF(ISERROR(x), fallback, x)`). */
  private val ErrorPredicates: Set[String] = Set("ISERROR", "ISERR", "ISNA")

  extension (wb: Workbook)

    /**
     * Seed every data table on every sheet (system clock), honoring the book's own
     * `<calcPr iterate="1"/>` for tables that depend on a reference cycle.
     */
    def seedDataTables(): XLResult[Workbook] =
      Right(seedWorkbook(wb, None, Clock.system, declaredIterative(wb))._1)

    /** Seed every data table on one sheet: `Left(SheetNotFound)` for a bad name. */
    @annotation.targetName("seedDataTablesSheet")
    def seedDataTables(sheetName: SheetName): XLResult[Workbook] =
      seedDataTables(sheetName, None, Clock.system)

    /**
     * Seed the data tables on one sheet whose record ref intersects `only` (`None` = all), with an
     * explicit clock for volatile source formulas (read ONCE — one seeding run is one volatile
     * generation). Iterative settings auto-derive from the book's `<calcPr>`.
     */
    @annotation.targetName("seedDataTablesScoped")
    def seedDataTables(
      sheetName: SheetName,
      only: Option[CellRange],
      clock: Clock
    ): XLResult[Workbook] =
      seedDataTables(sheetName, only, clock, declaredIterative(wb))

    /**
     * GH-453: seed every data table with explicit iterative settings (system clock) — overrides
     * whatever the book's `<calcPr>` declares (including its absence).
     */
    @annotation.targetName("seedDataTablesIterative")
    def seedDataTables(iterative: IterativeCalc): XLResult[Workbook] =
      Right(seedWorkbook(wb, None, Clock.system, Some(iterative))._1)

    /**
     * GH-453: scoped seeding with explicit iterative settings — `None` disables iteration entirely
     * (circular tables then skip with a warning in the report-producing variants).
     */
    @annotation.targetName("seedDataTablesScopedIterative")
    def seedDataTables(
      sheetName: SheetName,
      only: Option[CellRange],
      clock: Clock,
      iterative: Option[IterativeCalc]
    ): XLResult[Workbook] =
      if wb.sheets.exists(_.name == sheetName) then
        Right(seedWorkbook(wb, Some((sheetName, only)), clock, iterative)._1)
      else Left(XLError.SheetNotFound(sheetName.value))

    /**
     * GH-453: whole-book seeding that also returns the per-table warnings (circular tables skipped
     * for lack of iteration, non-converged axis combinations). Iterative settings auto-derive from
     * the book's `<calcPr>`.
     */
    def seedDataTablesReport(): XLResult[DataTableSeedReport] =
      seedDataTablesReport(Clock.system, declaredIterative(wb))

    /** GH-453: whole-book seeding report with an explicit clock and iterative settings. */
    @annotation.targetName("seedDataTablesReportIterative")
    def seedDataTablesReport(
      clock: Clock,
      iterative: Option[IterativeCalc]
    ): XLResult[DataTableSeedReport] =
      val (seeded, warnings) = seedWorkbook(wb, None, clock, iterative)
      Right(DataTableSeedReport(seeded, warnings))

  /** The book's own iterative-calculation declaration, if any. */
  private def declaredIterative(wb: Workbook): Option[IterativeCalc] =
    wb.metadata.calcPr.filter(_.iterativeCalculation).map(IterativeCalc.fromCalcPr)

  /** Workbook-level graph facts shared by every group of one seeding run (GH-453). */
  private final case class CycleContext(
    deps: Map[QualifiedRef, Set[QualifiedRef]],
    dependents: Map[QualifiedRef, Set[QualifiedRef]],
    core: Set[QualifiedRef]
  )

  /**
   * Seed the scoped sheets (`None` = whole book) in workbook order, collecting warnings.
   *
   * The dependency graph derives at most once per run, lazily — only when a record group exists on
   * a scoped sheet. Reusing it across groups is sound: seeding writes only plain interior values
   * and record caches, neither of which are graph nodes (`fromWorkbookBounded` excludes DataTable
   * kinds).
   */
  private def seedWorkbook(
    wb: Workbook,
    scope: Option[(SheetName, Option[CellRange])],
    clock: Clock,
    iterative: Option[IterativeCalc]
  ): (Workbook, Vector[SeedTableWarning]) =
    val targets: Vector[SheetName] = scope match
      case Some((name, _)) => Vector(name)
      case None => wb.sheets.map(_.name)
    val only = scope.flatMap(_._2)
    // GH-453 (review follow-up): ONE seeding run is ONE volatile generation — the clock is pinned
    // once per run and used for axis values, member fixpoints AND source formulas, on both the
    // iterated and the acyclic paths. Lazy: a run that seeds nothing never reads the clock.
    lazy val pinnedClock: Clock = Clock.fixed(clock.today(), clock.now())
    lazy val cycles: CycleContext =
      val (deps, dependents) = DependencyGraph.fromWorkbookBounded(wb)
      CycleContext(deps, dependents, DependencyGraph.qualifiedCyclicNodes(deps))
    targets.foldLeft((wb, Vector.empty[SeedTableWarning])) { case ((accWb, warns), name) =>
      accWb.sheets.find(_.name == name) match
        case None => (accWb, warns)
        case Some(sheet) =>
          val groups =
            recordGroups(sheet).filter { case (ref, _) => only.forall(_.intersects(ref)) }
          if groups.isEmpty then (accWb, warns)
          else
            val (seeded, groupWarns) =
              groups.foldLeft((sheet, Vector.empty[SeedTableWarning])) {
                case ((acc, ws), (_, kind)) =>
                  val (next, w) =
                    seedGroup(accWb.put(acc), acc, kind, pinnedClock, iterative, cycles)
                  (next, ws ++ w)
              }
            // Only a real change may mark the sheet modified — a no-op seed (already-correct
            // caches, or nothing but skips) must keep the verbatim-copy write fidelity.
            val nextWb = if seeded == sheet then accWb else accWb.put(seeded)
            (nextWb, warns ++ groupWarns)
    }

  /**
   * Record groups by record ref, row-major; the representative kind comes from the group's
   * row-major-first cell (deterministic under any Map iteration order).
   */
  private def recordGroups(sheet: Sheet): Vector[(CellRange, FormulaKind.DataTable)] =
    sheet.cells.values.toVector
      .flatMap { cell =>
        cell.value match
          case CellValue.Formula(_, _, dt: FormulaKind.DataTable) => Vector((cell, dt))
          case _ => Vector.empty
      }
      .sortBy { case (cell, _) => (cell.ref.row.index0, cell.ref.col.index0) }
      .distinctBy { case (_, dt) => dt.ref }
      .map { case (_, dt) => (dt.ref, dt) }
      .sortBy { case (ref, _) => (ref.start.row.index0, ref.start.col.index0) }

  private def seedGroup(
    wb: Workbook,
    sheet: Sheet,
    kind: FormulaKind.DataTable,
    clock: Clock,
    iterative: Option[IterativeCalc],
    cycles: => CycleContext
  ): (Sheet, Vector[SeedTableWarning]) =
    val interior = kind.ref
    val startRow = interior.start.row.index0
    val startCol = interior.start.col.index0
    val oversized = interior.width.toLong * interior.height.toLong > MaxInteriorArea
    val inputs: Option[(ARef, Option[ARef])] =
      if kind.del1 || kind.del2 || startRow < 1 || startCol < 1 || oversized then None
      else
        (kind.r1, kind.r2) match
          case (Some(r1), Some(r2)) if kind.dt2D => Some((r1, Some(r2)))
          case (Some(r1), _) if !kind.dt2D => Some((r1, None))
          case _ => None // missing required input(s)
    inputs match
      case None => (sheet, Vector.empty) // skip untouched
      case Some((input1, input2)) =>
        // GH-453 gate: does the source formula transitively depend on a reference cycle?
        val ctx = cycles
        val srcQ = sourceRefs(kind).map(r => QualifiedRef(sheet.name, r))
        // GH-493/GH-494: the precedent closure is needed on BOTH lanes now — the cyclic gate reads
        // it, and the acyclic lane re-derives its cone from it (it is no longer sound to skip the
        // walk just because the book declares no cycle).
        val closure = DependencyGraph.qualifiedTransitivePrecedents(ctx.deps, srcQ) ++ srcQ
        val relevantCore = ctx.core.intersect(closure)
        if relevantCore.isEmpty then
          seedGroupAcyclic(wb, sheet, kind, input1, input2, clock, ctx, closure, srcQ)
        else
          iterative match
            case None =>
              // Refuse loudly, never Left: the table stays untouched (interiors uncached, so
              // the data-table-unseeded lint still fires) and the report names the cycle.
              val warning =
                SeedTableWarning.CircularNotIterated(
                  sheet.name,
                  interior,
                  renderRefs(relevantCore)
                )
              (sheet, Vector(warning))
            case Some(it) =>
              // GH-453 (review follow-up): the iterated path's worst case costs
              // interiorCells × maxIter × |members| member-evaluations, and maxIter comes from
              // the FILE. A table past the budget skips untouched — loudly, via Skipped.
              val interiorCells = interior.width.toLong * interior.height.toLong
              val rounds = math.max(1, it.maxIter)
              val cost = BigInt(interiorCells) * BigInt(rounds) * BigInt(relevantCore.size)
              if cost > BigInt(MaxIterativeSeedBudget) then
                val reason =
                  s"iterative seeding budget exceeded: $interiorCells interior cells x " +
                    s"$rounds max rounds x ${relevantCore.size} cycle members > " +
                    s"$MaxIterativeSeedBudget member-evaluations"
                val warning =
                  SeedTableWarning.Skipped(sheet.name, interior, interiorCells.toInt, reason)
                (sheet, Vector(warning))
              else
                seedGroupIterative(
                  wb,
                  sheet,
                  kind,
                  input1,
                  input2,
                  clock,
                  it,
                  ctx,
                  relevantCore,
                  closure,
                  srcQ
                )

  /**
   * GH-493/GH-494: the source formula's precedent cone, in topological order, split by WHEN it has
   * to be re-derived (see the object scaladoc).
   *
   * @param base
   *   input-independent precedents that carry no cache — the range readers would silently DROP
   *   them, so they are materialized ONCE per table onto the group's temp sheets
   * @param beforeCycle
   *   precedents the axis substitution moves that are NOT downstream of the cyclic core — resolved
   *   right after the axis overlay, so a fixpoint reading them sees the substituted values
   * @param afterCycle
   *   precedents downstream of the cyclic core — resolved after the fixpoint folds its members in
   */
  private final case class WhatIfCone(
    base: List[(QualifiedRef, Int)],
    beforeCycle: List[(QualifiedRef, Int)],
    afterCycle: List[(QualifiedRef, Int)]
  ):
    /** Per-axis-combination cell evaluations this cone costs. */
    def perCombinationSize: Long = beforeCycle.size.toLong + afterCycle.size.toLong

  /**
   * Build the [[WhatIfCone]] for one table group. `core` is the relevant cyclic core (empty on the
   * acyclic lane): its members belong to the fixpoint, never to the cone.
   */
  private def whatIfCone(
    wb: Workbook,
    sheetIndex: Map[SheetName, Int],
    ctx: CycleContext,
    closure: Set[QualifiedRef],
    srcQ: Set[QualifiedRef],
    inputQ: Set[QualifiedRef],
    core: Set[QualifiedRef]
  ): WhatIfCone =
    // Source cells evaluate as expressions (never written back) and input cells carry the axis
    // value; both are excluded, as are the cycle members the Jacobi fixpoint owns.
    val precedents = closure -- srcQ -- inputQ -- core
    val moved = DependencyGraph.qualifiedTransitiveDependents(ctx.dependents, inputQ ++ core)
    val axisDependent = precedents.intersect(moved)
    // An input-INDEPENDENT precedent only needs re-deriving when it carries no cache: a cached one
    // reads through its cache (pinned-cache doctrine) and cannot move under the substitution.
    val uncached = (precedents -- axisDependent).filter(q => isUncachedFormula(wb, q))
    val members = axisDependent ++ uncached
    val deps =
      members.iterator.map(q => q -> ctx.deps.getOrElse(q, Set.empty).intersect(members)).toMap
    val back =
      members.iterator
        .map(q => q -> ctx.dependents.getOrElse(q, Set.empty).intersect(members))
        .toMap
    // Defensive `Nil`: the induced subgraph is acyclic by construction (cycle members are excluded
    // on the iterated lane, and the acyclic lane only reaches here with an acyclic closure).
    val ordered = DependencyGraph.qualifiedTopologicalSort(deps, back).getOrElse(Nil)
    val indexed = ordered.flatMap(q => sheetIndex.get(q.sheet).map(idx => (q, idx)))
    val downstreamOfCore =
      if core.isEmpty then Set.empty[QualifiedRef]
      else DependencyGraph.qualifiedTransitiveDependents(ctx.dependents, core)
    WhatIfCone(
      base = indexed.filter((q, _) => uncached.contains(q)),
      beforeCycle =
        indexed.filter((q, _) => axisDependent.contains(q) && !downstreamOfCore.contains(q)),
      afterCycle =
        indexed.filter((q, _) => axisDependent.contains(q) && downstreamOfCore.contains(q))
    )

  /** A formula cell with no cache — the shape the range readers drop (GH-494). */
  private def isUncachedFormula(wb: Workbook, q: QualifiedRef): Boolean =
    wb(q.sheet).toOption.flatMap(_.cells.get(q.ref)).map(_.value) match
      // GH-430: a DataTable record's cache IS its value; an uncached one pins to Empty, never
      // re-derives.
      case Some(CellValue.Formula(_, None, _: FormulaKind.DataTable)) => false
      case Some(CellValue.Formula(_, None, _)) => true
      case _ => false

  /**
   * The temp sheets a cone resolution produced, plus every cone cell it could NOT re-derive (GH-493
   * review rework): such a cell stays on its stale cache, which is precisely the FLAT-grid
   * mechanism GH-493 filed, so the table must report it rather than seed silently.
   */
  private final case class ConeResolution(sheets: Vector[Sheet], unresolved: Set[QualifiedRef])

  /**
   * Evaluate `cone` in topological order on the temp sheets, writing each value back before its
   * dependents run — the seeder's copy of `recalculateImpl`'s `evalPass`, including its tolerance:
   * a cell that fails to evaluate keeps whatever it had and the pass continues. Unlike `evalPass`
   * the failures are COUNTED, not merely tolerated (see [[ConeResolution]]).
   */
  private def resolveCone(
    wb: Workbook,
    sheets: Vector[Sheet],
    cone: List[(QualifiedRef, Int)],
    clock: Clock
  ): ConeResolution =
    cone.foldLeft(ConeResolution(sheets, Set.empty)) { case (acc, (q, idx)) =>
      val temp = wb.copy(sheets = acc.sheets)
      SheetEvaluator.evaluateCell(acc.sheets(idx))(q.ref, clock, Some(temp)) match
        case Right(value) =>
          acc.copy(sheets = acc.sheets.updated(idx, acc.sheets(idx).put(q.ref, value)))
        case Left(_) => acc.copy(unresolved = acc.unresolved + q)
    }

  /**
   * GH-493 (review rework): one [[SeedTableWarning.ConeUnresolved]] per table naming the DISTINCT
   * cone refs that could not be re-derived — de-duplicated across axis combinations, so a 400-cell
   * grid over one broken precedent warns once. Never [[SeedTableWarning.Skipped]]: the interiors
   * here were seeded, they may just not reflect the substitution.
   */
  private def coneWarning(
    name: SheetName,
    interior: CellRange,
    unresolved: Set[QualifiedRef]
  ): Vector[SeedTableWarning] =
    if unresolved.isEmpty then Vector.empty
    else
      Vector(
        SeedTableWarning.ConeUnresolved(
          name,
          interior,
          unresolved.size,
          renderRefs(unresolved)
        )
      )

  /**
   * GH-493/GH-494: the acyclic what-if lane — axis overlay, cone re-derivation, source evaluation,
   * error-guard probe. Everything happens on temp sheets; the real workbook is never touched.
   */
  private def seedGroupAcyclic(
    wb: Workbook,
    sheet: Sheet,
    kind: FormulaKind.DataTable,
    input1: ARef,
    input2: Option[ARef],
    clock: Clock,
    ctx: CycleContext,
    closure: Set[QualifiedRef],
    srcQ: Set[QualifiedRef]
  ): (Sheet, Vector[SeedTableWarning]) =
    val interior = kind.ref
    val sheetIndex: Map[SheetName, Int] = wb.sheets.zipWithIndex.map((s, i) => s.name -> i).toMap
    sheetIndex.get(sheet.name) match
      case None => (sheet, Vector.empty) // defensive: wb always contains `sheet`
      case Some(tableIdx) =>
        val inputQ = (Set(input1) ++ input2).map(r => QualifiedRef(sheet.name, r))
        val cone = whatIfCone(wb, sheetIndex, ctx, closure, srcQ, inputQ, Set.empty)
        val coneBudget = coneBudgetExceeded(sheet.name, interior, cone)
        if coneBudget.nonEmpty then (sheet, coneBudget)
        else
          val probes = guardProbes(sheet, kind)
          // Input-independent uncached precedents are the same for every combination.
          val prepared = resolveCone(wb, wb.sheets, cone.base, clock)
          val (seeded, guarded, guardName, unresolved) =
            interior.cellsRowMajor.foldLeft((sheet, 0, Option.empty[String], prepared.unresolved)) {
              case ((acc, fired, named, unres), cellRef) =>
                computeCell(
                  wb,
                  sheet,
                  kind,
                  cellRef,
                  input1,
                  input2,
                  clock,
                  prepared.sheets,
                  tableIdx,
                  cone,
                  probes
                ) match
                  case None => (acc, fired, named, unres) // tolerant: leave the cell untouched
                  case Some(outcome) =>
                    (
                      writeInterior(acc, cellRef, outcome.value),
                      if outcome.guard.isDefined then fired + 1 else fired,
                      named.orElse(outcome.guard),
                      unres ++ outcome.unresolved
                    )
            }
          (
            seeded,
            coneWarning(sheet.name, interior, unresolved) ++
              guardWarning(sheet.name, interior, guarded, guardName)
          )

  /**
   * GH-493/GH-494: the cone's per-combination cost against [[MaxConeSeedBudget]] — a non-empty
   * result is the [[SeedTableWarning.Skipped]] the table must report INSTEAD of being seeded.
   */
  private def coneBudgetExceeded(
    name: SheetName,
    interior: CellRange,
    cone: WhatIfCone
  ): Vector[SeedTableWarning] =
    val interiorCells = interior.width.toLong * interior.height.toLong
    val cost = BigInt(interiorCells) * BigInt(cone.perCombinationSize)
    if cost <= BigInt(MaxConeSeedBudget) then Vector.empty
    else
      val reason =
        s"what-if cone budget exceeded: $interiorCells interior cells x " +
          s"${cone.perCombinationSize} axis-dependent precedents > " +
          s"$MaxConeSeedBudget cell evaluations"
      Vector(SeedTableWarning.Skipped(name, interior, interiorCells.toInt, reason))

  /** GH-494: one [[SeedTableWarning.ErrorGuardFired]] per table, or nothing when no guard fired. */
  private def guardWarning(
    name: SheetName,
    interior: CellRange,
    fired: Int,
    guard: Option[String]
  ): Vector[SeedTableWarning] =
    if fired <= 0 then Vector.empty
    else
      Vector(
        SeedTableWarning.ErrorGuardFired(name, interior, fired, guard.getOrElse(""))
      )

  /**
   * GH-453: per-axis-combination Jacobi fixpoint for a table whose source formula depends on the
   * cyclic core. The REAL workbook is never touched — everything happens on temp sheets:
   *
   *   1. temp sheets strip exactly `stripSet` — the cells BETWEEN the what-if inputs / cycle and
   *      the source formula (their loaded caches are stale under the substitution); cycle members
   *      are never stripped (the fixpoint overlays them) and DataTable record caches stay pinned
   *      per GH-430 (`stripFormulaCaches` guards them);
   *   2. per interior cell: axis values evaluate against the group's base sheet exactly like the
   *      acyclic path, overlay into the input cells on the temp sheets, the pre-cycle cone resolves
   *      (GH-493/GH-494), the relevant cycle members fixpoint via
   *      [[WorkbookEvaluator.jacobiFixpoint]] (previous-round reads, 0-seeded, strict
   *      |Δ| < maxChange, the run's pinned clock), converged values fold in as plain values, the
   *      post-cycle cone resolves off them, and the source formula evaluates last;
   *   3. a Left anywhere stays tolerant — that cell is left untouched and seeding continues, but
   *      the table reports ONE [[SeedTableWarning.Skipped]] counting the unseeded cells (a
   *      fully-skipped table must never read as a clean run).
   */
  private def seedGroupIterative(
    wb: Workbook,
    sheet: Sheet,
    kind: FormulaKind.DataTable,
    input1: ARef,
    input2: Option[ARef],
    clock: Clock,
    iterative: IterativeCalc,
    ctx: CycleContext,
    relevantCore: Set[QualifiedRef],
    closure: Set[QualifiedRef],
    srcQ: Set[QualifiedRef]
  ): (Sheet, Vector[SeedTableWarning]) =
    val sheetIndex: Map[SheetName, Int] = wb.sheets.zipWithIndex.map((s, i) => s.name -> i).toMap
    sheetIndex.get(sheet.name) match
      case None => (sheet, Vector.empty) // defensive: wb always contains `sheet`
      case Some(tableIdx) =>
        val inputQ = (Set(input1) ++ input2).map(r => QualifiedRef(sheet.name, r))
        val stripSet =
          DependencyGraph
            .qualifiedTransitiveDependents(ctx.dependents, inputQ ++ relevantCore)
            .intersect(closure) -- relevantCore
        val stripBySheet = stripSet.groupMap(_.sheet)(_.ref)
        val stripped: Vector[Sheet] = wb.sheets.map { s =>
          val toStrip = stripBySheet.getOrElse(s.name, Set.empty)
          if toStrip.isEmpty then s else SheetEvaluator.stripFormulaCaches(s, toStrip)
        }
        // GH-493/GH-494: stripping alone is not enough — a stripped precedent VANISHES from any
        // range argument (the readers drop undecodable cells and shrink the array). The cone is
        // re-derived per combination; its input-independent uncached half resolves once, here.
        val cone = whatIfCone(wb, sheetIndex, ctx, closure, srcQ, inputQ, relevantCore)
        val coneBudget = coneBudgetExceeded(sheet.name, kind.ref, cone)
        if coneBudget.nonEmpty then (sheet, coneBudget)
        else
          seedIteratedInterior(
            wb,
            sheet,
            kind,
            input1,
            input2,
            clock,
            iterative,
            ctx,
            relevantCore,
            inputQ,
            sheetIndex,
            tableIdx,
            resolveCone(wb, stripped, cone.base, clock),
            cone
          )

  /** The interior fold of the iterated lane, once the cone and temp sheets are prepared. */
  private def seedIteratedInterior(
    wb: Workbook,
    sheet: Sheet,
    kind: FormulaKind.DataTable,
    input1: ARef,
    input2: Option[ARef],
    clock: Clock,
    iterative: IterativeCalc,
    ctx: CycleContext,
    relevantCore: Set[QualifiedRef],
    inputQ: Set[QualifiedRef],
    sheetIndex: Map[SheetName, Int],
    tableIdx: Int,
    prepared: ConeResolution,
    cone: WhatIfCone
  ): (Sheet, Vector[SeedTableWarning]) =
    // The what-if substitution PINS the input cells (Excel semantics): a cycle running
    // through an input is broken there, so the inputs are never Jacobi members — they stay
    // the plain axis values overlaid in step (2). (stripSet already excludes relevantCore,
    // hence the pinned inputs.) The reduced set may even be acyclic after pinning;
    // jacobiFixpoint still converges (prev-round reads settle in <= |members|+1 rounds).
    val members: List[(QualifiedRef, Int, String)] =
      (relevantCore -- inputQ).toList
        .flatMap { q =>
          sheetIndex.get(q.sheet).map { idx =>
            val expr = wb(q.sheet).toOption.flatMap(_.cells.get(q.ref)).map(_.value) match
              case Some(CellValue.Formula(e, _, _)) => e
              case _ => q.ref.toA1
            (q, idx, expr)
          }
        }
        .sortBy((q, _, _) => (q.sheet.value, q.ref.toA1))
    val probes = guardProbes(sheet, kind)
    // The clock arrives pinned from seedWorkbook: one seeding run is one volatile
    // generation, like one recalculation (GH-373) — axis values, member fixpoints, cone
    // cells and source formulas all read the same instant.
    val (seeded, unconverged, skipped, guarded, guardName, unresolved) =
      kind.ref.cellsRowMajor.foldLeft(
        (sheet, 0, 0, 0, Option.empty[String], prepared.unresolved)
      ) { case ((acc, fails, skips, fired, named, unres), cellRef) =>
        val computed = computeCellIterative(
          wb,
          sheet,
          kind,
          cellRef,
          input1,
          input2,
          clock,
          iterative,
          prepared.sheets,
          tableIdx,
          members,
          cone,
          probes
        )
        computed match
          case None => (acc, fails, skips + 1, fired, named, unres) // tolerant: leave it untouched
          case Some(outcome) =>
            (
              writeInterior(acc, cellRef, outcome.value),
              if outcome.converged then fails else fails + 1,
              skips,
              if outcome.guard.isDefined then fired + 1 else fired,
              named.orElse(outcome.guard),
              unres ++ outcome.unresolved
            )
      }
    val notConverged =
      if unconverged > 0 then
        Vector(
          SeedTableWarning.NotConverged(
            sheet.name,
            kind.ref,
            unconverged,
            math.max(1, iterative.maxIter)
          )
        )
      else Vector.empty
    // GH-453 (review follow-up): unseeded cells must be VISIBLE in the report — one Skipped
    // per table counting them (axis-value or cycle-member evaluation failures).
    val skippedWarning =
      if skipped > 0 then
        Vector(
          SeedTableWarning.Skipped(
            sheet.name,
            kind.ref,
            skipped,
            "axis value or cycle-member evaluation failed"
          )
        )
      else Vector.empty
    (
      seeded,
      notConverged ++ skippedWarning ++ coneWarning(sheet.name, kind.ref, unresolved) ++
        guardWarning(sheet.name, kind.ref, guarded, guardName)
    )

  /**
   * One interior cell's circular what-if value, its convergence verdict (GH-453) and the error
   * guard that fired for it, if any (GH-494). `clock` is the run's pinned clock — axis, cone,
   * fixpoint and source evaluations share one volatile generation.
   */
  private def computeCellIterative(
    wb: Workbook,
    sheet: Sheet,
    kind: FormulaKind.DataTable,
    cellRef: ARef,
    input1: ARef,
    input2: Option[ARef],
    clock: Clock,
    iterative: IterativeCalc,
    prepared: Vector[Sheet],
    tableIdx: Int,
    members: List[(QualifiedRef, Int, String)],
    cone: WhatIfCone,
    probes: Map[ARef, TExpr[?]]
  ): Option[CellOutcome] =
    val (sourceRef, overlayRefs) = whatIfGeometry(kind, cellRef, input1, input2)
    sheet.cells.get(sourceRef).map(_.value) match
      case Some(CellValue.Formula(expression, _, _)) =>
        // (1) axis values evaluate against the group's base sheet, exactly like the acyclic path
        val overlays = overlayRefs.foldLeft(Option(Vector.empty[(ARef, CellValue)])) {
          case (acc, (inputRef, axisRef)) =>
            acc.flatMap { vs =>
              SheetEvaluator
                .evaluateCell(sheet)(axisRef, clock, Some(wb))
                .toOption
                .map(axisValue => vs :+ (inputRef -> axisValue))
            }
        }
        overlays.flatMap { axisValues =>
          // (2) overlay the axis values into the input cells on the PREPARED temp sheets, then
          // resolve the cone the fixpoint READS (GH-493/GH-494) — upstream of the cycle, so its
          // members see substituted, decodable precedents rather than stripped ones
          val overlaid = axisValues.foldLeft(prepared) { case (sheets, (inputRef, v)) =>
            sheets.updated(tableIdx, sheets(tableIdx).put(inputRef, v))
          }
          val upstream = resolveCone(wb, overlaid, cone.beforeCycle, clock)
          // (3) fixpoint the relevant cycle members under this axis combination
          val (results, converged, _) =
            WorkbookEvaluator.jacobiFixpoint(wb, upstream.sheets, members, iterative, clock, None)
          val sourceQ = QualifiedRef(sheet.name, sourceRef)
          if members.exists((q, _, _) => q == sourceQ) then
            // The source formula itself was fixpointed. Re-evaluating it after folding the final
            // member values would advance it by one extra Jacobi round.
            results
              .get(sourceQ)
              .flatMap(_.toOption)
              .map(value => CellOutcome(value, converged, None, upstream.unresolved))
          else
            // (4) fold member values in as plain values, resolve the cone downstream of the cycle,
            // then evaluate the source formula
            val folded = members.foldLeft(Option(upstream.sheets)) { case (accOpt, (q, idx, _)) =>
              accOpt.flatMap { sheets =>
                results.get(q) match
                  case Some(Right(v)) => Some(sheets.updated(idx, sheets(idx).put(q.ref, v)))
                  case _ => None // tolerant: a failing member leaves this cell untouched
              }
            }
            folded.flatMap { sheets =>
              val resolved = resolveCone(wb, sheets, cone.afterCycle, clock)
              val tempWb = wb.copy(sheets = resolved.sheets)
              val tempSheet = resolved.sheets(tableIdx)
              SheetEvaluator
                .evaluateFormula(tempSheet)(expression, clock, Some(tempWb), None)
                .toOption
                .map { value =>
                  val guard = firedGuard(probes.get(sourceRef), tempSheet, tempWb, clock)
                  CellOutcome(
                    value,
                    converged,
                    guard,
                    upstream.unresolved ++ resolved.unresolved
                  )
                }
            }
        }
      case _ =>
        // Empty or non-formula source cell: Excel caches 0 (fixture-verified bytes).
        Some(CellOutcome(CellValue.Number(BigDecimal(0)), converged = true, None, Set.empty))

  /**
   * One interior cell's what-if outcome: the value to bank, whether the cycle converged (always
   * `true` off the iterated lane), the error guard that banked its fallback for this cell if any
   * (GH-494), and the cone cells that could not be re-derived under this substitution (GH-493
   * review rework).
   */
  private final case class CellOutcome(
    value: CellValue,
    converged: Boolean,
    guard: Option[String],
    unresolved: Set[QualifiedRef]
  )

  /**
   * One interior cell's what-if outcome on the acyclic lane; `None` leaves the cell untouched
   * (tolerant).
   */
  private def computeCell(
    wb: Workbook,
    sheet: Sheet,
    kind: FormulaKind.DataTable,
    cellRef: ARef,
    input1: ARef,
    input2: Option[ARef],
    clock: Clock,
    prepared: Vector[Sheet],
    tableIdx: Int,
    cone: WhatIfCone,
    probes: Map[ARef, TExpr[?]]
  ): Option[CellOutcome] =
    val (sourceRef, overlayRefs) = whatIfGeometry(kind, cellRef, input1, input2)
    sheet.cells.get(sourceRef).map(_.value) match
      case Some(CellValue.Formula(expression, _, _)) =>
        // Axis values evaluate against the group's base sheet, so results are order-independent.
        val overlays = overlayRefs.foldLeft(Option(prepared)) { case (acc, (inputRef, axisRef)) =>
          acc.flatMap { sheets =>
            SheetEvaluator
              .evaluateCell(sheet)(axisRef, clock, Some(wb))
              .toOption
              .map(axisValue => sheets.updated(tableIdx, sheets(tableIdx).put(inputRef, axisValue)))
          }
        }
        overlays.flatMap { overlaid =>
          // GH-493/GH-494: re-derive the cone under THIS substitution before the source formula
          // reads it — a cached intermediate would otherwise answer with its base value and an
          // uncached one would vanish from any range argument.
          val resolved = resolveCone(wb, overlaid, cone.beforeCycle ++ cone.afterCycle, clock)
          val tempWb = wb.copy(sheets = resolved.sheets)
          val tempSheet = resolved.sheets(tableIdx)
          SheetEvaluator
            .evaluateFormula(tempSheet)(expression, clock, Some(tempWb), None)
            .toOption
            .map(value =>
              CellOutcome(
                value,
                converged = true,
                firedGuard(probes.get(sourceRef), tempSheet, tempWb, clock),
                resolved.unresolved
              )
            )
        }
      case _ =>
        // Empty or non-formula source cell: Excel caches 0 (fixture-verified bytes).
        Some(CellOutcome(CellValue.Number(BigDecimal(0)), converged = true, None, Set.empty))

  /**
   * GH-494: each of the group's source formulas parsed ONCE per table, kept only when it carries an
   * error guard at all (a guard-free formula never needs the path walk). An unparseable source
   * formula simply carries no probe — evaluation would fail on it too, and the interior stays
   * untouched.
   */
  private def guardProbes(
    sheet: Sheet,
    kind: FormulaKind.DataTable
  ): Map[ARef, TExpr[?]] =
    sourceRefs(kind).iterator.flatMap { r =>
      sheet.cells.get(r).map(_.value) match
        case Some(CellValue.Formula(expression, _, _)) =>
          FormulaParser.parse(expression).toOption.filter(containsErrorGuard).map(expr => r -> expr)
        case _ => None
    }.toMap

  /**
   * GH-494: the first guard that ACTUALLY fired for this interior cell, rendered for the warning
   * (capped at 120 chars). `None` means no fallback was banked.
   *
   * "Actually fired" is decided STRUCTURALLY, top-down from the root of the source formula, by
   * walking only the path Excel itself evaluated (see [[firstFiredGuard]]) — never by comparing a
   * fallback's standalone value to the value being banked. Value equality is wrong in both
   * directions: it misses every guard that is not the root of the formula (`IFERROR(1/A1,0)+5`
   * banks 5, which no arm equals alone) and every `IF(ISERROR(x), <cell ref>, x)` (the ref does not
   * reproduce the banked value when evaluated on its own), while false-positiving whenever an
   * untaken fallback coincidentally equals what the taken path produced.
   */
  private def firedGuard(
    probe: Option[TExpr[?]],
    sheet: Sheet,
    wb: Workbook,
    clock: Clock
  ): Option[String] =
    probe
      .flatMap(expr => firstFiredGuard(expr, sheet, wb, clock))
      .map(guarded => FormulaPrinter.print(guarded, includeEquals = false).take(120))

  /**
   * GH-494: the first error guard on the EVALUATED path, as the protected expression it guards.
   *
   *   - `IFERROR(a, b)` / `IFNA(a, b)`: `a` erroring under this substitution IS the guard firing —
   *     `b` is never entered, since a fallback that was never selected contributed nothing;
   *   - `IF(cond, t, e)` where `cond` is `ISERROR`/`ISERR`/`ISNA` over `x`: `x` erroring IS the
   *     guard firing. Otherwise the condition evaluates (Excel truthiness, the same evaluator and
   *     pinned clock as the seeding run) and only the SELECTED branch is walked — a condition that
   *     fails to evaluate selects neither;
   *   - every other node: all children are on the evaluated path and are walked left to right.
   *
   * The walk is per interior cell because the verdict depends on the substituted temp sheet; the
   * AST it walks is parsed once per table ([[guardProbes]]).
   */
  private def firstFiredGuard(
    expr: TExpr[?],
    sheet: Sheet,
    wb: Workbook,
    clock: Clock
  ): Option[TExpr[?]] =
    expr match
      case call: TExpr.Call[?] =>
        val args = callArgExprs(call)
        call.spec.name.toUpperCase match
          case "IFERROR" | "IFNA" =>
            args match
              case guarded :: _ =>
                if guardResolvesToError(guarded, sheet, wb, clock) then Some(guarded)
                else firstFiredGuard(guarded, sheet, wb, clock)
              case Nil => None
          case "IF" =>
            args match
              case cond :: branches =>
                firedPredicate(cond, sheet, wb, clock)
                  .orElse(firstFiredGuard(cond, sheet, wb, clock))
                  .orElse(
                    branchTaken(cond, sheet, wb, clock)
                      .flatMap(taken => branches.drop(if taken then 0 else 1).headOption)
                      .flatMap(branch => firstFiredGuard(branch, sheet, wb, clock))
                  )
              case Nil => None
          case _ => firstFiredIn(args, sheet, wb, clock)
      case other => firstFiredIn(children(other), sheet, wb, clock)

  /** The first fired guard among a node's on-path children, left to right. */
  private def firstFiredIn(
    nodes: List[TExpr[?]],
    sheet: Sheet,
    wb: Workbook,
    clock: Clock
  ): Option[TExpr[?]] =
    nodes.iterator.flatMap(node => firstFiredGuard(node, sheet, wb, clock)).nextOption()

  /**
   * GH-494: an `IF` condition of the form `ISERROR(x)` (or `ISERR`/`ISNA`) whose `x` resolves to an
   * error — the house-standard guarded headline, firing into the TRUE arm. The parser wraps a
   * function-shaped condition in a Boolean coercion, hence [[unwrapTransparent]].
   */
  private def firedPredicate(
    cond: TExpr[?],
    sheet: Sheet,
    wb: Workbook,
    clock: Clock
  ): Option[TExpr[?]] =
    unwrapTransparent(cond) match
      case c: TExpr.Call[?] if ErrorPredicates.contains(c.spec.name.toUpperCase) =>
        callArgExprs(c).find(guarded => guardResolvesToError(guarded, sheet, wb, clock))
      case _ => None

  /**
   * Excel truthiness of an `IF` condition under this substitution — the same scalar coercion the
   * `IF` implementation itself applies. `None` when the condition cannot be evaluated or coerced:
   * neither branch is then on the evaluated path.
   */
  private def branchTaken(
    cond: TExpr[?],
    sheet: Sheet,
    wb: Workbook,
    clock: Clock
  ): Option[Boolean] =
    Evaluator.instance.eval(cond, sheet, clock, Some(wb), None) match
      case Left(_) => None
      case Right(result) =>
        val scalar = result match
          case array: ArrayResult => ScalarCoercion.collapseArray(array)
          case other => other
        ScalarCoercion.coerce("IF condition", scalar, BindingCoercion.Bool) match
          case Right(taken: Boolean) => Some(taken)
          case _ => None

  /** ISERROR's own verdict on a guarded expression: a Left OR an error VALUE (see `iserror`). */
  private def guardResolvesToError(
    probe: TExpr[?],
    sheet: Sheet,
    wb: Workbook,
    clock: Clock
  ): Boolean =
    Evaluator.instance.eval(probe, sheet, clock, Some(wb), None) match
      case Left(_) => true
      case Right(value) =>
        EvalResult.toCellValue(value) match
          case CellValue.Error(_) => true
          case _ => false

  /**
   * GH-494: does this formula carry an error guard ANYWHERE — `IFERROR`, `IFNA`, or `IF` over an
   * `ISERROR`/`ISERR`/`ISNA` condition? Purely structural: the per-cell path walk only runs for
   * source formulas that could fire at all, so a guard-free corner costs nothing extra.
   */
  private def containsErrorGuard(expr: TExpr[?]): Boolean =
    isGuardNode(expr) || children(expr).exists(containsErrorGuard)

  /** Is this node itself an error guard (see [[containsErrorGuard]])? */
  private def isGuardNode(expr: TExpr[?]): Boolean =
    expr match
      case call: TExpr.Call[?] =>
        call.spec.name.toUpperCase match
          case "IFERROR" | "IFNA" => true
          case "IF" =>
            callArgExprs(call).headOption.map(unwrapTransparent) match
              case Some(c: TExpr.Call[?]) => ErrorPredicates.contains(c.spec.name.toUpperCase)
              case _ => false
          case _ => false
      case _ => false

  /**
   * The sub-expressions of a node. Structural cases recurse; everything else (literals, refs,
   * ranges, bindings) is a leaf for guard purposes.
   */
  private def children(expr: TExpr[?]): List[TExpr[?]] =
    expr match
      case call: TExpr.Call[?] => callArgExprs(call)
      case TExpr.Add(l, r) => List(l, r)
      case TExpr.Sub(l, r) => List(l, r)
      case TExpr.Mul(l, r) => List(l, r)
      case TExpr.Div(l, r) => List(l, r)
      case TExpr.Pow(l, r) => List(l, r)
      case TExpr.Concat(l, r) => List(l, r)
      case TExpr.Eq(l, r) => List(l, r)
      case TExpr.Neq(l, r) => List(l, r)
      case TExpr.Lt(l, r) => List(l, r)
      case TExpr.Lte(l, r) => List(l, r)
      case TExpr.Gt(l, r) => List(l, r)
      case TExpr.Gte(l, r) => List(l, r)
      case TExpr.ToInt(e) => List(e)
      case TExpr.UnaryPlus(e) => List(e)
      case TExpr.Percent(e) => List(e)
      case TExpr.DateToSerial(e) => List(e)
      case TExpr.DateTimeToSerial(e) => List(e)
      case TExpr.Coerced(inner, _) => List(inner)
      case TExpr.Let(bindings, body) => bindings.map((_, value) => value) :+ body
      case _ => Nil

  /** The expression-shaped arguments of a call (range/cells positions cannot hold a guard). */
  private def callArgExprs(call: TExpr.Call[?]): List[TExpr[?]] =
    call.spec.argSpec.toValues(call.args).collect { case ArgValue.Expr(e) => e }

  /** Peel the analysis-transparent wrappers the parser adds around a nested expression. */
  @annotation.tailrec
  private def unwrapTransparent(expr: TExpr[?]): TExpr[?] =
    expr match
      case TExpr.Coerced(inner, _) => unwrapTransparent(inner)
      case TExpr.UnaryPlus(inner) => unwrapTransparent(inner)
      case other => other

  /**
   * D4 geometry shared by both substitution paths: the source formula cell for this interior cell,
   * and which axis value feeds which input cell.
   */
  private def whatIfGeometry(
    kind: FormulaKind.DataTable,
    cellRef: ARef,
    input1: ARef,
    input2: Option[ARef]
  ): (ARef, Vector[(ARef, ARef)]) =
    val interior = kind.ref
    val startRow = interior.start.row.index0
    val startCol = interior.start.col.index0
    val col = cellRef.col.index0
    val row = cellRef.row.index0
    val topAxis = ARef.from0(col, startRow - 1)
    val leftAxis = ARef.from0(startCol - 1, row)
    val sourceRef =
      if kind.dt2D then ARef.from0(startCol - 1, startRow - 1) // corner
      else if kind.dtr then ARef.from0(startCol - 1, row) // row-oriented: left of THIS row
      else ARef.from0(col, startRow - 1) // column-oriented: above THIS column
    val overlayRefs: Vector[(ARef, ARef)] =
      input2 match
        case Some(colInput) => Vector(input1 -> topAxis, colInput -> leftAxis) // 2-D
        case None if kind.dtr => Vector(input1 -> topAxis) // 1-D row-oriented
        case None => Vector(input1 -> leftAxis) // 1-D column-oriented
    (sourceRef, overlayRefs)

  /** Every distinct source-formula cell the interior's what-if substitution reads. */
  private def sourceRefs(kind: FormulaKind.DataTable): Set[ARef] =
    val interior = kind.ref
    val startRow = interior.start.row.index0
    val startCol = interior.start.col.index0
    if kind.dt2D then Set(ARef.from0(startCol - 1, startRow - 1))
    else if kind.dtr then
      (startRow to interior.end.row.index0).map(r => ARef.from0(startCol - 1, r)).toSet
    else (startCol to interior.end.col.index0).map(c => ARef.from0(c, startRow - 1)).toSet

  /** Shape-preserving interior write shared by both paths (see the object scaladoc). */
  private def writeInterior(acc: Sheet, cellRef: ARef, value: CellValue): Sheet =
    acc.cells.get(cellRef).map(_.value) match
      case Some(record @ CellValue.Formula(_, _, _: FormulaKind.DataTable)) =>
        acc.put(cellRef, record.copy(cachedValue = Some(value))) // shape-preserving
      case Some(_: CellValue.Formula) => acc // never overwrite a real formula
      case _ => acc.put(cellRef, value)

  /**
   * Qualified refs rendered `'Sheet'!A1`, sorted, capped at 8 — the payload shape shared by the
   * cycle members of [[SeedTableWarning.CircularNotIterated]] and the unresolved cone cells of
   * [[SeedTableWarning.ConeUnresolved]].
   */
  private def renderRefs(core: Set[QualifiedRef]): Vector[String] =
    core.toVector
      .sortBy(q => (q.sheet.value, q.ref.toA1))
      .take(8)
      .map(q => s"${SheetName.quoteForFormula(q.sheet.value)}!${q.ref.toA1}")
