package com.tjclp.xl.formula.eval

import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cells.{CellValue, FormulaKind}
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.formula.Clock
import com.tjclp.xl.formula.graph.DependencyGraph
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
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
   * GH-453 (review follow-up): `cells` interior cells were left UNSEEDED on the iterated path — an
   * axis value or cycle member failed to evaluate for those cells, or the whole table was skipped
   * by the iterative-seeding budget guard (see `DataTableSeeder.MaxIterativeSeedBudget`). `reason`
   * is a short human-readable cause. Without this case a fully-skipped table reads as a clean run.
   */
  case Skipped(sheet: SheetName, ref: CellRange, cells: Int, reason: String)

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
 * Upstream formulas keep the landed pinned-cache doctrine: a cached precedent reads via its cache.
 * Cells are computed row-major against the group's base sheet, so results are order-independent
 * within a group; groups seed in row-major record-ref order, sheets in workbook order.
 *
 * GH-453 — circular books: a table whose source formula transitively depends on a reference cycle
 * cannot use the pinned-cache substitution (the what-if never propagates through the cycle and
 * every interior would silently seed the base value — FLAT). Such tables gate three ways:
 *   - source precedents disjoint from the cyclic core: today's pinned-cache path, bit-identical;
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
 * posture (see [[IterativeCalc]]): the acyclic path is bit-identical with or without settings, the
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
        val closure =
          if ctx.core.isEmpty then Set.empty[QualifiedRef]
          else DependencyGraph.qualifiedTransitivePrecedents(ctx.deps, srcQ) ++ srcQ
        val relevantCore = ctx.core.intersect(closure)
        if relevantCore.isEmpty then
          // Acyclic path: today's pinned-cache substitution, bit-identical (no stripping).
          val seeded = interior.cellsRowMajor.foldLeft(sheet) { (acc, cellRef) =>
            computeCell(wb, sheet, kind, cellRef, input1, input2, clock) match
              case None => acc // tolerant: leave the cell untouched
              case Some(value) => writeInterior(acc, cellRef, value)
          }
          (seeded, Vector.empty)
        else
          iterative match
            case None =>
              // Refuse loudly, never Left: the table stays untouched (interiors uncached, so
              // the data-table-unseeded lint still fires) and the report names the cycle.
              val warning =
                SeedTableWarning.CircularNotIterated(
                  sheet.name,
                  interior,
                  renderCycle(relevantCore)
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
                  closure
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
   *      acyclic path, overlay into the input cells on the temp sheets, the relevant cycle members
   *      fixpoint via [[WorkbookEvaluator.jacobiFixpoint]] (previous-round reads, 0-seeded, strict
   *      |Δ| < maxChange, the run's pinned clock), converged values fold in as plain values, and
   *      the source formula evaluates off them;
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
    closure: Set[QualifiedRef]
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
        val prepared: Vector[Sheet] = wb.sheets.map { s =>
          val toStrip = stripBySheet.getOrElse(s.name, Set.empty)
          if toStrip.isEmpty then s else SheetEvaluator.stripFormulaCaches(s, toStrip)
        }
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
        // The clock arrives pinned from seedWorkbook: one seeding run is one volatile
        // generation, like one recalculation (GH-373) — axis values, member fixpoints and
        // source formulas all read the same instant.
        val (seeded, unconverged, skipped) =
          kind.ref.cellsRowMajor.foldLeft((sheet, 0, 0)) { case ((acc, fails, skips), cellRef) =>
            val computed = computeCellIterative(
              wb,
              sheet,
              kind,
              cellRef,
              input1,
              input2,
              clock,
              iterative,
              prepared,
              tableIdx,
              members
            )
            computed match
              case None => (acc, fails, skips + 1) // tolerant: leave the cell untouched
              case Some((value, converged)) =>
                (writeInterior(acc, cellRef, value), if converged then fails else fails + 1, skips)
          }
        val notConverged =
          if unconverged > 0 then
            Vector(
              SeedTableWarning.NotConverged(sheet.name, kind.ref, unconverged, iterative.maxIter)
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
        (seeded, notConverged ++ skippedWarning)

  /**
   * One interior cell's circular what-if value plus its convergence verdict (GH-453). `clock` is
   * the run's pinned clock — axis, fixpoint and source evaluations share one volatile generation.
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
    members: List[(QualifiedRef, Int, String)]
  ): Option[(CellValue, Boolean)] =
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
          // (2) overlay the axis values into the input cells on the PREPARED temp sheets
          val overlaid = axisValues.foldLeft(prepared) { case (sheets, (inputRef, v)) =>
            sheets.updated(tableIdx, sheets(tableIdx).put(inputRef, v))
          }
          // (3) fixpoint the relevant cycle members under this axis combination
          val (results, converged, _) =
            WorkbookEvaluator.jacobiFixpoint(wb, overlaid, members, iterative, clock, None)
          // (4) fold member values in as plain values, then evaluate the source formula
          val folded = members.foldLeft(Option(overlaid)) { case (accOpt, (q, idx, _)) =>
            accOpt.flatMap { sheets =>
              results.get(q) match
                case Some(Right(v)) => Some(sheets.updated(idx, sheets(idx).put(q.ref, v)))
                case _ => None // tolerant: a failing member leaves this cell untouched
            }
          }
          folded.flatMap { sheets =>
            SheetEvaluator
              .evaluateFormula(sheets(tableIdx))(
                expression,
                clock,
                Some(wb.copy(sheets = sheets)),
                None
              )
              .toOption
              .map(value => (value, converged))
          }
        }
      case _ =>
        // Empty or non-formula source cell: Excel caches 0 (fixture-verified bytes).
        Some((CellValue.Number(BigDecimal(0)), true))

  /** One interior cell's what-if value; `None` leaves the cell untouched (tolerant). */
  private def computeCell(
    wb: Workbook,
    sheet: Sheet,
    kind: FormulaKind.DataTable,
    cellRef: ARef,
    input1: ARef,
    input2: Option[ARef],
    clock: Clock
  ): Option[CellValue] =
    val (sourceRef, overlayRefs) = whatIfGeometry(kind, cellRef, input1, input2)
    sheet.cells.get(sourceRef).map(_.value) match
      case Some(CellValue.Formula(expression, _, _)) =>
        val overlaid = overlayRefs.foldLeft(Option(sheet)) { case (acc, (inputRef, axisRef)) =>
          acc.flatMap { s =>
            SheetEvaluator
              .evaluateCell(sheet)(axisRef, clock, Some(wb))
              .toOption
              .map(axisValue => s.put(inputRef, axisValue))
          }
        }
        overlaid.flatMap { s =>
          SheetEvaluator
            .evaluateFormula(s)(expression, clock, Some(wb.put(s)), None)
            .toOption
        }
      case _ =>
        // Empty or non-formula source cell: Excel caches 0 (fixture-verified bytes).
        Some(CellValue.Number(BigDecimal(0)))

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

  /** Cycle members rendered `'Sheet'!A1`, sorted, capped at 8 (GH-453 warning payload). */
  private def renderCycle(core: Set[QualifiedRef]): Vector[String] =
    core.toVector
      .sortBy(q => (q.sheet.value, q.ref.toA1))
      .take(8)
      .map(q => s"${SheetName.quoteForFormula(q.sheet.value)}!${q.ref.toA1}")
