package com.tjclp.xl.formula.eval

import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.functions.{FunctionSpec, FunctionSpecs}
import com.tjclp.xl.formula.graph.DependencyGraph
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.formula.printer.FormulaPrinter
import com.tjclp.xl.formula.parser.{FormulaParser, ParseError}
import com.tjclp.xl.formula.{Clock, Rng}

import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cells.{ArrayMode, Cell, CellError, CellValue, FormulaKind}
import com.tjclp.xl.codec.CodecError
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.patch.Patch
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.numfmt.NumFmt
import com.tjclp.xl.workbooks.Workbook
import java.time.{LocalDate, LocalDateTime}

/**
 * Extension methods for evaluating formulas in Excel sheets.
 *
 * Provides high-level API for formula evaluation:
 *   - Parse formula strings and evaluate against sheet
 *   - Evaluate formula cells (CellValue.Formula)
 *   - Bulk evaluation of all formulas in sheet
 *
 * Design principles:
 *   - Pure functional (no mutations, no side effects)
 *   - Total error handling (XLResult[A] = Either[XLError, A])
 *   - Clock parameter for deterministic date/time functions
 *   - Type conversion: TExpr results → CellValue
 *
 * Example:
 * {{{
 * import com.tjclp.xl.formula.{*, given}
 *
 * // Evaluate formula string
 * sheet.evaluateFormula("=SUM(A1:A10)") // XLResult[CellValue]
 *
 * // Evaluate cell with formula
 * sheet.evaluateCell(ref"B1") // XLResult[CellValue]
 *
 * // Evaluate all formulas: Workbook(sheet).recalculate()
 * }}}
 */
object SheetEvaluator:
  /**
   * Evaluate `formula` as the plain cell at `at`, without changing the input workbook. Places an
   * uncached formula, evaluates its value-read precedent closure, and checks both static cycles and
   * run-time reads of the placement (even under IFERROR). Overrides apply to this sheet and
   * invalidate their cached readers across the workbook. Overriding `at` is refused. The supplied
   * sheet replaces its namesake in the evaluation copy. Added, changed and removed cell values also
   * invalidate their cached readers. Supply a recalculated workbook for structural edits such as
   * resized tables.
   *
   * Unlike `evaluateFormula(..., currentCell = Some(at))`, this includes closure and cycle
   * orchestration. A failure is [[FormulaEvaluationError.Cycle]], the existing XLError with the
   * complete path, or [[FormulaEvaluationError.Failed]] with the XLError alone: the Left is not an
   * `XLError` so a cycle keeps its typed path, and `.left.map(_.error)` is the `XLResult`. The
   * formula comes first and the cell after, as `xl eval <formula> --at <ref>` reads them and as
   * `evaluateArrayWithClosure(…, formula, origin, …)` does. `clock` and `rng` are what TODAY/NOW
   * and RAND/RANDBETWEEN draw from, in the closure and in the formula, as for
   * `evaluateFormula(formula, clock, rng)`; one `rng` serves the whole evaluation, so a seeded one
   * makes it reproducible. Static closure analysis is conservative over ranges. Cached cross-sheet
   * cycles and indirect record-inspection cycles remain outside this API's guarantees (GH-729),
   * matching CLI eval.
   */
  def evaluateAt(
    workbook: Workbook,
    sheet: Sheet,
    formula: String,
    at: ARef,
    overrides: Map[ARef, CellValue] = Map.empty,
    clock: Clock = Clock.system,
    rng: Rng = Rng.system
  ): Either[FormulaEvaluationError, CellValue] =
    FormulaEvaluation.scalar(workbook, sheet, formula, overrides, Some(at), clock, rng)

  /**
   * Positionless dynamic-array evaluation (top-left value) after the same value-read closure as
   * [[evaluateAt]]. No cell is placed or guarded. Unrelated cycles are not visited.
   */
  def evaluateWithClosure(
    workbook: Workbook,
    sheet: Sheet,
    formula: String,
    overrides: Map[ARef, CellValue] = Map.empty,
    clock: Clock = Clock.system,
    rng: Rng = Rng.system
  ): Either[FormulaEvaluationError, CellValue] =
    FormulaEvaluation.scalar(workbook, sheet, formula, overrides, None, clock, rng)

  /**
   * Evaluate the same positionless closure, returning the spilled sheet and range. `origin` anchors
   * the output; it does not place a plain formula or enable implicit intersection. Spill tracking
   * is shared between the precedent fold and the final array evaluation.
   */
  def evaluateArrayWithClosure(
    workbook: Workbook,
    sheet: Sheet,
    formula: String,
    origin: ARef,
    overrides: Map[ARef, CellValue] = Map.empty,
    clock: Clock = Clock.system,
    rng: Rng = Rng.system
  ): Either[FormulaEvaluationError, (Sheet, CellRange)] =
    FormulaEvaluation.array(workbook, sheet, formula, origin, overrides, clock, rng)

  /**
   * #692: a formula cell's computed value and, when that value is an Excel error the evaluator
   * raised with a diagnostic, the reason — what [[com.tjclp.xl.formula.eval.RecalcResult]] names
   * the cell with.
   */
  private[formula] final case class CellOutcome(value: CellValue, reason: Option[String])

  extension (sheet: Sheet)
    /**
     * Evaluate formula string against this sheet.
     *
     * Parses formula, evaluates against sheet, converts result to CellValue.
     *
     * @param formula
     *   Excel formula string (with or without leading =)
     * @param clock
     *   Clock for date/time functions (defaults to system clock)
     * @param workbook
     *   Pass `Some(wb)` iff the formula references other sheets (`='Other'!A1`); intra-sheet
     *   formulas don't need it. Or use `wb.evaluateFormula(formula, onSheet)` which wires the
     *   context automatically.
     * @param currentCell
     *   `Some(cell)`: evaluate as the plain (legacy) formula stored at that cell, as Excel 365
     *   evaluates a plain `<f>` — a multi-cell reference in a value position is implicitly
     *   intersected with the cell (`=A1:A10*2` at row 5 reads A5). `None` (the default): evaluate
     *   as the dynamic-array formula typed into a new Excel 365 cell, returning its top-left value
     *   (`=SUM(A1:A10*B1:B10)` is the array sum). `sheet.evaluateCell(ref)` evaluates a stored cell
     *   by its own position and kind. This low-level method supplies position only; use
     *   `SheetEvaluator.evaluateAt` for placement, precedent evaluation and cycle protection.
     * @return
     *   Either XLError or evaluated CellValue
     *
     * Example:
     * {{{
     * sheet.evaluateFormula("=A1+B2") // Right(CellValue.Number(15))
     * sheet.evaluateFormula("=SUM(A1:A10)") // Right(CellValue.Number(55))
     * sheet.evaluateFormula("=TODAY()") // Right(CellValue.DateTime(...))
     * }}}
     */
    def evaluateFormula(
      formula: String,
      clock: Clock = Clock.system,
      workbook: Option[Workbook] = None,
      currentCell: Option[ARef] = None
    ): XLResult[CellValue] =
      evaluateFormulaWith(sheet, formula, Evaluator.instance, clock, workbook, currentCell)

    /**
     * Evaluate a formula with an explicit randomness source (GH-115) — deterministic
     * RAND/RANDBETWEEN via `Rng.seeded(seed)`.
     *
     * Note: explicit overloads instead of a default rng parameter — extension methods with default
     * arguments crash the compiler when merged through the formulaExports wildcard export (see the
     * DependentRecalculation note in exports.scala).
     */
    @annotation.targetName("evaluateFormulaWithRng")
    def evaluateFormula(formula: String, clock: Clock, rng: Rng): XLResult[CellValue] =
      evaluateFormulaWith(sheet, formula, Evaluator.instance(rng), clock, None, None)

    /** Evaluate a formula with an explicit randomness source and workbook context (GH-115). */
    @annotation.targetName("evaluateFormulaWithRngWorkbook")
    def evaluateFormula(
      formula: String,
      clock: Clock,
      rng: Rng,
      workbook: Option[Workbook]
    ): XLResult[CellValue] =
      evaluateFormulaWith(sheet, formula, Evaluator.instance(rng), clock, workbook, None)

    /**
     * Evaluate cell at ref (if it contains formula).
     *
     * If cell contains CellValue.Formula, parses and evaluates it against the current sheet state.
     * Formula dependencies that are themselves formulas will be referenced as-is (unevaluated).
     *
     * For formulas with dependencies on other formulas, recalculate the workbook
     * (`Workbook(sheet).recalculate()`), which evaluates every formula in dependency order.
     *
     * @param ref
     *   Cell reference to evaluate
     * @param clock
     *   Clock for date/time functions (defaults to system clock)
     * @return
     *   Either XLError or evaluated CellValue
     *
     * Example:
     * {{{
     * // A1 contains "=B1+C1" where B1=10, C1=20
     * sheet.evaluateCell(ref"A1") // Right(CellValue.Number(30))
     *
     * // D1 contains plain number
     * sheet.evaluateCell(ref"D1") // Right(CellValue.Number(42)) - unchanged
     * }}}
     */
    def evaluateCell(
      ref: ARef,
      clock: Clock = Clock.system,
      workbook: Option[Workbook] = None
    ): XLResult[CellValue] =
      evaluateCellWithEvaluator(sheet, ref, Evaluator.instance, clock, workbook)

    /** Evaluate a formula cell with an explicit randomness source (GH-115). */
    @annotation.targetName("evaluateCellWithRng")
    def evaluateCell(ref: ARef, clock: Clock, rng: Rng): XLResult[CellValue] =
      evaluateCell(ref, clock, rng, None)

    /** Evaluate a formula cell with explicit randomness and workbook context (GH-115). */
    @annotation.targetName("evaluateCellWithRngWorkbook")
    def evaluateCell(
      ref: ARef,
      clock: Clock,
      rng: Rng,
      workbook: Option[Workbook]
    ): XLResult[CellValue] =
      evaluateCellWithEvaluator(sheet, ref, Evaluator.instance(rng), clock, workbook)

    /**
     * Evaluate the formulas of a range cell by cell — the total counterpart of
     * [[evaluateForRange]], and what `xl view --eval` renders.
     *
     * The same formulas evaluate (the range's formulas plus their transitive formula precedents, on
     * this sheet and, through `workbook`, on every other (#891: other sheets' caches are no longer
     * read as they stand); every formula of the sheet when it has dynamic references), in
     * dependency order, but a formula that cannot evaluate no longer fails the range: it is listed
     * in [[RangeEvalResult.failures]] and the formulas depending on it are listed in
     * [[RangeEvalResult.blocked]], never evaluated (no stale cache is mixed into a live value). A
     * cycle fails its members — and blocks their dependents — only when it lies inside that
     * closure; a cycle elsewhere on the sheet no longer affects the range. Never throws.
     *
     * Note: explicit overloads instead of default parameters — extension methods with default
     * arguments crash the compiler when merged through the formulaExports wildcard export.
     *
     * @param range
     *   the cell range whose formula values to report
     * @param clock
     *   Clock for date/time functions
     * @param workbook
     *   workbook context for cross-sheet references
     */
    def evaluateForRangePerCell(
      range: CellRange,
      clock: Clock,
      workbook: Option[Workbook]
    ): RangeEvalResult =
      evaluateForRangesPerCell(sheet, Vector(range), clock, workbook)

    /** [[evaluateForRangePerCell]] with the system clock and no workbook context. */
    @annotation.targetName("evaluateForRangePerCellDefault")
    def evaluateForRangePerCell(range: CellRange): RangeEvalResult =
      evaluateForRangesPerCell(sheet, Vector(range), Clock.system, None)

    /**
     * Evaluate an array formula and spill results into adjacent cells.
     *
     * Array formulas like TRANSPOSE return multiple values that "spill" into adjacent cells
     * starting from the origin cell. This method:
     *   1. Parses and evaluates the formula
     *   2. If result is ArrayResult, applies PutArray patch to spill values
     *   3. If result is single value, puts it at origin only
     *   4. Returns the updated sheet and the range of cells affected
     *
     * #750: a result that would spill past XFD1048576 (Excel's `#SPILL!`) or holds more than
     * [[ArrayResult.MaxCells]] cells is refused with a FormulaError, checked from its shape before
     * a cell is built.
     *
     * @param formula
     *   Excel formula string (with or without leading =)
     * @param originRef
     *   The cell where the array formula is entered (spill starts here)
     * @param clock
     *   Clock for date/time functions (defaults to system clock)
     * @param workbook
     *   Optional workbook context for cross-sheet formula references
     * @return
     *   Either XLError or tuple of (updated sheet, affected range)
     *
     * Example:
     * {{{
     * // TRANSPOSE a 2x3 range into a 3x2 result
     * sheet.evaluateArrayFormula("=TRANSPOSE(A1:C2)", ref"E1")
     * // Returns (updatedSheet, CellRange(E1:F3)) with transposed values
     * }}}
     */
    def evaluateArrayFormula(
      formula: String,
      originRef: ARef,
      clock: Clock = Clock.system,
      workbook: Option[Workbook] = None
    ): XLResult[(Sheet, CellRange)] =
      evaluateArrayFormulaImpl(sheet, formula, originRef, Evaluator.arrayInstance, clock, workbook)

    /** Array formula evaluation with an explicit randomness source (GH-115). */
    @annotation.targetName("evaluateArrayFormulaWithRng")
    def evaluateArrayFormula(
      formula: String,
      originRef: ARef,
      clock: Clock,
      rng: Rng
    ): XLResult[(Sheet, CellRange)] =
      evaluateArrayFormulaImpl(sheet, formula, originRef, Evaluator.arrayInstance(rng), clock, None)

    /**
     * Put a formula at ref, inheriting the number format of its referenced cells (GH-184).
     *
     * Excel applies the referenced cells' format automatically when a formula is entered into a
     * General cell (`=B2-B3` over currency cells displays as `$400,000.00`); plain `put` keeps
     * formula cells at General. This opt-in variant parses the formula, infers a format via
     * [[FormulaFormatting.inferFormatFromReferences]], and — matching Excel — applies it only when
     * the target cell's current format is General (an explicit target format is preserved, as are
     * its other style properties).
     *
     * Note: explicit overloads instead of a default workbook parameter — extension methods with
     * default arguments crash the compiler when merged through the formulaExports wildcard export
     * (see the DependentRecalculation note in exports.scala).
     *
     * @param ref
     *   Target cell for the formula
     * @param formula
     *   Excel formula string (with or without leading =)
     * @return
     *   Left on parse failure; Right(updated sheet) otherwise (inference itself never fails —
     *   unresolvable or unformatted references simply contribute nothing)
     */
    def putFormulaInheriting(ref: ARef, formula: String): XLResult[Sheet] =
      putFormulaInheritingImpl(sheet, ref, formula, None)

    /**
     * Workbook-aware [[putFormulaInheriting]]: cross-sheet references (`=Data!B2`) resolve their
     * formats against the given workbook.
     */
    @annotation.targetName("putFormulaInheritingWorkbook")
    def putFormulaInheriting(ref: ARef, formula: String, workbook: Workbook): XLResult[Sheet] =
      putFormulaInheritingImpl(sheet, ref, formula, Some(workbook))

  // ========== Implementation shared by the public overloads ==========

  /**
   * What a range evaluation evaluates: the ranges' formula cells, the sheet's dependency graph, the
   * ranges' static closure (their formulas plus their transitive formula precedents), the target
   * formulas (the closure, or every formula when the sheet has dynamic references) and the sheet's
   * dynamic cells.
   */
  private def isFormula(value: CellValue): Boolean = value match
    case _: CellValue.Formula => true
    case _ => false

  /**
   * The per-cell evaluation of the formulas in `ranges` behind [[evaluateForRangePerCell]] (one
   * range) and `view --eval`'s pictures ([[LiveRender]]: the window and the cells its conditional
   * formatting reads). The targets are the range's formulas and every formula they read, directly
   * or transitively, on `sheet` and on every other sheet of the workbook (#891: other sheets'
   * caches used to be read as they stood, so a window showed values a recalculation would not).
   * Cycle members inside the targets fail as circular and block their dependents; the rest
   * evaluates in the canonical order against the threaded sheets. A failure blocks its transitive
   * dependents, and every failed or blocked cell loses its cache on its threaded sheet, so a
   * dynamic reader (INDIRECT, invisible to the static graph) re-derives it and fails too instead of
   * reading a stale value.
   *
   * Dynamic references: when `sheet` holds one, every formula of `sheet` is a target (GH-274), and
   * so is every formula of another sheet that has a dynamic reader among the targets, each with its
   * precedents. What such a reader reads on a sheet the targets never reach is read as it stands.
   *
   * Reported: the values of the range's own formulas; the failures of the range's static closure,
   * `sheet`'s first and in row-major order, then other sheets' in grid order; and the blocked cells
   * of that closure on `sheet`. With dynamic references more formulas are targets, but one that
   * feeds the range only through a dynamic read makes that reader — in the closure — fail on its
   * own, so an unrelated failure elsewhere stays silent. A cycle's members are listed when they lie
   * on `sheet`; a reader of a cycle that lies wholly on another sheet is blocked. In a book whose
   * calcPr iterates, cycle members keep their caches and are not failures; when one of them holds
   * no cache, the result is projected from an iterative `recalculate` of the book.
   *
   * #891: the graph questions go to the workbook's recalculation graph (cross-sheet cycles, defined
   * names, #504's precise lookups); the sheet-level graph and its hash-map Tarjan and Kahn are
   * gone. The evaluation is a recalculation pass over the targets: it reads the graph build's
   * parses and the pass's caches (column slices, the criteria index, aggregates), so a window whose
   * closure is a whole model costs what recalculating that model costs, not a scan of every
   * criteria range per formula.
   */
  private[eval] def evaluateForRangesPerCell(
    sheet: Sheet,
    ranges: Vector[CellRange],
    clock: Clock,
    workbook: Option[Workbook]
  ): RangeEvalResult =
    evaluateForRangesPerCellWithMemo(sheet, ranges, clock, workbook, new Evaluator.AggregateMemo)

  /**
   * [[evaluateForRangesPerCell]] with the pass's aggregate memo supplied by the caller: the parity
   * laws run one window with the per-pass caches on and off, and read the memo's counters.
   */
  private[formula] def evaluateForRangesPerCellWithMemo(
    sheet: Sheet,
    ranges: Vector[CellRange],
    clock: Clock,
    workbook: Option[Workbook],
    memo: Evaluator.AggregateMemo
  ): RangeEvalResult =
    // a small range is a handful of point lookups: many of them (a formula rule's reads at every
    // cell of a render window) never make the scan below quadratic
    val (small, large) = ranges.partition(r => r.width.toLong * r.height <= 256L)
    val points = small.iterator.flatMap(_.cells).toSet
    def covered(ref: ARef): Boolean = points.contains(ref) || large.exists(_.contains(ref))
    val book = workbook.fold(Workbook(sheet))(_.put(sheet))
    val counters = memo.counters
    lazy val graph =
      counters.time(RecalcCounters.Phase.GraphBuild)(RecalcGraph.build(book, counters))
    def at(ref: ARef): QualifiedRef = QualifiedRef(sheet.name, ref)
    def formulasOf(target: Sheet): Iterator[QualifiedRef] =
      target.cells.iterator
        .collect {
          case cell if isFormula(cell.value) => QualifiedRef(target.name, cell.ref)
        }
        .filter(graph.isFormula)
    val rangeFormulaCells = sheet.cells.iterator.collect {
      case cell @ Cell(ref, _, _, _)
          if isFormula(cell.value) && covered(ref) && graph.isFormula(at(ref)) =>
        ref
    }.toSet
    if rangeFormulaCells.isEmpty then RangeEvalResult(Map.empty, Vector.empty, Vector.empty)
    else
      val window = rangeFormulaCells.map(at)
      val sheetsByName = book.sheets.iterator.map(s => s.name -> s).toMap
      val core = graph.cyclicCore
      // #891 (Excel parity): a book whose calcPr iterates holds converged values in its cycles, and
      // Excel shows them. Its cycle members are neither failures nor re-evaluated here: they keep
      // their caches and everything downstream evaluates off them. Otherwise a cycle fails its
      // members and blocks what reads them.
      val iterating = book.metadata.calcPr.exists(_.iterativeCalculation)
      // That holds only if every member holds a cache. A member without (a library-authored or
      // `--no-recalc` book; Excel never writes one) would be evaluated on demand by its first
      // reader, ahead of its own precedents' slots and off their stale caches, and an aggregate
      // computed that early would be kept for the pass. Such a book is iterated: the window shows
      // what `recalculate` computes, at the cost of one recalculation.
      def holdsNoCache(q: QualifiedRef): Boolean =
        sheetsByName.get(q.sheet).flatMap(_.cells.get(q.ref)).map(_.value) match
          case Some(CellValue.Formula(_, None, _)) => true
          case _ => false
      val unsettled =
        book.metadata.calcPr.filter(_.iterativeCalculation && core.exists(holdsNoCache))
      // The cached members of an iterating book are pinned: sources of the values the file holds
      // and nothing else. No walk passes through one (what feeds a cycle is not a precedent of
      // what reads it, and a failure behind a cycle blocks nothing in front of it), none is
      // stripped, and one that reads dynamically is not a dynamic reader of this pass.
      val pinned = if iterating && unsettled.isEmpty then core else Set.empty[QualifiedRef]
      // the static closure, on this sheet and every other
      val closure = window ++ graph.precedents(window, pinned)
      val dynamic = DependencyGraph.dynamicCells(book, graph.parses) -- pinned
      // GH-274: a dynamic reader's targets are invisible to the static walk, so correctness needs
      // every formula of its sheet computed in order (correctness over narrowness; results still
      // report only the range). `sheet` widens when it holds any dynamic reference, as it always
      // has; another sheet when one of its dynamic readers is a target. Each widening brings the
      // new formulas' precedents, which may reach a dynamic reader on a further sheet.
      @annotation.tailrec
      def widen(targets: Set[QualifiedRef], widened: Set[SheetName]): Set[QualifiedRef] =
        val wanting = (dynamic.iterator.collect {
          case q if q.sheet == sheet.name || targets.contains(q) => q.sheet
        }.toSet) -- widened
        if wanting.isEmpty then targets
        else
          val added = wanting.iterator.flatMap(sheetsByName.get).flatMap(formulasOf).toSet
          widen(targets ++ added ++ graph.precedents(added, pinned), widened ++ wanting)
      val targets = if dynamic.isEmpty then closure else widen(closure, Set.empty)
      def dependentsOf(refs: Set[QualifiedRef]): Set[QualifiedRef] =
        (graph.dependents(refs -- pinned, pinned) & targets) -- refs -- pinned
      def formulaText(q: QualifiedRef): String =
        sheetsByName.get(q.sheet).flatMap(_.cells.get(q.ref)).map(_.value) match
          case Some(CellValue.Formula(expression, _, _)) => expression
          case _ => q.ref.toA1
      def failure(q: QualifiedRef, reason: String): CellEvalError =
        CellEvalError(q.sheet, q.ref, XLError.FormulaError(formulaText(q), reason))

      def ordered: RangeEvalResult =
        val cyclic = if iterating then Set.empty[QualifiedRef] else core & targets
        // every target downstream of any cycle of the book: the canonical order below leaves them
        // all out
        val coreBlocked =
          if iterating then Set.empty[QualifiedRef] else graph.dependents(core, Set.empty)
        val cycleBlocked = (coreBlocked & targets) -- cyclic
        val live = targets -- cyclic -- cycleBlocked -- (if iterating then core else Set.empty)

        val positions = book.sheets.iterator.zipWithIndex.map((s, i) => s.name -> i).toMap
        def stripped(sheets: Vector[Sheet], refs: Set[QualifiedRef]): Vector[Sheet] =
          refs.groupMap(_.sheet)(_.ref).foldLeft(sheets) { case (acc, (name, cells)) =>
            positions.get(name).fold(acc)(i => acc.updated(i, stripFormulaCaches(acc(i), cells)))
          }

        type State = (Vector[Sheet], Map[ARef, CellValue], Vector[CellEvalError], Set[QualifiedRef])
        val start: State =
          (
            stripped(book.sheets, cyclic ++ cycleBlocked),
            Map.empty,
            // a cycle's members are listed on the range's own sheet (as before #891)
            cyclic.toVector.filter(_.sheet == sheet.name).map(failure(_, "Circular reference")),
            cycleBlocked
          )
        // GH-695: record spill anchors' arrays only for a book that has `x#` readers
        if readsSpills(sheet, workbook) then memo.trackSpills()
        // #790: evaluation reuses the build's parse of each formula
        memo.passCache.foreach(_.useParses(graph.parses))
        val (_, values, failures, blocked) = graph.order(core ++ coreBlocked) match
          // Unreachable: the cyclic core and its dependents left the order. Stay total anyway.
          case scala.util.Left(unordered) =>
            val (sheets, reported, failed, blockedSoFar) = start
            val residual = live.toVector.map(
              failure(_, s"Unresolvable order: ${unordered.remaining.size} formulas left")
            )
            (sheets, reported, failed ++ residual, blockedSoFar)
          case scala.util.Right(sorted) =>
            val (startSheets, reported, failed, blockedSoFar) = start
            val ordered = sorted.order.filter(live)
            // GH-274: dynamic cells and their static dependents evaluate last, caches stripped
            val bucket =
              if dynamic.isEmpty then Set.empty[QualifiedRef]
              else (dynamic & live) ++ dependentsOf(dynamic)
            val evalOrder = ordered.filterNot(bucket) ++ ordered.filter(bucket)
            val initial = stripped(startSheets, bucket)
            // #767: one slot per target in evaluation order. Formulas outside the targets keep the
            // caches the file holds, which this pass never changes; the stripped cells (the dynamic
            // bucket, cycle members and the cells behind them) refuse every slice over them. A cell
            // a failure strips later holds no cache when a slice over it is built, and is refused.
            val slots = evalOrder.iterator.zipWithIndex.toMap
            memo.passCache.foreach(
              _.begin(Readiness.build(slots, bucket ++ cyclic ++ cycleBlocked))
            )
            def evaluatorFor(q: QualifiedRef): Evaluator =
              Evaluator.recalculationInstance(
                Rng.system,
                memo,
                WorkbookEvaluator.callerPosition(slots, bucket, q),
                Some((q.sheet, q.ref))
              )
            evalOrder.foldLeft[State]((initial, reported, failed, blockedSoFar)) {
              case (state @ (_, _, _, blockedNow), q) if blockedNow(q) => state
              case (state @ (sheets, results, failedNow, blockedNow), q) =>
                positions.get(q.sheet) match
                  case None => state
                  case Some(position) =>
                    val evaluated = counters.time(RecalcCounters.Phase.Eval) {
                      evaluateCellWithEvaluator(
                        sheets(position),
                        q.ref,
                        evaluatorFor(q),
                        clock,
                        workbook.map(_ => book.copy(sheets = sheets))
                      )
                    }
                    evaluated match
                      case scala.util.Right(value) =>
                        if threadChangesStructure(sheets(position), q.ref) then
                          memo.passCache.foreach(_.structureChanged(q.sheet))
                        val next =
                          if q.sheet == sheet.name && rangeFormulaCells.contains(q.ref) then
                            results + (q.ref -> value)
                          else results
                        val threaded = threadComputed(sheets(position), q.ref, value)
                        (sheets.updated(position, threaded), next, failedNow, blockedNow)
                      case scala.util.Left(error) =>
                        val newlyBlocked = dependentsOf(Set(q)) -- blockedNow
                        (
                          stripped(sheets, newlyBlocked + q),
                          results,
                          failedNow :+ CellEvalError(q.sheet, q.ref, error),
                          blockedNow ++ newlyBlocked
                        )
            }
        RangeEvalResult(
          values,
          sortedFailures(
            failures.filter(failed => closure(QualifiedRef(failed.sheet, failed.ref)))
          ),
          blocked.iterator
            .filter(q => q.sheet == sheet.name && closure(q))
            .map(_.ref)
            .toVector
            .sortBy(rowMajor)
        )
      def iterated(calc: IterativeCalc): RangeEvalResult =
        val full = WorkbookEvaluator.recalculate(book)(clock, calc)
        val errored = full.errors.iterator.map(e => QualifiedRef(e.sheet, e.ref)).toSet
        // a cell behind another's failure is blocked, as in the ordered pass; a cycle's members
        // fail together and are each a failure
        val blockedCells = (errored & closure).filter { q =>
          val upstream = graph.precedents(Set(q), Set.empty) & errored
          (if core(q) then upstream -- core else upstream - q).nonEmpty
        }
        val computed = full.evaluated.getOrElse(sheet.name, Map.empty)
        val after = full.workbook.sheets.find(_.name == sheet.name)
        def shown(ref: ARef): Option[CellValue] =
          computed
            .get(ref)
            .orElse(after.flatMap(_.cells.get(ref)).map(_.value).collect {
              case CellValue.Formula(_, Some(cached), _) => cached
            })
        RangeEvalResult(
          rangeFormulaCells.iterator
            .filterNot(ref => errored(at(ref)))
            .flatMap(ref => shown(ref).map(ref -> _))
            .toMap,
          sortedFailures(
            full.errors.filter { e =>
              val q = QualifiedRef(e.sheet, e.ref)
              closure(q) && !blockedCells(q)
            }
          ),
          blockedCells.iterator.filter(_.sheet == sheet.name).map(_.ref).toVector.sortBy(rowMajor),
          // #917: the cycles the range depends on that did not converge, as recalc reports them.
          // What a dynamic reader reads is not in the static closure, so with one in the closure
          // every unconverged cycle is reported: saying too much beats a silent stale value. One
          // the range does not reach (the widened `targets` hold every one of this sheet) cannot
          // change what the range shows, and reports nothing.
          if (closure & dynamic).nonEmpty then full.unconverged
          else
            full.unconverged.filter(
              _.members.exists((name, ref) => closure(QualifiedRef(name, ref)))
            )
        )
      def rowMajor(ref: ARef): (Int, Int) = (ref.row.index0, ref.col.index0)
      // `sheet`'s failures first and in row-major order, then other sheets' in grid order
      def sortedFailures(failures: Vector[CellEvalError]): Vector[CellEvalError] =
        val (here, elsewhere) = failures.partition(_.sheet == sheet.name)
        here.sortBy(failed => rowMajor(failed.ref)) ++
          elsewhere.sortBy(failed => QualifiedRef.gridKey(QualifiedRef(failed.sheet, failed.ref)))
      unsettled.map(IterativeCalc.fromCalcPr).fold(ordered)(iterated)

  /**
   * Internal cell-evaluation boundary for WorkbookEvaluator.
   *
   * Supplying an evaluator lets one recalculation generation carry a narrowly scoped aggregate memo
   * through every cell and parallel worker. Public overloads call the same boundary with a normal
   * uncached evaluator, so their semantics remain unchanged.
   */
  private[formula] def evaluateCellWithEvaluator(
    sheet: Sheet,
    ref: ARef,
    evaluator: => Evaluator,
    clock: Clock,
    workbook: Option[Workbook]
  ): XLResult[CellValue] =
    evaluateCellOutcome(sheet, ref, evaluator, clock, workbook).map(_.value)

  /**
   * #692: [[evaluateCellWithEvaluator]] with the evaluator's reason kept beside an Excel error
   * value — the diagnostic the GH-344 promotion would otherwise drop, for the recalc summary.
   */
  private[formula] def evaluateCellOutcome(
    sheet: Sheet,
    ref: ARef,
    evaluator: => Evaluator,
    clock: Clock,
    workbook: Option[Workbook]
  ): XLResult[CellOutcome] =
    evaluateCellOutcomeParsing(sheet, ref, evaluator, clock, workbook)(
      parseCounted(_, _, sheet.name, Some(ref))
    )

  /**
   * #891: [[evaluateCellOutcome]] with the cell's parse supplied by the caller: `parse` receives
   * the formula text and the cell's evaluator, at most once. Data-table seeding evaluates the same
   * cone once per interior cell and hands in the parse its graph build already made.
   */
  private[eval] def evaluateCellOutcomeParsing(
    sheet: Sheet,
    ref: ARef,
    evaluator: => Evaluator,
    clock: Clock,
    workbook: Option[Workbook]
  )(parse: (String, Evaluator) => XLResult[TExpr[?]]): XLResult[CellOutcome] =
    sheet.cells.valueAt(ref) match
      case value @ CellValue.Formula(expr, _, kind) =>
        // #790: one parse of the cell serves both the pin check and the evaluation
        lazy val cellEvaluator = evaluator
        lazy val parsedCell = parse(expr, cellEvaluator)
        pinnedCache(value, _ => parsedCell) match
          // GH-353/GH-430: pinned-cache semantics — the Excel-written cache IS the value
          case Some(cached) => scala.util.Right(CellOutcome(cached, None))
          case None =>
            // an ArrayFormula record (CSE or dynamic-array anchor) evaluates as an array and the
            // anchor holds element (0,0); a plain formula evaluates as a scalar cell
            kind match
              // GH-935 review: merged cells inside a dynamic array's spill range block it, and
              // Excel shows #SPILL! at the anchor (Microsoft documents merged cells as a cause)
              case FormulaKind.ArrayFormula(extent, _, _, _: ArrayMode.Dynamic)
                  if sheet.mergedRanges.exists(_.intersects(extent)) =>
                scala.util.Right(
                  CellOutcome(CellValue.Error(CellError.Spill), Some("spill range is merged"))
                )
              case _: FormulaKind.ArrayFormula =>
                val arrayEvaluator = cellEvaluator.withArrayResults
                parsedCell.flatMap { parsed =>
                  val raw = arrayEvaluator.eval(parsed, sheet, clock, workbook, Some(ref))
                  val result = cellOutcome(expr, raw)
                  // GH-695: the raw result goes to the generation memo for this generation's `x#`
                  // readers, which read it as they read an anchor they evaluate themselves; the
                  // threaded sheet keeps only element (0,0) as the anchor's cache
                  if result.isRight then
                    arrayEvaluator.generationMemo.foreach(_.recordSpill(sheet.name, ref, raw))
                  result
                }
              // Pass the current cell ref for ROW()/COLUMN() without arguments
              case _ =>
                parsedCell.flatMap(parsed =>
                  cellOutcome(expr, cellEvaluator.eval(parsed, sheet, clock, workbook, Some(ref)))
                )
      case other => scala.util.Right(CellOutcome(other, None))

  private def evaluateArrayFormulaImpl(
    sheet: Sheet,
    formula: String,
    originRef: ARef,
    evaluator: Evaluator,
    clock: Clock,
    workbook: Option[Workbook]
  ): XLResult[(Sheet, CellRange)] =
    for
      result <- evaluateArrayRaw(sheet, formula, originRef, evaluator, clock, workbook)

      // Handle array vs scalar result
      updated <- result match
        case ar: ArrayResult if ar.isEmpty =>
          val patch = Patch.PutArray(originRef, Vector.empty)
          val endRef = originRef.shift(ar.cols - 1, ar.rows - 1)
          scala.util.Right((Patch.applyPatch(sheet, patch), CellRange(originRef, endRef)))
        case ar: ArrayResult =>
          // #750: Excel's #SPILL! for a result that runs past XFD1048576, and xl's write bound,
          // both checked from the shape before a cell is built
          def blocked(blocker: DynamicArrayAuthoring.SpillBlocker) =
            val full = if formula.startsWith("=") then formula else s"=$formula"
            scala.util.Left(
              DynamicArrayAuthoring.SpillBlocked(full, originRef, None, blocker).toXLError
            )
          DynamicArrayAuthoring.spillExtent(originRef, ar.rows, ar.cols) match
            case None => blocked(DynamicArrayAuthoring.SpillBlocker.OffGrid(ar.rows, ar.cols))
            case Some(_) if ar.cellCount > ArrayResult.MaxCells =>
              blocked(
                DynamicArrayAuthoring.SpillBlocker.TooLarge(ar.rows, ar.cols, ArrayResult.MaxCells)
              )
            case Some(extent) =>
              ar.toDense("spill") match
                case scala.util.Left(err) =>
                  scala.util.Left(evalErrorToXLError(err, Some(formula)))
                case scala.util.Right(grid) =>
                  val patch = Patch.PutArray(originRef, grid)
                  scala.util.Right((Patch.applyPatch(sheet, patch), extent))
        case other =>
          val cv = EvalResult.toCellValue(other)
          scala.util.Right((sheet.put(originRef, cv), CellRange(originRef, originRef)))
    yield updated

  /**
   * Parse and evaluate `formula` at `originRef` with `evaluator` (array mode for the callers here),
   * without touching the sheet. GH-344: a Left carrying an Excel error VALUE promotes to that error
   * as the result; host failures stay loud Lefts.
   */
  private def evaluateArrayRaw(
    sheet: Sheet,
    formula: String,
    originRef: ARef,
    evaluator: Evaluator,
    clock: Clock,
    workbook: Option[Workbook]
  ): XLResult[Any] =
    parseFormula(formula).flatMap { expr =>
      evaluator.eval(expr, sheet, clock, workbook, Some(originRef)) match
        case scala.util.Right(value) => scala.util.Right(value)
        case scala.util.Left(evalError) =>
          EvalError.toErrorValue(evalError) match
            case Some(code) => scala.util.Right(CellValue.Error(code): Any)
            case None => scala.util.Left(evalErrorToXLError(evalError, Some(formula)))
    }

  /**
   * GH-714: the array value of `formula` entered at `originRef` — what Excel 365 computes for the
   * formula typed into that cell (no implicit intersection) — as the array of values its spill
   * would show, without patching anything. A scalar result is a 1x1 array. Each element is the
   * value a spill cell stores: a cached formula record reads as its cache, and an empty element
   * reads 0 (a formula cell is never blank, the [[cellResult]] rule). #750: the array may be sparse
   * and of any shape; a caller checks its `cellCount` before building its cells.
   */
  private[xl] def evaluateArrayValues(
    sheet: Sheet,
    formula: String,
    originRef: ARef,
    evaluator: Evaluator,
    clock: Clock,
    workbook: Option[Workbook]
  ): XLResult[ArrayResult] =
    evaluateArrayRaw(sheet, formula, originRef, evaluator.withArrayResults, clock, workbook)
      .map {
        case ar: ArrayResult => ar.mapPureLazy(spillValue)
        case other => ArrayResult.single(spillValue(EvalResult.toCellValue(other)))
      }

  /**
   * GH-714: `formula` evaluated as the plain (legacy) cell at `ref` would be — references in value
   * positions intersected with the cell — with an explicit evaluator (randomness source).
   */
  private[xl] def evaluatePlainAt(
    sheet: Sheet,
    formula: String,
    ref: ARef,
    evaluator: Evaluator,
    clock: Clock,
    workbook: Option[Workbook]
  ): XLResult[CellValue] =
    evaluateFormulaWith(sheet, formula, evaluator, clock, workbook, Some(ref))

  /** The value one element of an evaluated array stores in its cell (see [[cellResult]]). */
  private def spillValue(value: CellValue): CellValue =
    effectiveValue(value) match
      case CellValue.Empty => CellValue.Number(BigDecimal(0))
      case other => other

  // ========== Helper Functions ==========

  private def putFormulaInheritingImpl(
    sheet: Sheet,
    ref: ARef,
    formula: String,
    workbook: Option[Workbook]
  ): XLResult[Sheet] =
    // GH-479: the model's one canonical rule (trim, one leading '=' off, trim) is applied first, so
    // a padded `" = B2 - B3 "` parses and is stored exactly as fx / the CLI would store it
    val canonical = CellValue.canonicalFormulaText(formula)
    FormulaParser
      .parse(canonical)
      .left
      .map(parseError =>
        XLError.FormulaError(formula, s"Parse error: ${ParseError.describe(parseError)}")
      )
      .map { expr =>
        val withFormula = sheet.put(ref, CellValue.Formula(canonical))
        val currentStyle =
          withFormula.cells.get(ref).flatMap(_.styleId).flatMap(withFormula.styleRegistry.get)
        // Excel parity: inherit only into a General-formatted target (formats are inferred
        // against the pre-put sheet so a self-reference can't observe the formula being written)
        FormulaFormatting.inferFormatFromReferences(expr, sheet, workbook) match
          case Some(fmt) if currentStyle.forall(_.numFmt == NumFmt.General) =>
            import com.tjclp.xl.sheets.styleSyntax.withCellStyle
            withFormula.withCellStyle(
              ref,
              currentStyle.getOrElse(CellStyle.default).withNumFmt(fmt)
            )
          case _ => withFormula
      }

  /**
   * Shared parse → evaluate → CellValue pipeline, parameterized by evaluator (GH-115: rng).
   *
   * GH-344: THE boundary-promotion site — a Left classified as an Excel error VALUE by
   * [[EvalError.toErrorValue]] becomes `Right(CellValue.Error(code))` here, funneling every
   * `evaluateFormula` overload, `evaluateCell`, `evaluateForRangePerCell`, `wb.evaluateFormula`,
   * `recalculate`, and the CLI. Headline contract: `sheet.evaluateFormula("=1/0")` is
   * `Right(CellValue.Error(Div0))`. Host failures (parse, missing workbook/sheet, unknown function,
   * cycles) stay loud Lefts.
   */
  private def evaluateFormulaWith(
    sheet: Sheet,
    formula: String,
    evaluator: Evaluator,
    clock: Clock,
    workbook: Option[Workbook],
    currentCell: Option[ARef]
  ): XLResult[CellValue] =
    evaluateFormulaOutcome(sheet, formula, evaluator, clock, workbook, currentCell).map(_.value)

  /** [[evaluateFormulaWith]] keeping the reason beside an error value (#692). */
  private def evaluateFormulaOutcome(
    sheet: Sheet,
    formula: String,
    evaluator: Evaluator,
    clock: Clock,
    workbook: Option[Workbook],
    currentCell: Option[ARef]
  ): XLResult[CellOutcome] =
    // A formula with a cell position evaluates as that plain cell would: Excel's legacy formula,
    // references in value positions implicitly intersected. Without one there is no cell to
    // intersect with, so it evaluates as the same formula typed into a new Excel 365 cell would:
    // as an array, showing its top-left value. `@range` still needs the position.
    val positioned = if currentCell.isDefined then evaluator else evaluator.withArrayResults
    parseCounted(formula, evaluator, sheet.name, currentCell).flatMap(expr =>
      cellOutcome(formula, positioned.eval(expr, sheet, clock, workbook, currentCell))
    )

  /**
   * #767: [[parseFormula]] timed as the `Parse` phase of the evaluator's pass. #790: the formula of
   * `cell` on `sheet` reuses the parse the pass's graph build made of it.
   */
  private def parseCounted(
    formula: String,
    evaluator: Evaluator,
    sheet: SheetName,
    cell: Option[ARef]
  ): XLResult[TExpr[?]] =
    parseFormulaCell(formula, sheet, cell, evaluator.generationMemo)

  /**
   * #790: [[parseFormula]] of the formula of `cell` on `sheet`, reusing the parse the graph build
   * of `memo`'s pass made of it; a fresh parse is counted and timed in `memo`'s counters.
   */
  private[eval] def parseFormulaCell(
    formula: String,
    sheet: SheetName,
    cell: Option[ARef],
    memo: Option[Evaluator.AggregateMemo]
  ): XLResult[TExpr[?]] =
    Evaluator.parseFormulaCell(formula, sheet, cell, memo).left.map(parseFailure(formula, _))

  /** The parse half of [[evaluateFormulaWith]]: a parse failure is the `Parse error:` XLError. */
  private[eval] def parseFormula(formula: String): XLResult[TExpr[?]] =
    FormulaParser.parse(formula).left.map(parseFailure(formula, _))

  private def parseFailure(formula: String, parseError: ParseError): XLError =
    XLError.FormulaError(formula, s"Parse error: ${ParseError.describe(parseError)}")

  /**
   * The evaluate half of [[evaluateFormulaWith]] for an ALREADY PARSED formula — GH-537: an
   * iterative fixpoint parses each member once and evaluates its `TExpr` every round. `formulaText`
   * is the source text the diagnostics quote; it must be the text `expr` was parsed from. The
   * GH-344 boundary promotion lives here and nowhere else.
   */
  private[eval] def evaluateParsedWith(
    sheet: Sheet,
    formulaText: String,
    expr: TExpr[?],
    evaluator: Evaluator,
    clock: Clock,
    workbook: Option[Workbook],
    currentCell: Option[ARef]
  ): XLResult[CellValue] =
    cellResult(formulaText, evaluator.eval(expr, sheet, clock, workbook, currentCell))

  /** A cached formula record's value (nested caches unwrapped); any other value as it is. */
  @annotation.tailrec
  private def effectiveValue(value: CellValue): CellValue = value match
    case CellValue.Formula(_, Some(cached), _) => effectiveValue(cached)
    case other => other

  /** A formula's raw evaluation as the cell value it stores (the GH-344 boundary promotion). */
  private def cellResult(formulaText: String, raw: Either[EvalError, Any]): XLResult[CellValue] =
    cellOutcome(formulaText, raw).map(_.value)

  /**
   * [[cellResult]] with the reason an error value was raised with (#692): the context of a Left
   * promoted here — a lookup's miss, a division's operands — which the value alone cannot carry.
   */
  private def cellOutcome(formulaText: String, raw: Either[EvalError, Any]): XLResult[CellOutcome] =
    raw match
      case scala.util.Right(value) =>
        // A formula cell is never blank in Excel: a result that is a reference to an empty cell
        // (INDEX, INDIRECT, OFFSET, CHOOSE, a lookup) reads 0, as `=Z1` already does. Only the
        // cell's final value changes — inside a formula the reference stays blank (ISBLANK, COUNTA).
        // GH-695: a function that returns a referenced cell itself (IFERROR, IFNA, the lookups)
        // can hand back a cached formula record — a threaded spill anchor, or a precedent this
        // pass did not re-evaluate; the cell stores that record's value, never the record.
        effectiveValue(EvalResult.toCellValue(value)) match
          case CellValue.Empty =>
            scala.util.Right(CellOutcome(CellValue.Number(BigDecimal(0)), None))
          case other => scala.util.Right(CellOutcome(other, None))
      case scala.util.Left(evalError) =>
        EvalError.toErrorValue(evalError) match
          case Some(code) =>
            scala.util.Right(CellOutcome(CellValue.Error(code), EvalError.reason(evalError)))
          case None => scala.util.Left(evalErrorToXLError(evalError, Some(formulaText)))

  /**
   * GH-353: Excel closed-workbook semantics — Some(cached) when this is a formula cell whose
   * expression touches an external workbook ([2]Book1!A1) AND Excel wrote a cached value.
   *
   * Such cells are PINNED: the cached value is the only source of truth (the external workbook is
   * not loaded, so re-evaluation can only fail) and must be preserved verbatim. Callers skip
   * evaluation and use the cache directly; dependents then read it exactly as they read any cached
   * precedent. Uncached external formulas return None and fall through to normal evaluation, which
   * yields a clear per-cell error (Evaluator.externalRefUnsupported).
   *
   * The '[' pre-filter keeps the common recalculation path parse-free (same trick as
   * DependencyGraph.dynamicCells).
   *
   * Parse-failure fallback: some external shapes are beyond xl's parser (e.g. an external range in
   * an XLOOKUP array slot, or external defined names `[2]!name`). Re-evaluating such a cell can
   * only produce a parse error, and clearing its cache would destroy the sole source of truth, so a
   * CACHED formula that fails to parse but lexically carries the external-workbook prefix (`[n]…` /
   * `'[n]…`) is pinned too. A false positive (a `[n]`-looking substring in an unparseable
   * non-external formula) merely keeps that cell's existing cache — the same value every earlier xl
   * version reported for it.
   */
  private[formula] def pinnedExternalCache(value: CellValue): Option[CellValue] =
    pinnedExternalCache(value, FormulaParser.parse)

  /**
   * [[pinnedExternalCache]] with the formula text parsed by `parse` (#790: a recalculation passes
   * the parse it evaluates the cell with, so the cell is parsed once). `parse` runs only for a
   * cached formula whose text contains `[`.
   */
  private[formula] def pinnedExternalCache[E](
    value: CellValue,
    parse: String => Either[E, TExpr[?]]
  ): Option[CellValue] =
    value match
      case CellValue.Formula(expr, Some(cached), _) if expr.contains('[') =>
        parse(expr) match
          case scala.util.Right(ast) if TExpr.containsExternalRef(ast) => Some(cached)
          case scala.util.Right(_) => None
          case scala.util.Left(_) if externalPrefixPattern.matcher(expr).find() => Some(cached)
          case scala.util.Left(_) => None
      case _ => None

  /**
   * GH-430: the generalized GH-353 seam. A data-table record's cache is its only truthful value —
   * xl does not evaluate `TABLE(...)` (the record, not the text, is the formula), so evaluation
   * pins the Excel-written cache exactly like closed-workbook externals; an uncached record pins to
   * Empty rather than parse-failing. ArrayFormula kinds are NOT pinned: their text evaluates as an
   * array and the anchor takes element (0,0) (evaluateCellWithEvaluator).
   */
  private[eval] def pinnedCache(value: CellValue): Option[CellValue] =
    pinnedCache(value, FormulaParser.parse)

  /** [[pinnedCache]] with the external check's parse supplied ([[pinnedExternalCache]]). */
  private[eval] def pinnedCache[E](
    value: CellValue,
    parse: String => Either[E, TExpr[?]]
  ): Option[CellValue] =
    value match
      case CellValue.Formula(_, cached, _: FormulaKind.DataTable) =>
        Some(cached.getOrElse(CellValue.Empty))
      case _ => pinnedExternalCache(value, parse)

  /** GH-353: lexical external-workbook prefix — `[2]Book1!`, `'[3]Sheet Name'!`, `[2]!name`. */
  private val externalPrefixPattern = java.util.regex.Pattern.compile("""'?\[\d+\]""")

  /**
   * GH-695: one evaluator for a fold a caller runs itself and the formula it evaluates after it
   * (the CLI's `eval`/`evala` precedent pass), so `x#` in that formula reads the arrays the fold
   * computed. Use it with [[evaluateCellUsing]], [[evaluateFormulaUsing]] and
   * [[evaluateArrayFormulaUsing]], which match the public entry points.
   */
  private[xl] def spillTrackingEvaluator(
    formula: String,
    workbook: Workbook,
    rng: Rng = Rng.system
  ): Evaluator =
    Evaluator.spillTrackingInstance(
      rng,
      Evaluator.mayReadSpills(formula) || Evaluator.mayReadSpills(workbook)
    )

  /**
   * GH-715: an evaluation of a formula placed at `at` on the sheet named `sheet`: `evaluator` is
   * [[spillTrackingEvaluator]] with that cell guarded. Any read of it during the evaluation, by
   * whatever path computed it (INDIRECT, OFFSET, a defined name, another sheet's formula), fails
   * instead of answering with the cell's old content, and [[wasRead]] records it even when the
   * failure is swallowed (IFERROR takes any failure). For the reads to reach the guard, the caller
   * puts an uncached formula at `at` ([[placed]]) in the sheet it evaluates on and in the workbook
   * it passes. Single-use, like the guard it holds.
   */
  private[xl] final class Placement private[SheetEvaluator] (
    val evaluator: Evaluator,
    guard: Evaluator.PlacementGuard
  ):
    /** Whether the evaluation read the placement: the plain cell there would be circular. */
    def wasRead: Boolean = guard.read

  /** A [[Placement]] for `formula` at `at` on the sheet named `sheet` of `workbook`. */
  private[xl] def placement(
    formula: String,
    workbook: Workbook,
    sheet: SheetName,
    at: ARef,
    rng: Rng = Rng.system
  ): Placement =
    val (evaluator, guard) = Evaluator.placedInstance(
      rng,
      Evaluator.mayReadSpills(formula) || Evaluator.mayReadSpills(workbook),
      sheet,
      at
    )
    new Placement(evaluator, guard)

  /**
   * `sheet` with `formula` placed, uncached, at `at`: what a [[Placement]] needs there. Stored
   * without its leading `=`, as every formula record is.
   */
  private[xl] def placed(sheet: Sheet, at: ARef, formula: String): Sheet =
    sheet.put(at, CellValue.Formula(formula.stripPrefix("="), None))

  /** GH-695: whether a sheet-level fold should record anchors' results for `x#` readers. */
  private def readsSpills(sheet: Sheet, workbook: Option[Workbook]): Boolean =
    workbook.fold(Evaluator.mayReadSpills(sheet))(Evaluator.mayReadSpills)

  private[xl] def evaluateCellUsing(
    sheet: Sheet,
    ref: ARef,
    evaluator: Evaluator,
    workbook: Option[Workbook],
    clock: Clock = Clock.system
  ): XLResult[CellValue] =
    evaluateCellWithEvaluator(sheet, ref, evaluator, clock, workbook)

  /** `currentCell` as in the public `evaluateFormula`: `Some(cell)` is the plain cell there. */
  private[xl] def evaluateFormulaUsing(
    sheet: Sheet,
    formula: String,
    evaluator: Evaluator,
    workbook: Option[Workbook],
    clock: Clock = Clock.system,
    currentCell: Option[ARef] = None
  ): XLResult[CellValue] =
    evaluateFormulaWith(sheet, formula, evaluator, clock, workbook, currentCell)

  private[xl] def evaluateArrayFormulaUsing(
    sheet: Sheet,
    formula: String,
    originRef: ARef,
    evaluator: Evaluator,
    workbook: Option[Workbook],
    clock: Clock = Clock.system
  ): XLResult[(Sheet, CellRange)] =
    evaluateArrayFormulaImpl(sheet, formula, originRef, evaluator.withArrayResults, clock, workbook)

  /**
   * GH-695: write a computed value back into a threaded evaluation sheet. A formula cell keeps its
   * record with the value as its cache, so a later `x#` read in the same fold still finds a spill
   * anchor (a constant is `#REF!`) — an ArrayFormula record, whose result the fold recorded in its
   * generation memo (see [[evaluateCellWithEvaluator]]), or a Normal one, which `x#` evaluates as
   * an array — and a plain read reuses the one computation (a cached formula reads as its cache).
   * Every other cell becomes the plain value, as before.
   */
  private[xl] def threadComputed(sheet: Sheet, ref: ARef, value: CellValue): Sheet =
    sheet.cells.valueAt(ref) match
      case f: CellValue.Formula =>
        value match
          // only an uncached record handed back by reference (`IFERROR(X1,0)` over an X1 not yet
          // evaluated; `cellResult` unwraps cached ones): the value is unknown, so the cell stays
          // uncached and a reader evaluates it on demand
          case _: CellValue.Formula => sheet.put(ref, f.copy(cachedValue = None))
          case computed => sheet.put(ref, f.copy(cachedValue = Some(computed)))
      case _ => sheet.put(ref, value)

  /**
   * #767: whether [[threadComputed]] at `ref` may change which cells of `sheet` are populated: it
   * does unless the cell holds a formula record, whose cache alone it rewrites.
   */
  private[eval] def threadChangesStructure(sheet: Sheet, ref: ARef): Boolean =
    sheet.cells.valueAt(ref) match
      case _: CellValue.Formula => false
      case _ => true

  /**
   * GH-274: strip stale formula caches from the given cells.
   *
   * Used on the deferred dynamic bucket before threading an evaluation fold: a dynamic read
   * (INDIRECT) of a not-yet-evaluated bucket cell then recursively evaluates the formula fresh
   * (depth-guarded) instead of trusting a previous generation's cache. Normal and ArrayFormula
   * records are both evaluatable; DataTable records alone retain their pinned cache, which is their
   * only source of truth. Only the threaded temp sheet is affected — final cache write-back
   * overlays computed results on the original sheet.
   */
  private[eval] def stripFormulaCaches(sheet: Sheet, refs: Set[ARef]): Sheet =
    val at = IArray.newBuilder[ARef]
    val stripped = IArray.newBuilder[CellValue]
    refs.foreach { r =>
      sheet.cells.valueAt(r) match
        // GH-430: a DataTable cache is the sole source of truth. Graph analysis normally keeps
        // these records out of the dynamic bucket; keep this guard as structural safety.
        case CellValue.Formula(_, Some(_), _: FormulaKind.DataTable) => ()
        case f @ CellValue.Formula(_, Some(_), _) =>
          at += r
          stripped += f.copy(cachedValue = None)
        case _ => ()
    }
    sheet.withValues(at.result(), stripped.result())

  /**
   * Convert EvalError to XLError for integration.
   *
   * @param error
   *   The evaluation error
   * @param formulaContext
   *   Optional formula string for context
   * @return
   *   XLError with detailed message
   */
  private def evalErrorToXLError(error: EvalError, formulaContext: Option[String]): XLError =
    val contextStr = formulaContext.map(f => s" in formula: $f").getOrElse("")

    error match
      case EvalError.DivByZero(num, denom) =>
        XLError.FormulaError(
          formulaContext.getOrElse(""),
          s"Division by zero: $num / $denom$contextStr"
        )
      case EvalError.CodecFailed(ref, codecErr) =>
        val codecMsg = codecErr match
          case CodecError.TypeMismatch(expected, actual) =>
            s"Expected $expected, got $actual"
          case CodecError.ParseError(value, targetType, detail) =>
            s"Cannot parse '$value' as $targetType: $detail"
        XLError.FormulaError(
          formulaContext.getOrElse(""),
          s"Type mismatch at $ref: $codecMsg$contextStr"
        )
      case EvalError.CircularRef(cycle) =>
        val cyclePath = cycle
          .map { ref =>
            // Format ARef to A1 notation
            // (Simplified - in production would use ARef.toA1 extension)
            s"cell_${ref}"
          }
          .mkString(" → ")
        XLError.FormulaError(
          formulaContext.getOrElse(""),
          s"Circular reference detected: $cyclePath"
        )
      // #829: a dynamic cycle reads at the cell boundary as Tarjan's static cycle does — the
      // reason alone, beside `Circular reference` — never as a generic evaluation failure
      case EvalError.EvalFailed(reason, _) if Evaluator.isDynamicCycle(error) =>
        XLError.FormulaError(formulaContext.getOrElse(""), reason)
      case EvalError.EvalFailed(reason, ctx) =>
        val fullContext = (ctx.toList ++ formulaContext.toList).mkString(", ")
        XLError.FormulaError(
          formulaContext.getOrElse(""),
          s"Evaluation failed: $reason${if fullContext.nonEmpty then s" ($fullContext)" else ""}"
        )
      // GH-344: unreachable after boundary promotion (toErrorValue promotes every ErrorValue),
      // kept total for direct callers of this table.
      case EvalError.ErrorValue(err, ctx) =>
        XLError.FormulaError(
          formulaContext.getOrElse(""),
          ctx.fold(err.toExcel)(c => s"${err.toExcel} ($c)")
        )
      case other =>
        XLError.FormulaError(
          formulaContext.getOrElse(""),
          s"Evaluation error: $other$contextStr"
        )
