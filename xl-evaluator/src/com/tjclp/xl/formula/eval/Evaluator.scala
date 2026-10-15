package com.tjclp.xl.formula.eval

import java.util.concurrent.atomic.AtomicLong

import com.tjclp.xl.formula.ast.{BinarySpine, BindingCoercion, TExpr}
import com.tjclp.xl.formula.functions.{
  ArgValue,
  ArrayLift,
  EvalContext,
  FunctionSpec,
  FunctionSpecs,
  LiftSlot,
  ReferenceFunctionNames,
  ReferenceOperators,
  SumProductTail
}
import com.tjclp.xl.formula.graph.DependencyGraph
import com.tjclp.xl.formula.printer.FormulaPrinter
import com.tjclp.xl.formula.parser.{FormulaParser, ParseError}
import com.tjclp.xl.formula.{Clock, Rng}

import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.{Cell, CellError, CellValue, FormulaKind}
import com.tjclp.xl.workbooks.{DefinedName, Workbook}
import com.tjclp.xl.SheetName
import com.tjclp.xl.syntax.* // Extension methods for Sheet.get, CellRange.cells, ARef.toA1

import scala.collection.concurrent.TrieMap
import scala.math.BigDecimal
import scala.util.boundary
import scala.util.boundary.break

/**
 * Pure functional formula evaluator.
 *
 * Evaluates TExpr AST against a Sheet, returning either an error or the computed value. All
 * evaluation is total - no exceptions thrown, no side effects.
 *
 * Laws satisfied:
 *   1. Literal identity: eval(Lit(x)) == Right(x)
 *   2. Arithmetic laws: eval(Add(Lit(a), Lit(b))) == Right(a + b)
 *   3. Eager logicals (GH-344): AND/OR evaluate every argument left-to-right — the first failure
 *      wins (Excel does not short-circuit logical functions)
 *   4. Totality: eval always returns Either[EvalError, A] (never throws)
 *
 * Example:
 * {{{
 * val expr = TExpr.Add(TExpr.Lit(BigDecimal(10)), TExpr.Ref(ref"A1", TExpr.decodeNumeric))
 * val evaluator = Evaluator.instance
 * evaluator.eval(expr, sheet) match
 *   case Right(result) => println(s"Result: $$result")
 *   case Left(error) => println(s"Error: $$error")
 * }}}
 */
trait Evaluator:
  /**
   * Evaluate expression against sheet.
   *
   * Every factory on the companion returns a guarded evaluator: a throwable escaping evaluation (an
   * internal defect or a failing Clock/Rng capability) comes back as `EvalError.EvalFailed` naming
   * it, and a StackOverflowError likewise; OutOfMemoryError and InterruptedException propagate.
   *
   * @param expr
   *   The expression to evaluate
   * @param sheet
   *   The sheet providing cell values
   * @param clock
   *   Clock for date/time functions (defaults to system clock)
   * @param workbook
   *   Optional workbook for cross-sheet references (defaults to None)
   * @param currentCell
   *   The cell the formula belongs to: ROW()/COLUMN() with no arguments read it, and a plain
   *   (legacy) evaluator — [[Evaluator.instance]] — intersects a multi-cell reference in a value
   *   position with it (`=A1:A10*2` at row 5 reads A5). Without it, that intersection is the loud
   *   `@` failure; evaluate a formula that has no cell with `sheet.evaluateFormula`, which
   *   evaluates it as a dynamic-array formula typed into a new cell
   * @return
   *   Either evaluation error or computed value
   */
  def eval[A](
    expr: TExpr[A],
    sheet: Sheet,
    clock: Clock = Clock.system,
    workbook: Option[Workbook] = None,
    currentCell: Option[ARef] = None
  ): Either[EvalError, A]

  /**
   * This evaluator in array mode — what an ArrayFormula record (a CSE or dynamic-array anchor)
   * evaluates with, so its value is the array's element (0,0), the Excel anchor. Implementations
   * outside the evaluator keep their own mode.
   */
  private[formula] def withArrayResults: Evaluator = this

  /**
   * GH-695: the generation memo this evaluator carries, if any — where an evaluation fold records
   * each spill anchor's computed array for the generation's `x#` readers.
   */
  private[formula] def generationMemo: Option[Evaluator.AggregateMemo] = None

object Evaluator:
  /**
   * Default evaluator instance.
   *
   * Pure functional implementation with short-circuit evaluation for And/Or.
   */
  def instance: Evaluator = TotalEvaluator(new EvaluatorImpl())

  /**
   * Evaluator instance with an explicit randomness source (GH-115).
   *
   * Use `Evaluator.instance(Rng.seeded(seed))` for deterministic RAND/RANDBETWEEN.
   */
  def instance(rng: Rng): Evaluator = TotalEvaluator(new EvaluatorImpl(rng = rng))

  /**
   * Evaluator instance that knows the workbook's saved location (GH-424).
   *
   * CELL("filename") reports `directory[file]sheet` from this path; without one it returns the
   * empty string — Excel's behavior for a never-saved workbook.
   */
  def instance(workbookPath: Option[String]): Evaluator =
    TotalEvaluator(new EvaluatorImpl(workbookPath = workbookPath))

  /**
   * Internal evaluator for one workbook-recalculation generation.
   *
   * Public/direct evaluation deliberately has no aggregate memo: its one-shot semantics and
   * allocation profile stay unchanged. WorkbookEvaluator alone creates this capability and shares
   * it across the generation's sequential or parallel cell evaluations. #767: `position` is the
   * slot the evaluated cell occupies in the pass's order, for one cell's evaluation; the pass makes
   * one such evaluator per cell.
   */
  private[formula] def recalculationInstance(
    rng: Rng,
    aggregateMemo: AggregateMemo,
    position: EvalPosition = EvalPosition.None
  ): Evaluator =
    TotalEvaluator(
      new EvaluatorImpl(rng = rng, aggregateMemo = Some(aggregateMemo), passPosition = position)
    )

  /**
   * GH-695: [[instance]] plus a generation memo that records spill anchors' results (when
   * `trackSpills`) and caches no aggregates — for the sheet-level evaluation folds, which thread
   * computed values like a recalculation but never had the aggregate memo.
   */
  private[formula] def spillTrackingInstance(rng: Rng, trackSpills: Boolean): Evaluator =
    TotalEvaluator(new EvaluatorImpl(rng = rng, aggregateMemo = Some(spillTracking(trackSpills))))

  private def spillTracking(trackSpills: Boolean): AggregateMemo =
    val memo = new AggregateMemo(cachesAggregates = false)
    if trackSpills then memo.trackSpills()
    memo

  /**
   * GH-715: [[spillTrackingInstance]] for a formula placed at `at` on the sheet named `sheet`: the
   * evaluator carries a [[PlacementGuard]] for that cell, which every memo of the evaluation
   * inherits, so a read of the placement by any path is a `CircularRef`, and the returned guard
   * records the read even when IFERROR swallows that error (IFERROR takes any failure). The caller
   * puts an uncached formula at `at`, in the sheet it evaluates on and in the workbook it passes,
   * so reads reach a memo.
   */
  private[formula] def placedInstance(
    rng: Rng,
    trackSpills: Boolean,
    sheet: SheetName,
    at: ARef
  ): (Evaluator, PlacementGuard) =
    val guard = new PlacementGuard(sheet, at)
    // No depth-0 memo: one would outlive every evaluation it serves (the CLI's fold makes a new
    // Sheet per step, and the memo keys by Sheet identity). Each read starts its own memo, which
    // inherits the guard, and the guard's one record of a read outlives them all. The trade-off: an
    // uncached precedent two top-level reads share evaluates once per read, not once per pass. The
    // CLI fold has already evaluated and cached the precedents the static check sees, so only
    // run-time reads (INDIRECT, OFFSET, names, other sheets) can pay it.
    val evaluator = TotalEvaluator(
      new EvaluatorWithDepth(
        0,
        rng = rng,
        memo = None,
        aggregateMemo = Some(spillTracking(trackSpills)),
        guard = Some(guard)
      )
    )
    (evaluator, guard)

  /**
   * GH-695: whether any formula in the book (cells or defined names) may read a spill — `x#` or its
   * stored `ANCHORARRAY` spelling. A cheap textual over-approximation (`#REF!` matches too): it
   * only decides whether an evaluation fold records anchors' results for `x#` readers.
   */
  private[formula] def mayReadSpills(wb: Workbook): Boolean =
    wb.sheets.exists(mayReadSpills) || wb.metadata.definedNames.exists(n =>
      mayReadSpills(n.formula)
    )

  private[formula] def mayReadSpills(sheet: Sheet): Boolean =
    sheet.cells.valuesIterator.exists(_.value match
      case CellValue.Formula(text, _, _) => mayReadSpills(text)
      case _ => false)

  private[formula] def mayReadSpills(formula: String): Boolean =
    formula.indexOf('#') >= 0 || formula.toUpperCase(java.util.Locale.ROOT).contains("ANCHORARRAY")

  /**
   * Evaluator instance that allows array results to propagate.
   *
   * Used for array formula evaluation where arithmetic over ranges should spill arrays.
   */
  def arrayInstance: Evaluator = TotalEvaluator(new EvaluatorImpl(allowArrayResults = true))

  /** Array-result evaluator with an explicit randomness source (GH-115). */
  def arrayInstance(rng: Rng): Evaluator =
    TotalEvaluator(new EvaluatorImpl(allowArrayResults = true, rng = rng))

  /**
   * #750, for the sparse-equals-dense laws: [[arrayInstance]] whose ranges are all read one way
   * ([[ArrayArithmetic.RangeRead]]: through their populated box, or position by position). The
   * choice rides on the pass memo, which every derived evaluator of the evaluation shares, so it is
   * an explicit value rather than ambient state. The memo is the evaluator's own: use the instance
   * for one evaluation of one sheet.
   */
  private[formula] def arrayInstanceReading(rng: Rng, read: ArrayArithmetic.RangeRead): Evaluator =
    TotalEvaluator(
      new EvaluatorWithDepth(
        0,
        allowArrayResults = true,
        rng = rng,
        memo = Some(new EvalMemo(rangeRead = read)),
        guard = None
      )
    )

  /**
   * Convenience method for direct evaluation (forwards to instance.eval).
   */
  def eval[A](
    expr: TExpr[A],
    sheet: Sheet,
    clock: Clock = Clock.system,
    workbook: Option[Workbook] = None,
    currentCell: Option[ARef] = None
  ): Either[EvalError, A] =
    instance.eval(expr, sheet, clock, workbook, currentCell)

  // Helper methods for consistent cross-sheet error messages
  private[formula] def missingWorkbookError(refStr: String, isRange: Boolean = false): EvalError =
    val refType = if isRange then "range" else "reference"
    EvalError.EvalFailed(
      s"Cross-sheet $refType $refStr requires workbook context, but none was provided.",
      None
    )

  /**
   * GH-353: external-workbook references cannot be resolved (the external workbook is not loaded).
   * Closed-workbook semantics live OUTSIDE the evaluator: a formula cell bearing an external ref
   * keeps its Excel-written cached value (SheetEvaluator.pinnedExternalCache) and is never
   * re-evaluated; this error surfaces only for direct evaluation or cells with no cache, and
   * propagates to dependents as a normal per-cell error.
   */
  private[formula] def externalRefUnsupported(refStr: String): EvalError =
    EvalError.EvalFailed(
      s"External workbook reference $refStr cannot be resolved: external workbooks are not " +
        "loaded. Excel's cached value is used when the cell has one; recalculate in Excel to " +
        "refresh it.",
      None
    )

  private[formula] def sheetNotFoundError(
    sheetName: SheetName,
    err: com.tjclp.xl.error.XLError
  ): EvalError =
    EvalError.EvalFailed(
      s"Sheet '${sheetName.value}' not found in workbook: ${err.message}",
      None
    )

  /**
   * Resolve a RangeLocation to its target (sheet, range) pair — THE single static resolution
   * boundary for range-typed argument slots (GH-394).
   *
   * For Local ranges, returns the current sheet with the carried range. For CrossSheet ranges,
   * looks up the target sheet in the workbook context. For Name locations the carried range is
   * unknown until here: the name resolves against the workbook's name table (sheet-scoped SHADOWS
   * workbook-scoped, case-insensitive — [[lookupDefinedName]]; a sheet-qualified name looks up as
   * seen from its qualifier), its refersTo text parses, and range/cell-shaped targets yield the
   * resolved pair (a single-cell target acts as a 1×1 range, like Excel). Name→name chains follow
   * hop by hop like expression-position NameRef (GH-411), guarded by `resolvingNames` — a chain
   * that revisits a member is a clean cycle error, one deeper than [[NameWalk.MaxDepth]] names a
   * clean too-deep error (#691). Other non-range targets (constants, formulas) are a clean per-cell
   * #VALUE! error — never a MatchError. This is the STATIC resolution: a function's range slot, and
   * the typed [[TExpr.Aggregate]] node's, go through [[resolveRangeSlot]] (GH-710), which also
   * evaluates a name computing a reference and reads a LET name's reference.
   *
   * The callers that stay static do so on purpose: a written cross-sheet reference or range (the
   * SheetRange/SheetRef operand paths, a defined name's own refersTo hops), which no evaluation can
   * change; the callers that try the static form first and evaluate when it fails (OFFSET's base,
   * the `@` operand, the reference-argument path for a name, a lifted slot's name); and the spill
   * operator's anchor (`x#` takes one cell, written or named). A LET name never reaches them as a
   * location: `ArgSpec.sumProductArg` and those paths carry it as an expression.
   *
   * @param resolvingNames
   *   the GH-384 name-cycle guard (the resolved names on the current resolution path, #691);
   *   callers inside the evaluator pass their ambient guard so a name's refersTo cannot re-enter
   *   itself through a range slot.
   * @param memo
   *   #767: the pass's aggregate memo, whose [[PassCache]] caches static name resolutions and whose
   *   counters count name parses; None outside a recalculation
   */
  private[formula] def resolveRangeLocation(
    location: TExpr.RangeLocation,
    currentSheet: Sheet,
    workbook: Option[Workbook],
    resolvingNames: NameTrail = NameTrail.empty,
    memo: Option[AggregateMemo] = None
  ): Either[EvalError, (Sheet, CellRange)] =
    location match
      case TExpr.RangeLocation.Local(range, _) =>
        Right((currentSheet, range))
      case TExpr.RangeLocation.CrossSheet(sheetName, range, _) =>
        workbook match
          case None =>
            // GH-280: quote cell-ref-shaped sheet names so diagnostics read unambiguously
            val refStr = s"${SheetName.quoteForFormula(sheetName.value)}!${range.toA1}"
            Left(missingWorkbookError(refStr, isRange = true))
          case Some(wb) =>
            wb(sheetName) match
              case Left(err) => Left(sheetNotFoundError(sheetName, err))
              case Right(targetSheet) => Right((targetSheet, range))
      // GH-353: the target workbook is not loaded — same friendly error as a direct external
      // ref; cells CONTAINING such calls are pinned to their Excel cache upstream and never
      // reach this point
      case loc @ TExpr.RangeLocation.External(_, _, _, _) =>
        Left(externalRefUnsupported(loc.toA1))
      // GH-394: a defined name in a range slot resolves through the name table
      case TExpr.RangeLocation.Name(name, scope) =>
        resolveNameToRange(name, scope, currentSheet, workbook, resolvingNames, memo)
      // GH-710: a LET name's reference lives in the LET environment, which only a function's
      // EvalContext carries (EvalContext.resolveRange)
      case TExpr.RangeLocation.Binding(name) =>
        Left(
          EvalError.ErrorValue(
            CellError.Value,
            Some(s"LET name '$name' resolves only inside its LET")
          )
        )
      // GH-612: an error in a range slot (SUM(#REF!)) IS the error value it names
      case TExpr.RangeLocation.Error(error, _) => Left(EvalError.ErrorValue(error))

  /**
   * GH-394: resolve a defined name used in a RANGE-typed argument slot to its (sheet, range).
   *
   * Mirrors [[EvaluatorWithDepth.evalNameRef]]'s lookup (missing workbook, unknown name, and
   * unparseable refersTo produce the same message shapes) but restricts the TARGET to reference
   * shapes: ranges, sheet-qualified ranges, single cells (1×1 ranges), and — GH-411, like
   * expression position — further defined names, followed hop by hop under the `resolvingNames`
   * cycle guard (each hop looks up as seen from its predecessor's defining sheet, the evalNameRef
   * convention). Anything else — constants, formulas — is Excel's #VALUE! as a clean error value.
   * The defining-sheet convention matches evalNameRef: sheet-scoped names resolve their unqualified
   * refersTo against their scope sheet.
   */
  private def resolveNameToRange(
    name: String,
    scope: Option[SheetName],
    currentSheet: Sheet,
    workbook: Option[Workbook],
    resolvingNames: NameTrail,
    memo: Option[AggregateMemo]
  ): Either[EvalError, (Sheet, CellRange)] =
    resolveNameRangeTarget(name, scope, currentSheet, workbook, resolvingNames, memo).flatMap {
      case NameRangeTarget.Range(target, range) => Right((target, range))
      case NameRangeTarget.Computed(notARange) => Left(notARange)
    }

  /**
   * GH-710: what a defined name in a range slot names statically — a range, or a value only
   * evaluation can compute (the name, or the last name of its chain, is bound to a formula or a
   * constant). Every failure on the way — no workbook, an unknown or unparseable name, a name cycle
   * or a chain too deep (#691), a name bound to an error literal, a missing sheet — is a `Left`, so
   * [[resolveRangeSlot]] evaluates a name only when its definition really computes something.
   */
  private[formula] enum NameRangeTarget:
    case Range(target: Sheet, range: CellRange)

    /**
     * The definition is a formula, which may compute a reference; `notARange` is the static
     * diagnostic. A constant is a static `Left` instead: no reference can come out of it.
     */
    case Computed(notARange: EvalError)

  /**
   * #733: a definition that is a constant — a scalar literal, signed or with `%` (`0.05`, `-0.05`,
   * `"x"`, `TRUE`, `5%`). Its value can never be a reference, and in a range slot it is `#VALUE!`
   * whatever it holds. An array constant is not one: a 1×1 `{#N/A}` evaluates to its error.
   */
  private def isConstant(expr: TExpr[?]): Boolean = expr match
    case TExpr.Lit(_: ArrayResult) => false
    case TExpr.Lit(_) => true
    case TExpr.UnaryPlus(inner) => isConstant(inner)
    case TExpr.Percent(inner) => isConstant(inner)
    // unary minus parses as `0 - x`
    case TExpr.Sub(left, right) => isConstant(left) && isConstant(right)
    case _ => false

  /**
   * #733: whether a defined name in a range slot computes its reference — bound, directly or
   * through a chain of names, to a formula (OFFSET, INDIRECT, IF over references) — so only
   * evaluation knows the cells it names. A name bound to a range, a cell or a constant is not, and
   * neither is a name that fails to resolve (it reads nothing: its evaluation fails).
   */
  private[formula] def nameComputesRange(
    name: String,
    scope: Option[SheetName],
    currentSheet: Sheet,
    workbook: Workbook
  ): Boolean =
    resolveNameRangeTarget(name, scope, currentSheet, Some(workbook), NameTrail.empty) match
      case Right(NameRangeTarget.Computed(_)) => true
      case _ => false

  /**
   * #767: the static walk, with the pass's cache in front of it. A walk that starts with an empty
   * trail (no enclosing name resolution, so no cycle refusal can depend on the path) and ends in a
   * range is recorded in the [[PassCache]] and answered from it afterwards, its sheet bound again
   * from `workbook`. A computed name, an error and a walk inside another name's resolution are
   * never cached.
   */
  private def resolveNameRangeTarget(
    name: String,
    scope: Option[SheetName],
    currentSheet: Sheet,
    workbook: Option[Workbook],
    resolvingNames: NameTrail,
    memo: Option[AggregateMemo] = None
  ): Either[EvalError, NameRangeTarget] =
    def untraced(traced: TracedTarget): NameRangeTarget = traced match
      case TracedTarget.Range(target, range, _) => NameRangeTarget.Range(target, range)
      case TracedTarget.Computed(notARange) => NameRangeTarget.Computed(notARange)
    (memo.flatMap(_.passCache), workbook) match
      case (Some(passCache), Some(wb)) if resolvingNames.path.isEmpty =>
        passCache.cachedName(name, scope, currentSheet, wb) match
          case Some((target, range)) => Right(NameRangeTarget.Range(target, range))
          case None =>
            tracedNameTarget(name, scope, currentSheet, workbook, resolvingNames, memo).map {
              traced =>
                traced match
                  case TracedTarget.Range(_, range, origin) =>
                    passCache.recordName(name, scope, currentSheet, wb, origin, range)
                  case TracedTarget.Computed(_) => ()
                untraced(traced)
            }
      case _ =>
        tracedNameTarget(name, scope, currentSheet, workbook, resolvingNames, memo).map(untraced)

  /**
   * #767: a [[NameRangeTarget]] that also says where its sheet came from, so a cached resolution
   * binds the same sheet again: the reader's own sheet, or the workbook's sheet of that name.
   */
  private enum TracedTarget:
    case Range(target: Sheet, range: CellRange, origin: PassCache.NameOrigin)
    case Computed(notARange: EvalError)

  private def tracedNameTarget(
    name: String,
    scope: Option[SheetName],
    currentSheet: Sheet,
    workbook: Option[Workbook],
    resolvingNames: NameTrail,
    memo: Option[AggregateMemo]
  ): Either[EvalError, TracedTarget] =
    def crossSheet(
      location: TExpr.RangeLocation,
      definingSheet: Sheet,
      guard: NameTrail
    ): Either[EvalError, TracedTarget] =
      resolveRangeLocation(location, definingSheet, workbook, guard, memo)
        .map((target, range) =>
          TracedTarget.Range(target, range, PassCache.NameOrigin.Book(target.name))
        )
    workbook match
      case None =>
        Left(
          EvalError.EvalFailed(
            s"Defined name '$name' requires workbook context, but none was provided.",
            None
          )
        )
      case Some(wb) =>
        val lookupFrom = scope.getOrElse(currentSheet.name)
        lookupDefinedName(wb, lookupFrom, name) match
          case None =>
            Left(EvalError.EvalFailed(s"Name '$name' is not defined in this workbook.", None))
          case Some(dn) =>
            val scopeSheet = definedNameScope(wb, dn)
            val definingSheet = scopeSheet.getOrElse(currentSheet)
            // the reader's sheet of a chain hop is this name's defining sheet
            val definingOrigin: PassCache.NameOrigin =
              scopeSheet.fold(PassCache.NameOrigin.Reader)(s => PassCache.NameOrigin.Book(s.name))
            def throughHop(traced: TracedTarget): TracedTarget = traced match
              case TracedTarget.Range(target, range, PassCache.NameOrigin.Reader) =>
                TracedTarget.Range(target, range, definingOrigin)
              case other => other
            val node = NameWalk.Node(dn, definingSheet.name)
            val guard = resolvingNames.enter(node)
            resolvingNames.refusal(name, node).toLeft(()).flatMap { _ =>
              parseNameBody(dn.formula, memo) match
                case Left(parseErr) =>
                  Left(
                    EvalError.EvalFailed(
                      s"Defined name '$name' has an unparseable definition '${dn.formula}': " +
                        s"${ParseError.toXLError(parseErr, dn.formula).message}",
                      None
                    )
                  )
                case Right(target) =>
                  def range(cells: CellRange): Either[EvalError, TracedTarget] =
                    Right(TracedTarget.Range(definingSheet, cells, definingOrigin))
                  target match
                    case TExpr.RangeRef(cells, _) => range(cells)
                    case TExpr.SheetRange(sheetName, cells, _) =>
                      crossSheet(
                        TExpr.RangeLocation.CrossSheet(sheetName, cells),
                        definingSheet,
                        guard
                      )
                    // Single-cell targets act as 1×1 ranges (names routinely point at one cell).
                    // A standalone ref at the top of a parsed formula surfaces as the TYPED
                    // Ref/SheetRef (the resolved-value rewrite); Poly forms are kept for safety.
                    case TExpr.Ref(at, _, _) => range(CellRange(at, at))
                    case TExpr.PolyRef(at, _) => range(CellRange(at, at))
                    case TExpr.SheetRef(sheetName, at, _, _) =>
                      crossSheet(
                        TExpr.RangeLocation.CrossSheet(sheetName, CellRange(at, at)),
                        definingSheet,
                        guard
                      )
                    case TExpr.SheetPolyRef(sheetName, at, _) =>
                      crossSheet(
                        TExpr.RangeLocation.CrossSheet(sheetName, CellRange(at, at)),
                        definingSheet,
                        guard
                      )
                    // GH-411: name→name chains follow like expression position, each hop as seen
                    // from its predecessor's defining sheet, under the shared cycle guard
                    case TExpr.NameRef(next) =>
                      tracedNameTarget(next, None, definingSheet, workbook, guard, memo)
                        .map(throughHop)
                    case TExpr.SheetNameRef(qualifier, next) =>
                      tracedNameTarget(next, Some(qualifier), definingSheet, workbook, guard, memo)
                        .map(throughHop)
                    // GH-630: a name bound to an error literal IS that error (Excel: SUM(bad) with
                    // bad = #N/A is #N/A), so COUNT/COUNTA can triage it as an error argument
                    case TExpr.ErrorLit(err, _) =>
                      Left(
                        EvalError.ErrorValue(err, Some(s"Defined name '$name' is ${err.toExcel}"))
                      )
                    case other =>
                      // Excel: a non-reference name in a range position is #VALUE! — unless the
                      // formula computes a reference (GH-710, resolveRangeSlot). A constant
                      // (`rate = 0.05`, `-0.05`) never does: it is #VALUE! here, never evaluated
                      val notARange = EvalError.ErrorValue(
                        CellError.Value,
                        Some(s"name '$name' does not refer to a range (refersTo: ${dn.formula})")
                      )
                      if isConstant(other) then Left(notARange)
                      else Right(TracedTarget.Computed(notARange))
            }

  /**
   * A defined name's body, parsed once per pass when the pass has a [[PassCache]] (#767), else on
   * every call; each actual parse counts as a `NameParses`.
   */
  private[formula] def parseNameBody(
    formula: String,
    memo: Option[AggregateMemo]
  ): Either[ParseError, TExpr[?]] =
    memo.flatMap(_.passCache) match
      case Some(passCache) => passCache.parsedNameBody(formula)(FormulaParser.parse)
      case None =>
        memo.foreach(_.counters.increment(RecalcCounters.Counter.NameParses))
        FormulaParser.parse(formula)

  /**
   * The (sheet, range) a range-typed argument slot names — THE resolver for a function's range slot
   * (`EvalContext.resolveRange`) and the typed [[TExpr.Aggregate]] node's (GH-710). A written range
   * or cell and a name bound to one resolve statically ([[resolveRangeLocation]]). A name whose
   * definition is a formula is evaluated (`evalName`) for the reference it computes — OFFSET,
   * INDIRECT, INDEX, IF or CHOOSE over references, a dynamic range — and a LET name is the
   * reference it is bound to (`bindings`). The error such a formula evaluates to is the slot's
   * error (`#REF!` for an OFFSET off the grid); a name or LET name holding a value is `#VALUE!`, as
   * in Excel, with the static diagnostic. An error the evaluation itself produces (a bare `#VALUE!`
   * from `OFFSET(A1,"x",0)`, a cycle) is the slot's as it is. A name that fails statically
   * (unknown, cyclic, too deep, bound to an error literal or a constant) is never evaluated.
   *
   * @param evalName
   *   evaluates a defined name (as written, with its qualifier) for the reference it computes, a
   *   [[RangeOperand]] when it computes one. The static walk has already run, so the evaluator's
   *   `evalNameRef(…, asReference = true)` is the right callback: it does not resolve again.
   */
  private[formula] def resolveRangeSlot(
    location: TExpr.RangeLocation,
    currentSheet: Sheet,
    workbook: Option[Workbook],
    resolvingNames: NameTrail,
    bindings: Map[String, Any],
    evalName: (String, Option[SheetName]) => Either[EvalError, Any],
    memo: Option[AggregateMemo] = None
  ): Either[EvalError, (Sheet, CellRange)] =
    resolveRangeSlotTraced(
      location,
      currentSheet,
      workbook,
      resolvingNames,
      bindings,
      evalName,
      memo
    )
      .map((target, range, _) => (target, range))

  /**
   * [[resolveRangeSlot]], also saying whether the slot resolved statically (#767): a written range
   * or cell, or a name bound to one, which a pass may serve from a [[ColumnSlice]]. A name that
   * computes its reference and a LET name are not static.
   */
  private[formula] def resolveRangeSlotTraced(
    location: TExpr.RangeLocation,
    currentSheet: Sheet,
    workbook: Option[Workbook],
    resolvingNames: NameTrail,
    bindings: Map[String, Any],
    evalName: (String, Option[SheetName]) => Either[EvalError, Any],
    memo: Option[AggregateMemo]
  ): Either[EvalError, (Sheet, CellRange, Boolean)] =
    def referenceOf(value: Any, notARange: EvalError): Either[EvalError, (Sheet, CellRange)] =
      value match
        case RangeOperand(target, range) => Right((target, range))
        case CellValue.Error(error) => Left(EvalError.ErrorValue(error))
        // a named formula evaluates as an array: its error arrives as a 1×1 array
        case single: ArrayResult if single.rows == 1 && single.cols == 1 =>
          single(0, 0) match
            case CellValue.Error(error) => Left(EvalError.ErrorValue(error))
            case _ => Left(notARange)
        case _ => Left(notARange)
    location match
      case TExpr.RangeLocation.Name(name, scope) =>
        resolveNameRangeTarget(name, scope, currentSheet, workbook, resolvingNames, memo)
          .flatMap {
            case NameRangeTarget.Range(target, range) => Right((target, range, true))
            case NameRangeTarget.Computed(notARange) =>
              evalName(name, scope)
                .flatMap(referenceOf(_, notARange))
                .map((target, range) => (target, range, false))
          }
      case TExpr.RangeLocation.Binding(name) =>
        val notARange = EvalError.ErrorValue(
          CellError.Value,
          Some(s"LET name '$name' does not refer to a range")
        )
        val bound = bindings.get(name) match
          case Some(value) => referenceOf(value, notARange)
          // Unreachable from formula text: the parser emits Binding only for a LET name in scope,
          // and evalLet binds every name before its body runs. A hand-built AST can still get here.
          case None => Left(EvalError.EvalFailed(s"LET name '$name' is not in scope", None))
        bound.map((target, range) => (target, range, false))
      case _ =>
        resolveRangeLocation(location, currentSheet, workbook, resolvingNames, memo)
          .map((target, range) => (target, range, true))

  // ===== GH-384: defined-name resolution =====

  /**
   * Look up a defined name visible from `currentSheet`.
   *
   * OOXML/Excel scoping: a sheet-scoped entry (localSheetId is a POSITIONAL index into `wb.sheets`)
   * SHADOWS a workbook-scoped entry of the same identifier; lookup is case-insensitive like Excel.
   * (GH-384's issue text said "workbook-scoped first" — that contradicts OOXML §18.2.5/Excel
   * behavior, so shadowing is implemented instead.)
   */
  private[formula] def lookupDefinedName(
    wb: Workbook,
    currentSheet: SheetName,
    name: String
  ): Option[DefinedName] =
    val sheetIdx = wb.sheets.indexWhere(_.name == currentSheet)
    lookupDefinedNameAt(wb, Option.when(sheetIdx >= 0)(sheetIdx), name)

  /**
   * [[lookupDefinedName]] with the sheet position already resolved — for callers that loop over
   * (sheet × name) pairs and can compute positions once instead of an O(sheets) `indexWhere` per
   * lookup (DependencyGraph.dynamicCells makes sheets × names such calls per recalculation).
   *
   * Callers must thread positions computed once; a lazy position map on Workbook cannot replace
   * this. Recalculation copies the workbook per evaluated cell (`wb.copy(sheets = …)`), so a
   * per-instance lazy val would be rebuilt by every copy — the cost it was meant to remove.
   * (WorkbookMetadata.definedNameIndex survives those copies only because the metadata FIELD is
   * shared by reference.)
   */
  private[formula] def lookupDefinedNameAt(
    wb: Workbook,
    sheetIdx: Option[Int],
    name: String
  ): Option[DefinedName] =
    wb.metadata.definedNameIndex.resolve(name, sheetIdx)

  /**
   * The sheet a defined name's refersTo evaluates against when the name is sheet-scoped; None for
   * workbook-scoped names (callers fall back to the referencing formula's sheet — refersTo text is
   * almost always fully qualified anyway).
   */
  private[formula] def definedNameScope(wb: Workbook, dn: DefinedName): Option[Sheet] =
    dn.localSheetId.flatMap(idx => wb.sheets.lift(idx))

  /**
   * #691: the defined names resolving on the current evaluation path — the cycle guard and the
   * depth cap — and the [[NameMemo]] of the outermost name being evaluated, once one is open.
   *
   * The guard keys on resolved nodes ([[NameWalk.Node]]: the definition and the sheet its body
   * evaluates against), never on the name's text: a sheet-local `X` defined as `T!X` reaches T's
   * own `X`, a different node, and evaluates it.
   */
  private[formula] final case class NameTrail(path: Set[NameWalk.Node], memo: Option[NameMemo]):
    def enter(node: NameWalk.Node): NameTrail = copy(path = path + node)

    /**
     * Why `node`, what the reference `name` resolves to, cannot resolve on this path: it is on the
     * path already (a cycle), or the path holds [[NameWalk.MaxDepth]] names (a chain too deep —
     * unresolvable, as every static name walk treats it). A refusal answers for this path only, so
     * it taints the open memo ([[NameMemo.taint]]).
     */
    def refusal(name: String, node: NameWalk.Node): Option[EvalError] =
      val refused =
        if path.contains(node) then
          Some(EvalError.EvalFailed(s"$NameCycleDetected while resolving '$name'.", None))
        else if path.sizeIs >= NameWalk.MaxDepth then
          Some(
            EvalError.EvalFailed(
              s"Defined name chain too deep while resolving '$name': more than " +
                s"${NameWalk.MaxDepth} names, each referring to the next.",
              None
            )
          )
        else None
      if refused.isDefined then memo.foreach(_.taint())
      refused

  private[formula] object NameTrail:
    val empty: NameTrail = NameTrail(Set.empty, None)

  /**
   * #691: the values of the defined names one outermost name evaluation reaches, so a name its
   * definition reaches along several paths — a diamond, each name reading the next twice — is
   * evaluated once per reading context instead of once per path (2^depth).
   *
   * A memo opens at the outermost name (an empty [[NameTrail]]) and threads, inside the trail,
   * through every evaluator derived within that name's evaluation; it never outlives it. Keys: the
   * resolved node, the read (`asReference`), the evaluator's array mode and the formula's cell
   * (implicit intersection), with the defining sheet and the workbook compared by identity.
   *
   * Reuse gives exactly what evaluating again would, the path-guarded evaluation without the memo:
   *   - A value that drew from the RNG is never stored: each reference to a name over RAND draws
   *     anew (`Rnd-Rnd` is not 0). The outermost name's evaluator draws through [[counting]].
   *   - The trail's refusals (a cycle, a chain past [[NameWalk.MaxDepth]]) are the only way the
   *     path reaches a value, so the first one [[taint]]s the memo: nothing is stored or served
   *     after it. Before it, a value is the same on any path where evaluating it again would refuse
   *     nothing — a name body evaluates in array mode with fresh bindings whatever reads it, so it
   *     takes the same branches — and a hit is served only where its chain still fits under the
   *     cap: each entry carries its height, the names its evaluation went below it.
   *
   * Not thread-safe, like [[EvalMemo]]: an evaluation is single-threaded and the memo never escapes
   * it.
   */
  private[formula] final class NameMemo:
    private val values = scala.collection.mutable.HashMap.empty[NameMemo.Key, NameMemo.Entry]
    private val draws = new AtomicLong(0L)
    private val tainted = new java.util.concurrent.atomic.AtomicBoolean(false)
    // the deepest name (its path size, itself included) entered by the computation in progress
    private val deepest = new java.util.concurrent.atomic.AtomicInteger(0)

    /** `rng`, counting its draws for this memo. */
    def counting(rng: Rng): Rng =
      val underlying = rng
      new Rng:
        def nextDouble(): Double =
          draws.incrementAndGet()
          underlying.nextDouble()

    /** A refusal happened on some path: from now on nothing is stored or served. */
    def taint(): Unit = tainted.set(true)

    /**
     * The value of `key`, evaluated by `compute` at `depth` (the path's size with the name on it)
     * unless an entry may be served there.
     */
    def getOrCompute(key: NameMemo.Key, definingSheet: Sheet, workbook: Workbook, depth: Int)(
      compute: => Either[EvalError, Any]
    ): Either[EvalError, Any] =
      values.get(key) match
        case Some(hit)
            if !tainted.get() && (hit.definingSheet eq definingSheet) &&
              (hit.workbook eq workbook) && depth + hit.height <= NameWalk.MaxDepth =>
          deepest.accumulateAndGet(depth + hit.height, (a, b) => a.max(b))
          hit.value
        case _ =>
          // Compute BEFORE storing: the computation recurses into this memo for other names.
          val outer = deepest.getAndSet(depth)
          val before = draws.get()
          val computed = compute
          val height = deepest.get() - depth
          deepest.accumulateAndGet(outer, (a, b) => a.max(b))
          if !tainted.get() && draws.get() == before then
            values.update(key, NameMemo.Entry(definingSheet, workbook, height, computed))
          computed

  private[formula] object NameMemo:
    final case class Key(
      node: NameWalk.Node,
      asReference: Boolean,
      arrayMode: Boolean,
      cell: Option[ARef]
    )

    /** `height`: the most names below the entry's own on a chain its evaluation followed. */
    final case class Entry(
      definingSheet: Sheet,
      workbook: Workbook,
      height: Int,
      value: Either[EvalError, Any]
    )

  /**
   * Excel's implicit intersection of `range` with the formula's own cell — the `@` operator
   * (GH-604) and a multi-cell reference in a lifted scalar slot of a plain formula cell: a single
   * cell is itself; a one-row range yields the cell in the formula's column, a one-column range the
   * cell in the formula's row; anything else — a 2-D range, or a vector the formula's row/column
   * does not cross — is `#VALUE!`. Without the formula's position (ad-hoc evaluation) it is a loud
   * failure.
   */
  private[formula] def implicitIntersection(
    range: CellRange,
    current: Option[ARef]
  ): Either[EvalError, ARef] =
    if range.width == 1 && range.height == 1 then Right(range.start)
    else
      current match
        case None =>
          Left(
            EvalError.EvalFailed(
              s"@${range.toA1} (implicit intersection) needs the formula's cell position",
              Some("@range")
            )
          )
        case Some(cell) =>
          val col = cell.col.index0
          val row = cell.row.index0
          if range.height == 1 && col >= range.colStart.index0 && col <= range.colEnd.index0 then
            Right(ARef.from0(col, range.rowStart.index0))
          else if range.width == 1 && row >= range.rowStart.index0 && row <= range.rowEnd.index0
          then Right(ARef.from0(range.colStart.index0, row))
          else
            Left(
              EvalError.ErrorValue(
                CellError.Value,
                Some(s"@${range.toA1}: no cell in the formula's row or column (${cell.toA1})")
              )
            )

  /**
   * The functions that return a reference (Excel's IF, IFS, CHOOSE and SWITCH return the reference
   * they select; OFFSET, INDIRECT, INDEX and, GH-713, XLOOKUP compute one; GH-669: so does the
   * intersection operator ` `, and GH-713 the range operator `:`): in a reference position their
   * result reaches the consumer whole, in a plain cell's value position it is implicitly
   * intersected. Built from `ReferenceFunctionNames` (the selectors, and every registered
   * `FunctionSpec.Referencing`, which the drift law ties to the registry) and the two operators,
   * which are unregistered. Static, so the parser's reference-shape check never initializes the
   * registry.
   */
  private[formula] val referenceFunctions: Set[String] =
    ReferenceFunctionNames.selecting ++ ReferenceFunctionNames.computing +
      ReferenceOperators.IntersectionName + ReferenceOperators.RangeName

  /**
   * GH-669: an operator node (binary, `%` or unary `+`) whose operands are only array constants,
   * scalar literals, error literals and single-cell references — at least one an array constant.
   * Such an operation reads the same in a plain cell and in array mode except for Excel's legacy
   * rule that its array survives, so the evaluator takes it in array mode. A chain's spine is
   * walked in a loop (BinarySpine).
   */
  /** GH-669: an array constant, or an operation over array constants (see below). */
  private[formula] def isArrayConstant(expr: TExpr[?]): Boolean = expr match
    case TExpr.Lit(_: ArrayResult) => true
    case other => arrayConstantOperation(other)

  private[formula] def arrayConstantOperation(expr: TExpr[?]): Boolean =
    // None: not such an operation; Some(sawArray)
    def scan(e: TExpr[?]): Option[Boolean] = e match
      case TExpr.Lit(_: ArrayResult) => Some(true)
      case TExpr.Lit(_) | TExpr.ErrorLit(_, _) | TExpr.Ref(_, _, _) | TExpr.PolyRef(_, _) |
          TExpr.SheetRef(_, _, _, _) | TExpr.SheetPolyRef(_, _, _) =>
        Some(false)
      case TExpr.Coerced(inner, _) => scan(inner)
      case TExpr.UnaryPlus(inner) => scan(inner)
      case TExpr.Percent(inner) => scan(inner)
      case chain if BinarySpine.isBinary(chain) =>
        BinarySpine.operandList(chain).foldLeft(Option(false)) { (saw, operand) =>
          saw.flatMap(before => scan(operand).map(_ || before))
        }
      case _ => None
    val isOperator = expr match
      case TExpr.UnaryPlus(_) | TExpr.Percent(_) => true
      case other => BinarySpine.isBinary(other)
    isOperator && scan(expr).contains(true)

  /**
   * Whether an expression may denote a reference rather than a value: a cell, a name, a LET name, a
   * LET expression, or a call to a function that returns a reference. A plain cell's LET binds such
   * an expression's reference itself; a LET body that computes a value still binds that value.
   */
  private[formula] def denotesReference(expr: TExpr[?]): Boolean = expr match
    case _: TExpr.Ref[?] | _: TExpr.PolyRef | _: TExpr.SheetRef[?] | _: TExpr.SheetPolyRef |
        _: TExpr.NameRef | _: TExpr.SheetNameRef | _: TExpr.BindingRef |
        _: TExpr.CoercedBindingRef[?] | _: TExpr.Let[?] =>
      true
    case call: TExpr.Call[?] => referenceFunctions.contains(call.spec.name)
    case TExpr.Coerced(inner, _) => denotesReference(inner)
    case _ => false

  /**
   * Functions whose expression argument must be a reference (readsBindingAsRange). Lazy, so this
   * object's initialization never forces the function registry's.
   */
  private lazy val ReferenceBaseFunctions: Set[String] = Set(FunctionSpecs.offset.name)

  /**
   * GH-710: whether an expression reads a LET name in a range slot (`COUNTIF(r, ">2")`), or holds a
   * call whose expression argument must be a reference (OFFSET's base: `OFFSET(c, 1, 0)` with `c`
   * bound to B5). An array evaluation's LET binds a reference-denoting value as the reference
   * itself only then — those positions need the reference, and every other read of the name
   * materializes it as before. Excel's LET always binds the reference, so binding it is never
   * wrong: without the OFFSET clause, `LET(c,B5,OFFSET(c,1,0))` was `#VALUE!` in array mode unless
   * an unrelated range slot in the same LET read a LET name.
   *
   * A walk for functions that inspect a reference's position or shape (ROW, COLUMN) is a trigger of
   * the same kind (`inspectsReference`); `evalLet` binds references when either answers true.
   *
   * The walk tracks neither which name a slot reads nor shadowing: a range-slot read of any LET
   * name in the subtree, an inner LET's own or one that shadows an outer name, makes the outer LET
   * bind its references too. That over-approximation is harmless: every value read of a bound
   * reference materializes it, so the result is the same either way.
   *
   * The walk must descend through every wrapper a range slot can sit under (the binary operators,
   * `Coerced`, `UnaryPlus`, `Percent`, `ToInt`, the date serials): a wrapper it misses turns
   * `COUNTIF(r, …)` into `#VALUE!` in array mode. DynamicNameRangeSlotSpec pins one formula per
   * wrapper. `evalLet` runs it once per LET evaluation, over that LET's subtree, so nested LETs
   * walk their inner subtrees again (size × depth per evaluation). It is a pure function of the AST
   * and could be computed once per `Let` node at parse time if profiling ever shows the cost
   * (#744).
   */
  private[formula] def readsBindingAsRange(expr: TExpr[?]): Boolean =
    // TODO(#744): cache per Let node (or compute at parse time) instead of walking per evaluation;
    // and memoize a computed name's reference per recalculation, rather than evaluating it again
    // at every range slot that reads it (resolveRangeSlot).
    // Exhaustive on purpose, with no wildcard: a new TExpr case is a "match may not be exhaustive"
    // warning here until it is classified as a wrapper to descend through or a leaf.
    expr match
      case call: TExpr.Call[?] =>
        ReferenceBaseFunctions.contains(call.spec.name) ||
        call.spec.argSpec.toValues(call.args).exists {
          case ArgValue.Range(TExpr.RangeLocation.Binding(_)) => true
          case ArgValue.Expr(arg) => readsBindingAsRange(arg)
          case ArgValue.Range(_) | ArgValue.Cells(_) => false
        }
      case chain @ (_: TExpr.Add | _: TExpr.Sub | _: TExpr.Mul | _: TExpr.Div | _: TExpr.Pow |
          _: TExpr.Concat | _: TExpr.Eq[?] | _: TExpr.Neq[?] | _: TExpr.Lt[?] | _: TExpr.Lte[?] |
          _: TExpr.Gt[?] | _: TExpr.Gte[?]) =>
        BinarySpine.operandList(chain).exists(readsBindingAsRange)
      case TExpr.Let(letBindings, body) =>
        letBindings.exists((_, value) => readsBindingAsRange(value)) || readsBindingAsRange(body)
      case TExpr.Aggregate(_, location) =>
        location match
          case TExpr.RangeLocation.Binding(_) => true
          case _ => false
      case TExpr.Coerced(inner, _) => readsBindingAsRange(inner)
      case TExpr.UnaryPlus(inner) => readsBindingAsRange(inner)
      case TExpr.Percent(inner) => readsBindingAsRange(inner)
      case TExpr.ToInt(inner) => readsBindingAsRange(inner)
      case TExpr.DateToSerial(inner) => readsBindingAsRange(inner)
      case TExpr.DateTimeToSerial(inner) => readsBindingAsRange(inner)
      // leaves: no sub-expression, so no range slot below them
      case _: TExpr.Lit[?] | _: TExpr.Ref[?] | _: TExpr.PolyRef | _: TExpr.SheetRef[?] |
          _: TExpr.SheetPolyRef | _: TExpr.ExternalRef | _: TExpr.ExternalRange |
          _: TExpr.RangeRef | _: TExpr.SheetRange | _: TExpr.ErrorLit | TExpr.Missing |
          _: TExpr.BindingRef | _: TExpr.CoercedBindingRef[?] | _: TExpr.NameRef |
          _: TExpr.SheetNameRef =>
        false

  /**
   * Whether `expr` reads a reference's position or shape anywhere (a call flagged
   * `requiresReferenceArguments`: ROW, COLUMN, ROWS, COLUMNS), so an array-mode LET binds its
   * reference-denoting values as references. The walk covers every node with children: calls, the
   * binary operators (unary minus is `0 - x`), the unary wrappers, nested LETs and coercions; array
   * constants are literals and hold no reference.
   */
  private[formula] def inspectsReference(expr: TExpr[?]): Boolean = expr match
    case call: TExpr.Call[?] =>
      call.spec.flags.requiresReferenceArguments || call.spec.argSpec.toValues(call.args).exists {
        case ArgValue.Expr(arg) => inspectsReference(arg)
        case _ => false
      }
    case chain if BinarySpine.isBinary(chain) =>
      BinarySpine.operandList(chain).exists(inspectsReference)
    case TExpr.Let(letBindings, body) =>
      letBindings.exists((_, value) => inspectsReference(value)) || inspectsReference(body)
    case TExpr.Coerced(inner, _) => inspectsReference(inner)
    case TExpr.UnaryPlus(inner) => inspectsReference(inner)
    case TExpr.Percent(inner) => inspectsReference(inner)
    case TExpr.ToInt(inner) => inspectsReference(inner)
    case TExpr.DateToSerial(inner) => inspectsReference(inner)
    case TExpr.DateTimeToSerial(inner) => inspectsReference(inner)
    case _ => false

  /** Maximum recursion depth for cross-sheet formula evaluation (GH-161 cycle protection). */
  private val MaxCrossSheetRecursionDepth = 100

  // The failures a cycle with no placement guard surfaces as, shared with [[isCycleFailure]]
  private val CrossSheetDepthExceeded = "Cross-sheet formula recursion depth exceeded"
  private val SpillDepthExceeded = "Spill reference recursion depth exceeded"
  private val NameCycleDetected = "Defined name cycle detected"

  /**
   * GH-715: whether `error` is how a cycle surfaces outside a positioned evaluation: a recursion
   * depth guard (cells, or a spill reference, reading each other) or a defined-name cycle. Each is
   * an `EvalFailed` the cycle's readers see; it is not an internal failure.
   *
   * Matched by message prefix, the three raise sites sharing the constants above: a dedicated
   * marker would be a new case or field on the published `EvalError`, a binary break.
   * LookupUncachedKeysSpec pins each raise site, so a rewording that loses the match fails there.
   * TODO(#751): a structured cycle marker on `EvalError` replaces the prefixes at the next
   * binary-breaking release.
   */
  private[formula] def isCycleFailure(error: EvalError): Boolean = error match
    case EvalError.EvalFailed(message, _) =>
      message.startsWith(CrossSheetDepthExceeded) || message.startsWith(SpillDepthExceeded) ||
      message.startsWith(NameCycleDetected)
    case _ => false

  /** Resolve range cells with the same cache, recursion and context rules as scalar references. */
  private[formula] def cellValueReader(
    targetSheet: Sheet,
    clock: Clock,
    workbook: Option[Workbook],
    depth: Int,
    rng: Rng,
    memo: EvalMemo,
    workbookPath: Option[String],
    aggregateMemo: Option[AggregateMemo]
  ): ARef => Either[EvalError, CellValue] = at =>
    targetSheet(at).value match
      case CellValue.Formula(_, Some(cached), _) => Right(cached)
      case CellValue.Formula(_, None, _: FormulaKind.DataTable) => Right(CellValue.Empty)
      case CellValue.Formula(expression, None, kind) =>
        memo.getOrCompute(targetSheet, at) {
          evalCrossSheetFormula(
            expression,
            targetSheet,
            clock,
            workbook,
            depth,
            rng,
            memo,
            workbookPath,
            aggregateMemo,
            Some(at),
            isArrayRecord(kind)
          )
        }
      case value => Right(value)

  /** An ArrayFormula record (CSE or dynamic-array anchor) evaluates in array mode. */
  private[formula] def isArrayRecord(kind: FormulaKind): Boolean = kind match
    case _: FormulaKind.ArrayFormula => true
    case _ => false

  /**
   * Evaluate a formula string from a cross-sheet reference (GH-161).
   *
   * When a cross-sheet reference points to a formula cell without a cached value, we need to
   * recursively parse and evaluate that formula against the target sheet.
   *
   * @param formulaStr
   *   The formula string (without leading =)
   * @param targetSheet
   *   The sheet containing the formula cell
   * @param clock
   *   Clock for date/time functions
   * @param workbook
   *   Workbook context for nested cross-sheet references
   * @param depth
   *   Current recursion depth (for cycle protection)
   * @param memo
   *   the evaluation's memo, which carries its placement guard (GH-715). No default: a fresh memo
   *   here would carry no guard, and the placement could evaluate as a precedent of itself
   * @param arrayRecord
   *   the cell is an ArrayFormula record: it evaluates in array mode and reads as element (0,0)
   * @return
   *   Either evaluation error or computed CellValue
   */
  private[formula] def evalCrossSheetFormula(
    formulaStr: String,
    targetSheet: Sheet,
    clock: Clock,
    workbook: Option[Workbook],
    depth: Int = 0,
    rng: Rng = Rng.system,
    memo: EvalMemo,
    workbookPath: Option[String] = None,
    aggregateMemo: Option[AggregateMemo] = None,
    currentCell: Option[ARef] = None,
    arrayRecord: Boolean = false
  ): Either[EvalError, CellValue] =
    boundary:
      // GH-161 review: Add recursion depth limit to prevent stack overflow on circular refs
      if depth > MaxCrossSheetRecursionDepth then
        break(
          Left(
            EvalError.EvalFailed(
              s"$CrossSheetDepthExceeded (max: $MaxCrossSheetRecursionDepth). Possible circular reference.",
              None
            )
          )
        )
      // GH-715: a positioned evaluation's placement is never evaluated as a precedent of itself
      currentCell.flatMap(memo.guard(targetSheet, _)).foreach(error => break(Left(error)))

      val counters = aggregateMemo.fold(RecalcCounters.disabled)(_.counters)
      counters.time(RecalcCounters.Phase.Parse)(FormulaParser.parse(formulaStr)) match
        case Left(parseErr) =>
          Left(
            EvalError.EvalFailed(
              s"Failed to parse cross-sheet formula: ${ParseError.toXLError(parseErr, formulaStr).message}",
              None
            )
          )
        case Right(expr) =>
          // Recursively evaluate with depth-aware evaluator (GH-161 cycle protection).
          // The referenced formula is a separate lexical unit: the rng threads through, but LET
          // bindings never leak across formula boundaries (fresh empty environment). The memo
          // threads too (GH-346): every cell in this recursion tree evaluates at most once, and
          // the workbook path (GH-424) so nested CELL("filename") calls see the same location.
          new EvaluatorWithDepth(
            depth + 1,
            allowArrayResults = arrayRecord,
            rng = rng,
            memo = Some(memo),
            workbookPath = workbookPath,
            aggregateMemo = aggregateMemo,
            guard = memo.placement
          )
            .eval(expr, targetSheet, clock, workbook, currentCell) match
            case Right(result) => Right(EvalResult.toCellValue(result))
            // GH-344: an error-computing precedent delivers its Excel error VALUE to readers
            // (the cell-mediated cascade); host failures stay loud Lefts.
            case Left(evalError) =>
              EvalError.toErrorValue(evalError) match
                case Some(code) => Right(CellValue.Error(code))
                case None => Left(evalError)

  /**
   * GH-655: evaluate a spill anchor's formula text as an ARRAY at the anchor's own position on its
   * sheet — the value of the spill reference `anchor#` when the file carries no cached spill. The
   * result is the raw evaluation (an ArrayResult for an array-valued formula, a scalar otherwise);
   * the caller decides what a scalar means. The depth guard, memo, rng and workbook path thread
   * exactly as in [[evalCrossSheetFormula]]; LET bindings do not cross the formula boundary.
   */
  private[formula] def evalSpillFormula(
    formulaStr: String,
    targetSheet: Sheet,
    clock: Clock,
    workbook: Option[Workbook],
    depth: Int,
    rng: Rng,
    memo: EvalMemo,
    workbookPath: Option[String],
    aggregateMemo: Option[AggregateMemo],
    anchor: ARef
  ): Either[EvalError, Any] =
    if depth > MaxCrossSheetRecursionDepth then
      Left(
        EvalError.EvalFailed(
          s"$SpillDepthExceeded (max: $MaxCrossSheetRecursionDepth). " +
            s"Possible circular reference through ${anchor.toA1}#.",
          None
        )
      )
    else if memo.guard(targetSheet, anchor).isDefined then
      // GH-715: the placement's own spill, `D5#` placed at D5, reads the placement
      Left(EvalError.CircularRef(List(anchor)))
    else
      FormulaParser.parse(formulaStr) match
        case Left(parseErr) =>
          Left(
            EvalError.EvalFailed(
              s"${anchor.toA1}#: the anchor's formula does not parse: " +
                ParseError.toXLError(parseErr, formulaStr).message,
              None
            )
          )
        case Right(expr) =>
          new EvaluatorWithDepth(
            depth + 1,
            allowArrayResults = true,
            rng = rng,
            memo = Some(memo),
            workbookPath = workbookPath,
            aggregateMemo = aggregateMemo,
            guard = memo.placement
          ).eval(expr, targetSheet, clock, workbook, Some(anchor))

  /**
   * GH-715: the cell a positioned evaluation places its formula in (`at` on the sheet named
   * `sheet`), and whether anything read it. The evaluator carries it (every derived evaluator gets
   * it, a required constructor argument of [[EvaluatorWithDepth]]) as well as the memo, so a reader
   * that has to start a fresh memo still guards the placement, and every memo of the evaluation
   * shares the one record of the read. Single-threaded, like the memo.
   */
  private[formula] final class PlacementGuard(sheet: SheetName, at: ARef):
    // Local mutation like EvalMemo's map: one evaluation pass, one thread
    @SuppressWarnings(Array("org.wartremover.warts.Var"))
    private var wasRead = false

    /** Whether anything read the placement. */
    def read: Boolean = wasRead

    /** The read's error when `ref` on `onSheet` is the placement, recording the read. */
    def check(onSheet: Sheet, ref: ARef): Option[EvalError] =
      // case-insensitive, as Excel compares sheet names and as the rest of GH-715 does
      val guarded = ref == at && onSheet.name.value.equalsIgnoreCase(sheet.value)
      if guarded then wasRead = true
      Option.when(guarded)(EvalError.CircularRef(List(ref)))

  /**
   * GH-346: per-pass memo for recursively evaluated uncached formula cells.
   *
   * Recursive evaluation of a `CellValue.Formula(_, None)` reference re-derived the referenced cell
   * once per PATH through the dependency graph — exponential on multi-branch recursive chains (an
   * LBO debt schedule: each period's balance read by several formulas, chained period over period).
   * The memo makes each distinct cell evaluate at most once per evaluation pass, which is also
   * Excel's one-computation-per-cell-per-recalc semantics.
   *
   * Scope and purity: a memo is created where recursion begins (an uncached reference or a function
   * reading uncached cells at depth 0) and threads only through the derived `EvaluatorWithDepth`
   * instances and `EvalContext`s of that pass, so it never outlives a single top-level evaluation
   * and is invisible at the public API — local mutation only, the same pattern as the iterative
   * Tarjan engine in DependencyGraph. Sharing is per recursion tree, not per top-level eval:
   * sibling depth-0 references (`=A1+A1`) each start their own memo and re-walk the shared subtree
   * once each — bounded and linear per subtree, a constant-factor cost accepted to keep the memo's
   * lifetime trivially safe. Entries key by Sheet IDENTITY, then ref: two same-named Sheet
   * snapshots (a recalc temp sheet vs the original workbook copy) memoize independently, so a stale
   * snapshot can never serve values for a newer one.
   *
   * Not thread-safe: an evaluation pass is single-threaded and the memo never escapes it — revisit
   * if evaluation is ever parallelized internally.
   *
   * Errors memoize too — a cycle otherwise re-burns the full recursion-depth guard once per path,
   * which is itself exponential. A depth-guard error consequently records by first-visit order;
   * that is observable only on chains deeper than the guard, where results were already
   * path-dependent (and where whole-workbook `recalculate()`, which orders iteratively, is the
   * supported path).
   *
   * GH-715: `placement` guards one cell, the cell a positioned evaluation places its formula in
   * ([[placedInstance]]). That cell holds an uncached formula, so every read of it, whatever
   * computed it (INDIRECT, OFFSET, a defined name, another sheet's formula), comes here, and is the
   * circular reference the plain cell there would be: `CircularRef(List(at))`, a one-cell path no
   * dependency-graph cycle has, never the cell's old content.
   */
  private[formula] final class EvalMemo(
    val placement: Option[PlacementGuard] = None,
    // #750: how ArrayArithmetic.readRange reads a range in this pass; Auto (by its size) except in
    // the sparse-equals-dense laws, through arrayInstanceReading
    val rangeRead: ArrayArithmetic.RangeRead = ArrayArithmetic.RangeRead.Auto
  ):
    private val bySheet =
      new java.util.IdentityHashMap[
        Sheet,
        scala.collection.mutable.HashMap[ARef, Either[EvalError, CellValue]]
      ]()

    def getOrCompute(sheet: Sheet, at: ARef)(
      compute: => Either[EvalError, CellValue]
    ): Either[EvalError, CellValue] =
      guard(sheet, at).fold(memoized(sheet, at, compute))(Left(_))

    private val boxes =
      new java.util.IdentityHashMap[
        Sheet,
        scala.collection.mutable.HashMap[CellRange, Option[CellRange]]
      ]()

    /**
     * #750: the populated box of `range` on `sheet` ([[ArrayArithmetic.readRange]]), found once per
     * pass: a formula that reads one whole column twice scans the sheet's keys once.
     */
    def populatedBox(sheet: Sheet, range: CellRange)(
      compute: => Option[CellRange]
    ): Option[CellRange] =
      val forSheet = boxes.computeIfAbsent(sheet, _ => scala.collection.mutable.HashMap.empty)
      forSheet.getOrElseUpdate(range, compute)

    /**
     * The guarded placement's error when `at` on `sheet` is it. Reads that evaluate an uncached
     * formula without this memo's lookup ([[evalCrossSheetFormula]], [[evalSpillFormula]]) ask too.
     */
    def guard(sheet: Sheet, at: ARef): Option[EvalError] =
      placement.flatMap(_.check(sheet, at))

    private def memoized(
      sheet: Sheet,
      at: ARef,
      compute: => Either[EvalError, CellValue]
    ): Either[EvalError, CellValue] =
      val forSheet =
        bySheet.computeIfAbsent(sheet, _ => scala.collection.mutable.HashMap.empty)
      forSheet.get(at) match
        case Some(hit) => hit
        case None =>
          // Compute BEFORE storing (never inside a map callback): the computation itself may
          // recurse into this memo for other cells of the same sheet.
          val computed = compute
          forSheet.update(at, computed)
          computed

  /** Observable counters for focused aggregate-memo tests and profiling. */
  private[formula] final case class AggregateMemoStats(
    hits: Long,
    fills: Long,
    bypasses: Long,
    entries: Int
  )

  /**
   * Distinguishes the two aggregate evaluators, whose coercion contracts are not interchangeable.
   */
  private[formula] enum AggregateMemoMode:
    case FunctionCall
    case TypedNode

  /**
   * One calculation generation's memo for a single raw-range aggregate.
   *
   * A memo instance belongs to exactly one ONE-PASS region of a workbook recalculation. Keys
   * therefore use the logical sheet name rather than retaining every immutable `Sheet` snapshot:
   * the formula graph guarantees that every formula cell in an explicit range is evaluated before
   * its aggregate consumer, after which that range cannot change again in the one-pass region. This
   * is what lets same-sheet sequential consumers reuse a result even though each cache write
   * creates a new `Sheet` object. Range anchors are intentionally absent because they affect
   * formula dragging, never values read.
   *
   * GH-537: the one-pass regions are (a) a non-iterative generation as a whole, (b) every acyclic
   * `Straight` segment of an iterative condensation walk — a reader of a range touching cyclic cell
   * C has a graph edge to C, so it is ordered after C's component has finalized, and members that
   * failed are stripped to uncached formulas, which `cacheable` refuses — and (c) ONE Jacobi round
   * inside a fixpoint, where every member is overlaid cached and results apply at round end. A memo
   * must never outlive its region: the fixpoint engine takes a fresh one per round, and a
   * Gauss–Seidel sweep (members changing mid-round) takes none.
   *
   * A per-key entry synchronizes the first eligible fold. This gives parallel waves single-flight
   * behavior instead of allowing every cold reader to scan the same range before a TrieMap
   * `putIfAbsent` wins. Ineligible ranges are marked once, then every reader computes outside the
   * monitor. Results are immutable `Either[EvalError, BigDecimal]` values; errors memoize exactly
   * like successes.
   *
   * #767: `counters` is the pass's [[RecalcCounters]]; every production pass uses the disabled
   * instance, and only the measured seams of `WorkbookEvaluator` supply an enabled one. `passCache`
   * holds the pass's used ranges, static name resolutions and column slices ([[PassCache]]); it
   * exists when the memo caches aggregates and `columnCache` is on (the default; parity tests turn
   * it off), with `limits` bounding what it holds. `criteriaIndex` (on by default; parity tests
   * turn it off) lets criteria calls answer from the slices' equality indexes.
   */
  private[formula] final class AggregateMemo(
    cachesAggregates: Boolean = true,
    val counters: RecalcCounters = RecalcCounters.disabled,
    columnCache: Boolean = true,
    limits: PassCache.Limits = PassCache.DefaultLimits,
    criteriaIndex: Boolean = true
  ):
    private val entries = TrieMap.empty[AggregateMemoKey, AggregateMemoEntry]

    /** #767: the pass's caches; None for a memo that caches nothing or has them switched off. */
    val passCache: Option[PassCache] =
      Option.when(cachesAggregates && columnCache)(new PassCache(limits, counters, criteriaIndex))

    /**
     * GH-711: SUMPRODUCT's blank tail asks whether a defined name reaches a random draw. Building
     * that verdict walks the name table (a resolver, a sheet canonicaliser, a parse of every body
     * reached), so the generation builds it once and every formula reuses it. A generation never
     * changes the name table or the sheet list, only cells; the entry is still keyed on both, so a
     * memo that meets another workbook rebuilds instead of answering for the wrong one. The walk
     * memoizes verdicts in a mutable map, so a parallel wave queries it under the entry's monitor.
     */
    private val drawingNames =
      new java.util.concurrent.atomic.AtomicReference[Option[NameVerdicts]](None)

    def namesReachingDraw(
      wb: Workbook,
      build: Workbook => (SheetName, Option[SheetName], String) => NameWalk.Reach
    ): (SheetName, Option[SheetName], String) => NameWalk.Reach =
      val names = wb.metadata.definedNames
      val sheets = wb.sheets.map(_.name)
      // get-then-set, not compare-and-set: two readers that miss together both build, which
      // costs a duplicate walk and never a wrong verdict (each builds from the workbook it holds)
      val entry = drawingNames.get() match
        case Some(known) if known.matches(names, sheets) => known
        case _ =>
          val built = NameVerdicts(names, sheets, build(wb))
          drawingNames.set(Some(built))
          built
      entry.verdict

    /**
     * GH-695: each spill anchor's raw evaluation as this generation's fold computed it — an array
     * is its spill, a scalar is not a spill anchor, an error is that error, exactly as `x#` reads
     * an anchor it evaluates itself. The threaded sheet caches only the anchor's top-left element,
     * and the recorded extent's cells hold the previous generation's spill (or nothing), so an `x#`
     * reader looks here first. Off until [[trackSpills]]: a book without `x#` readers keeps no
     * arrays.
     */
    private val spills = TrieMap.empty[(SheetName, ARef), Either[EvalError, Any]]
    private val tracksSpills = new java.util.concurrent.atomic.AtomicBoolean(false)

    def trackSpills(): Unit = tracksSpills.set(true)

    def recordSpill(sheet: SheetName, anchor: ARef, raw: Either[EvalError, Any]): Unit =
      if tracksSpills.get() then spills.update((sheet, anchor), raw)

    def recordedSpill(sheet: SheetName, anchor: ARef): Option[Either[EvalError, Any]] =
      spills.get((sheet, anchor))
    private val hitCount = new AtomicLong(0L)
    private val fillCount = new AtomicLong(0L)
    private val bypassCount = new AtomicLong(0L)

    def getOrCompute(
      targetSheet: Sheet,
      range: CellRange,
      aggregatorId: String,
      mode: AggregateMemoMode
    )(
      compute: => Either[EvalError, BigDecimal]
    ): Either[EvalError, BigDecimal] =
      if !cachesAggregates then compute
      else
        val key = AggregateMemoKey(targetSheet.name, range.start, range.end, aggregatorId, mode)
        val entry = entries.getOrElseUpdate(key, new AggregateMemoEntry)
        entry.lookupOrCompute(cacheable(targetSheet, range))(compute) match
          case AggregateMemoLookup.Hit(value) =>
            hitCount.incrementAndGet()
            value
          case AggregateMemoLookup.Filled(value) =>
            fillCount.incrementAndGet()
            value
          case AggregateMemoLookup.Bypass =>
            bypassCount.incrementAndGet()
            compute

    def stats: AggregateMemoStats =
      AggregateMemoStats(
        hits = hitCount.get(),
        fills = fillCount.get(),
        bypasses = bypassCount.get(),
        entries = entries.size
      )

    private def cacheable(targetSheet: Sheet, range: CellRange): Boolean =
      def isUncachedFormula(value: CellValue): Boolean = value match
        case CellValue.Formula(_, None, _) => true
        case _ => false

      val containsUncachedFormula =
        if range.cellCount <= targetSheet.cells.size.toLong then
          range.cells.exists(ref =>
            targetSheet.cells.get(ref).exists(c => isUncachedFormula(c.value))
          )
        else
          targetSheet.cells.iterator.exists { case (ref, cell) =>
            range.contains(ref) && isUncachedFormula(cell.value)
          }
      !containsUncachedFormula

  /** Generation-local logical key; anchors and transient immutable Sheet identities are absent. */
  private final case class AggregateMemoKey(
    targetSheet: SheetName,
    start: ARef,
    end: ARef,
    aggregatorId: String,
    mode: AggregateMemoMode
  )

  private enum AggregateMemoState:
    case Unchecked
    case Uncacheable
    case Cached(value: Either[EvalError, BigDecimal])

  private enum AggregateMemoLookup:
    case Hit(value: Either[EvalError, BigDecimal])
    case Filled(value: Either[EvalError, BigDecimal])
    case Bypass

  /** One workbook's name verdicts, as [[AggregateMemo.namesReachingDraw]] caches them. */
  private final class NameVerdicts(
    names: Vector[DefinedName],
    sheets: Vector[SheetName],
    walk: (SheetName, Option[SheetName], String) => NameWalk.Reach
  ):
    def matches(otherNames: Vector[DefinedName], otherSheets: Vector[SheetName]): Boolean =
      ((names eq otherNames) || names == otherNames) && sheets == otherSheets

    // one lookup at a time holds the monitor: the walk's verdict map is not thread-safe, and a
    // lookup is a map hit once the name has been walked
    val verdict: (SheetName, Option[SheetName], String) => NameWalk.Reach =
      (from, qualifier, name) => synchronized(walk(from, qualifier, name))

  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  private final class AggregateMemoEntry:
    private var state: AggregateMemoState = AggregateMemoState.Unchecked

    def lookupOrCompute(
      cacheable: => Boolean
    )(
      compute: => Either[EvalError, BigDecimal]
    ): AggregateMemoLookup = synchronized {
      state match
        case AggregateMemoState.Cached(value) => AggregateMemoLookup.Hit(value)
        case AggregateMemoState.Uncacheable => AggregateMemoLookup.Bypass
        case AggregateMemoState.Unchecked =>
          if cacheable then
            val value = compute
            state = AggregateMemoState.Cached(value)
            AggregateMemoLookup.Filled(value)
          else
            state = AggregateMemoState.Uncacheable
            AggregateMemoLookup.Bypass
    }

/**
 * GH-344: the single typed-result → CellValue table, shared by the cross-sheet recursion
 * (Evaluator.evalCrossSheetFormula) and the public evaluation boundary (SheetEvaluator) so a value
 * reads identically wherever it lands in a cell.
 */
private[formula] object EvalResult:
  /**
   * Convert a typed evaluation result to its CellValue form. Scalar context: an ArrayResult
   * collapses to its top-left value (Excel non-array entry), Empty when empty.
   */
  def toCellValue(result: Any): CellValue = result match
    case cv: CellValue => cv // IFERROR and similar functions return CellValue directly
    case bd: BigDecimal => CellValue.Number(bd)
    case s: String => CellValue.Text(s)
    case b: Boolean => CellValue.Bool(b)
    case i: Int => CellValue.Number(BigDecimal(i))
    case ld: java.time.LocalDate => CellValue.DateTime(ld.atStartOfDay())
    case ldt: java.time.LocalDateTime => CellValue.DateTime(ldt)
    case ar: ArrayResult => if ar.isEmpty then CellValue.Empty else ar(0, 0)
    case other => CellValue.Text(other.toString)

/**
 * The evaluation boundary every [[Evaluator]] factory returns (#681 item 5): the outermost frame of
 * one evaluation, containing a throwable that escapes it as `EvalError.EvalFailed`
 * ([[EvalDefect]]). The implementation recurses through its own unguarded `eval` and derives
 * unguarded [[EvaluatorWithDepth]] instances, so this is the only frame that catches — no formula
 * construct (IFERROR, ISERROR, the array broadcasts) can swallow a defect as an ordinary Left.
 */
private final class TotalEvaluator(underlying: Evaluator) extends Evaluator:
  def eval[A](
    expr: TExpr[A],
    sheet: Sheet,
    clock: Clock = Clock.system,
    workbook: Option[Workbook] = None,
    currentCell: Option[ARef] = None
  ): Either[EvalError, A] =
    EvalDefect.guard(currentCell)(underlying.eval(expr, sheet, clock, workbook, currentCell))

  override private[formula] def withArrayResults: Evaluator =
    TotalEvaluator(underlying.withArrayResults)

  override private[formula] def generationMemo: Option[Evaluator.AggregateMemo] =
    underlying.generationMemo

/**
 * Private implementation of Evaluator.
 *
 * Implements all TExpr cases with proper error handling and short-circuit semantics.
 *
 * @param bindings
 *   GH-193: the LET environment — values of in-scope bindings, keyed by declared name. Threaded
 *   through instance state so every recursive eval (including function-argument evaluation via
 *   EvalContext) sees the same environment.
 * @param rng
 *   GH-115: randomness capability for RAND/RANDBETWEEN, threaded like bindings so derived
 *   evaluators (array args, cross-sheet recursion, LET bodies) draw from the same source.
 * @param resolvingNames
 *   GH-384: the defined names currently being resolved on this evaluation path — the name→name
 *   cycle guard, keyed on the resolved name (#691, [[Evaluator.NameTrail]]). A refersTo chain that
 *   revisits a member (aa → bb → aa) is a clean per-cell error instead of unbounded recursion, and
 *   so is one deeper than [[NameWalk.MaxDepth]] names. It carries the outermost name's value memo
 *   ([[Evaluator.NameMemo]]). Cell-mediated cycles (name → cell → name) are covered separately by
 *   the depth-guarded cross-sheet recursion.
 * @param workbookPath
 *   GH-424: the workbook's saved location if the embedder knows one, surfaced to functions via
 *   EvalContext (CELL("filename")). None reproduces Excel's pre-save behavior.
 */
private class EvaluatorImpl(
  allowArrayResults: Boolean = false,
  bindings: Map[String, Any] = Map.empty,
  rng: Rng = Rng.system,
  resolvingNames: Evaluator.NameTrail = Evaluator.NameTrail.empty,
  workbookPath: Option[String] = None,
  aggregateMemo: Option[Evaluator.AggregateMemo] = None,
  // SUMPRODUCT's shared extent, ambient below it: scope documented on EvalContext.Operands
  arrayRangeBounds: Option[CellRange => CellRange] = None,
  // #767: the pass slot of a top-level evaluation at its own slot; derived evaluators of the same
  // formula keep it, a precedent evaluated on demand (a deeper evaluator) does not
  passPosition: EvalPosition = EvalPosition.None
) extends Evaluator:
  /** Current recursion depth for cross-sheet formula evaluation. */
  protected def currentDepth: Int = 0

  /**
   * GH-346: the current pass's memo for recursively evaluated uncached formula cells — None at the
   * top level (a memo is created at the first point recursion can begin and threads through every
   * derived evaluator of the pass via EvaluatorWithDepth).
   */
  protected def memoOpt: Option[Evaluator.EvalMemo] = None

  /** GH-715: the placement a positioned evaluation guards; None outside one. */
  protected def placementGuard: Option[Evaluator.PlacementGuard] = None

  /**
   * The pass's memo, or a fresh one where recursion begins here. GH-715: a fresh memo inherits the
   * evaluator's placement guard, so the guard never depends on a memo having been threaded.
   */
  protected def passMemo: Evaluator.EvalMemo =
    memoOpt.getOrElse(new Evaluator.EvalMemo(placementGuard))

  /** One workbook recalculation generation's raw-range aggregate memo, absent for public eval. */
  protected def aggregateMemoOpt: Option[Evaluator.AggregateMemo] = aggregateMemo

  override private[formula] def generationMemo: Option[Evaluator.AggregateMemo] = aggregateMemo

  /** #767: the pass's work counters; disabled outside a measured recalculation. */
  protected def passCounters: RecalcCounters =
    aggregateMemo.fold(RecalcCounters.disabled)(_.counters)

  override private[formula] def withArrayResults: Evaluator =
    new EvaluatorImpl(
      allowArrayResults = true,
      bindings,
      rng,
      resolvingNames,
      workbookPath,
      aggregateMemo,
      arrayRangeBounds = arrayRangeBounds,
      passPosition = passPosition
    )
  // Suppress asInstanceOf warning for GADT type handling (required for type parameter erasure)
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  def eval[A](
    expr: TExpr[A],
    sheet: Sheet,
    clock: Clock = Clock.system,
    workbook: Option[Workbook] = None,
    currentCell: Option[ARef] = None
  ): Either[EvalError, A] =
    // @unchecked: GADT exhaustivity - PolyRef should be resolved before evaluation
    (expr: @unchecked) match
      // ===== PolyRef Handling (Same-Sheet Reference) =====
      //
      // PolyRef should be resolved to typed Ref during parsing (see resolveTopLevelPolyRef
      // in FormulaParser). If we reach this case, it means a PolyRef escaped resolution,
      // which is a programming error. Return an error instead of using unsafe asInstanceOf.
      //
      case TExpr.PolyRef(at, _) =>
        Left(
          EvalError.EvalFailed(
            s"Unresolved PolyRef at ${(at: ARef).toA1} - should have been resolved during parsing",
            None
          )
        )

      // ===== Sheet-Qualified References (Cross-Sheet) =====
      //
      // SheetPolyRef should be resolved to typed SheetRef during parsing (see
      // resolveTopLevelPolyRef in FormulaParser). If we reach this case, it means
      // a SheetPolyRef escaped resolution, which is a programming error.
      // Return an error instead of using unsafe asInstanceOf.
      //
      case TExpr.SheetPolyRef(sheetName, at, _) =>
        Left(
          EvalError.EvalFailed(
            s"Unresolved SheetPolyRef at ${SheetName.quoteForFormula(sheetName.value)}!${(at: ARef).toA1} - should have been resolved during parsing",
            None
          )
        )

      case TExpr.SheetRef(sheetName, at, _, decode) =>
        // SheetRef: resolve cell from target sheet in workbook
        workbook match
          case None =>
            // GH-280: quote cell-ref-shaped sheet names so diagnostics read unambiguously
            val refStr = s"${SheetName.quoteForFormula(sheetName.value)}!${(at: ARef).toA1}"
            Left(Evaluator.missingWorkbookError(refStr))
          case Some(wb) =>
            wb(sheetName) match
              case Left(err) =>
                Left(Evaluator.sheetNotFoundError(sheetName, err))
              case Right(targetSheet) =>
                val cell = targetSheet(at)
                // GH-161: Handle formula cells without cached values by recursively evaluating
                cell.value match
                  // GH-430: TABLE(...) display text is not evaluable — the decoder path treats
                  // the uncached record like any cached formula (pinned semantics, no parse)
                  case CellValue.Formula(_, None, _: FormulaKind.DataTable) =>
                    decodeOrCarried(at, cell, decode, Some(sheetName))
                  case CellValue.Formula(formulaStr, None, kind) =>
                    // Formula has no cached value - parse and evaluate against target sheet
                    // GH-161 review: Apply decoder to Cell with evaluated result (type-safe)
                    // GH-161 review: Pass currentDepth for cycle protection
                    // GH-346: memoized per pass — a cell evaluates once, not once per path
                    val memo = passMemo
                    memo
                      .getOrCompute(targetSheet, at) {
                        // at its own position: an implicit intersection there reads its own row
                        Evaluator.evalCrossSheetFormula(
                          formulaStr,
                          targetSheet,
                          clock,
                          workbook,
                          currentDepth,
                          rng,
                          memo,
                          workbookPath,
                          aggregateMemoOpt,
                          Some(at),
                          Evaluator.isArrayRecord(kind)
                        )
                      }
                      .flatMap(evaluatedValue =>
                        decodeOrCarried(at, Cell(at, evaluatedValue), decode, Some(sheetName))
                      )
                  case _ =>
                    // Cached formula or non-formula cell - use decoder
                    decodeOrCarried(at, cell, decode, Some(sheetName))

      // ===== GH-353: External-Workbook References =====
      //
      // The target workbook is not loaded, so these can never evaluate. Cells CONTAINING them
      // are pinned to their Excel-written cached value upstream (SheetEvaluator) and never
      // reach this point; direct evaluation and uncached cells get a clear per-cell error.
      case TExpr.ExternalRef(index, name, at, _) =>
        Left(Evaluator.externalRefUnsupported(s"[$index]$name!${(at: ARef).toA1}"))
      case TExpr.ExternalRange(index, name, range, _) =>
        Left(Evaluator.externalRefUnsupported(s"[$index]$name!${range.toA1}"))

      // A range where a value is needed: its values in array mode, the implicitly intersected
      // cell in a plain cell (`=A1:A10` in row 5 is A5, as Excel computes a legacy formula)
      case TExpr.SheetRange(sheetName, range, _) =>
        Evaluator
          .resolveRangeLocation(
            TExpr.RangeLocation.CrossSheet(sheetName, range),
            sheet,
            workbook,
            resolvingNames
          )
          .flatMap { case (targetSheet, _) =>
            derefRange(range, targetSheet, clock, workbook, currentCell)
          }
          .asInstanceOf[Either[EvalError, A]]

      case TExpr.RangeRef(range, _) =>
        derefRange(range, sheet, clock, workbook, currentCell).asInstanceOf[Either[EvalError, A]]

      // ===== Literals =====
      case TExpr.Lit(value) =>
        // Literal: return value directly (identity law)
        Right(value)

      // GH-612: an error literal IS the error value it names — it travels the Left channel like
      // any other Excel error value and promotes to CellValue.Error at the cell boundary
      case TExpr.ErrorLit(error, _) =>
        Left(EvalError.ErrorValue(error))

      // GH-603: an omitted argument is a blank — every argument slot that holds a value wraps it
      // in Coerced at parse time (0, "", FALSE per target; value slots read 0: IF(TRUE,,5) is 0)
      case TExpr.Missing =>
        Right(CellValue.Empty).asInstanceOf[Either[EvalError, A]]

      // ===== Cell References =====
      case TExpr.Ref(at, _, decode) =>
        // Ref: resolve cell, decode value with codec
        // Note: sheet(at) returns empty cell if not present, decode handles empty cells
        val cell = sheet(at)
        // GH-208: Handle same-sheet formula cells without cached values by recursively evaluating
        cell.value match
          // GH-430: TABLE(...) display text is not evaluable — decoder path, pinned semantics
          case CellValue.Formula(_, None, _: FormulaKind.DataTable) =>
            decodeOrCarried(at, cell, decode)
          case CellValue.Formula(formulaStr, None, kind) =>
            // GH-346: memoized per pass — a cell evaluates once, not once per path
            val memo = passMemo
            memo
              .getOrCompute(sheet, at) {
                // at its own position: an implicit intersection there reads its own row
                Evaluator
                  .evalCrossSheetFormula(
                    formulaStr,
                    sheet,
                    clock,
                    workbook,
                    currentDepth,
                    rng,
                    memo,
                    workbookPath,
                    aggregateMemoOpt,
                    Some(at),
                    Evaluator.isArrayRecord(kind)
                  )
              }
              .flatMap(evaluatedValue => decodeOrCarried(at, Cell(at, evaluatedValue), decode))
          case _ =>
            decodeOrCarried(at, cell, decode)

      // GH-669: an operator over array constants keeps its array in a plain cell, as Excel's legacy
      // evaluation does — only a REFERENCE is implicitly intersected there, and an array constant
      // is a value, so `SUM({1,2}*2)` is 6 and `={1,2}*10` shows 10. Such an operation (its
      // operands array constants, scalar literals and single cells, which read the same in either
      // mode) evaluates in array mode; the cell boundary still shows an array's top-left value.
      case op if !allowArrayResults && Evaluator.arrayConstantOperation(op) =>
        withArrayResults.eval(op, sheet, clock, workbook, currentCell)

      // ===== Arithmetic Operators =====
      // These support array arithmetic with broadcasting when operands are ranges or array results
      case TExpr.Add(x, y) =>
        evalArithmetic(x, y, ArrayArithmetic.add, sheet, clock, workbook, currentCell)
          .asInstanceOf[Either[EvalError, A]]

      case TExpr.Sub(x, y) =>
        evalArithmetic(x, y, ArrayArithmetic.sub, sheet, clock, workbook, currentCell)
          .asInstanceOf[Either[EvalError, A]]

      case TExpr.Mul(x, y) =>
        evalArithmetic(x, y, ArrayArithmetic.mul, sheet, clock, workbook, currentCell)
          .asInstanceOf[Either[EvalError, A]]

      case TExpr.Div(x, y) =>
        evalArithmetic(x, y, ArrayArithmetic.div, sheet, clock, workbook, currentCell)
          .asInstanceOf[Either[EvalError, A]]

      case TExpr.Pow(x, y) =>
        evalArithmetic(x, y, ArrayArithmetic.pow, sheet, clock, workbook, currentCell)
          .asInstanceOf[Either[EvalError, A]]

      // GH-374: unary plus is the identity — preserved in the AST only so the printer can
      // replicate source text byte-for-byte; at evaluation time the operand's value (scalar,
      // array, whatever) passes through untouched.
      case TExpr.UnaryPlus(e) =>
        eval(e, sheet, clock, workbook, currentCell)

      // GH-355: postfix percent is ÷100 routed through the same array machinery as the binary
      // operators, so range operands broadcast elementwise (=A1:A3% divides each cell by 100)
      // and scalars stay exact (BigDecimal(10)/100 = 0.1).
      case TExpr.Percent(e) =>
        evalArithmetic(
          e,
          TExpr.Lit(BigDecimal(100)),
          ArrayArithmetic.div,
          sheet,
          clock,
          workbook,
          currentCell
        ).asInstanceOf[Either[EvalError, A]]

      // ===== String Operators =====
      case TExpr.Concat(x, y) =>
        evalConcat(x, y, sheet, clock, workbook, currentCell).asInstanceOf[Either[EvalError, A]]

      // ===== Comparison Operators =====
      // GH-197: Use evalComparison for array-aware comparisons
      // GH-335: polymorphic Excel ordering (number < text < logical, case-insensitive text,
      // empty coerces) — op interprets the comparison sign from compareCellValues
      case TExpr.Lt(x, y) =>
        evalComparison(x, y, _ < 0, sheet, clock, workbook, currentCell)
          .asInstanceOf[Either[EvalError, A]]

      case TExpr.Lte(x, y) =>
        evalComparison(x, y, _ <= 0, sheet, clock, workbook, currentCell)
          .asInstanceOf[Either[EvalError, A]]

      case TExpr.Gt(x, y) =>
        evalComparison(x, y, _ > 0, sheet, clock, workbook, currentCell)
          .asInstanceOf[Either[EvalError, A]]

      case TExpr.Gte(x, y) =>
        evalComparison(x, y, _ >= 0, sheet, clock, workbook, currentCell)
          .asInstanceOf[Either[EvalError, A]]

      case TExpr.Eq(x, y) =>
        // Equality: polymorphic comparison (array-aware)
        evalEqualityComparison(x, y, negate = false, sheet, clock, workbook, currentCell)
          .asInstanceOf[Either[EvalError, A]]

      case TExpr.Neq(x, y) =>
        // Inequality: polymorphic comparison (array-aware)
        evalEqualityComparison(x, y, negate = true, sheet, clock, workbook, currentCell)
          .asInstanceOf[Either[EvalError, A]]

      // ===== Type Conversions =====
      case TExpr.ToInt(expr) =>
        // ToInt: total conversion to Int. The operand is statically BigDecimal, but erased
        // upstream casts can deliver other runtime values — evaluate as Any and coerce with the
        // shared Integer conventions (GH-307: fractional values TRUNCATE toward zero like Excel,
        // numeric text parses, anything else is a clean Left) per GH-193 totality.
        // An array operand (array mode) collapses to its top-left value first: an integer slot
        // takes one integer.
        eval(expr.asInstanceOf[TExpr[Any]], sheet, clock, workbook, currentCell)
          .flatMap {
            case ar: ArrayResult =>
              ScalarCoercion.collapseTo("ToInt", ar, Some(BindingCoercion.Integer))
            case value => ScalarCoercion.coerce("ToInt", value, BindingCoercion.Integer)
          }
          .asInstanceOf[Either[EvalError, A]]

      // ===== Date/Time Conversions =====
      case TExpr.DateToSerial(dateExpr) =>
        evalDateSerial(dateExpr, sheet, clock, workbook, currentCell)
          .asInstanceOf[Either[EvalError, A]]

      case TExpr.DateTimeToSerial(dtExpr) =>
        evalDateSerial(dtExpr, sheet, clock, workbook, currentCell)
          .asInstanceOf[Either[EvalError, A]]

      case TExpr.Aggregate(aggregatorId, location) =>
        // Use Aggregator typeclass to evaluate any registered aggregate function
        Aggregator.lookup(aggregatorId) match
          case None =>
            Left(EvalError.EvalFailed(s"Unknown aggregator: $aggregatorId", None))
          case Some(agg) =>
            // GH-411: the ambient name-cycle guard rides along so a name's refersTo cannot
            // re-enter itself through a range slot. GH-710: a dynamic name and a LET name resolve
            // as in a function's range slot
            Evaluator.resolveRangeSlot(
              location,
              sheet,
              workbook,
              resolvingNames,
              bindings,
              evalNameRef(_, _, sheet, clock, workbook, currentCell, asReference = true),
              aggregateMemo
            ) match
              // GH-630: a slot that IS an error value (`COUNT(#REF!)`, a name bound to `#N/A`)
              // follows the aggregate's error-argument policy, exactly as the Call form does
              case Left(err) => errorArgumentResult(agg, err)
              case Right((targetSheet, range)) =>
                aggregateMemoOpt match
                  case Some(memo) =>
                    memo.getOrCompute(
                      targetSheet,
                      range,
                      agg.name,
                      Evaluator.AggregateMemoMode.TypedNode
                    ) {
                      evalAggregateNode(agg, range, targetSheet, clock, workbook)
                    }
                  case None => evalAggregateNode(agg, range, targetSheet, clock, workbook)

      case call: TExpr.Call[?] =>
        evalCall(call, sheet, clock, workbook, currentCell, selectsReference = false)
          .asInstanceOf[Either[EvalError, A]]

      // ===== GH-193: LET lexical bindings =====
      // Cast the Either container, not the value: BindingRef extends TExpr[Nothing], so the GADT
      // match refines A to Nothing and a value-level cast would compile to a throwing
      // cast-to-Nothing (same reason the arithmetic cases cast their Either results).
      case TExpr.BindingRef(name) =>
        (bindings.get(name) match
          // a reference a plain cell's LET bound, where a value is needed: its intersected cell
          // (its values in array mode, as a SUMPRODUCT in the body reads it)
          case Some(RangeOperand(target, range)) =>
            derefRange(range, target, clock, workbook, currentCell)
          case Some(value) => Right(value)
          case None =>
            // Parser-prevented: BindingRef is only emitted for lexically resolved names
            Left(EvalError.EvalFailed(s"LET name '$name' is not in scope", None))
        ).asInstanceOf[Either[EvalError, A]]

      // A binding used in a typed argument position (rewritten from BindingRef by the as*Expr
      // coercion boundary): coerce the bound value totally — Left(TypeMismatch) when
      // uncoercible — so consuming functions never checkcast a mistyped value and throw.
      // Arrays pass through in array mode (SUM(t) aggregates a bound TRANSPOSE) and collapse to
      // their top-left value in scalar mode (GH-302); a bound reference intersects there.
      case TExpr.CoercedBindingRef(name, target) =>
        evalCoercedBinding(
          name,
          target,
          collapse = !allowArrayResults,
          clock,
          workbook,
          currentCell
        )
          .asInstanceOf[Either[EvalError, A]]

      // GH-302/GH-306: a runtime-polymorphic expression in a typed argument position — evaluate,
      // then coerce the runtime value totally per the target's conventions. Same array policy as
      // CoercedBindingRef; evalArg and evalMaybeArray override the collapse policy positionally.
      case TExpr.Coerced(inner, target) =>
        evalCoercedExpr(
          inner,
          target,
          sheet,
          clock,
          workbook,
          currentCell,
          collapse = !allowArrayResults
        ).asInstanceOf[Either[EvalError, A]]

      case TExpr.Let(letBindings, body) =>
        evalLet(letBindings, body, sheet, clock, workbook, currentCell)
          .asInstanceOf[Either[EvalError, A]]

      // ===== GH-384: defined-name references =====
      // Resolved here (not at parse) because only the workbook knows the name table. Same
      // Either-container cast rationale as BindingRef (TExpr[Nothing] refinement).
      case TExpr.NameRef(name) =>
        evalNameRef(name, scope = None, sheet, clock, workbook, currentCell)
          .asInstanceOf[Either[EvalError, A]]

      // GH-394: sheet-qualified name (=Model!case) — identical resolution, but the lookup runs
      // as seen from the QUALIFYING sheet (its sheet-scoped names shadow workbook-scoped ones)
      case TExpr.SheetNameRef(qualifier, name) =>
        evalNameRef(name, scope = Some(qualifier), sheet, clock, workbook, currentCell)
          .asInstanceOf[Either[EvalError, A]]

  /**
   * A function call. `selectsReference`: the call's result feeds a reference position (an
   * aggregate's argument), so IF, IFS, CHOOSE and SWITCH keep the reference they select whole.
   *
   * GH-712: SUMPRODUCT's shared extent reaches a call only along the path its static trimming
   * (`TExpr.transformRanges`) takes — operators, LET and the calls whose flags answer
   * `entersArrayRangeBounds` (lifted calls; ROW, COLUMN). Any other call (MAX, COUNT, INDEX's
   * shape, IF, a nested SUMPRODUCT) evaluates its arguments outside it, so
   * `SUMPRODUCT((A:A>0)*MAX(ROW(A:A)))` reads MAX over every row, as Excel does, rather than over
   * the used ones.
   */
  private def evalCall(
    call: TExpr.Call[?],
    sheet: Sheet,
    clock: Clock,
    workbook: Option[Workbook],
    currentCell: Option[ARef],
    selectsReference: Boolean
  ): Either[EvalError, Any] =
    if arrayRangeBounds.isDefined && !call.spec.flags.entersArrayRangeBounds then
      outsideArrayRangeBounds.evalCallNode(
        call,
        sheet,
        clock,
        workbook,
        currentCell,
        selectsReference
      )
    else evalCallNode(call, sheet, clock, workbook, currentCell, selectsReference)

  /** This evaluator without SUMPRODUCT's shared extent (see [[evalCall]]). */
  private lazy val outsideArrayRangeBounds: EvaluatorImpl =
    new EvaluatorWithDepth(
      currentDepth,
      allowArrayResults,
      bindings,
      rng,
      memoOpt,
      resolvingNames,
      workbookPath,
      aggregateMemoOpt,
      placementGuard,
      arrayRangeBounds = None,
      passPosition = passPosition
    )

  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  private def evalCallNode(
    call: TExpr.Call[?],
    sheet: Sheet,
    clock: Clock,
    workbook: Option[Workbook],
    currentCell: Option[ARef],
    selectsReference: Boolean
  ): Either[EvalError, Any] =
    // GH-302: scalar argument positions COLLAPSE ArrayResults (implicit intersection:
    // top-left value, Empty when empty) instead of rejecting them. Typed positions
    // (Coerced/CoercedBindingRef) collapse-then-coerce so the value matches the position's
    // type regardless of the evaluator's array mode; Any/CellValue positions collapse to the
    // raw CellValue (consumers go through ExprValue.from).
    //
    // Any other node collapses through its STATIC scalar kind (TExpr.scalarKind): in array
    // mode an arithmetic node (statically TExpr[BigDecimal], kept bare in numeric slots)
    // evaluates to an ArrayResult, and its top-left element must reach the function as a
    // number or a Left — never a raw CellValue the body would cast and throw on. Only the
    // Either container is cast.
    def evalArg[A](expr: TExpr[A]): Either[EvalError, A] =
      val result: Either[EvalError, Any] = expr match
        case TExpr.Coerced(inner, target) =>
          evalCoercedExpr(inner, target, sheet, clock, workbook, currentCell, collapse = true)
        case TExpr.CoercedBindingRef(name, target) =>
          evalCoercedBinding(name, target, collapse = true, clock, workbook, currentCell)
        case other =>
          (eval(other, sheet, clock, workbook, currentCell): Either[EvalError, Any]).flatMap {
            case ar: ArrayResult =>
              val kind = TExpr.scalarKind(other)
              ScalarCoercion.collapseTo(kind.fold("argument")(coercionLabel), ar, kind)
            case value => Right(value)
          }
      result.asInstanceOf[Either[EvalError, A]]
    // GH-346: functions that read uncached formula cells (aggregates, array materialization)
    // recurse through the same memo, so a shared precedent evaluates once per pass. Created
    // here when the pass has none yet — the ctx is per Call node, so it cannot leak across
    // top-level evaluations.
    val callMemo = passMemo
    // GH-197: Array-aware evaluator for functions like SUMPRODUCT that accept array expressions.
    // GH-193: carries the LET environment and recursion depth so array-evaluated arguments
    // (e.g. SUM over a range-valued binding) still resolve in-scope names.
    def evalArrayArg(
      expr: TExpr[Any],
      bounds: Option[CellRange => CellRange] = arrayRangeBounds
    ): Either[EvalError, Any] =
      new EvaluatorWithDepth(
        currentDepth,
        allowArrayResults = true,
        bindings,
        rng,
        Some(callMemo),
        resolvingNames, // GH-384: the name-cycle guard survives array-argument evaluation
        workbookPath,
        aggregateMemoOpt,
        placementGuard,
        arrayRangeBounds = bounds,
        passPosition = passPosition
      )
        .eval(expr, sheet, clock, workbook, currentCell)

    val ctx = EvalContext(
      sheet,
      clock,
      workbook,
      [A] => (expr: TExpr[A]) => evalArg(expr),
      expr => evalArrayArg(expr),
      currentCell,
      currentDepth,
      bindings,
      rng,
      Some(callMemo),
      workbookPath,
      aggregateMemoOpt,
      placementGuard = placementGuard,
      operands = EvalContext.Operands(
        arrayMode = allowArrayResults,
        referenceArg = Some(referenceArgument(_, sheet, clock, workbook, currentCell)),
        selectsReference = selectsReference,
        nameTrail = resolvingNames,
        nameReference =
          Some(evalNameRef(_, _, sheet, clock, workbook, currentCell, asReference = true)),
        arrayRangeBounds = arrayRangeBounds,
        boundedArrayArg = Some((expr, bound) => evalArrayArg(expr, Some(bound))),
        passPosition = passPosition
      )
    )
    // a reference position asks a function that returns a reference (`FunctionSpec.reference`)
    // for the reference itself, unread — only once its lifted slots are planned, so an array in
    // one (`SUM(INDEX(r,{1;2}))`) still lifts, per element, as values, and a scalar one is
    // evaluated once
    def unlifted(args: call.spec.Args): Either[EvalError, Any] =
      if selectsReference then call.spec.reference(args, ctx).getOrElse(call.spec.eval(args, ctx))
      else call.spec.eval(args, ctx)
    call.spec.flags.lift match
      case ArrayLift.Off => unlifted(call.args)
      // in array mode the result may be an ArrayResult: its consumers handle one (the
      // array-mode typing invariant)
      case lift @ ArrayLift.On(_, _) =>
        evalLiftedCall(call, lift, ctx, sheet, clock, workbook, currentCell)(unlifted)

  /**
   * An argument in Excel's reference operand class ([[EvalContext.evalReferenceArg]]), and the
   * reference an expression denotes wherever Excel needs one ([[EvalContext.evalReference]]). A
   * reference reaches its consumer whole, unread (a [[RangeOperand]]): a cell or range, a name
   * bound to or computing one, a LET name bound to one, a LET body that is one, the reference IF,
   * IFS, CHOOSE or SWITCH select, and the one OFFSET, INDIRECT and INDEX return. Any other
   * expression evaluates in this evaluator's mode, its own result not collapsed: in a plain cell
   * `SUM(ABS(A1:A10))` in row 5 is ABS(A5), and `SUM(SEQUENCE(3))` still folds all three.
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  private def referenceArgument(
    expr: TExpr[Any],
    sheet: Sheet,
    clock: Clock,
    workbook: Option[Workbook],
    currentCell: Option[ARef]
  ): Either[EvalError, Any] =
    def whole(location: TExpr.RangeLocation): Either[EvalError, Any] =
      Evaluator
        .resolveRangeLocation(location, sheet, workbook, resolvingNames)
        .map((target, range) => RangeOperand(target, range))
    def cell(at: ARef): TExpr.RangeLocation = TExpr.RangeLocation.Local(CellRange(at, at))
    def name(definedName: String, scope: Option[SheetName]): Either[EvalError, Any] =
      Evaluator.resolveRangeLocation(
        TExpr.RangeLocation.Name(definedName, scope),
        sheet,
        workbook,
        resolvingNames
      ) match
        case Right((target, range)) => Right(RangeOperand(target, range))
        // a name computing a reference is that reference; a named formula, its value (its array)
        case Left(_) =>
          evalNameRef(definedName, scope, sheet, clock, workbook, currentCell, asReference = true)
    (expr: TExpr[?]) match
      case TExpr.RangeRef(range, _) => Right(RangeOperand(sheet, range))
      case TExpr.SheetRange(sheetName, range, _) =>
        whole(TExpr.RangeLocation.CrossSheet(sheetName, range))
      // one cell is a reference too: `SUM(IF(TRUE,C1,0))` skips a text C1 as `SUM(C1)` does, and
      // `AND(C1)` ignores it as `AND(C1:C1)` does
      case TExpr.Ref(at, _, _) => whole(cell(at))
      case TExpr.PolyRef(at, _) => whole(cell(at))
      case TExpr.SheetRef(sheetName, at, _, _) =>
        whole(TExpr.RangeLocation.CrossSheet(sheetName, CellRange(at, at)))
      case TExpr.SheetPolyRef(sheetName, at, _) =>
        whole(TExpr.RangeLocation.CrossSheet(sheetName, CellRange(at, at)))
      case TExpr.NameRef(n) => name(n, None)
      case TExpr.SheetNameRef(qualifier, n) => name(n, Some(qualifier))
      // a LET name's value — the reference it was bound to, whole — and a LET whose body is the
      // reference argument
      case TExpr.BindingRef(bound) =>
        bindings
          .get(bound)
          .toRight(EvalError.EvalFailed(s"LET name '$bound' is not in scope", None))
      case TExpr.CoercedBindingRef(bound, target) =>
        evalCoercedBinding(
          bound,
          target,
          collapse = false,
          clock,
          workbook,
          currentCell,
          asReference = true
        )
      case TExpr.Let(letBindings, body) =>
        evalLet(letBindings, body, sheet, clock, workbook, currentCell, referenceBody = true)
      // the slot's coercion is the consumer's: a reference, an array or a value reaches it as is,
      // and the consumer applies Excel's rule for each (an aggregate's rule for a value typed
      // into its argument list: `SUM("5")` is 5, `SUM("x")` is #VALUE!)
      case TExpr.Coerced(inner, _) =>
        referenceArgument(inner.asInstanceOf[TExpr[Any]], sheet, clock, workbook, currentCell)
      case call: TExpr.Call[?] if Evaluator.referenceFunctions.contains(call.spec.name) =>
        evalCall(call, sheet, clock, workbook, currentCell, selectsReference = true)
      // `+range` is an operator with a value operand, not a reference (the ToolPak rule relies on
      // it): in a plain cell it is intersected like any operand
      case other => eval(other.asInstanceOf[TExpr[Any]], sheet, clock, workbook, currentCell)

  // ===== Excel array lifting (FunctionFlags.lift) =====

  /**
   * A call to a function flagged [[FunctionFlags.lift]]: every lifted scalar slot is classified
   * once ([[classifyLiftedSlot]]) — so each is evaluated at most once and RAND draws once — then
   * the function runs once with the slots' values substituted, or, when a slot holds an array, once
   * per element of the broadcast shape (an axis of extent 1 repeats, a mismatched extent pads #N/A,
   * the GH-344 4b rule). Scalar-certain slots take the unlifted path untouched.
   */
  private def evalLiftedCall(
    call: TExpr.Call[?],
    lift: ArrayLift.On,
    ctx: EvalContext,
    sheet: Sheet,
    clock: Clock,
    workbook: Option[Workbook],
    currentCell: Option[ARef]
  )(unlifted: call.spec.Args => Either[EvalError, Any]): Either[EvalError, Any] =
    import ArrayLifting.SlotValue
    val slots = call.spec.argSpec.scalarSlots(call.args).zipWithIndex.map {
      case ((expr, kind), position) => (expr, kind, lift.slots.forall(_.contains(position)))
    }
    // array mode lifts anything that may be an array; a plain cell intersects references and,
    // GH-669, lifts over an array constant — a value, which Excel's legacy evaluation keeps
    // (`SUM(COUNTIF(A1:A10,{1,2,3}))`, `SUM(ABS({-1,-2}))` need no array formula)
    def mayLift(expr: TExpr[?], kind: LiftSlot): Boolean =
      if allowArrayResults then !ArrayLifting.isScalarCertain(expr, bindings)
      else
        val peeled = ArrayLifting.peel(expr, kind)._1
        ArrayLifting.referenceLocation(peeled).isDefined || Evaluator.isArrayConstant(peeled)
    if !slots.exists((expr, kind, lifted) => lifted && mayLift(expr, kind)) then unlifted(call.args)
    else
      val planned = slots.foldLeft[Either[EvalError, Vector[(TExpr[?], LiftSlot, SlotValue)]]](
        Right(Vector.empty)
      ) {
        case (Left(err), _) => Left(err)
        case (Right(acc), (expr, kind, lifted)) =>
          val classified =
            if lifted && mayLift(expr, kind) then
              classifyLiftedSlot(
                call.spec.name,
                lift.rangeRefs,
                expr,
                kind,
                sheet,
                clock,
                workbook,
                currentCell
              )
            else Right((kind, SlotValue.Keep))
          classified.map((slotKind, value) => acc :+ (expr, slotKind, value))
      }
      planned.flatMap { plan =>
        // the slot expressions for one call; `elements` holds one element per array slot, in order
        def argsFor(elements: Vector[CellValue]) =
          val (replacements, _) =
            plan.foldLeft((Vector.empty[TExpr[?]], elements)) {
              case ((acc, remaining), (expr, kind, value)) =>
                value match
                  case SlotValue.Keep => (acc :+ expr, remaining)
                  case SlotValue.Scalar(v, fromCell) =>
                    (acc :+ ArrayLifting.substitute(v, kind, fromCell), remaining)
                  case SlotValue.ErrorArg(code) => (acc :+ TExpr.ErrorLit(code), remaining)
                  case SlotValue.Elements(_, fromCells) =>
                    val element = remaining.headOption.getOrElse(CellValue.Empty)
                    (acc :+ ArrayLifting.substitute(element, kind, fromCells), remaining.drop(1))
            }
          call.spec.argSpec.replaceScalarSlots(call.args, replacements.toList)._1
        val arrays = plan.collect { case (_, _, SlotValue.Elements(array, _)) => array }
        if arrays.isEmpty then unlifted(argsFor(Vector.empty))
        else
          // #750: a sparse array's fill cells share one call; a call that draws a random number
          // (itself, or in an argument it evaluates per element) must draw for every cell instead
          def drawsPerElement: Boolean =
            val evaluatedOnce = plan.collect {
              case (
                    expr,
                    _,
                    SlotValue.Scalar(_, _) | SlotValue.Elements(_, _) | SlotValue.ErrorArg(_)
                  ) =>
                expr
            }
            lazy val nameDraws = namedDraws(sheet, workbook)
            SumProductTail.draws(call.spec) ||
            call.spec.argSpec.toValues(call.args).exists {
              case ArgValue.Expr(e) =>
                !evaluatedOnce.exists(_ eq e) &&
                WorkbookAudit.callsWhere(e, nameDraws(_, _), SumProductTail.draws)
              case ArgValue.Range(TExpr.RangeLocation.Name(name, scope)) => nameDraws(scope, name)
              case _ => false
            }
          val perCell = !arrays.forall(_.isDense) && drawsPerElement
          ArrayArithmetic.broadcastN(arrays, perCell) { elements =>
            ArrayLifting.elementResult(call.spec.eval(argsFor(elements), ctx))
          }
      }

  /**
   * Classify one lifted slot, returning the kind its substitutes take (the peeled wrapper's target
   * — [[ArrayLifting.peel]]) and its value.
   *
   * A reference (a range, or a name bound to one; never under unary plus) resolves: one cell is
   * read as a scalar (a name bound to one cell reads as a reference to that cell would); a
   * multi-cell reference is `#VALUE!` for the Analysis ToolPak lineage (`rangeRefs = false`),
   * materializes in array mode, and in a plain cell intersects the formula's cell like `@` — a
   * failed intersection reaches the function as `#VALUE!` (ISERROR sees it), a missing position is
   * loud. A computed slot evaluates once in array mode and is kept as is in a plain cell, where its
   * own references have already intersected (an array value keeps its top-left, GH-302). An error
   * value becomes an error argument; a host failure keeps the slot's expression, so the function
   * evaluates it lazily exactly as it would unlifted (IFERROR still catches it).
   */
  private def classifyLiftedSlot(
    fnName: String,
    rangeRefs: Boolean,
    expr: TExpr[?],
    kind: LiftSlot,
    sheet: Sheet,
    clock: Clock,
    workbook: Option[Workbook],
    currentCell: Option[ARef]
  ): Either[EvalError, (LiftSlot, ArrayLifting.SlotValue)] =
    import ArrayLifting.SlotValue
    val (peeled, peeledKind) = ArrayLifting.peel(expr, kind)
    def readCell(target: Sheet, at: ARef): Either[EvalError, CellValue] =
      Evaluator.cellValueReader(
        target,
        clock,
        workbook,
        currentDepth,
        rng,
        passMemo,
        workbookPath,
        aggregateMemoOpt
      )(at)
    def reference(
      location: TExpr.RangeLocation,
      target: Sheet,
      range: CellRange
    ): Either[EvalError, SlotValue] =
      if range.width == 1 && range.height == 1 then
        readCell(target, range.start).map(SlotValue.Scalar(_, fromCell = true))
      else if !rangeRefs then Left(ArrayLifting.rangeReferenceRefused(fnName, location))
      else if allowArrayResults then
        materializeRange(range, target, clock, workbook).map(ArrayLifting.fromArray(_, true))
      else
        Evaluator.implicitIntersection(range, currentCell) match
          case Right(at) => readCell(target, at).map(SlotValue.Scalar(_, fromCell = true))
          case Left(failure) =>
            EvalError
              .toErrorValue(failure)
              .map(_ => SlotValue.ErrorArg(CellError.Value))
              .toRight(failure)
    def computed: SlotValue =
      // GH-669: a plain cell computes an array constant's slot in array mode (see mayLift)
      if !allowArrayResults && Evaluator.isArrayConstant(peeled) then
        withArrayResults.eval(peeled, sheet, clock, workbook, currentCell) match
          case Right(array: ArrayResult) => ArrayLifting.fromArray(array, fromCells = false)
          case Right(value) => SlotValue.Scalar(value, fromCell = false)
          case Left(failure) =>
            EvalError.toErrorValue(failure).fold(SlotValue.Keep)(SlotValue.ErrorArg(_))
      else if !allowArrayResults then SlotValue.Keep
      else
        evalMaybeArray(peeled, sheet, clock, workbook, currentCell) match
          case Right(array: ArrayResult) => ArrayLifting.fromArray(array, fromCells = false)
          case Right(value) => SlotValue.Scalar(value, fromCell = false)
          case Left(failure) =>
            EvalError.toErrorValue(failure).fold(SlotValue.Keep)(SlotValue.ErrorArg(_))
    val value = ArrayLifting.referenceLocation(peeled) match
      case Some(location) =>
        Evaluator.resolveRangeLocation(location, sheet, workbook, resolvingNames) match
          case Right((target, range)) => reference(location, target, range)
          case Left(_) =>
            location match
              // a name bound to a constant or formula is a computed value, not a reference
              case TExpr.RangeLocation.Name(_, _) => Right(computed)
              case _ => Right(SlotValue.Keep)
      case None => Right(computed)
    value.map {
      case SlotValue.Keep => (kind, SlotValue.Keep)
      case other => (peeledKind, other)
    }

  /**
   * GH-711/#750: whether a defined name, referenced from `sheet`, reaches a call that draws a
   * random number ([[SumProductTail.draws]]); built on first use, once per recalculation
   * generation.
   */
  private def namedDraws(
    sheet: Sheet,
    workbook: Option[Workbook]
  ): (Option[SheetName], String) => Boolean =
    workbook match
      case None => (_, _) => false
      case Some(wb) =>
        def build(book: Workbook) =
          WorkbookAudit.namesReaching(book, WorkbookAudit.callsWhere(_, _, SumProductTail.draws))
        val reach = aggregateMemoOpt.fold(build(wb))(_.namesReachingDraw(wb, build))
        (scope, name) => reach(sheet.name, scope, name) != NameWalk.Reach.Absent

  // ===== GH-193: LET evaluation =====

  /**
   * Evaluate LET bindings left-to-right (each against the environment so far), then the body with
   * the full environment. A failing binding short-circuits with the binding name in the message.
   *
   * A binding evaluates in the formula's own mode and keeps a reference a reference, so LET never
   * changes a value: plain(LET(x, e, b)) is plain(b with e for x). Range-shaped binding values were
   * substituted into the body by the parser (never materialized — a whole-column binding would
   * allocate millions of cells). In a plain cell any other binding that denotes a reference (a
   * cell, a name, a LET name bound to one, OFFSET/INDIRECT/INDEX, IF/CHOOSE selecting one) binds
   * the reference itself, a [[RangeOperand]]: intersected where the body reads it as a value, whole
   * where it reads it as a reference (an aggregate, ROWS), materialized in array mode (SUMPRODUCT).
   * Array mode binds the reference as well when a range slot in the LET reads a LET name
   * (`COUNTIF(r, ">2")`, GH-710): the slot takes it whole, every value read materializes it. Every
   * other binding is its value — in a plain cell its references intersect, while an array-returning
   * call (TRANSPOSE) still binds its array.
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  private def evalLet(
    letBindings: List[(String, TExpr[?])],
    body: TExpr[?],
    sheet: Sheet,
    clock: Clock,
    workbook: Option[Workbook],
    currentCell: Option[ARef],
    referenceBody: Boolean = false
  ): Either[EvalError, Any] =
    def scoped(env: Map[String, Any]): EvaluatorImpl =
      new EvaluatorWithDepth(
        currentDepth,
        allowArrayResults,
        env,
        rng,
        memoOpt,
        resolvingNames,
        workbookPath,
        aggregateMemoOpt,
        placementGuard,
        arrayRangeBounds = arrayRangeBounds,
        passPosition = passPosition
      )
    // An array evaluation binds references when a range slot reads a LET name (GH-710) or when
    // something inspects a reference's position or shape (GH-712). One flag for the whole LET,
    // deliberately: a reference-denoting binding read as a value reads the same whether it is bound
    // as the reference or as its array (an aggregate skips an array's text and logical elements as
    // it skips a reference's — RowColumnArraySpec pins SUM, COUNT and COUNTA), and Excel's LET
    // always binds the reference, so binding one is never wrong.
    lazy val bindsReferences = !allowArrayResults ||
      letBindings.exists((_, value) =>
        Evaluator.readsBindingAsRange(value) || Evaluator.inspectsReference(value)
      ) ||
      Evaluator.readsBindingAsRange(body) || Evaluator.inspectsReference(body)
    val envResult = letBindings.foldLeft[Either[EvalError, Map[String, Any]]](Right(bindings)) {
      case (Left(err), _) => Left(err)
      case (Right(env), (name, valueExpr)) =>
        valueExpr match
          case _: TExpr.RangeRef | _: TExpr.SheetRange => Right(env)
          case _ =>
            val evaluated =
              if bindsReferences && Evaluator.denotesReference(valueExpr) then
                scoped(env).referenceArgument(
                  valueExpr.asInstanceOf[TExpr[Any]],
                  sheet,
                  clock,
                  workbook,
                  currentCell
                )
              else
                // Bare cell refs resolve to the cell's effective value (cached formula extracted,
                // Empty → 0) — same treatment as top-level refs and equality operands (GH-233).
                scoped(env).eval(
                  TExpr.asResolvedValueExpr(valueExpr).asInstanceOf[TExpr[Any]],
                  sheet,
                  clock,
                  workbook,
                  currentCell
                )
            evaluated match
              case Right(value) => Right(env + (name -> unwrapBindingValue(value)))
              case Left(err) =>
                // GH-344: a binding that computes an Excel error VALUE binds it as a VALUE and
                // continues — `=LET(x,1/0,x)` is #DIV/0! while `=LET(x,1/0,IFERROR(x,5))` is 5,
                // Excel-exact. Host failures keep the loud wrap naming the binding.
                EvalError.toErrorValue(err) match
                  case Some(code) => Right(env + (name -> (CellValue.Error(code): Any)))
                  case None =>
                    Left(
                      EvalError.EvalFailed(
                        s"LET binding '$name': ${EvalError.toXLError(err).message}",
                        None
                      )
                    )
    }
    envResult.flatMap { env =>
      // in a reference position (SUM(LET(r, A1:A3, r))) the body is a reference argument;
      // elsewhere a range-valued body is a value like any range (its values in array mode, its
      // intersected cell in a plain cell)
      if referenceBody then
        scoped(env).referenceArgument(
          body.asInstanceOf[TExpr[Any]],
          sheet,
          clock,
          workbook,
          currentCell
        )
      else
        scoped(env).eval(
          TExpr.asResolvedValueExpr(body).asInstanceOf[TExpr[Any]],
          sheet,
          clock,
          workbook,
          currentCell
        )
    }

  // ===== GH-384: defined-name resolution =====

  /**
   * Resolve a defined-name reference to its value.
   *
   * Lookup (sheet-scoped SHADOWS workbook-scoped, case-insensitive — see
   * Evaluator.lookupDefinedName), then parse the refersTo text and evaluate it in the DEFINING
   * context: sheet-scoped names evaluate against their scope sheet, workbook-scoped names against
   * the referencing formula's sheet (refersTo text is almost always fully qualified, so the ambient
   * sheet rarely matters). A name bound to a reference is that reference, exactly like a literal
   * range: its values in array mode, its intersected cell in a plain cell's value position; a
   * reference position reads it whole. A named formula is an array context, as in Excel; one that
   * computes a reference (OFFSET, INDIRECT, INDEX, IF/CHOOSE selecting one — a dynamic range, a
   * scenario switch) is that reference, read the same way. `asReference`: the reference position's
   * read — a reference as an unread [[RangeOperand]].
   *
   * Every failure mode is a clean Left: no workbook context (the SheetRef posture), unknown name,
   * unparseable refersTo, name→name cycles (via `resolvingNames`, keyed on the resolved name) and
   * chains deeper than [[NameWalk.MaxDepth]] names (#691). The environment for the refersTo body is
   * fresh (no LET bindings leak into a name's definition), and the result unwraps CellValue
   * wrappers to primitives exactly like LET binding values so names compose with
   * arithmetic/comparison/text machinery. #691: values are memoised inside the outermost name's
   * evaluation ([[Evaluator.NameMemo]]), so a diamond of names evaluates each once, not once per
   * path.
   */
  private def evalNameRef(
    name: String,
    scope: Option[SheetName],
    sheet: Sheet,
    clock: Clock,
    workbook: Option[Workbook],
    currentCell: Option[ARef],
    asReference: Boolean = false
  ): Either[EvalError, Any] =
    workbook match
      case None =>
        Left(
          EvalError.EvalFailed(
            s"Defined name '$name' requires workbook context, but none was provided.",
            None
          )
        )
      case Some(wb) =>
        // GH-394: a sheet-qualified name (=Model!case) looks up AS SEEN FROM its qualifier
        Evaluator.lookupDefinedName(wb, scope.getOrElse(sheet.name), name) match
          case None =>
            Left(
              EvalError.EvalFailed(
                s"Name '$name' is not defined in this workbook.",
                None
              )
            )
          case Some(dn) =>
            val definingSheet = Evaluator.definedNameScope(wb, dn).getOrElse(sheet)
            val node = NameWalk.Node(dn, definingSheet.name)
            resolvingNames.refusal(name, node) match
              case Some(refused) => Left(refused)
              case None =>
                // #691: the outermost name opens the memo its whole evaluation shares, and
                // counts the draws made under it so a value that drew from the RNG is not reused
                val memo = resolvingNames.memo.getOrElse(new Evaluator.NameMemo)
                val trail = Evaluator.NameTrail(resolvingNames.path + node, Some(memo))
                val bodyRng = if resolvingNames.memo.isEmpty then memo.counting(rng) else rng
                val key = Evaluator.NameMemo.Key(node, asReference, allowArrayResults, currentCell)
                memo.getOrCompute(key, definingSheet, wb, trail.path.size) {
                  evalNameBody(
                    name,
                    dn,
                    definingSheet,
                    trail,
                    bodyRng,
                    clock,
                    workbook,
                    currentCell,
                    asReference
                  )
                }

  /**
   * The value of `dn`, the definition a reference to `name` resolved to, evaluated against
   * `definingSheet` under `trail` (the path with `dn` on it) — [[evalNameRef]] after the guard.
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  private def evalNameBody(
    name: String,
    dn: DefinedName,
    definingSheet: Sheet,
    trail: Evaluator.NameTrail,
    bodyRng: Rng,
    clock: Clock,
    workbook: Option[Workbook],
    currentCell: Option[ARef],
    asReference: Boolean
  ): Either[EvalError, Any] =
    Evaluator.parseNameBody(dn.formula, aggregateMemoOpt) match
      case Left(parseErr) =>
        Left(
          EvalError.EvalFailed(
            s"Defined name '$name' has an unparseable definition '${dn.formula}': " +
              s"${ParseError.toXLError(parseErr, dn.formula).message}",
            None
          )
        )
      case Right(target) =>
        // a reference, read by the position: whole in a reference position, else as a
        // value (derefRange: its values in array mode, its intersected cell in a plain cell)
        def read(target: Sheet, range: CellRange): Either[EvalError, Any] =
          if asReference then Right(RangeOperand(target, range))
          else derefRange(range, target, clock, workbook, currentCell)
        target match
          case TExpr.RangeRef(range, _) => read(definingSheet, range)
          case TExpr.SheetRange(sheetName, range, _) =>
            Evaluator
              .resolveRangeLocation(
                TExpr.RangeLocation.CrossSheet(sheetName, range),
                definingSheet,
                workbook,
                trail
              )
              .flatMap { case (targetSheet, _) => read(targetSheet, range) }
          case other =>
            // Bare refs resolve to the cell's effective value (cached formula
            // extracted, Empty → 0) like top-level refs; the derived evaluator
            // carries the cycle guard and a FRESH binding environment (LET
            // bindings never leak into a name's definition). A named formula is an
            // array context, as in Excel; one that computes a reference keeps it.
            // GH-712: the definition is its own formula, outside SUMPRODUCT's static
            // trimming, so it evaluates without the shared extent (a name computing
            // `MAX(ROW(Sheet1!$A:$A))` is 1,048,576); a reference it returns is
            // bounded by the ROW or COLUMN that reads it.
            val resolved = TExpr.asResolvedValueExpr(other)
            val derived: EvaluatorImpl = new EvaluatorWithDepth(
              currentDepth,
              allowArrayResults = true,
              Map.empty,
              bodyRng,
              memoOpt,
              trail,
              workbookPath,
              aggregateMemoOpt,
              placementGuard,
              arrayRangeBounds = None,
              passPosition = passPosition
            )
            // Keep references through name aliases and LET bodies too. Materializing
            // an alias here loses its position, so a plain consumer would read the
            // top-left value instead of intersecting the reference at its own row.
            derived
              .referenceArgument(
                resolved.asInstanceOf[TExpr[Any]],
                definingSheet,
                clock,
                workbook,
                currentCell
              )
              .flatMap {
                case RangeOperand(targetSheet, range) => read(targetSheet, range)
                case value => Right(unwrapBindingValue(value))
              }

  /**
   * Fold one raw range for the [[TExpr.Aggregate]] node. Mirrors the FunctionSpec variadic
   * aggregate's raw-range branch, including the GH-344 item 6 error gate — carried error CELLS
   * propagate as the element's Excel error VALUE per the aggregator's policy (COUNT skips them) —
   * pinning the `Aggregate(id, r) ≡ Call(spec, r)` law.
   */
  /**
   * GH-630: the aggregate's answer when its ONE range slot resolved to an Excel error value: COUNT
   * 0, COUNTA 1, everything else the error ([[Aggregator.onErrorArgument]]). A host failure (no
   * workbook, an unknown name, a cycle) is not an error value and stays loud.
   */
  private def errorArgumentResult[Acc](
    agg: Aggregator[Acc],
    err: EvalError
  ): Either[EvalError, BigDecimal] =
    EvalError
      .toErrorValue(err)
      .flatMap(_ => Aggregator.onErrorArgument(agg, agg.empty))
      .toRight(err)
      .flatMap(agg.finalizeWithError)

  private def evalAggregateNode[Acc](
    agg: Aggregator[Acc],
    range: CellRange,
    targetSheet: Sheet,
    clock: Clock,
    workbook: Option[Workbook]
  ): Either[EvalError, BigDecimal] =
    val readCell = Evaluator.cellValueReader(
      targetSheet,
      clock,
      workbook,
      currentDepth,
      rng,
      passMemo,
      workbookPath,
      aggregateMemoOpt
    )
    val result = range.cells.foldLeft[Either[EvalError, Acc]](Right(agg.empty)) { (accE, at) =>
      accE.flatMap { acc =>
        readCell(at).flatMap { value =>
          val cell = Cell(at, value)
          if agg.countsNonEmpty then
            // COUNTA mode: count any non-empty cell (error cells are non-empty)
            cell.value match
              case CellValue.Empty => Right(acc)
              case _ => Right(agg.combine(acc, BigDecimal(1)))
          else if agg.countsEmpty then
            // COUNTBLANK mode: count only empty cells
            cell.value match
              case CellValue.Empty => Right(agg.combine(acc, BigDecimal(1)))
              case _ => Right(acc)
          else
            ArrayArithmetic.carriedError(cell.value) match
              case Some(err) if agg.propagatesErrors =>
                Left(
                  EvalError.ErrorValue(
                    err,
                    Some(s"${agg.name}: array element contains ${err.toExcel} error")
                  )
                )
              case Some(_) => Right(acc) // COUNT: errors are not numbers
              case None =>
                // Standard mode: only process numeric values
                TExpr.decodeNumeric(cell) match
                  case Right(value) => Right(agg.combine(acc, value))
                  case Left(_) => Right(acc) // Skip non-numeric cells
        }
      }
    }
    // Finalize and return the result (may return error for AVERAGE on empty range)
    result.flatMap(agg.finalizeWithError)

  /**
   * GH-344: decode a cell for a typed reference position, falling back to the Excel error VALUE the
   * cell carries when the decode refuses — `=A1+1` over a #REF! cell is #REF!, not a type mismatch.
   *
   * #709: a refusal over a text cell (plain, rich, or a formula's cached text) is `#VALUE!` —
   * `=ABS(A1)` over "abc", `=IF(A1,1,0)` over "abc" — as Excel computes it, WHICHEVER decoder
   * refused and for whatever reason: the built-in decoders read numeric text first, so a text
   * refusal at a reference boundary is always Excel's `#VALUE!`, and a caller's decoder passed
   * through `TExpr.ref(at, decode)` gets the same treatment (it no longer sees `CodecFailed` over
   * text). Decode failures over other values keep their loud CodecFailed; positions whose decoders
   * accept error cells (decodeCellValue/decodeResolvedValue: ISERROR, IFERROR, bare `=A1`) never
   * reach the fallback.
   *
   * The date slot's out-of-range serial and the integer slot's out-of-range value are matched
   * first, by instance (ScalarCoercion.DateSerialOutOfRange, ScalarCoercion.IntOutOfRange —
   * produced only by decodeAsDateScalar and decodeAsInt): they are error values over numbers too,
   * and the text fallback would discard the reason. Over text, a caller decoder's own reason is
   * kept in the diagnostic, bounded (ScalarCoercion.decoderReason); the built-in decoders' wording
   * is unchanged. `sheet` names a sheet-qualified reference's sheet in the diagnostic (`Other!C5`).
   */
  private def decodeOrCarried[B](
    at: ARef,
    cell: Cell,
    decode: Cell => Either[com.tjclp.xl.codec.CodecError, B],
    sheet: Option[SheetName] = None
  ): Either[EvalError, B] =
    def label: String =
      sheet.fold(at.toA1)(s => s"${SheetName.quoteForFormula(s.value)}!${at.toA1}")
    decode(cell).left.map { codecErr =>
      if codecErr eq ScalarCoercion.DateSerialOutOfRange then
        ScalarCoercion.dateSerialOutOfRange(label, ScalarCoercion.shownValue(cell.value))
      else if codecErr eq ScalarCoercion.IntOutOfRange then
        ScalarCoercion.intOutOfRange(label, ScalarCoercion.shownValue(cell.value))
      else
        ArrayArithmetic.carriedError(cell.value) match
          case Some(err) => EvalError.ErrorValue(err)
          case None =>
            ScalarCoercion.cellText(cell.value) match
              case Some(text) =>
                val because = ScalarCoercion.decoderReason(codecErr).fold("")(r => s" ($r)")
                EvalError.ErrorValue(
                  CellError.Value,
                  Some(
                    s"$label: text '${ScalarCoercion.quotedText(text)}' in a typed argument$because"
                  )
                )
              case None => EvalError.CodecFailed(at, codecErr)
    }

  /**
   * GH-344: the strict-position guard for Any-typed operands — Left(ErrorValue) when the evaluated
   * operand carries an Excel error VALUE, Right(()) otherwise.
   */
  private def carriedOperandError(label: String, value: Any): Either[EvalError, Unit] =
    ArrayArithmetic.carriedError(ArrayArithmetic.anyToCellValue(value)) match
      case Some(err) => Left(EvalError.ErrorValue(err, Some(label)))
      case None => Right(())

  /**
   * `&`: both operands evaluate like arithmetic operands (evalMaybeArray — a bare range
   * materializes, a Coerced operand passes its array through). An array on either side broadcasts
   * element-wise ([[ArrayArithmetic.concatElement]]: a carried error wins, left first; mismatched
   * shapes pad #N/A) and collapses to its top-left text in scalar mode, the operators' shared
   * GH-302 policy. Two scalars join under the decodeAsString conventions (GH-193: operands are
   * statically String but erased casts can deliver other runtime values, so they are read as Any
   * and coerced totally); GH-344: a scalar operand carrying an Excel error VALUE propagates it
   * instead of stringifying it. #681: a non-finite Double operand (a programmatic AST's) reads as
   * the `#NUM!` it is through `anyToCellValue` on both branches — the scalar guard below and the
   * array's elements — so the text "NaN" never reaches a cell.
   */
  private def evalConcat(
    xExpr: TExpr[String],
    yExpr: TExpr[String],
    sheet: Sheet,
    clock: Clock,
    workbook: Option[Workbook],
    currentCell: Option[ARef]
  ): Either[EvalError, Any] =
    val label = "text concatenation"
    def asArray(value: Any): ArrayResult = value match
      case ar: ArrayResult => ar
      case scalar => ArrayResult.single(ArrayArithmetic.anyToCellValue(scalar))
    def joined(elements: Vector[CellValue]): Either[EvalError, CellValue] =
      Right(
        ArrayArithmetic.concatElement(
          elements.headOption.getOrElse(CellValue.Empty),
          elements.lift(1).getOrElse(CellValue.Empty)
        )
      )
    for
      xv <- evalMaybeArray(xExpr, sheet, clock, workbook, currentCell)
      yv <- evalMaybeArray(yExpr, sheet, clock, workbook, currentCell)
      result <- (xv, yv) match
        case (_: ArrayResult, _) | (_, _: ArrayResult) =>
          ArrayArithmetic
            .broadcastN(Vector(asArray(xv), asArray(yv)))(joined)
            .flatMap(collapseUnlessArrayMode(label, BindingCoercion.Text))
        case (x, y) =>
          for
            _ <- carriedOperandError(label, x)
            _ <- carriedOperandError(label, y)
          yield ScalarCoercion.concatText(x) + ScalarCoercion.concatText(y)
    yield result

  /**
   * DateToSerial / DateTimeToSerial, total over whatever the operand evaluates to: a date or
   * date-time becomes its Excel serial; an array (a date function's result in array mode) maps
   * element-wise to serials in array mode — error elements carry, uncoercible elements become error
   * elements — and collapses to its top-left serial in scalar mode; anything else coerces under the
   * Numeric conventions (dates ARE numbers), so an erased-cast operand is a clean Left.
   */
  private def evalDateSerial(
    operand: TExpr[?],
    sheet: Sheet,
    clock: Clock,
    workbook: Option[Workbook],
    currentCell: Option[ARef]
  ): Either[EvalError, Any] =
    val label = "date serial"
    def serialElement(cv: CellValue): CellValue =
      ArrayArithmetic.carriedError(cv) match
        case Some(err) => CellValue.Error(err)
        case None =>
          ScalarCoercion.coerce(label, cv, BindingCoercion.Numeric) match
            case Right(serial: BigDecimal) => CellValue.Number(serial)
            case Right(_) => CellValue.Error(CellError.Value)
            case Left(e) => CellValue.Error(EvalError.toCellError(e).getOrElse(CellError.Value))
    eval(operand.asInstanceOf[TExpr[Any]], sheet, clock, workbook, currentCell).flatMap {
      case ar: ArrayResult if allowArrayResults =>
        Right(ar.mapPure(serialElement))
      case ar: ArrayResult => ScalarCoercion.collapseTo(label, ar, Some(BindingCoercion.Numeric))
      case value => ScalarCoercion.coerce(label, value, BindingCoercion.Numeric)
    }

  /**
   * Unwrap a CellValue binding result to its primitive so bound values compose with arithmetic,
   * comparison, and text machinery exactly like literals do. DateTime unwraps to its Excel serial
   * number (dates ARE numbers in Excel's value model).
   */
  private def unwrapBindingValue(value: Any): Any = value match
    case CellValue.Number(n) => n
    case CellValue.Text(s) => s
    case CellValue.Bool(b) => b
    case CellValue.DateTime(dt) => BigDecimal(CellValue.dateTimeToExcelSerial(dt))
    case CellValue.Empty => BigDecimal(0)
    case CellValue.Formula(_, Some(cached), _) => unwrapBindingValue(cached)
    case other => other

  /**
   * Total coercion of a bound value into a typed argument position (TExpr.CoercedBindingRef).
   *
   * Scalar conventions live in the shared [[ScalarCoercion]] table (the GH-193 precedent,
   * generalized by GH-306). Array policy: pass through when `collapse` is false (array operand
   * positions, evalArrayExpr aggregation) so SUM over a bound TRANSPOSE still aggregates; otherwise
   * collapse to the top-left value and coerce (GH-302). A bound reference is read by the position:
   * whole for a reference position (`asReference`), otherwise as a value ([[derefRange]]: its
   * values in array mode, its intersected cell in a plain cell). Uncoercible values produce
   * Left(TypeMismatch) naming the binding; never a ClassCastException downstream.
   */
  private def evalCoercedBinding(
    name: String,
    target: BindingCoercion,
    collapse: Boolean,
    clock: Clock,
    workbook: Option[Workbook],
    currentCell: Option[ARef],
    asReference: Boolean = false
  ): Either[EvalError, Any] =
    def coerced(value: Any): Either[EvalError, Any] = value match
      case arr: ArrayResult if !collapse => Right(arr)
      case arr: ArrayResult =>
        ScalarCoercion.coerce(s"LET binding '$name'", ScalarCoercion.collapseArray(arr), target)
      case other => ScalarCoercion.coerce(s"LET binding '$name'", other, target)
    bindings.get(name) match
      case None =>
        // Parser-prevented: emitted only for lexically resolved names
        Left(EvalError.EvalFailed(s"LET name '$name' is not in scope", None))
      case Some(ref: RangeOperand) if asReference => Right(ref)
      case Some(RangeOperand(targetSheet, range)) =>
        derefRange(range, targetSheet, clock, workbook, currentCell).flatMap(coerced)
      case Some(value) => coerced(value)

  /**
   * GH-302/GH-306: evaluate a [[TExpr.Coerced]] wrapper — the inner expression evaluates with this
   * evaluator (LET environment, rng and depth preserved), then the runtime value coerces totally
   * per the target. Arrays pass through when `collapse` is false (operand positions: broadcasting)
   * and collapse to top-left before coercion otherwise (scalar positions).
   */
  private def evalCoercedExpr(
    inner: TExpr[Any],
    target: BindingCoercion,
    sheet: Sheet,
    clock: Clock,
    workbook: Option[Workbook],
    currentCell: Option[ARef],
    collapse: Boolean
  ): Either[EvalError, Any] =
    eval(inner, sheet, clock, workbook, currentCell).flatMap {
      case arr: ArrayResult if !collapse => Right(arr)
      case arr: ArrayResult =>
        ScalarCoercion.coerce(coercionLabel(target), ScalarCoercion.collapseArray(arr), target)
      case value => ScalarCoercion.coerce(coercionLabel(target), value, target)
    }

  /** Position description for Coerced error messages. */
  private def coercionLabel(target: BindingCoercion): String = target match
    case BindingCoercion.Text => "text argument"
    case BindingCoercion.Integer => "integer argument"
    case BindingCoercion.Bool => "boolean argument"
    case BindingCoercion.Numeric => "numeric argument"
    case BindingCoercion.Date => "date argument"

  // ===== Array Arithmetic Helpers =====

  /**
   * Evaluate expression, allowing ArrayResult or RangeRef results.
   *
   * Unlike eval(), this method handles RangeRef by converting to ArrayResult, enabling array
   * arithmetic with broadcasting.
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  private def materializeRange(
    range: CellRange,
    targetSheet: Sheet,
    clock: Clock,
    workbook: Option[Workbook]
  ): Either[EvalError, ArrayResult] =
    val memo = passMemo
    ArrayArithmetic.readRange(
      range,
      targetSheet,
      Evaluator.cellValueReader(
        targetSheet,
        clock,
        workbook,
        currentDepth,
        rng,
        memo,
        workbookPath,
        aggregateMemoOpt
      ),
      Some(memo)
    )

  /**
   * A reference where a value is needed (Excel's value operand class). In array mode, the range's
   * values. In a plain cell (scalar mode), the one cell Excel's implicit intersection picks
   * ([[Evaluator.implicitIntersection]]: the formula's row of a column, its column of a row), as a
   * 1×1 array that every scalar consumer already collapses and coerces for its position. A range
   * the formula's row or column does not cross is `#VALUE!`, an Excel error value that IFERROR and
   * ISERROR see; without the formula's position, a multi-cell range is the loud `@` failure.
   */
  private def derefRange(
    range: CellRange,
    targetSheet: Sheet,
    clock: Clock,
    workbook: Option[Workbook],
    currentCell: Option[ARef]
  ): Either[EvalError, ArrayResult] =
    if allowArrayResults then materializeRange(range, targetSheet, clock, workbook)
    else
      Evaluator
        .implicitIntersection(range, currentCell)
        .flatMap(at => materializeRange(CellRange(at, at), targetSheet, clock, workbook))

  private def evalMaybeArray(
    expr: TExpr[?],
    sheet: Sheet,
    clock: Clock,
    workbook: Option[Workbook],
    currentCell: Option[ARef]
  ): Either[EvalError, Any] =
    expr match
      // an operand is a value position: a range broadcasts in array mode and is implicitly
      // intersected in a plain cell (derefRange)
      case TExpr.RangeRef(range, _) =>
        derefRange(range, sheet, clock, workbook, currentCell)
      case TExpr.SheetRange(sheetName, range, _) =>
        Evaluator
          .resolveRangeLocation(
            TExpr.RangeLocation.CrossSheet(sheetName, range),
            sheet,
            workbook,
            resolvingNames
          )
          .flatMap { case (targetSheet, _) =>
            derefRange(range, targetSheet, clock, workbook, currentCell)
          }
      // GH-374: unary plus is transparent in operand positions — =+A1:A3*10 broadcasts
      // exactly like =A1:A3*10
      case TExpr.UnaryPlus(inner) =>
        evalMaybeArray(inner, sheet, clock, workbook, currentCell)
      // GH-476: an untyped ref reaching an OPERAND position resolves here rather than falling to
      // eval's "programming error" arm. Only the top level of a formula and typed argument
      // positions coerce PolyRefs during parsing, so `=IF(1=1,+S2!G1,0)` (unary plus around a
      // cross-sheet ref inside a function argument — the house `=+` style) arrived raw and died
      // with "Unresolved SheetPolyRef".
      case _: TExpr.PolyRef | _: TExpr.SheetPolyRef =>
        eval(
          TExpr.asResolvedValueExpr(expr).asInstanceOf[TExpr[Any]],
          sheet,
          clock,
          workbook,
          currentCell
        )
      // GH-302: coerced nodes in OPERAND positions pass ArrayResults through (so
      // =INDIRECT("A1:A3")*10 broadcasts exactly like =A1:A3*10) and coerce scalars totally
      // (so ="16"&"" or a text call result still enters arithmetic per the Numeric table).
      case TExpr.Coerced(inner, target) =>
        evalCoercedExpr(inner, target, sheet, clock, workbook, currentCell, collapse = false)
      case TExpr.CoercedBindingRef(name, target) =>
        evalCoercedBinding(name, target, collapse = false, clock, workbook, currentCell)
      // GH-680: a chained operand — `a+b+…+y`, the left operand of the next operator — evaluates
      // its left spine in one loop rather than one recursion per operator, so a chain as long as
      // the parser admits cannot exhaust the stack. Every binary operator evaluates its left
      // operand first through this method and returns its error unchanged, so evaluating the
      // innermost node, then each node above with its left operand replaced by the value below
      // (a literal evaluates to exactly that value), is the recursive evaluation step for step.
      case chained if BinarySpine.isChained(chained) =>
        BinarySpine.unwind(chained, BinarySpine.isBinary)._2 match
          case innermost :: above =>
            above.foldLeft(evalMaybeArray(innermost, sheet, clock, workbook, currentCell)) {
              (below, node) =>
                below.flatMap { value =>
                  val step = BinarySpine.rebuild(node, TExpr.Lit(value), BinarySpine.right(node))
                  evalMaybeArray(step, sheet, clock, workbook, currentCell)
                }
            }
          // unreachable: a chained node unwinds to at least itself
          case Nil => eval(chained.asInstanceOf[TExpr[Any]], sheet, clock, workbook, currentCell)
      case other =>
        eval(other.asInstanceOf[TExpr[Any]], sheet, clock, workbook, currentCell)

  /**
   * Convert evaluation result to ArrayOperand.
   */
  private def toOperand(value: Any, sheet: Sheet): Either[EvalError, ArrayArithmetic.ArrayOperand] =
    value match
      case bd: BigDecimal => Right(ArrayArithmetic.ArrayOperand.Scalar(bd))
      case ar: ArrayResult => Right(ArrayArithmetic.ArrayOperand.Array(ar))
      // GH-196: Coerce booleans to numeric in arithmetic (TRUE→1, FALSE→0)
      case b: Boolean =>
        Right(ArrayArithmetic.ArrayOperand.Scalar(if b then BigDecimal(1) else BigDecimal(0)))
      case i: Int => Right(ArrayArithmetic.ArrayOperand.Scalar(BigDecimal(i)))
      case l: Long => Right(ArrayArithmetic.ArrayOperand.Scalar(BigDecimal(l)))
      case d: Double if d.isFinite => Right(ArrayArithmetic.ArrayOperand.Scalar(BigDecimal(d)))
      // GH-193: date values coerce to their Excel serial in arithmetic (dates ARE numbers),
      // e.g. a LET binding holding TODAY() used as `d+1`.
      case ld: java.time.LocalDate =>
        Right(
          ArrayArithmetic.ArrayOperand.Scalar(
            BigDecimal(CellValue.dateTimeToExcelSerial(ld.atStartOfDay()))
          )
        )
      case ldt: java.time.LocalDateTime =>
        Right(ArrayArithmetic.ArrayOperand.Scalar(BigDecimal(CellValue.dateTimeToExcelSerial(ldt))))
      // GH-337: a scalar error VALUE (e.g. a carried #N/A) enters array arithmetic as a 1x1
      // error element and broadcasts across the other operand instead of failing the formula.
      case cv @ CellValue.Error(_) =>
        Right(ArrayArithmetic.ArrayOperand.Array(ArrayResult.single(cv)))
      case _ => Left(EvalError.TypeMismatch("arithmetic", "number or array", value.toString))

  /**
   * GH-302: operator positions in scalar mode collapse array results to their top-left value,
   * consistent with scalar ARGUMENT positions — =INDIRECT("A1")+1 works exactly like
   * =ABS(INDIRECT("A1")). A range operand reaches here already intersected with the formula's cell
   * (derefRange: =A1:A3*10 in row 2 is A2*10), so only an array value is collapsed. Array mode
   * passes the array through for spill/broadcast.
   */
  private def collapseUnlessArrayMode(
    label: String,
    target: BindingCoercion
  )(arr: ArrayResult): Either[EvalError, Any] =
    if allowArrayResults then Right(arr)
    else ScalarCoercion.coerce(label, ScalarCoercion.collapseArray(arr), target)

  /**
   * Evaluate binary arithmetic with array support.
   *
   * Handles:
   *   - Scalar * Scalar -> Scalar (fast path)
   *   - Scalar * Array -> Array (broadcast)
   *   - Array * Scalar -> Array (broadcast)
   *   - Array * Array -> Array (element-wise with broadcasting)
   *   - RangeRef -> automatically converted to Array
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  private def evalArithmetic(
    xExpr: TExpr[BigDecimal],
    yExpr: TExpr[BigDecimal],
    op: ArrayArithmetic.BinaryOp,
    sheet: Sheet,
    clock: Clock,
    workbook: Option[Workbook],
    currentCell: Option[ARef]
  ): Either[EvalError, Any] =
    for
      xVal <- evalMaybeArray(xExpr, sheet, clock, workbook, currentCell)
      yVal <- evalMaybeArray(yExpr, sheet, clock, workbook, currentCell)
      xOp <- toOperand(xVal, sheet)
      yOp <- toOperand(yVal, sheet)
      result <- ArrayArithmetic.broadcast(xOp, yOp, op)
      output <- result match
        case ArrayArithmetic.ArrayOperand.Scalar(v) => Right(v)
        case ArrayArithmetic.ArrayOperand.Array(arr) =>
          collapseUnlessArrayMode("arithmetic", BindingCoercion.Numeric)(arr)
    yield output

  /**
   * GH-197/GH-335: Evaluate an ordered comparison (< <= > >=) with array support.
   *
   * Operands evaluate polymorphically and compare with Excel's total order
   * (ArrayArithmetic.compareCellValues): text lexicographic/case-insensitive, number < text <
   * logical across types, empty coercing to the other operand's zero value. When either operand is
   * a range or array the comparison broadcasts elementwise into an ArrayResult of booleans; two
   * scalars return a plain Boolean.
   *
   * @param op
   *   interprets the comparison sign (e.g. `_ < 0` for Lt)
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  private def evalComparison(
    xExpr: TExpr[?],
    yExpr: TExpr[?],
    op: Int => Boolean,
    sheet: Sheet,
    clock: Clock,
    workbook: Option[Workbook],
    currentCell: Option[ARef]
  ): Either[EvalError, Any] =
    def broadcastToBool(left: ArrayResult, right: ArrayResult): Either[EvalError, Any] =
      ArrayArithmetic
        .broadcastOrderedCompare(left, right, op)
        .flatMap(collapseUnlessArrayMode("comparison", BindingCoercion.Bool))
    def single(scalar: Any): ArrayResult =
      ArrayResult.single(ArrayArithmetic.anyToCellValue(scalar))
    for
      xVal <- evalMaybeArray(xExpr, sheet, clock, workbook, currentCell)
      yVal <- evalMaybeArray(yExpr, sheet, clock, workbook, currentCell)
      result <- (xVal, yVal) match
        // At least one array -> element-wise comparison with broadcasting
        case (lArr: ArrayResult, rArr: ArrayResult) => broadcastToBool(lArr, rArr)
        case (lArr: ArrayResult, scalar) => broadcastToBool(lArr, single(scalar))
        case (scalar, rArr: ArrayResult) => broadcastToBool(single(scalar), rArr)
        // Both scalars -> plain boolean (fast path)
        case (x, y) =>
          ArrayArithmetic
            .compareCellValues(
              ArrayArithmetic.anyToCellValue(x),
              ArrayArithmetic.anyToCellValue(y)
            )
            .map(op)
    yield result

  /**
   * GH-197: Evaluate equality/inequality with array support.
   *
   * Unlike evalComparison (which is numeric), this handles polymorphic equality for strings,
   * numbers, booleans. Enables patterns like `(A1:A3="Yes")*B1:B3`.
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  private def evalEqualityComparison[A](
    xExpr: TExpr[A],
    yExpr: TExpr[A],
    negate: Boolean,
    sheet: Sheet,
    clock: Clock,
    workbook: Option[Workbook],
    currentCell: Option[ARef]
  ): Either[EvalError, Any] =
    for
      xVal <- evalMaybeArray(xExpr.asInstanceOf[TExpr[Any]], sheet, clock, workbook, currentCell)
      yVal <- evalMaybeArray(yExpr.asInstanceOf[TExpr[Any]], sheet, clock, workbook, currentCell)
      result <- (xVal, yVal) match
        // Array vs Array -> element-wise comparison
        case (lArr: ArrayResult, rArr: ArrayResult) =>
          ArrayArithmetic
            .broadcastEqualityCompare(lArr, Left(rArr), negate)
            .flatMap(collapseUnlessArrayMode("comparison", BindingCoercion.Bool))
        // Left is array, right is scalar -> element-wise comparison
        case (arr: ArrayResult, scalar) =>
          ArrayArithmetic
            .broadcastEqualityCompare(arr, Right(scalar), negate)
            .flatMap(collapseUnlessArrayMode("comparison", BindingCoercion.Bool))
        // Left is scalar, right is array -> create 1x1 array and broadcast
        case (scalar, arr: ArrayResult) =>
          val scalarArr = ArrayResult.single(ArrayArithmetic.anyToCellValue(scalar))
          ArrayArithmetic
            .broadcastEqualityCompare(scalarArr, Left(arr), negate)
            .flatMap(collapseUnlessArrayMode("comparison", BindingCoercion.Bool))
        // Both scalars -> plain boolean (fast path).
        // GH-234: use the same case-insensitive/coercing semantics as the array path
        // (ArrayArithmetic.cellValueEquals) so scalar and array equality agree with Excel
        // (e.g. ="A"="a" -> TRUE). Previously used raw `x == y` (case-sensitive).
        // GH-344: an error OPERAND propagates before equality is judged (left first, matching
        // equalityElement) — errors are never "equal" or "unequal". cellValueEquals itself stays
        // total fold-to-false: CriteriaMatcher and the lookups depend on error cells in
        // criteria/lookup ranges simply never matching.
        case (x, y) =>
          val xcv = ArrayArithmetic.anyToCellValue(x)
          val ycv = ArrayArithmetic.anyToCellValue(y)
          ArrayArithmetic.carriedError(xcv).orElse(ArrayArithmetic.carriedError(ycv)) match
            case Some(err) => Left(EvalError.ErrorValue(err, Some("comparison")))
            case None =>
              val eq = ArrayArithmetic.cellValueEquals(xcv, ycv)
              Right(if negate then !eq else eq)
    yield result

/**
 * Depth-aware evaluator for cross-sheet formula cycle protection (GH-161).
 *
 * Extends EvaluatorImpl but tracks recursion depth. When a SheetRef with uncached formula triggers
 * recursive evaluation, the depth is passed through to detect infinite loops. Also carries the LET
 * environment (GH-193) so derived evaluators preserve in-scope bindings.
 */
private class EvaluatorWithDepth(
  depth: Int,
  allowArrayResults: Boolean = false,
  bindings: Map[String, Any] = Map.empty,
  rng: Rng = Rng.system,
  memo: Option[Evaluator.EvalMemo] = None,
  resolvingNames: Evaluator.NameTrail = Evaluator.NameTrail.empty,
  workbookPath: Option[String] = None,
  aggregateMemo: Option[Evaluator.AggregateMemo] = None,
  // GH-715: required, so every derived evaluator decides it; see Evaluator.PlacementGuard
  guard: Option[Evaluator.PlacementGuard],
  arrayRangeBounds: Option[CellRange => CellRange] = None,
  passPosition: EvalPosition = EvalPosition.None
) extends EvaluatorImpl(
      allowArrayResults,
      bindings,
      rng,
      resolvingNames,
      workbookPath,
      aggregateMemo,
      arrayRangeBounds = arrayRangeBounds,
      passPosition = passPosition
    ):
  override protected def currentDepth: Int = depth
  override protected def memoOpt: Option[Evaluator.EvalMemo] = memo
  override protected def placementGuard: Option[Evaluator.PlacementGuard] = guard
  override private[formula] def withArrayResults: Evaluator =
    new EvaluatorWithDepth(
      depth,
      allowArrayResults = true,
      bindings,
      rng,
      memo,
      resolvingNames,
      workbookPath,
      aggregateMemo,
      guard,
      arrayRangeBounds = arrayRangeBounds,
      passPosition = passPosition
    )
