package com.tjclp.xl.formula.eval

import com.tjclp.xl.formula.ast.{BindingCoercion, TExpr}
import com.tjclp.xl.formula.functions.{FunctionSpec, FunctionSpecs, EvalContext}
import com.tjclp.xl.formula.graph.DependencyGraph
import com.tjclp.xl.formula.printer.FormulaPrinter
import com.tjclp.xl.formula.parser.{FormulaParser, ParseError}
import com.tjclp.xl.formula.{Clock, Rng}

import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.{Cell, CellValue, FormulaKind}
import com.tjclp.xl.workbooks.{DefinedName, Workbook}
import com.tjclp.xl.SheetName
import com.tjclp.xl.syntax.* // Extension methods for Sheet.get, CellRange.cells, ARef.toA1
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
   * @param expr
   *   The expression to evaluate
   * @param sheet
   *   The sheet providing cell values
   * @param clock
   *   Clock for date/time functions (defaults to system clock)
   * @param workbook
   *   Optional workbook for cross-sheet references (defaults to None)
   * @param currentCell
   *   Optional current cell reference (for ROW()/COLUMN() with no arguments)
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

object Evaluator:
  /**
   * Default evaluator instance.
   *
   * Pure functional implementation with short-circuit evaluation for And/Or.
   */
  def instance: Evaluator = new EvaluatorImpl()

  /**
   * Evaluator instance with an explicit randomness source (GH-115).
   *
   * Use `Evaluator.instance(Rng.seeded(seed))` for deterministic RAND/RANDBETWEEN.
   */
  def instance(rng: Rng): Evaluator = new EvaluatorImpl(rng = rng)

  /**
   * Evaluator instance that knows the workbook's saved location (GH-424).
   *
   * CELL("filename") reports `directory[file]sheet` from this path; without one it returns the
   * empty string — Excel's behavior for a never-saved workbook.
   */
  def instance(workbookPath: Option[String]): Evaluator =
    new EvaluatorImpl(workbookPath = workbookPath)

  /**
   * Evaluator instance that allows array results to propagate.
   *
   * Used for array formula evaluation where arithmetic over ranges should spill arrays.
   */
  def arrayInstance: Evaluator = new EvaluatorImpl(allowArrayResults = true)

  /** Array-result evaluator with an explicit randomness source (GH-115). */
  def arrayInstance(rng: Rng): Evaluator = new EvaluatorImpl(allowArrayResults = true, rng = rng)

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
   * Resolve a RangeLocation to its target (sheet, range) pair — THE single resolution boundary for
   * range-typed argument slots (GH-394).
   *
   * For Local ranges, returns the current sheet with the carried range. For CrossSheet ranges,
   * looks up the target sheet in the workbook context. For Name locations the carried range is
   * unknown until here: the name resolves against the workbook's name table (sheet-scoped SHADOWS
   * workbook-scoped, case-insensitive — [[lookupDefinedName]]; a sheet-qualified name looks up as
   * seen from its qualifier), its refersTo text parses, and range/cell-shaped targets yield the
   * resolved pair (a single-cell target acts as a 1×1 range, like Excel). Name→name chains follow
   * hop by hop like expression-position NameRef (GH-411), guarded by `resolvingNames` — a chain
   * that revisits a member is a clean cycle error. Other non-range targets (constants, formulas)
   * are a clean per-cell #VALUE! error — never a MatchError.
   *
   * @param resolvingNames
   *   the GH-384 name-cycle guard (UPPERCASED names on the current resolution path); callers inside
   *   the evaluator pass their ambient guard so a name's refersTo cannot re-enter itself through a
   *   range slot.
   */
  private[formula] def resolveRangeLocation(
    location: TExpr.RangeLocation,
    currentSheet: Sheet,
    workbook: Option[Workbook],
    resolvingNames: Set[String] = Set.empty
  ): Either[EvalError, (Sheet, CellRange)] =
    location match
      case TExpr.RangeLocation.Local(range) =>
        Right((currentSheet, range))
      case TExpr.RangeLocation.CrossSheet(sheetName, range) =>
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
      case loc @ TExpr.RangeLocation.External(_, _, _) =>
        Left(externalRefUnsupported(loc.toA1))
      // GH-394: a defined name in a range slot resolves through the name table
      case TExpr.RangeLocation.Name(name, scope) =>
        resolveNameToRange(name, scope, currentSheet, workbook, resolvingNames)

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
    resolvingNames: Set[String]
  ): Either[EvalError, (Sheet, CellRange)] =
    workbook match
      case None =>
        Left(
          EvalError.EvalFailed(
            s"Defined name '$name' requires workbook context, but none was provided.",
            None
          )
        )
      case Some(_) if resolvingNames.contains(name.toUpperCase) =>
        Left(
          EvalError.EvalFailed(
            s"Defined name cycle detected while resolving '$name'.",
            None
          )
        )
      case Some(wb) =>
        val guard = resolvingNames + name.toUpperCase
        val lookupFrom = scope.getOrElse(currentSheet.name)
        lookupDefinedName(wb, lookupFrom, name) match
          case None =>
            Left(EvalError.EvalFailed(s"Name '$name' is not defined in this workbook.", None))
          case Some(dn) =>
            FormulaParser.parse(dn.formula) match
              case Left(parseErr) =>
                Left(
                  EvalError.EvalFailed(
                    s"Defined name '$name' has an unparseable definition '${dn.formula}': " +
                      s"${ParseError.toXLError(parseErr, dn.formula).message}",
                    None
                  )
                )
              case Right(target) =>
                val definingSheet = definedNameScope(wb, dn).getOrElse(currentSheet)
                target match
                  case TExpr.RangeRef(range) => Right((definingSheet, range))
                  case TExpr.SheetRange(sheetName, range) =>
                    resolveRangeLocation(
                      TExpr.RangeLocation.CrossSheet(sheetName, range),
                      definingSheet,
                      workbook,
                      guard
                    )
                  // Single-cell targets act as 1×1 ranges (names routinely point at one cell).
                  // A standalone ref at the top of a parsed formula surfaces as the TYPED
                  // Ref/SheetRef (the resolved-value rewrite); Poly forms are kept for safety.
                  case TExpr.Ref(at, _, _) => Right((definingSheet, CellRange(at, at)))
                  case TExpr.PolyRef(at, _) => Right((definingSheet, CellRange(at, at)))
                  case TExpr.SheetRef(sheetName, at, _, _) =>
                    resolveRangeLocation(
                      TExpr.RangeLocation.CrossSheet(sheetName, CellRange(at, at)),
                      definingSheet,
                      workbook,
                      guard
                    )
                  case TExpr.SheetPolyRef(sheetName, at, _) =>
                    resolveRangeLocation(
                      TExpr.RangeLocation.CrossSheet(sheetName, CellRange(at, at)),
                      definingSheet,
                      workbook,
                      guard
                    )
                  // GH-411: name→name chains follow like expression position, each hop as seen
                  // from its predecessor's defining sheet, under the shared cycle guard
                  case TExpr.NameRef(next) =>
                    resolveNameToRange(next, None, definingSheet, workbook, guard)
                  case TExpr.SheetNameRef(qualifier, next) =>
                    resolveNameToRange(next, Some(qualifier), definingSheet, workbook, guard)
                  case _ =>
                    // Excel: a non-reference name in a range position is #VALUE!
                    Left(
                      EvalError.ErrorValue(
                        com.tjclp.xl.cells.CellError.Value,
                        Some(s"name '$name' does not refer to a range (refersTo: ${dn.formula})")
                      )
                    )

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
    val names = wb.metadata.definedNames
    val sheetIdx = wb.sheets.indexWhere(_.name == currentSheet)
    val sheetScoped =
      if sheetIdx >= 0 then
        names.find(dn => dn.name.equalsIgnoreCase(name) && dn.localSheetId.contains(sheetIdx))
      else None
    sheetScoped.orElse(names.find(dn => dn.name.equalsIgnoreCase(name) && dn.localSheetId.isEmpty))

  /**
   * The sheet a defined name's refersTo evaluates against when the name is sheet-scoped; None for
   * workbook-scoped names (callers fall back to the referencing formula's sheet — refersTo text is
   * almost always fully qualified anyway).
   */
  private[formula] def definedNameScope(wb: Workbook, dn: DefinedName): Option[Sheet] =
    dn.localSheetId.flatMap(idx => wb.sheets.lift(idx))

  /** Maximum recursion depth for cross-sheet formula evaluation (GH-161 cycle protection). */
  private val MaxCrossSheetRecursionDepth = 100

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
    memo: EvalMemo = new EvalMemo,
    workbookPath: Option[String] = None
  ): Either[EvalError, CellValue] =
    boundary:
      // GH-161 review: Add recursion depth limit to prevent stack overflow on circular refs
      if depth > MaxCrossSheetRecursionDepth then
        break(
          Left(
            EvalError.EvalFailed(
              s"Cross-sheet formula recursion depth exceeded (max: $MaxCrossSheetRecursionDepth). Possible circular reference.",
              None
            )
          )
        )

      FormulaParser.parse(formulaStr) match
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
            rng = rng,
            memo = Some(memo),
            workbookPath = workbookPath
          )
            .eval(expr, targetSheet, clock, workbook) match
            case Right(result) => Right(EvalResult.toCellValue(result))
            // GH-344: an error-computing precedent delivers its Excel error VALUE to readers
            // (the cell-mediated cascade); host failures stay loud Lefts.
            case Left(evalError) =>
              EvalError.toErrorValue(evalError) match
                case Some(code) => Right(CellValue.Error(code))
                case None => Left(evalError)

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
   */
  private[formula] final class EvalMemo:
    private val bySheet =
      new java.util.IdentityHashMap[
        Sheet,
        scala.collection.mutable.HashMap[ARef, Either[EvalError, CellValue]]
      ]()

    def getOrCompute(sheet: Sheet, at: ARef)(
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
 *   GH-384: UPPERCASED defined names currently being resolved on this evaluation path — the
 *   name→name cycle guard. A refersTo chain that revisits a member (aa → bb → aa) is a clean
 *   per-cell error instead of unbounded recursion. Cell-mediated cycles (name → cell → name) are
 *   covered separately by the depth-guarded cross-sheet recursion.
 * @param workbookPath
 *   GH-424: the workbook's saved location if the embedder knows one, surfaced to functions via
 *   EvalContext (CELL("filename")). None reproduces Excel's pre-save behavior.
 */
private class EvaluatorImpl(
  allowArrayResults: Boolean = false,
  bindings: Map[String, Any] = Map.empty,
  rng: Rng = Rng.system,
  resolvingNames: Set[String] = Set.empty,
  workbookPath: Option[String] = None
) extends Evaluator:
  /** Current recursion depth for cross-sheet formula evaluation. */
  protected def currentDepth: Int = 0

  /**
   * GH-346: the current pass's memo for recursively evaluated uncached formula cells — None at the
   * top level (a memo is created at the first point recursion can begin and threads through every
   * derived evaluator of the pass via EvaluatorWithDepth).
   */
  protected def memoOpt: Option[Evaluator.EvalMemo] = None
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
                    decodeOrCarried(at, cell, decode)
                  case CellValue.Formula(formulaStr, None, _) =>
                    // Formula has no cached value - parse and evaluate against target sheet
                    // GH-161 review: Apply decoder to Cell with evaluated result (type-safe)
                    // GH-161 review: Pass currentDepth for cycle protection
                    // GH-346: memoized per pass — a cell evaluates once, not once per path
                    val memo = memoOpt.getOrElse(new Evaluator.EvalMemo)
                    memo
                      .getOrCompute(targetSheet, at) {
                        Evaluator.evalCrossSheetFormula(
                          formulaStr,
                          targetSheet,
                          clock,
                          workbook,
                          currentDepth,
                          rng,
                          memo,
                          workbookPath
                        )
                      }
                      .flatMap(evaluatedValue =>
                        decodeOrCarried(at, Cell(at, evaluatedValue), decode)
                      )
                  case _ =>
                    // Cached formula or non-formula cell - use decoder
                    decodeOrCarried(at, cell, decode)

      // ===== GH-353: External-Workbook References =====
      //
      // The target workbook is not loaded, so these can never evaluate. Cells CONTAINING them
      // are pinned to their Excel-written cached value upstream (SheetEvaluator) and never
      // reach this point; direct evaluation and uncached cells get a clear per-cell error.
      case TExpr.ExternalRef(index, name, at, _) =>
        Left(Evaluator.externalRefUnsupported(s"[$index]$name!${(at: ARef).toA1}"))
      case TExpr.ExternalRange(index, name, range) =>
        Left(Evaluator.externalRefUnsupported(s"[$index]$name!${range.toA1}"))

      case TExpr.SheetRange(sheetName, range) =>
        // SheetRange should be wrapped in a function (SUM, COUNT, etc.) before evaluation
        // GH-280: quote cell-ref-shaped sheet names so diagnostics read unambiguously
        val refStr = s"${SheetName.quoteForFormula(sheetName.value)}!${range.toA1}"
        Left(
          EvalError.EvalFailed(
            s"Cross-sheet range $refStr must be used within a function like SUM or COUNT.",
            None
          )
        )

      case TExpr.RangeRef(range) =>
        Left(
          EvalError.EvalFailed(
            s"Range ${range.toA1} must be used within a function like SUM or COUNT.",
            None
          )
        )

      // ===== Literals =====
      case TExpr.Lit(value) =>
        // Literal: return value directly (identity law)
        Right(value)

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
          case CellValue.Formula(formulaStr, None, _) =>
            // GH-346: memoized per pass — a cell evaluates once, not once per path
            val memo = memoOpt.getOrElse(new Evaluator.EvalMemo)
            memo
              .getOrCompute(sheet, at) {
                Evaluator
                  .evalCrossSheetFormula(
                    formulaStr,
                    sheet,
                    clock,
                    workbook,
                    currentDepth,
                    rng,
                    memo,
                    workbookPath
                  )
              }
              .flatMap(evaluatedValue => decodeOrCarried(at, Cell(at, evaluatedValue), decode))
          case _ =>
            decodeOrCarried(at, cell, decode)

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
        // Concatenate: join two strings. Operands are statically String, but erased upstream
        // casts (e.g. a numeric LET binding or a numeric-returning call coerced via
        // asStringExpr) can deliver non-String runtime values — evaluate as Any (a String-typed
        // binder would checkcast and throw) and coerce totally with the decodeAsString
        // conventions instead of crashing (GH-193). GH-344: an operand carrying an Excel error
        // VALUE (an Any-typed call result holding CellValue.Error) propagates the error instead
        // of silently stringifying it.
        for
          xv <- eval(x.asInstanceOf[TExpr[Any]], sheet, clock, workbook, currentCell)
          _ <- carriedOperandError("text concatenation", xv)
          yv <- eval(y.asInstanceOf[TExpr[Any]], sheet, clock, workbook, currentCell)
          _ <- carriedOperandError("text concatenation", yv)
        yield concatText(xv) + concatText(yv)

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
        eval(expr.asInstanceOf[TExpr[Any]], sheet, clock, workbook, currentCell)
          .flatMap(value => ScalarCoercion.coerce("ToInt", value, BindingCoercion.Integer))
          .asInstanceOf[Either[EvalError, A]]

      // ===== Date/Time Conversions =====
      case TExpr.DateToSerial(dateExpr) =>
        eval(dateExpr, sheet, clock, workbook, currentCell).map { date =>
          BigDecimal(CellValue.dateTimeToExcelSerial(date.atStartOfDay()))
        }

      case TExpr.DateTimeToSerial(dtExpr) =>
        eval(dtExpr, sheet, clock, workbook, currentCell).map { dt =>
          BigDecimal(CellValue.dateTimeToExcelSerial(dt))
        }

      case TExpr.Aggregate(aggregatorId, location) =>
        // Use Aggregator typeclass to evaluate any registered aggregate function
        Aggregator.lookup(aggregatorId) match
          case None =>
            Left(EvalError.EvalFailed(s"Unknown aggregator: $aggregatorId", None))
          case Some(agg) =>
            // GH-411: the ambient name-cycle guard rides along so a name's refersTo cannot
            // re-enter itself through a range slot
            Evaluator.resolveRangeLocation(location, sheet, workbook, resolvingNames).flatMap {
              case (targetSheet, range) => evalAggregateNode(agg, range, targetSheet)
            }

      case call: TExpr.Call[?] =>
        // GH-302: scalar argument positions COLLAPSE ArrayResults (implicit intersection:
        // top-left value, Empty when empty) instead of rejecting them. Typed positions
        // (Coerced/CoercedBindingRef) collapse-then-coerce so the value matches the position's
        // type regardless of the evaluator's array mode; Any/CellValue positions collapse to the
        // raw CellValue (consumers go through ExprValue.from).
        def evalArg[A](expr: TExpr[A]): Either[EvalError, A] =
          val result = expr match
            case TExpr.Coerced(inner, target) =>
              evalCoercedExpr(inner, target, sheet, clock, workbook, currentCell, collapse = true)
            case TExpr.CoercedBindingRef(name, target) =>
              evalCoercedBinding(name, target, collapse = true)
            case other => eval(other, sheet, clock, workbook, currentCell)
          result.map {
            case ar: ArrayResult => ScalarCoercion.collapseArray(ar).asInstanceOf[A]
            case value => value.asInstanceOf[A]
          }
        // GH-346: functions that read uncached formula cells (aggregates, array materialization)
        // recurse through the same memo, so a shared precedent evaluates once per pass. Created
        // here when the pass has none yet — the ctx is per Call node, so it cannot leak across
        // top-level evaluations.
        val callMemo = memoOpt.getOrElse(new Evaluator.EvalMemo)
        // GH-197: Array-aware evaluator for functions like SUMPRODUCT that accept array expressions.
        // GH-193: carries the LET environment and recursion depth so array-evaluated arguments
        // (e.g. SUM over a range-valued binding) still resolve in-scope names.
        def evalArrayArg(expr: TExpr[Any]): Either[EvalError, Any] =
          new EvaluatorWithDepth(
            currentDepth,
            allowArrayResults = true,
            bindings,
            rng,
            Some(callMemo),
            resolvingNames, // GH-384: the name-cycle guard survives array-argument evaluation
            workbookPath
          )
            .eval(expr, sheet, clock, workbook, currentCell)

        val ctx = EvalContext(
          sheet,
          clock,
          workbook,
          [A] => (expr: TExpr[A]) => evalArg(expr),
          evalArrayArg,
          currentCell,
          currentDepth,
          bindings,
          rng,
          Some(callMemo),
          workbookPath
        )
        call.spec.eval(call.args, ctx)

      // ===== GH-193: LET lexical bindings =====
      // Cast the Either container, not the value: BindingRef extends TExpr[Nothing], so the GADT
      // match refines A to Nothing and a value-level cast would compile to a throwing
      // cast-to-Nothing (same reason the arithmetic cases cast their Either results).
      case TExpr.BindingRef(name) =>
        (bindings.get(name) match
          case Some(value) => Right(value)
          case None =>
            // Parser-prevented: BindingRef is only emitted for lexically resolved names
            Left(EvalError.EvalFailed(s"LET name '$name' is not in scope", None))
        ).asInstanceOf[Either[EvalError, A]]

      // A binding used in a typed argument position (rewritten from BindingRef by the as*Expr
      // coercion boundary): coerce the bound value totally — Left(TypeMismatch) when
      // uncoercible — so consuming functions never checkcast a mistyped value and throw.
      // Arrays pass through in array mode (SUM(t) aggregates a bound TRANSPOSE) and collapse to
      // their top-left value in scalar mode (GH-302 implicit intersection).
      case TExpr.CoercedBindingRef(name, target) =>
        evalCoercedBinding(name, target, collapse = !allowArrayResults)
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

  // ===== GH-193: LET evaluation =====

  /**
   * Evaluate LET bindings left-to-right (each against the environment so far), then the body with
   * the full environment. A failing binding short-circuits with the binding name in the message.
   *
   * Range-shaped binding values were substituted into the body by the parser, so they are skipped
   * here (never materialized — a whole-column binding would allocate millions of cells). All other
   * values evaluate array-aware so array-producing calls (e.g. TRANSPOSE) can be bound.
   */
  private def evalLet(
    letBindings: List[(String, TExpr[?])],
    body: TExpr[?],
    sheet: Sheet,
    clock: Clock,
    workbook: Option[Workbook],
    currentCell: Option[ARef]
  ): Either[EvalError, Any] =
    val envResult = letBindings.foldLeft[Either[EvalError, Map[String, Any]]](Right(bindings)) {
      case (Left(err), _) => Left(err)
      case (Right(env), (name, valueExpr)) =>
        valueExpr match
          case _: TExpr.RangeRef | _: TExpr.SheetRange => Right(env)
          case _ =>
            // Bare cell refs resolve to the cell's effective value (cached formula extracted,
            // Empty → 0) — same treatment as top-level refs and equality operands (GH-233).
            val resolved = TExpr.asResolvedValueExpr(valueExpr)
            new EvaluatorWithDepth(
              currentDepth,
              allowArrayResults = true,
              env,
              rng,
              memoOpt,
              resolvingNames,
              workbookPath
            )
              .eval(resolved.asInstanceOf[TExpr[Any]], sheet, clock, workbook, currentCell) match
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
      body match
        // A range-valued body (e.g. LET(r, A1:A10, r) under SUM) yields an array in array
        // contexts; scalar contexts keep the standard "range must be used within a function"
        // error from eval below.
        case TExpr.RangeRef(range) if allowArrayResults =>
          Right(ArrayArithmetic.rangeToArray(range, sheet))
        case TExpr.SheetRange(sheetName, range) if allowArrayResults =>
          Evaluator
            .resolveRangeLocation(
              TExpr.RangeLocation.CrossSheet(sheetName, range),
              sheet,
              workbook,
              resolvingNames
            )
            .map { case (targetSheet, _) => ArrayArithmetic.rangeToArray(range, targetSheet) }
        case other =>
          val resolvedBody = TExpr.asResolvedValueExpr(other)
          new EvaluatorWithDepth(
            currentDepth,
            allowArrayResults,
            env,
            rng,
            memoOpt,
            resolvingNames,
            workbookPath
          )
            .eval(resolvedBody.asInstanceOf[TExpr[Any]], sheet, clock, workbook, currentCell)
    }

  // ===== GH-384: defined-name resolution =====

  /**
   * Resolve a defined-name reference to its value.
   *
   * Lookup (sheet-scoped SHADOWS workbook-scoped, case-insensitive — see
   * Evaluator.lookupDefinedName), then parse the refersTo text and evaluate it in the DEFINING
   * context: sheet-scoped names evaluate against their scope sheet, workbook-scoped names against
   * the referencing formula's sheet (refersTo text is almost always fully qualified, so the ambient
   * sheet rarely matters). Range-shaped targets materialize to an ArrayResult against the defining
   * sheet — aggregate positions consume it directly (=SUM(rev_range)), operand positions broadcast,
   * and scalar contexts collapse to the top-left value, exactly like a literal range.
   *
   * Every failure mode is a clean Left: no workbook context (the SheetRef posture), unknown name,
   * unparseable refersTo, and name→name cycles (via `resolvingNames`). The environment for the
   * refersTo body is fresh (no LET bindings leak into a name's definition), and the result unwraps
   * CellValue wrappers to primitives exactly like LET binding values so names compose with
   * arithmetic/comparison/text machinery.
   */
  private def evalNameRef(
    name: String,
    scope: Option[SheetName],
    sheet: Sheet,
    clock: Clock,
    workbook: Option[Workbook],
    currentCell: Option[ARef]
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
        val key = name.toUpperCase
        if resolvingNames.contains(key) then
          Left(
            EvalError.EvalFailed(
              s"Defined name cycle detected while resolving '$name'.",
              None
            )
          )
        else
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
              FormulaParser.parse(dn.formula) match
                case Left(parseErr) =>
                  Left(
                    EvalError.EvalFailed(
                      s"Defined name '$name' has an unparseable definition '${dn.formula}': " +
                        s"${ParseError.toXLError(parseErr, dn.formula).message}",
                      None
                    )
                  )
                case Right(target) =>
                  val definingSheet = Evaluator.definedNameScope(wb, dn).getOrElse(sheet)
                  target match
                    // Range-shaped names materialize like literal ranges in array positions:
                    // consumers collapse (scalar), broadcast (operands), or aggregate (SUM)
                    case TExpr.RangeRef(range) =>
                      Right(ArrayArithmetic.rangeToArray(range, definingSheet))
                    case TExpr.SheetRange(sheetName, range) =>
                      Evaluator
                        .resolveRangeLocation(
                          TExpr.RangeLocation.CrossSheet(sheetName, range),
                          definingSheet,
                          workbook,
                          resolvingNames + key
                        )
                        .map { case (targetSheet, _) =>
                          ArrayArithmetic.rangeToArray(range, targetSheet)
                        }
                    case other =>
                      // Bare refs resolve to the cell's effective value (cached formula
                      // extracted, Empty → 0) like top-level refs; the derived evaluator
                      // carries the cycle guard and a FRESH binding environment (LET
                      // bindings never leak into a name's definition).
                      val resolved = TExpr.asResolvedValueExpr(other)
                      new EvaluatorWithDepth(
                        currentDepth,
                        allowArrayResults = true,
                        Map.empty,
                        rng,
                        memoOpt,
                        resolvingNames + key,
                        workbookPath
                      )
                        .eval(
                          resolved.asInstanceOf[TExpr[Any]],
                          definingSheet,
                          clock,
                          workbook,
                          currentCell
                        )
                        .map(unwrapBindingValue)

  /**
   * Fold one raw range for the [[TExpr.Aggregate]] node. Mirrors the FunctionSpec variadic
   * aggregate's raw-range branch, including the GH-344 item 6 error gate — carried error CELLS
   * propagate as the element's Excel error VALUE per the aggregator's policy (COUNT skips them) —
   * pinning the `Aggregate(id, r) ≡ Call(spec, r)` law.
   */
  private def evalAggregateNode[Acc](
    agg: Aggregator[Acc],
    range: CellRange,
    targetSheet: Sheet
  ): Either[EvalError, BigDecimal] =
    val cells = range.cells.map(cellRef => targetSheet(cellRef))
    val result = cells.foldLeft[Either[EvalError, Acc]](Right(agg.empty)) { (accE, cell) =>
      accE.flatMap { acc =>
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
    // Finalize and return the result (may return error for AVERAGE on empty range)
    result.flatMap(agg.finalizeWithError)

  /**
   * GH-344: decode a cell for a typed reference position, falling back to the Excel error VALUE the
   * cell carries when the decode refuses — `=A1+1` over a #REF! cell is #REF!, not a type mismatch.
   * Decode failures over non-error values keep their loud CodecFailed; positions whose decoders
   * accept error cells (decodeCellValue/decodeResolvedValue: ISERROR, IFERROR, bare `=A1`) never
   * reach the fallback.
   */
  private def decodeOrCarried[B](
    at: ARef,
    cell: Cell,
    decode: Cell => Either[com.tjclp.xl.codec.CodecError, B]
  ): Either[EvalError, B] =
    decode(cell).left.map { codecErr =>
      ArrayArithmetic.carriedError(cell.value) match
        case Some(err) => EvalError.ErrorValue(err)
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
   * Total text coercion for '&' operands, mirroring the decodeAsString conventions (Number →
   * toString, Bool → TRUE/FALSE, DateTime → ISO, Empty → "").
   */
  private def concatText(value: Any): String = value match
    case s: String => s
    case b: Boolean => if b then "TRUE" else "FALSE"
    case bd: BigDecimal => bd.toString
    case i: Int => i.toString
    case ld: java.time.LocalDate => ld.toString
    case ldt: java.time.LocalDateTime => ldt.toString
    case CellValue.Text(s) => s
    case CellValue.Number(n) => n.toString
    case CellValue.Bool(b) => if b then "TRUE" else "FALSE"
    case CellValue.DateTime(dt) => dt.toString
    case CellValue.Empty => ""
    case other => other.toString

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
   * collapse to the top-left value and coerce (GH-302 implicit intersection). Uncoercible values
   * produce Left(TypeMismatch) naming the binding; never a ClassCastException downstream.
   */
  private def evalCoercedBinding(
    name: String,
    target: BindingCoercion,
    collapse: Boolean
  ): Either[EvalError, Any] =
    bindings.get(name) match
      case None =>
        // Parser-prevented: emitted only for lexically resolved names
        Left(EvalError.EvalFailed(s"LET name '$name' is not in scope", None))
      case Some(arr: ArrayResult) if !collapse => Right(arr)
      case Some(arr: ArrayResult) =>
        ScalarCoercion.coerce(s"LET binding '$name'", ScalarCoercion.collapseArray(arr), target)
      case Some(value) => ScalarCoercion.coerce(s"LET binding '$name'", value, target)

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
  private def evalMaybeArray(
    expr: TExpr[?],
    sheet: Sheet,
    clock: Clock,
    workbook: Option[Workbook],
    currentCell: Option[ARef]
  ): Either[EvalError, Any] =
    expr match
      case TExpr.RangeRef(range) =>
        // Convert range to ArrayResult directly
        Right(ArrayArithmetic.rangeToArray(range, sheet))
      case TExpr.SheetRange(sheetName, range) =>
        Evaluator
          .resolveRangeLocation(
            TExpr.RangeLocation.CrossSheet(sheetName, range),
            sheet,
            workbook,
            resolvingNames
          )
          .map { case (targetSheet, _) => ArrayArithmetic.rangeToArray(range, targetSheet) }
      // GH-374: unary plus is transparent in operand positions — =+A1:A3*10 broadcasts
      // exactly like =A1:A3*10
      case TExpr.UnaryPlus(inner) =>
        evalMaybeArray(inner, sheet, clock, workbook, currentCell)
      // GH-302: coerced nodes in OPERAND positions pass ArrayResults through (so
      // =INDIRECT("A1:A3")*10 broadcasts exactly like =A1:A3*10) and coerce scalars totally
      // (so ="16"&"" or a text call result still enters arithmetic per the Numeric table).
      case TExpr.Coerced(inner, target) =>
        evalCoercedExpr(inner, target, sheet, clock, workbook, currentCell, collapse = false)
      case TExpr.CoercedBindingRef(name, target) =>
        evalCoercedBinding(name, target, collapse = false)
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
      case d: Double => Right(ArrayArithmetic.ArrayOperand.Scalar(BigDecimal(d)))
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
   * GH-302: operator positions in scalar mode collapse array results to their top-left value
   * (implicit intersection), consistent with scalar ARGUMENT positions — =INDIRECT("A1")+1 works
   * exactly like =ABS(INDIRECT("A1")). Plain ranges collapse the same way (=A1:A3*10 → A1*10,
   * pinned in ArrayArithmeticSpec). Array mode passes the array through for spill/broadcast.
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
  resolvingNames: Set[String] = Set.empty,
  workbookPath: Option[String] = None
) extends EvaluatorImpl(allowArrayResults, bindings, rng, resolvingNames, workbookPath):
  override protected def currentDepth: Int = depth
  override protected def memoOpt: Option[Evaluator.EvalMemo] = memo
