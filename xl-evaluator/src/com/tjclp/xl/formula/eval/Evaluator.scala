package com.tjclp.xl.formula.eval

import com.tjclp.xl.formula.ast.{BindingCoercion, TExpr}
import com.tjclp.xl.formula.functions.{FunctionSpec, FunctionSpecs, EvalContext}
import com.tjclp.xl.formula.graph.DependencyGraph
import com.tjclp.xl.formula.printer.FormulaPrinter
import com.tjclp.xl.formula.parser.{FormulaParser, ParseError}
import com.tjclp.xl.formula.{Clock, Rng}

import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.workbooks.Workbook
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
 *   3. Short-circuit: And(Lit(false), error) == Right(false) (no error raised)
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

  private[formula] def sheetNotFoundError(
    sheetName: SheetName,
    err: com.tjclp.xl.error.XLError
  ): EvalError =
    EvalError.EvalFailed(
      s"Sheet '${sheetName.value}' not found in workbook: ${err.message}",
      None
    )

  /**
   * Resolve a RangeLocation to the target sheet.
   *
   * For Local ranges, returns the current sheet. For CrossSheet ranges, looks up the target sheet
   * in the workbook context.
   */
  private[formula] def resolveRangeLocation(
    location: TExpr.RangeLocation,
    currentSheet: Sheet,
    workbook: Option[Workbook]
  ): Either[EvalError, Sheet] =
    location match
      case TExpr.RangeLocation.Local(_) =>
        Right(currentSheet)
      case TExpr.RangeLocation.CrossSheet(sheetName, range) =>
        workbook match
          case None =>
            val refStr = s"${sheetName.value}!${range.toA1}"
            Left(missingWorkbookError(refStr, isRange = true))
          case Some(wb) =>
            wb(sheetName) match
              case Left(err) => Left(sheetNotFoundError(sheetName, err))
              case Right(targetSheet) => Right(targetSheet)

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
    rng: Rng = Rng.system
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
          // bindings never leak across formula boundaries (fresh empty environment).
          new EvaluatorWithDepth(depth + 1, rng = rng)
            .eval(expr, targetSheet, clock, workbook)
            .map { result =>
              // Convert typed result to CellValue
              result match
                case cv: CellValue => cv
                case bd: BigDecimal => CellValue.Number(bd)
                case s: String => CellValue.Text(s)
                case b: Boolean => CellValue.Bool(b)
                case i: Int => CellValue.Number(BigDecimal(i))
                case ld: java.time.LocalDate => CellValue.DateTime(ld.atStartOfDay())
                case ldt: java.time.LocalDateTime => CellValue.DateTime(ldt)
                // GH-274: an array-returning call (INDIRECT/OFFSET standalone) read through a
                // reference collapses to its top-left value (the SheetEvaluator.toCellValue
                // scalar-context convention) instead of stringifying.
                case ar: ArrayResult => if ar.isEmpty then CellValue.Empty else ar(0, 0)
                case other => CellValue.Text(other.toString)
            }

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
 */
private class EvaluatorImpl(
  allowArrayResults: Boolean = false,
  bindings: Map[String, Any] = Map.empty,
  rng: Rng = Rng.system
) extends Evaluator:
  /** Current recursion depth for cross-sheet formula evaluation. */
  protected def currentDepth: Int = 0
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
            s"Unresolved SheetPolyRef at ${sheetName.value}!${(at: ARef).toA1} - should have been resolved during parsing",
            None
          )
        )

      case TExpr.SheetRef(sheetName, at, _, decode) =>
        // SheetRef: resolve cell from target sheet in workbook
        workbook match
          case None =>
            val refStr = s"${sheetName.value}!${(at: ARef).toA1}"
            Left(Evaluator.missingWorkbookError(refStr))
          case Some(wb) =>
            wb(sheetName) match
              case Left(err) =>
                Left(Evaluator.sheetNotFoundError(sheetName, err))
              case Right(targetSheet) =>
                val cell = targetSheet(at)
                // GH-161: Handle formula cells without cached values by recursively evaluating
                cell.value match
                  case CellValue.Formula(formulaStr, None) =>
                    // Formula has no cached value - parse and evaluate against target sheet
                    // GH-161 review: Apply decoder to Cell with evaluated result (type-safe)
                    // GH-161 review: Pass currentDepth for cycle protection
                    Evaluator
                      .evalCrossSheetFormula(
                        formulaStr,
                        targetSheet,
                        clock,
                        workbook,
                        currentDepth,
                        rng
                      )
                      .flatMap { evaluatedValue =>
                        val resultCell = Cell(at, evaluatedValue)
                        decode(resultCell).left.map(codecErr => EvalError.CodecFailed(at, codecErr))
                      }
                  case _ =>
                    // Cached formula or non-formula cell - use decoder
                    decode(cell).left.map(codecErr => EvalError.CodecFailed(at, codecErr))

      case TExpr.SheetRange(sheetName, range) =>
        // SheetRange should be wrapped in a function (SUM, COUNT, etc.) before evaluation
        val refStr = s"${sheetName.value}!${range.toA1}"
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
          case CellValue.Formula(formulaStr, None) =>
            Evaluator
              .evalCrossSheetFormula(formulaStr, sheet, clock, workbook, currentDepth, rng)
              .flatMap { evaluatedValue =>
                val resultCell = Cell(at, evaluatedValue)
                decode(resultCell).left.map(codecErr => EvalError.CodecFailed(at, codecErr))
              }
          case _ =>
            decode(cell).left.map(codecErr => EvalError.CodecFailed(at, codecErr))

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

      // ===== String Operators =====
      case TExpr.Concat(x, y) =>
        // Concatenate: join two strings. Operands are statically String, but erased upstream
        // casts (e.g. a numeric LET binding or a numeric-returning call coerced via
        // asStringExpr) can deliver non-String runtime values — evaluate as Any (a String-typed
        // binder would checkcast and throw) and coerce totally with the decodeAsString
        // conventions instead of crashing (GH-193).
        for
          xv <- eval(x.asInstanceOf[TExpr[Any]], sheet, clock, workbook, currentCell)
          yv <- eval(y.asInstanceOf[TExpr[Any]], sheet, clock, workbook, currentCell)
        yield concatText(xv) + concatText(yv)

      // ===== Comparison Operators =====
      // GH-197: Use evalComparison for array-aware comparisons
      case TExpr.Lt(x, y) =>
        // Less than: numeric comparison (array-aware)
        evalComparison(x, y, _ < _, sheet, clock, workbook, currentCell)
          .asInstanceOf[Either[EvalError, A]]

      case TExpr.Lte(x, y) =>
        // Less than or equal: numeric comparison (array-aware)
        evalComparison(x, y, _ <= _, sheet, clock, workbook, currentCell)
          .asInstanceOf[Either[EvalError, A]]

      case TExpr.Gt(x, y) =>
        // Greater than: numeric comparison (array-aware)
        evalComparison(x, y, _ > _, sheet, clock, workbook, currentCell)
          .asInstanceOf[Either[EvalError, A]]

      case TExpr.Gte(x, y) =>
        // Greater than or equal: numeric comparison (array-aware)
        evalComparison(x, y, _ >= _, sheet, clock, workbook, currentCell)
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
            Evaluator.resolveRangeLocation(location, sheet, workbook).flatMap { targetSheet =>
              val cells = location.range.cells.map(cellRef => targetSheet(cellRef))
              // Fold over cells using the aggregator's combine function
              val result = cells.foldLeft(agg.empty) { (acc, cell) =>
                if agg.countsNonEmpty then
                  // COUNTA mode: count any non-empty cell
                  cell.value match
                    case CellValue.Empty => acc
                    case _ => agg.combine(acc, BigDecimal(1))
                else if agg.countsEmpty then
                  // COUNTBLANK mode: count only empty cells
                  cell.value match
                    case CellValue.Empty => agg.combine(acc, BigDecimal(1))
                    case _ => acc
                else
                  // Standard mode: only process numeric values
                  TExpr.decodeNumeric(cell) match
                    case Right(value) => agg.combine(acc, value)
                    case Left(_) if agg.skipNonNumeric => acc
                    case Left(_) => acc // Skip non-numeric cells
              }
              // Finalize and return the result (may return error for AVERAGE on empty range)
              agg.finalizeWithError(result)
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
        // GH-197: Array-aware evaluator for functions like SUMPRODUCT that accept array expressions.
        // GH-193: carries the LET environment and recursion depth so array-evaluated arguments
        // (e.g. SUM over a range-valued binding) still resolve in-scope names.
        def evalArrayArg(expr: TExpr[Any]): Either[EvalError, Any] =
          new EvaluatorWithDepth(currentDepth, allowArrayResults = true, bindings, rng)
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
          rng
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
            new EvaluatorWithDepth(currentDepth, allowArrayResults = true, env, rng)
              .eval(resolved.asInstanceOf[TExpr[Any]], sheet, clock, workbook, currentCell)
              .map(value => env + (name -> unwrapBindingValue(value)))
              .left
              .map { err =>
                EvalError.EvalFailed(
                  s"LET binding '$name': ${EvalError.toXLError(err).message}",
                  None
                )
              }
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
            .resolveRangeLocation(TExpr.RangeLocation.CrossSheet(sheetName, range), sheet, workbook)
            .map(targetSheet => ArrayArithmetic.rangeToArray(range, targetSheet))
        case other =>
          val resolvedBody = TExpr.asResolvedValueExpr(other)
          new EvaluatorWithDepth(currentDepth, allowArrayResults, env, rng)
            .eval(resolvedBody.asInstanceOf[TExpr[Any]], sheet, clock, workbook, currentCell)
    }

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
    case CellValue.Formula(_, Some(cached)) => unwrapBindingValue(cached)
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
            workbook
          )
          .map(targetSheet => ArrayArithmetic.rangeToArray(range, targetSheet))
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
      case _ => Left(EvalError.TypeMismatch("arithmetic", "number or array", value.toString))

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
          if allowArrayResults then Right(arr)
          else Left(EvalError.TypeMismatch("arithmetic", "number", "array"))
    yield output

  /**
   * GH-197: Evaluate comparison with array support.
   *
   * When either operand is a range or array:
   *   - Convert to arrays and perform element-wise comparison
   *   - Return ArrayResult of booleans
   *
   * When both operands are scalars:
   *   - Return plain Boolean (fast path)
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  private def evalComparison(
    xExpr: TExpr[BigDecimal],
    yExpr: TExpr[BigDecimal],
    op: ArrayArithmetic.CompareOp,
    sheet: Sheet,
    clock: Clock,
    workbook: Option[Workbook],
    currentCell: Option[ARef]
  ): Either[EvalError, Any] =
    for
      xVal <- evalMaybeArray(xExpr, sheet, clock, workbook, currentCell)
      yVal <- evalMaybeArray(yExpr, sheet, clock, workbook, currentCell)
      result <- (xVal, yVal) match
        // Both scalars -> plain boolean (fast path)
        case (x: BigDecimal, y: BigDecimal) =>
          Right(op(x, y))
        // At least one array -> element-wise comparison
        case _ =>
          for
            xOp <- toOperand(xVal, sheet)
            yOp <- toOperand(yVal, sheet)
            compared <- ArrayArithmetic.broadcastCompare(xOp, yOp, op)
            output <-
              if allowArrayResults then Right(compared)
              else Left(EvalError.TypeMismatch("comparison", "boolean", "array"))
          yield output
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
          ArrayArithmetic.broadcastEqualityCompare(lArr, Left(rArr), negate).flatMap { compared =>
            if allowArrayResults then Right(compared)
            else Left(EvalError.TypeMismatch("comparison", "boolean", "array"))
          }
        // Left is array, right is scalar -> element-wise comparison
        case (arr: ArrayResult, scalar) =>
          ArrayArithmetic.broadcastEqualityCompare(arr, Right(scalar), negate).flatMap { compared =>
            if allowArrayResults then Right(compared)
            else Left(EvalError.TypeMismatch("comparison", "boolean", "array"))
          }
        // Left is scalar, right is array -> create 1x1 array and broadcast
        case (scalar, arr: ArrayResult) =>
          val scalarArr = ArrayResult.single(ArrayArithmetic.anyToCellValue(scalar))
          ArrayArithmetic.broadcastEqualityCompare(scalarArr, Left(arr), negate).flatMap {
            compared =>
              if allowArrayResults then Right(compared)
              else Left(EvalError.TypeMismatch("comparison", "boolean", "array"))
          }
        // Both scalars -> plain boolean (fast path).
        // GH-234: use the same case-insensitive/coercing semantics as the array path
        // (ArrayArithmetic.cellValueEquals) so scalar and array equality agree with Excel
        // (e.g. ="A"="a" -> TRUE). Previously used raw `x == y` (case-sensitive).
        case (x, y) =>
          val eq = ArrayArithmetic.cellValueEquals(
            ArrayArithmetic.anyToCellValue(x),
            ArrayArithmetic.anyToCellValue(y)
          )
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
  rng: Rng = Rng.system
) extends EvaluatorImpl(allowArrayResults, bindings, rng):
  override protected def currentDepth: Int = depth
