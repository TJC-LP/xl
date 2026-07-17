package com.tjclp.xl.formula.graph

import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.functions.{FunctionSpecs, FunctionRegistry, ArgValue}
import com.tjclp.xl.formula.parser.FormulaParser
import com.tjclp.xl.formula.eval.{EvalError, Evaluator}

import com.tjclp.xl.workbooks.Workbook

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.CellRange
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.sheets.Sheet
import scala.annotation.{nowarn, tailrec}

/**
 * Dependency graph for formula cells.
 *
 * Represents the dependency relationships between cells containing formulas. Each node is a cell
 * reference (ARef), and each directed edge A → B means cell A depends on cell B (A uses B's value).
 *
 * Supports:
 *   - Cycle detection using Tarjan's strongly connected components algorithm
 *   - Topological sorting using Kahn's algorithm for correct evaluation order
 *   - Precedent/dependent queries for impact analysis
 *
 * Design principles:
 *   - Pure functional (no mutations, all operations return new data)
 *   - Total error handling (cycles reported via Either)
 *   - O(V + E) complexity for graph algorithms
 *   - O(1) lookups for precedent/dependent queries (Map-based adjacency lists)
 *
 * Example:
 * {{{
 * // Sheet with formulas: A1="=B1+C1", B1="=10", C1="=20"
 * val graph = DependencyGraph.fromSheet(sheet)
 * graph.precedents(ref"A1") // Set(B1, C1)
 * graph.dependents(ref"B1") // Set(A1)
 *
 * // Detect cycles
 * DependencyGraph.detectCycles(graph) // Right(()) - no cycles
 *
 * // Get evaluation order
 * DependencyGraph.topologicalSort(graph) // Right(List(B1, C1, A1))
 * }}}
 */
final case class DependencyGraph(
  // Forward edges: ref → cells this ref depends on
  dependencies: Map[ARef, Set[ARef]],
  // Reverse edges: ref → cells that depend on this ref
  dependents: Map[ARef, Set[ARef]]
)

object DependencyGraph:
  private def depsFromArgValues(
    values: List[ArgValue],
    exprDeps: TExpr[?] => Set[ARef],
    rangeDeps: TExpr.RangeLocation => Set[ARef],
    cellRangeDeps: CellRange => Set[ARef]
  ): Set[ARef] =
    values.foldLeft(Set.empty[ARef]) { (acc, value) =>
      value match
        case ArgValue.Expr(expr) => acc ++ exprDeps(expr)
        case ArgValue.Range(range) => acc ++ rangeDeps(range)
        case ArgValue.Cells(range) => acc ++ cellRangeDeps(range)
    }

  private def boundedCells(range: CellRange, bounds: Option[CellRange]): Set[ARef] =
    bounds match
      case Some(b) => range.intersect(b).map(_.cells.toSet).getOrElse(Set.empty)
      case None => range.cells.toSet

  /**
   * Build dependency graph from Sheet.
   *
   * Iterates through all cells, extracts references from Formula cells, and constructs the
   * dependency graph. Non-formula cells (constants, text, etc.) are ignored.
   *
   * @param sheet
   *   The sheet to analyze
   * @return
   *   Dependency graph with nodes for all formula cells and edges for all references
   *
   * Example:
   * {{{
   * val sheet = Sheet.empty
   *   .put(ref"A1", CellValue.Formula("=B1+C1"))
   *   .put(ref"B1", CellValue.Number(10))
   *   .put(ref"C1", CellValue.Formula("=D1*2"))
   *
   * val graph = DependencyGraph.fromSheet(sheet)
   * // graph.dependencies = Map(A1 -> Set(B1, C1), C1 -> Set(D1))
   * // graph.dependents = Map(B1 -> Set(A1), C1 -> Set(A1), D1 -> Set(C1))
   * }}}
   */
  def fromSheet(sheet: Sheet): DependencyGraph =
    // Get bounds once for all extractions - constrains full column/row ranges
    val bounds = sheet.usedRange

    val formulaCells = sheet.cells.flatMap { case (ref, cell) =>
      cell.value match
        case CellValue.Formula(expression, _) => Some(ref -> expression)
        case _ => None
    }

    // Build forward edges (dependencies) - use bounded extraction to avoid 1M+ cells
    val dependencies = formulaCells.map { case (ref, formulaStr) =>
      val deps = FormulaParser.parse(formulaStr) match
        case scala.util.Right(expr) => extractDependenciesBounded(expr, bounds)
        case scala.util.Left(_) => Set.empty[ARef] // Parse error: no dependencies
      ref -> deps
    }.toMap

    // Build reverse edges (dependents)
    val dependents = dependencies.foldLeft(Map.empty[ARef, Set[ARef]]) { case (acc, (ref, deps)) =>
      deps.foldLeft(acc) { (acc2, dep) =>
        acc2.updated(dep, acc2.getOrElse(dep, Set.empty) + ref)
      }
    }

    DependencyGraph(dependencies, dependents)

  /**
   * Check if an expression contains any cell references.
   *
   * GH-197: This is a structural check that doesn't enumerate cells in ranges. Use this for quick
   * boolean checks (e.g., "does this formula need a sheet?") instead of extractDependencies which
   * would enumerate 1M+ cells for full-column ranges.
   *
   * @param expr
   *   The expression to analyze
   * @return
   *   true if the expression contains any Ref, PolyRef, RangeRef, or cross-sheet references
   */
  @nowarn("msg=Unreachable case")
  def containsCellReferences[A](expr: TExpr[A]): Boolean =
    expr match
      // Cell references
      case TExpr.Ref(_, _, _) => true
      case TExpr.PolyRef(_, _) => true
      case TExpr.RangeRef(_) => true
      case TExpr.SheetRef(_, _, _, _) => true
      case TExpr.SheetPolyRef(_, _, _) => true
      case TExpr.SheetRange(_, _) => true
      // GH-353: external-workbook refs ARE cell references (they just live outside the workbook)
      case TExpr.ExternalRef(_, _, _, _) => true
      case TExpr.ExternalRange(_, _, _) => true
      case TExpr.Aggregate(_, _) => true

      // Function calls - check arguments
      case call: TExpr.Call[?] =>
        call.spec.argSpec
          .toValues(call.args)
          .exists {
            case ArgValue.Expr(e) => containsCellReferences(e)
            case ArgValue.Range(_) => true
            case ArgValue.Cells(_) => true
          }

      // Binary operators - check both sides
      case TExpr.Add(l, r) => containsCellReferences(l) || containsCellReferences(r)
      case TExpr.Sub(l, r) => containsCellReferences(l) || containsCellReferences(r)
      case TExpr.Mul(l, r) => containsCellReferences(l) || containsCellReferences(r)
      case TExpr.Div(l, r) => containsCellReferences(l) || containsCellReferences(r)
      case TExpr.Pow(l, r) => containsCellReferences(l) || containsCellReferences(r)
      case TExpr.Concat(l, r) => containsCellReferences(l) || containsCellReferences(r)
      case TExpr.Eq(l, r) => containsCellReferences(l) || containsCellReferences(r)
      case TExpr.Neq(l, r) => containsCellReferences(l) || containsCellReferences(r)
      case TExpr.Lt(l, r) => containsCellReferences(l) || containsCellReferences(r)
      case TExpr.Lte(l, r) => containsCellReferences(l) || containsCellReferences(r)
      case TExpr.Gt(l, r) => containsCellReferences(l) || containsCellReferences(r)
      case TExpr.Gte(l, r) => containsCellReferences(l) || containsCellReferences(r)

      // Unary operators
      case TExpr.ToInt(e) => containsCellReferences(e)
      case TExpr.UnaryPlus(e) => containsCellReferences(e)
      case TExpr.Percent(e) => containsCellReferences(e)
      case TExpr.DateToSerial(e) => containsCellReferences(e)
      case TExpr.DateTimeToSerial(e) => containsCellReferences(e)

      // GH-193: LET — check binding values and the body; BindingRef is a name, not a cell
      case TExpr.Let(bindings, body) =>
        bindings.exists((_, value) => containsCellReferences(value)) ||
        containsCellReferences(body)
      case TExpr.BindingRef(_) => false
      case TExpr.CoercedBindingRef(_, _) => false

      // GH-384: a defined name resolves to cells (or a formula over cells) at evaluation time —
      // it behaves like a reference for "does this formula read data?" checks
      case TExpr.NameRef(_) => true
      // GH-394: sheet-qualified names likewise
      case TExpr.SheetNameRef(_, _) => true

      // GH-306: runtime coercion wrapper — transparent for analysis
      case TExpr.Coerced(inner, _) => containsCellReferences(inner)

      // Literals and constants
      case TExpr.Lit(_) => false

  /**
   * GH-274: Check whether an expression contains a dynamic-reference function call.
   *
   * A dynamic reference is a `TExpr.Call` whose spec is flagged `FunctionFlags.dynamicDeps` (e.g.
   * INDIRECT): its data dependencies are decided by evaluated text, so the static graph cannot see
   * them. The walk recurses into Call arguments; `ArgValue.Range`/`Cells` arguments cannot contain
   * calls, and reference/literal leaves are never dynamic.
   */
  def containsDynamicReference[A](expr: TExpr[A]): Boolean =
    expr match
      case call: TExpr.Call[?] =>
        call.spec.flags.dynamicDeps || call.spec.argSpec
          .toValues(call.args)
          .exists {
            case ArgValue.Expr(e) => containsDynamicReference(e)
            case ArgValue.Range(_) => false
            case ArgValue.Cells(_) => false
          }
      case TExpr.Add(l, r) => containsDynamicReference(l) || containsDynamicReference(r)
      case TExpr.Sub(l, r) => containsDynamicReference(l) || containsDynamicReference(r)
      case TExpr.Mul(l, r) => containsDynamicReference(l) || containsDynamicReference(r)
      case TExpr.Div(l, r) => containsDynamicReference(l) || containsDynamicReference(r)
      case TExpr.Pow(l, r) => containsDynamicReference(l) || containsDynamicReference(r)
      case TExpr.Concat(l, r) => containsDynamicReference(l) || containsDynamicReference(r)
      case TExpr.Eq(l, r) => containsDynamicReference(l) || containsDynamicReference(r)
      case TExpr.Neq(l, r) => containsDynamicReference(l) || containsDynamicReference(r)
      case TExpr.Lt(l, r) => containsDynamicReference(l) || containsDynamicReference(r)
      case TExpr.Lte(l, r) => containsDynamicReference(l) || containsDynamicReference(r)
      case TExpr.Gt(l, r) => containsDynamicReference(l) || containsDynamicReference(r)
      case TExpr.Gte(l, r) => containsDynamicReference(l) || containsDynamicReference(r)
      case TExpr.ToInt(e) => containsDynamicReference(e)
      case TExpr.UnaryPlus(e) => containsDynamicReference(e)
      case TExpr.Percent(e) => containsDynamicReference(e)
      case TExpr.DateToSerial(e) => containsDynamicReference(e)
      case TExpr.DateTimeToSerial(e) => containsDynamicReference(e)
      // GH-306: runtime coercion wrapper — a coerced INDIRECT/OFFSET is still dynamic
      case TExpr.Coerced(inner, _) => containsDynamicReference(inner)
      // GH-193: LET — binding values and the body may carry dynamic calls
      case TExpr.Let(bindings, body) =>
        bindings.exists((_, value) => containsDynamicReference(value)) ||
        containsDynamicReference(body)
      case _ => false

  /**
   * GH-274: Find all formula cells in the sheet bearing dynamic references (INDIRECT et al).
   *
   * A case-insensitive substring pre-filter over the registry's dynamicDeps-flagged names runs
   * before any parsing, so sheets without dynamic functions pay no parse cost. Formulas that fail
   * to parse are not dynamic (matching `fromSheet`, where they contribute no edges).
   */
  def dynamicCells(sheet: Sheet): Set[ARef] =
    val names = FunctionRegistry.dynamicFunctionNames
    if names.isEmpty then Set.empty
    else
      sheet.cells.iterator.flatMap { case (ref, cell) =>
        cell.value match
          case CellValue.Formula(expression, _)
              if names.exists(n => expression.toUpperCase.contains(n)) =>
            FormulaParser.parse(expression) match
              case scala.util.Right(expr) if containsDynamicReference(expr) => Some(ref)
              case _ => None
          case _ => None
      }.toSet

  /**
   * GH-274: A dynamic cell set closed under "depends on me": the cells themselves plus every
   * transitive static dependent. `transitiveDependents` excludes its seeds, so the union is
   * required. This is the bucket `deferDynamic` moves to the end of evaluation order.
   */
  def dynamicClosure(graph: DependencyGraph, dynamic: Set[ARef]): Set[ARef] =
    dynamic ++ transitiveDependents(graph, dynamic)

  /**
   * GH-274: Stable partition of a topological order — non-closure cells first, closure cells last,
   * relative order preserved within each part.
   *
   * Lemma (order preservation): if `order` is a valid topological order of the graph and `closure`
   * is dependent-closed (contains every static dependent of each of its members), the result is
   * also a valid topological order. For any edge u-depends-on-v: if u and v are in the same part,
   * their relative order is preserved (stable partition); if u is deferred and v is not, v sits in
   * the front part, before u; v deferred with u not deferred is impossible — u depends on v, so u
   * is a dependent of v and dependent-closure would have pulled u into the closure. Property-tested
   * in DependencyGraphSpec.
   */
  def deferDynamic(order: List[ARef], closure: Set[ARef]): List[ARef] =
    if closure.isEmpty then order
    else order.filterNot(closure.contains) ++ order.filter(closure.contains)

  /**
   * Check if an expression contains any **unqualified** cell references.
   *
   * GH-210: Fully-qualified references (SheetRef, SheetPolyRef, SheetRange, CrossSheet aggregates)
   * already name their target sheet, so the formula doesn't require a `-s` flag. Only unqualified
   * refs (Ref, PolyRef, RangeRef, Local aggregates) need an ambient sheet context.
   *
   * @param expr
   *   The expression to analyze
   * @return
   *   true if the expression contains any unqualified cell reference
   */
  @nowarn("msg=Unreachable case")
  def containsUnqualifiedCellReferences[A](expr: TExpr[A]): Boolean =
    expr match
      // Unqualified cell references - need ambient sheet
      case TExpr.Ref(_, _, _) => true
      case TExpr.PolyRef(_, _) => true
      case TExpr.RangeRef(_) => true
      case TExpr.Aggregate(_, TExpr.RangeLocation.Local(_)) => true

      // Qualified cell references - sheet already specified
      case TExpr.SheetRef(_, _, _, _) => false
      case TExpr.SheetPolyRef(_, _, _) => false
      case TExpr.SheetRange(_, _) => false
      // GH-353: external refs are fully qualified (workbook + sheet) — no ambient sheet needed
      case TExpr.ExternalRef(_, _, _, _) => false
      case TExpr.ExternalRange(_, _, _) => false
      case TExpr.Aggregate(_, TExpr.RangeLocation.CrossSheet(_, _)) => false
      case TExpr.Aggregate(_, TExpr.RangeLocation.External(_, _, _)) => false
      // GH-394: an unqualified name's lookup depends on the ambient sheet (sheet-scoped names
      // shadow workbook-scoped ones); a sheet-qualified name carries its own context
      case TExpr.Aggregate(_, TExpr.RangeLocation.Name(_, scope)) => scope.isEmpty

      // Function calls - check arguments
      case call: TExpr.Call[?] =>
        call.spec.argSpec
          .toValues(call.args)
          .exists {
            case ArgValue.Expr(e) => containsUnqualifiedCellReferences(e)
            case ArgValue.Range(loc) =>
              loc match
                case TExpr.RangeLocation.Local(_) => true
                case TExpr.RangeLocation.CrossSheet(_, _) => false
                // GH-353: fully qualified (workbook + sheet) — no ambient sheet needed
                case TExpr.RangeLocation.External(_, _, _) => false
                // GH-394: unqualified name lookup depends on the ambient sheet
                case TExpr.RangeLocation.Name(_, scope) => scope.isEmpty
            case ArgValue.Cells(_) => true
          }

      // Binary operators - check both sides
      case TExpr.Add(l, r) =>
        containsUnqualifiedCellReferences(l) || containsUnqualifiedCellReferences(r)
      case TExpr.Sub(l, r) =>
        containsUnqualifiedCellReferences(l) || containsUnqualifiedCellReferences(r)
      case TExpr.Mul(l, r) =>
        containsUnqualifiedCellReferences(l) || containsUnqualifiedCellReferences(r)
      case TExpr.Div(l, r) =>
        containsUnqualifiedCellReferences(l) || containsUnqualifiedCellReferences(r)
      case TExpr.Pow(l, r) =>
        containsUnqualifiedCellReferences(l) || containsUnqualifiedCellReferences(r)
      case TExpr.Concat(l, r) =>
        containsUnqualifiedCellReferences(l) || containsUnqualifiedCellReferences(r)
      case TExpr.Eq(l, r) =>
        containsUnqualifiedCellReferences(l) || containsUnqualifiedCellReferences(r)
      case TExpr.Neq(l, r) =>
        containsUnqualifiedCellReferences(l) || containsUnqualifiedCellReferences(r)
      case TExpr.Lt(l, r) =>
        containsUnqualifiedCellReferences(l) || containsUnqualifiedCellReferences(r)
      case TExpr.Lte(l, r) =>
        containsUnqualifiedCellReferences(l) || containsUnqualifiedCellReferences(r)
      case TExpr.Gt(l, r) =>
        containsUnqualifiedCellReferences(l) || containsUnqualifiedCellReferences(r)
      case TExpr.Gte(l, r) =>
        containsUnqualifiedCellReferences(l) || containsUnqualifiedCellReferences(r)

      // Unary operators
      case TExpr.ToInt(e) => containsUnqualifiedCellReferences(e)
      case TExpr.UnaryPlus(e) => containsUnqualifiedCellReferences(e)
      case TExpr.Percent(e) => containsUnqualifiedCellReferences(e)
      case TExpr.DateToSerial(e) => containsUnqualifiedCellReferences(e)
      case TExpr.DateTimeToSerial(e) => containsUnqualifiedCellReferences(e)

      // GH-193: LET — check binding values and the body; BindingRef is a name, not a cell
      case TExpr.Let(bindings, body) =>
        bindings.exists((_, value) => containsUnqualifiedCellReferences(value)) ||
        containsUnqualifiedCellReferences(body)
      case TExpr.BindingRef(_) => false
      case TExpr.CoercedBindingRef(_, _) => false

      // GH-384: name resolution depends on the ambient sheet (sheet-scoped names SHADOW
      // workbook-scoped ones), so a name-bearing formula needs a sheet context
      case TExpr.NameRef(_) => true
      // GH-394: a sheet-qualified name carries its own lookup context
      case TExpr.SheetNameRef(_, _) => false

      // GH-306: runtime coercion wrapper — transparent for analysis
      case TExpr.Coerced(inner, _) => containsUnqualifiedCellReferences(inner)

      // Literals and constants
      case TExpr.Lit(_) => false

  /**
   * Extract all cell references from TExpr.
   *
   * Recursively traverses the expression AST and collects all cell references, including:
   *   - Single cell references (Ref)
   *   - Range references (RangeRef) expanded to all cells in range
   *
   * @param expr
   *   The expression to analyze
   * @return
   *   Set of all cell references used in the expression
   *
   * Example:
   * {{{
   * val expr = TExpr.Add(TExpr.Ref(ref"A1", ...), TExpr.Lit(10))
   * extractDependencies(expr) // Set(A1)
   *
   * val sumExpr = TExpr.sum(CellRange.parse("B1:B10").toOption.get)
   * extractDependencies(sumExpr) // Set(B1, B2, ..., B10)
   * }}}
   */
  // nowarn: Compiler incorrectly reports PolyRef as unreachable, but tests confirm it IS reached at runtime
  @nowarn("msg=Unreachable case")
  def extractDependencies[A](expr: TExpr[A]): Set[ARef] =
    expr match
      // Single cell reference
      case TExpr.Ref(at, _, _) => Set(at)

      // Polymorphic reference (type resolved at evaluation time)
      case TExpr.PolyRef(at, _) => Set(at)

      // Cross-sheet references return Set.empty in same-sheet dependency extraction.
      // This is intentional: extractDependencies builds intra-sheet graphs only.
      // For workbook-level dependency tracking, use extractQualifiedDependencies + fromWorkbook.
      case TExpr.SheetRef(_, _, _, _) => Set.empty
      case TExpr.SheetPolyRef(_, _, _) => Set.empty
      case TExpr.SheetRange(_, _) => Set.empty
      // GH-353: external-workbook refs target cells OUTSIDE the workbook — no edges ever
      case TExpr.ExternalRef(_, _, _, _) => Set.empty
      case TExpr.ExternalRange(_, _, _) => Set.empty
      case TExpr.RangeRef(range) =>
        range.cells.toSet

      case call: TExpr.Call[?] =>
        depsFromArgValues(
          call.spec.argSpec.toValues(call.args),
          expr => extractDependencies(expr),
          _.localCells,
          _.cells.toSet
        )

      // Recursive cases (binary operators)
      case TExpr.Add(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.Sub(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.Mul(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.Div(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.Pow(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.Concat(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.Eq(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.Neq(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.Lt(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.Lte(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.Gt(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.Gte(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.ToInt(expr) =>
        extractDependencies(expr) // Type conversion - extract from wrapped expr
      // GH-374: unary plus is transparent — dependencies under it feed recalc edges
      case TExpr.UnaryPlus(expr) => extractDependencies(expr)
      // GH-355: postfix percent — dependencies under it feed recalc edges
      case TExpr.Percent(expr) => extractDependencies(expr)
      case TExpr.Aggregate(_, location) => location.localCells

      // GH-193: LET — union of binding-value and body dependencies; BindingRef has none
      case TExpr.Let(bindings, body) =>
        bindings.foldLeft(extractDependencies(body)) { case (acc, (_, value)) =>
          acc ++ extractDependencies(value)
        }
      case TExpr.BindingRef(_) => Set.empty
      case TExpr.CoercedBindingRef(_, _) => Set.empty
      // GH-384: name targets live behind workbook metadata this sheet-level walk cannot see —
      // no intra-sheet edges (consistent with cross-sheet refs; the qualified extractor resolves).
      // GH-394: sheet-qualified names likewise — NOTE this means an edit to a name's TARGET does
      // not trigger sheet-level targeted recalc of name users; the workbook-level qualified
      // extractor below carries the real edges.
      case TExpr.NameRef(_) => Set.empty
      case TExpr.SheetNameRef(_, _) => Set.empty

      // GH-306: runtime coercion wrapper — transparent for analysis
      case TExpr.Coerced(inner, _) => extractDependencies(inner)

      // Literals and nullary functions (no dependencies)
      case TExpr.Lit(_) => Set.empty
      case TExpr.DateToSerial(dateExpr) => extractDependencies(dateExpr)
      case TExpr.DateTimeToSerial(dtExpr) => extractDependencies(dtExpr)

  /**
   * Extract all cell references from TExpr, bounded by the sheet's used range.
   *
   * This optimized version constrains full column/row references (like A:A or 1:1) to the
   * intersection with bounds, avoiding iteration over 1M+ cells. Use this when building dependency
   * graphs from sheets.
   *
   * @param expr
   *   The expression to analyze
   * @param bounds
   *   Optional bounding range (typically sheet.usedRange) to constrain full ranges
   * @return
   *   Set of all cell references used in the expression, bounded by the given range
   */
  @nowarn("msg=Unreachable case")
  def extractDependenciesBounded[A](expr: TExpr[A], bounds: Option[CellRange]): Set[ARef] =
    // Helper to bound a CellRange
    def boundRange(range: CellRange): Set[ARef] =
      bounds match
        case Some(b) => range.intersect(b).map(_.cells.toSet).getOrElse(Set.empty)
        case None => range.cells.toSet

    expr match
      // Single cell reference
      case TExpr.Ref(at, _, _) => Set(at)

      // Polymorphic reference (type resolved at evaluation time)
      case TExpr.PolyRef(at, _) => Set(at)

      // Cross-sheet references return Set.empty in same-sheet dependency extraction.
      case TExpr.SheetRef(_, _, _, _) => Set.empty
      case TExpr.SheetPolyRef(_, _, _) => Set.empty
      case TExpr.SheetRange(_, _) => Set.empty
      // GH-353: external-workbook refs target cells OUTSIDE the workbook — no edges ever
      case TExpr.ExternalRef(_, _, _, _) => Set.empty
      case TExpr.ExternalRange(_, _, _) => Set.empty
      case TExpr.RangeRef(range) => boundRange(range)

      // Recursive cases (binary operators)
      case TExpr.Add(l, r) =>
        extractDependenciesBounded(l, bounds) ++ extractDependenciesBounded(r, bounds)
      case TExpr.Sub(l, r) =>
        extractDependenciesBounded(l, bounds) ++ extractDependenciesBounded(r, bounds)
      case TExpr.Mul(l, r) =>
        extractDependenciesBounded(l, bounds) ++ extractDependenciesBounded(r, bounds)
      case TExpr.Div(l, r) =>
        extractDependenciesBounded(l, bounds) ++ extractDependenciesBounded(r, bounds)
      case TExpr.Pow(l, r) =>
        extractDependenciesBounded(l, bounds) ++ extractDependenciesBounded(r, bounds)
      case TExpr.Concat(l, r) =>
        extractDependenciesBounded(l, bounds) ++ extractDependenciesBounded(r, bounds)
      case TExpr.Eq(l, r) =>
        extractDependenciesBounded(l, bounds) ++ extractDependenciesBounded(r, bounds)
      case TExpr.Neq(l, r) =>
        extractDependenciesBounded(l, bounds) ++ extractDependenciesBounded(r, bounds)
      case TExpr.Lt(l, r) =>
        extractDependenciesBounded(l, bounds) ++ extractDependenciesBounded(r, bounds)
      case TExpr.Lte(l, r) =>
        extractDependenciesBounded(l, bounds) ++ extractDependenciesBounded(r, bounds)
      case TExpr.Gt(l, r) =>
        extractDependenciesBounded(l, bounds) ++ extractDependenciesBounded(r, bounds)
      case TExpr.Gte(l, r) =>
        extractDependenciesBounded(l, bounds) ++ extractDependenciesBounded(r, bounds)
      case TExpr.ToInt(expr) => extractDependenciesBounded(expr, bounds)
      case TExpr.UnaryPlus(expr) => extractDependenciesBounded(expr, bounds)
      case TExpr.Percent(expr) => extractDependenciesBounded(expr, bounds)
      case TExpr.Aggregate(_, location) => location.localCellsBounded(bounds)

      case call: TExpr.Call[?] =>
        depsFromArgValues(
          call.spec.argSpec.toValues(call.args),
          expr => extractDependenciesBounded(expr, bounds),
          loc => loc.localCellsBounded(bounds),
          range => boundedCells(range, bounds)
        )

      // GH-193: LET — union of binding-value and body dependencies; BindingRef has none
      case TExpr.Let(bindings, body) =>
        bindings.foldLeft(extractDependenciesBounded(body, bounds)) { case (acc, (_, value)) =>
          acc ++ extractDependenciesBounded(value, bounds)
        }
      case TExpr.BindingRef(_) => Set.empty
      case TExpr.CoercedBindingRef(_, _) => Set.empty
      // GH-384: name targets live behind workbook metadata this sheet-level walk cannot see —
      // no intra-sheet edges (consistent with cross-sheet refs; the qualified extractor resolves).
      // GH-394: sheet-qualified names likewise — NOTE this means an edit to a name's TARGET does
      // not trigger sheet-level targeted recalc of name users; the workbook-level qualified
      // extractor below carries the real edges.
      case TExpr.NameRef(_) => Set.empty
      case TExpr.SheetNameRef(_, _) => Set.empty

      // GH-306: runtime coercion wrapper — transparent for analysis
      case TExpr.Coerced(inner, _) => extractDependenciesBounded(inner, bounds)

      // Literals and nullary functions (no dependencies)
      case TExpr.Lit(_) => Set.empty
      case TExpr.DateToSerial(dateExpr) => extractDependenciesBounded(dateExpr, bounds)
      case TExpr.DateTimeToSerial(dtExpr) => extractDependenciesBounded(dtExpr, bounds)

  /**
   * Get cells this cell depends on (precedents).
   *
   * Returns the set of cells whose values are used in calculating this cell's value. If the cell
   * has no formula or is not in the graph, returns empty set.
   *
   * @param graph
   *   The dependency graph
   * @param ref
   *   The cell reference to query
   * @return
   *   Set of cells this cell depends on (may be empty)
   *
   * Example:
   * {{{
   * // A1 = "=B1+C1"
   * precedents(graph, ref"A1") // Set(B1, C1)
   * precedents(graph, ref"B1") // Set() - B1 is a constant
   * }}}
   */
  def precedents(graph: DependencyGraph, ref: ARef): Set[ARef] =
    graph.dependencies.getOrElse(ref, Set.empty)

  /**
   * Get cells that depend on this cell (dependents).
   *
   * Returns the set of cells that use this cell's value in their calculations. If no cells depend
   * on this cell, returns empty set.
   *
   * @param graph
   *   The dependency graph
   * @param ref
   *   The cell reference to query
   * @return
   *   Set of cells that depend on this cell (may be empty)
   *
   * Example:
   * {{{
   * // A1 = "=B1+C1", D1 = "=B1*2"
   * dependents(graph, ref"B1") // Set(A1, D1)
   * dependents(graph, ref"A1") // Set() - nothing depends on A1
   * }}}
   */
  def dependents(graph: DependencyGraph, ref: ARef): Set[ARef] =
    graph.dependents.getOrElse(ref, Set.empty)

  /**
   * Compute transitive dependencies for a set of cells.
   *
   * Given a set of starting cells, returns all cells that are directly or transitively depended
   * upon. This is useful for targeted evaluation - to evaluate only formulas in a range, we need to
   * also evaluate all cells they depend on (recursively).
   *
   * @param graph
   *   The dependency graph to traverse
   * @param refs
   *   The starting cell references
   * @return
   *   Set of all cells that the starting cells depend on (transitively)
   *
   * Example:
   * {{{
   * // A1="=B1+C1", B1="=D1", C1="=10", D1="=20"
   * // transitiveDependencies(graph, Set(A1))
   * // Returns: Set(B1, C1, D1) - all cells A1 depends on directly or indirectly
   * }}}
   */
  @scala.annotation.tailrec
  def transitiveDependencies(
    graph: DependencyGraph,
    refs: Set[ARef],
    visited: Set[ARef] = Set.empty
  ): Set[ARef] =
    val toVisit = refs -- visited
    if toVisit.isEmpty then visited
    else
      // Get direct dependencies of all cells in toVisit
      val directDeps = toVisit.flatMap(ref => graph.dependencies.getOrElse(ref, Set.empty))
      // Recurse with direct deps as new frontier
      transitiveDependencies(graph, directDeps, visited ++ toVisit)

  /**
   * Compute transitive dependents (reverse of transitiveDependencies).
   *
   * Returns all cells that depend on the given cells, directly or transitively. Used for eager
   * recalculation - when cell X changes, find all formulas to recalculate.
   *
   * @param graph
   *   The dependency graph to traverse
   * @param refs
   *   The starting cell references (modified cells)
   * @return
   *   Set of all cells that depend on the starting cells (transitively), excluding the starting
   *   refs themselves
   *
   * Example:
   * {{{
   * // A1="=10", B1="=A1*2", C1="=B1+5"
   * // transitiveDependents(graph, Set(A1))
   * // Returns: Set(B1, C1) - all cells affected when A1 changes
   * }}}
   */
  def transitiveDependents(
    graph: DependencyGraph,
    refs: Set[ARef]
  ): Set[ARef] =
    transitiveDependentsOf(graph.dependents, refs)

  /**
   * Reverse-reachability core shared by the sheet-level and workbook-level (qualified) variants —
   * generic in the node type (GH-346). Excludes the starting refs from the result.
   */
  private def transitiveDependentsOf[K](
    dependents: Map[K, Set[K]],
    originalRefs: Set[K]
  ): Set[K] =
    @scala.annotation.tailrec
    def impl(frontier: Set[K], visited: Set[K]): Set[K] =
      val toVisit = frontier -- visited
      if toVisit.isEmpty then visited -- originalRefs // Exclude starting refs
      else
        val directDeps = toVisit.flatMap(ref => dependents.getOrElse(ref, Set.empty))
        impl(directDeps, visited ++ toVisit)
    impl(originalRefs, Set.empty)

  /**
   * Iterative Tarjan SCC engine shared by `detectCycles`, `cyclicNodes`, and their qualified
   * (workbook-level) counterparts — generic in the node type (GH-346).
   *
   * Uses an explicit work stack instead of recursion: `recalculate` promises totality, and a deep
   * linear dependency chain (e.g. 100k sequential formulas) must not throw StackOverflowError.
   * Invokes `onScc` for every completed component in Tarjan pop order (deepest member first,
   * component root last).
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  private def foreachSccOf[K](dependencies: Map[K, Set[K]])(onScc: List[K] => Unit): Unit =
    var index = 0
    var indices = Map.empty[K, Int]
    var lowLinks = Map.empty[K, Int]
    var stack = List.empty[K]
    var onStack = Set.empty[K]
    // DFS frames: (node, successors not yet examined)
    var frames = List.empty[(K, List[K])]

    def push(v: K): Unit =
      indices = indices.updated(v, index)
      lowLinks = lowLinks.updated(v, index)
      index += 1
      stack = v :: stack
      onStack = onStack + v
      frames = (v, dependencies.getOrElse(v, Set.empty).toList) :: frames

    dependencies.keySet.foreach { root =>
      if !indices.contains(root) then
        push(root)
        while frames.nonEmpty do
          frames match
            case (v, w :: rest) :: tail =>
              frames = (v, rest) :: tail
              if !indices.contains(w) then push(w)
              else if onStack.contains(w) then
                lowLinks = lowLinks.updated(v, math.min(lowLinks(v), indices(w)))
            case (v, Nil) :: tail =>
              frames = tail
              if lowLinks(v) == indices(v) then
                val (sccTail, remaining) = stack.span(_ != v)
                stack = remaining.drop(1)
                val scc = sccTail :+ v
                onStack = onStack -- scc
                onScc(scc)
              frames match
                case (p, _) :: _ =>
                  lowLinks = lowLinks.updated(p, math.min(lowLinks(p), lowLinks(v)))
                case Nil => ()
            case Nil => () // unreachable: guarded by frames.nonEmpty
    }

  private def foreachScc(graph: DependencyGraph)(onScc: List[ARef] => Unit): Unit =
    foreachSccOf(graph.dependencies)(onScc)

  /**
   * Detect circular references using Tarjan's strongly connected components algorithm.
   *
   * A circular reference occurs when a cell's formula depends (directly or transitively) on its own
   * value. For example: A1="=B1", B1="=C1", C1="=A1" forms a cycle.
   *
   * Runs in O(V + E) with an explicit work stack (stack-safe on deep dependency chains).
   *
   * @param graph
   *   The dependency graph to analyze
   * @return
   *   Left(CircularRef) if a cycle is detected (includes cycle path), Right(()) if acyclic
   *
   * Example:
   * {{{
   * // No cycle: A1="=10", B1="=A1+5"
   * detectCycles(graph) // Right(())
   *
   * // Cycle: A1="=B1", B1="=A1"
   * detectCycles(graph) // Left(EvalError.CircularRef(List(A1, B1, A1)))
   * }}}
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  def detectCycles(graph: DependencyGraph): Either[EvalError.CircularRef, Unit] =
    var found = Option.empty[List[ARef]]
    foreachScc(graph) { scc =>
      if found.isEmpty then
        scc match
          case single :: Nil =>
            if graph.dependencies.get(single).exists(_.contains(single)) then
              found = Some(List(single, single)) // Self-loop: v -> v
          case multi =>
            val cycleNodes = multi.reverse
            found =
              Some(cycleNodes :+ cycleNodes.headOption.getOrElse(multi.last)) // close the cycle
    }
    found match
      case Some(cycle) => scala.util.Left(EvalError.CircularRef(cycle))
      case None => scala.util.Right(())

  /**
   * Collect every node that participates in a reference cycle.
   *
   * Unlike `detectCycles` (which fails fast on the first cycle), this identifies the complete set
   * of cells inside strongly connected components of size > 1, plus self-loops. Used by
   * `Workbook.recalculate` to isolate cyclic cells while still evaluating the acyclic remainder.
   *
   * Stack-safe: shares the iterative Tarjan engine with `detectCycles`.
   *
   * @return
   *   The set of cycle participants (empty when the graph is acyclic)
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  def cyclicNodes(graph: DependencyGraph): Set[ARef] =
    var cyclic = Set.empty[ARef]
    foreachScc(graph) { scc =>
      scc match
        case single :: Nil =>
          if graph.dependencies.get(single).exists(_.contains(single)) then cyclic = cyclic + single
        case multi => cyclic = cyclic ++ multi
    }
    cyclic

  /**
   * Kahn's-algorithm core shared by the sheet-level and workbook-level (qualified) sorts — generic
   * in the node type (GH-346). Only nodes present in `dependencies` (formula cells) are ordered;
   * edges to non-nodes (constants, empty cells) are ignored. Left carries the nodes that could not
   * be ordered (cycle participants and everything downstream of them).
   */
  private def kahnOrder[K](
    dependencies: Map[K, Set[K]],
    dependents: Map[K, Set[K]]
  ): Either[Set[K], List[K]] =
    // All formula cells (only process formulas, not constants)
    val allNodes = dependencies.keySet

    // If no formula cells, early exit
    if allNodes.isEmpty then scala.util.Right(List.empty[K])
    else
      // Calculate in-degree for each node (number of formula cells it depends on)
      // Only count dependencies on other formula cells, not constants
      val inDegree = allNodes.map { node =>
        val deps = dependencies.getOrElse(node, Set.empty)
        val formulaDeps = deps.filter(allNodes.contains)
        node -> formulaDeps.size
      }.toMap

      // Start with nodes that have in-degree 0 (no dependencies)
      val initialQueue = allNodes.filter(node => inDegree(node) == 0).toList

      @tailrec
      def process(
        queue: List[K],
        processedInDegree: Map[K, Int],
        acc: List[K]
      ): (List[K], Map[K, Int]) =
        queue match
          case Nil => (acc, processedInDegree)
          case node :: rest =>
            val deps = dependents.getOrElse(node, Set.empty).filter(allNodes.contains)
            val (nextInDegree, newlyZero) = deps.foldLeft((processedInDegree, List.empty[K])) {
              case ((degreeAcc, zeros), dep) =>
                val newInDegree = degreeAcc(dep) - 1
                val updatedDegree = degreeAcc.updated(dep, newInDegree)
                val nextZeros = if newInDegree == 0 then zeros :+ dep else zeros
                (updatedDegree, nextZeros)
            }
            process(rest ++ newlyZero, nextInDegree, acc :+ node)

      val (result, _) = process(initialQueue, inDegree, List.empty)

      // If all nodes are processed, graph is acyclic
      if result.size == allNodes.size then scala.util.Right(result)
      else scala.util.Left(allNodes -- result.toSet)

  /**
   * Topological sort using Kahn's algorithm.
   *
   * Returns a linear ordering of cells such that for every dependency A → B, cell B appears before
   * cell A in the ordering. This ensures formulas are evaluated in the correct order (dependencies
   * before dependents).
   *
   * Uses Kahn's algorithm which runs in O(V + E) time. The algorithm maintains a queue of nodes
   * with in-degree 0 and processes them in order, removing edges as it goes.
   *
   * @param graph
   *   The dependency graph to sort
   * @return
   *   Left(CircularRef) if a cycle is detected, Right(evaluation order) if acyclic
   *
   * Example:
   * {{{
   * // A1="=B1+C1", B1="=10", C1="=20"
   * topologicalSort(graph) // Right(List(B1, C1, A1))
   *
   * // A1="=B1", B1="=A1" (cycle)
   * topologicalSort(graph) // Left(EvalError.CircularRef(List(A1, B1, A1)))
   * }}}
   */
  def topologicalSort(graph: DependencyGraph): Either[EvalError.CircularRef, List[ARef]] =
    kahnOrder(graph.dependencies, graph.dependents) match
      case scala.util.Right(order) => scala.util.Right(order)
      case scala.util.Left(remainingNodes) =>
        // Cycle detected: find one cycle for error reporting
        val cycle = remainingNodes.headOption match
          case Some(start) =>
            // Follow dependencies to reconstruct cycle
            def findCycle(current: ARef, visited: Set[ARef]): List[ARef] =
              if visited.contains(current) then
                // Found cycle
                List(current)
              else
                graph.dependencies.getOrElse(current, Set.empty).headOption match
                  case Some(next) if remainingNodes.contains(next) =>
                    current :: findCycle(next, visited + current)
                  case _ => List(current)

            val cyclePath = findCycle(start, Set.empty)
            cyclePath.headOption.map(first => cyclePath :+ first).getOrElse(List.empty)
          case None => List.empty

        scala.util.Left(EvalError.CircularRef(cycle))

  // ===== Cross-Sheet Dependency Tracking =====

  /**
   * Cell reference qualified with sheet name for cross-sheet tracking.
   *
   * Used to track dependencies across sheets within a workbook. Each QualifiedRef uniquely
   * identifies a cell in the workbook by combining the sheet name and cell reference.
   *
   * Example:
   * {{{
   * val ref = QualifiedRef(SheetName.unsafe("Sales"), ref"A1")
   * // Represents Sales!A1
   * }}}
   */
  final case class QualifiedRef(sheet: SheetName, ref: ARef):
    // GH-280: quote cell-ref-shaped sheet names so a sheet literally named "A1" renders
    // unambiguously in diagnostics ('A1'!B2, not A1!B2)
    override def toString: String = s"${SheetName.quoteForFormula(sheet.value)}!${ref.toA1}"

  /**
   * Build dependency graph from Workbook (cross-sheet aware).
   *
   * Iterates through all sheets and cells, extracting references from Formula cells and
   * constructing a workbook-level dependency graph. Cross-sheet references are properly tracked
   * using QualifiedRef.
   *
   * @param workbook
   *   The workbook to analyze
   * @return
   *   Dependency graph with QualifiedRef nodes covering all sheets
   *
   * Example:
   * {{{
   * // Sheet1!A1 = "=Sheet2!B1", Sheet2!B1 = 10
   * val graph = DependencyGraph.fromWorkbook(workbook)
   * // graph contains: QualifiedRef(Sheet1, A1) -> Set(QualifiedRef(Sheet2, B1))
   * }}}
   */
  def fromWorkbook(
    workbook: com.tjclp.xl.workbooks.Workbook
  ): Map[QualifiedRef, Set[QualifiedRef]] =
    workbook.sheets.flatMap { sheet =>
      sheet.cells.flatMap { case (cellRef, cell) =>
        cell.value match
          case CellValue.Formula(expression, _) =>
            val deps = FormulaParser.parse(expression) match
              case scala.util.Right(expr) =>
                extractQualifiedDependencies(expr, sheet.name, workbook = Some(workbook))
              case scala.util.Left(_) => Set.empty[QualifiedRef]
            Some(QualifiedRef(sheet.name, cellRef) -> deps)
          case _ => None
      }
    }.toMap

  /**
   * GH-346: workbook-level dependency graph with ranges bounded by each TARGET sheet's used range —
   * the qualified analog of `fromSheet`'s bounded extraction (a full-column reference must not
   * enumerate 1M+ cells). Single refs stay exact. Ranges over sheets absent from the workbook (or
   * empty ones) contribute no edges: they cannot affect evaluation order, and the referencing
   * formula still errors per cell on evaluation.
   *
   * Returns forward edges (ref → its dependencies) and reverse edges (ref → its dependents), which
   * together drive `qualifiedTopologicalSort`.
   */
  def fromWorkbookBounded(
    workbook: com.tjclp.xl.workbooks.Workbook
  ): (Map[QualifiedRef, Set[QualifiedRef]], Map[QualifiedRef, Set[QualifiedRef]]) =
    val bounds: Map[SheetName, Option[CellRange]] =
      workbook.sheets.map(s => s.name -> s.usedRange).toMap
    val cellsFor: (SheetName, CellRange) => Set[QualifiedRef] = (sheet, range) =>
      bounds.get(sheet) match
        case Some(Some(b)) =>
          range.intersect(b) match
            case Some(clipped) => clipped.cells.map(ref => QualifiedRef(sheet, ref)).toSet
            case None => Set.empty
        case _ => Set.empty // empty or missing sheet: nothing to depend on

    val dependencies = workbook.sheets.flatMap { sheet =>
      sheet.cells.flatMap { case (cellRef, cell) =>
        cell.value match
          case CellValue.Formula(expression, _) =>
            val deps = FormulaParser.parse(expression) match
              case scala.util.Right(expr) =>
                extractQualifiedDependencies(expr, sheet.name, cellsFor, Some(workbook))
              case scala.util.Left(_) => Set.empty[QualifiedRef]
            Some(QualifiedRef(sheet.name, cellRef) -> deps)
          case _ => None
      }
    }.toMap

    val dependents =
      dependencies.foldLeft(Map.empty[QualifiedRef, Set[QualifiedRef]]) { case (acc, (ref, deps)) =>
        deps.foldLeft(acc) { (acc2, dep) =>
          acc2.updated(dep, acc2.getOrElse(dep, Set.empty) + ref)
        }
      }

    (dependencies, dependents)

  /**
   * GH-346: every node participating in a reference cycle of the workbook-level graph — the
   * qualified analog of `cyclicNodes`, sharing the iterative (stack-safe) Tarjan engine. Detects
   * cycles that span sheets, which per-sheet analysis cannot see.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  def qualifiedCyclicNodes(
    dependencies: Map[QualifiedRef, Set[QualifiedRef]]
  ): Set[QualifiedRef] =
    var cyclic = Set.empty[QualifiedRef]
    foreachSccOf(dependencies) {
      case single :: Nil =>
        if dependencies.get(single).exists(_.contains(single)) then cyclic = cyclic + single
      case multi => cyclic = cyclic ++ multi
    }
    cyclic

  /**
   * GH-346: transitive dependents over the workbook-level reverse edges (excludes the starting
   * refs) — the qualified analog of `transitiveDependents`.
   */
  def qualifiedTransitiveDependents(
    dependents: Map[QualifiedRef, Set[QualifiedRef]],
    refs: Set[QualifiedRef]
  ): Set[QualifiedRef] =
    transitiveDependentsOf(dependents, refs)

  /**
   * GH-346: Kahn order over the workbook-level graph. Left carries the nodes that could not be
   * ordered (cycle participants and their downstream); callers that prune cycles first only reach
   * Left as a defensive-totality path.
   */
  def qualifiedTopologicalSort(
    dependencies: Map[QualifiedRef, Set[QualifiedRef]],
    dependents: Map[QualifiedRef, Set[QualifiedRef]]
  ): Either[Set[QualifiedRef], List[QualifiedRef]] =
    kahnOrder(dependencies, dependents)

  /** Unbounded range expansion — the original `fromWorkbook` semantics. */
  private val unboundedQualifiedCells: (SheetName, CellRange) => Set[QualifiedRef] =
    (sheet, range) => range.cells.map(ref => QualifiedRef(sheet, ref)).toSet

  /**
   * Extract all qualified cell references from TExpr.
   *
   * Similar to extractDependencies but returns QualifiedRef to track cross-sheet references.
   * Same-sheet references are qualified with the current sheet name. Range expansion is delegated
   * to `cellsFor` so workbook-level callers can bound ranges by the TARGET sheet's used range
   * (GH-346) — the qualified analog of `extractDependenciesBounded`; single refs stay exact.
   *
   * @param expr
   *   The expression to analyze
   * @param currentSheet
   *   The sheet containing the formula (used for same-sheet ref qualification)
   * @param cellsFor
   *   Expands a range on a named sheet to qualified cells (possibly bounded)
   * @param workbook
   *   GH-384: name table for defined-name resolution (None disables name edges)
   * @param visitingNames
   *   GH-384: UPPERCASED names on the current resolution path — the name→name cycle guard
   * @return
   *   Set of qualified cell references used in the expression
   */
  @nowarn("msg=Unreachable case")
  private def extractQualifiedDependencies[A](
    expr: TExpr[A],
    currentSheet: SheetName,
    cellsFor: (SheetName, CellRange) => Set[QualifiedRef] = unboundedQualifiedCells,
    workbook: Option[Workbook] = None,
    visitingNames: Set[String] = Set.empty
  ): Set[QualifiedRef] =
    def locCells(location: TExpr.RangeLocation): Set[QualifiedRef] =
      location match
        case TExpr.RangeLocation.Local(range) => cellsFor(currentSheet, range)
        case TExpr.RangeLocation.CrossSheet(sheet, range) => cellsFor(sheet, range)
        // GH-353: external-workbook ranges target cells OUTSIDE the workbook — no edges ever
        case TExpr.RangeLocation.External(_, _, _) => Set.empty
        // GH-394: a name in a range slot resolves like the equivalent name EXPRESSION — its
        // target's cells contribute edges (qualified to the DEFINING sheet) so recalc orders
        // name-gated dependents correctly; a sheet-qualified name looks up relative to its
        // qualifier. Unresolvable names contribute no edges (evaluation reports the error).
        case TExpr.RangeLocation.Name(name, scope) =>
          scope match
            case None => go(TExpr.NameRef(name))
            case Some(qualifier) => go(TExpr.SheetNameRef(qualifier, name))

    def go(e: TExpr[?]): Set[QualifiedRef] =
      e match
        // Same-sheet references - qualify with current sheet
        case TExpr.Ref(at, _, _) => Set(QualifiedRef(currentSheet, at))
        case TExpr.PolyRef(at, _) => Set(QualifiedRef(currentSheet, at))
        case TExpr.RangeRef(range) => cellsFor(currentSheet, range)

        // Cross-sheet references - use target sheet
        case TExpr.SheetRef(sheet, at, _, _) => Set(QualifiedRef(sheet, at))
        case TExpr.SheetPolyRef(sheet, at, _) => Set(QualifiedRef(sheet, at))
        case TExpr.SheetRange(sheet, range) => cellsFor(sheet, range)
        // GH-353: external-workbook refs target cells OUTSIDE the workbook — no edges ever
        case TExpr.ExternalRef(_, _, _, _) => Set.empty
        case TExpr.ExternalRange(_, _, _) => Set.empty
        case TExpr.Add(l, r) => go(l) ++ go(r)
        case TExpr.Sub(l, r) => go(l) ++ go(r)
        case TExpr.Mul(l, r) => go(l) ++ go(r)
        case TExpr.Div(l, r) => go(l) ++ go(r)
        case TExpr.Pow(l, r) => go(l) ++ go(r)
        case TExpr.Concat(l, r) => go(l) ++ go(r)
        case TExpr.Eq(l, r) => go(l) ++ go(r)
        case TExpr.Neq(l, r) => go(l) ++ go(r)
        case TExpr.Lt(l, r) => go(l) ++ go(r)
        case TExpr.Lte(l, r) => go(l) ++ go(r)
        case TExpr.Gt(l, r) => go(l) ++ go(r)
        case TExpr.Gte(l, r) => go(l) ++ go(r)
        // Unary operators
        case TExpr.ToInt(x) => go(x)
        case TExpr.UnaryPlus(x) => go(x)
        case TExpr.Percent(x) => go(x)

        // Reference functions
        case TExpr.Aggregate(_, location) => locCells(location)

        case call: TExpr.Call[?] =>
          val values = call.spec.argSpec.toValues(call.args)
          values.foldLeft(Set.empty[QualifiedRef]) { (acc, value) =>
            value match
              case ArgValue.Expr(expr) => acc ++ go(expr)
              case ArgValue.Range(range) => acc ++ locCells(range)
              case ArgValue.Cells(range) => acc ++ cellsFor(currentSheet, range)
          }

        // GH-193: LET — union of binding-value and body dependencies; BindingRef has none
        case TExpr.Let(bindings, body) =>
          bindings.foldLeft(go(body)) { case (acc, (_, value)) => acc ++ go(value) }
        case TExpr.BindingRef(_) => Set.empty
        case TExpr.CoercedBindingRef(_, _) => Set.empty

        // GH-384: resolve the defined name and recurse into its refersTo, qualifying its
        // same-sheet refs to the DEFINING sheet — =IF(case=2, ...) contributes an edge to the
        // cells 'case' targets, so Kahn orders a computed toggle before its name-gated
        // dependents and Tarjan sees cycles routed through names. Name→name chains carry a
        // visited guard; unresolvable/unparseable names contribute no edges (evaluation
        // reports the per-cell error).
        case TExpr.NameRef(name) =>
          val key = name.toUpperCase
          if visitingNames.contains(key) then Set.empty
          else
            (for
              wb <- workbook
              dn <- Evaluator.lookupDefinedName(wb, currentSheet, name)
              target <- FormulaParser.parse(dn.formula).toOption
            yield
              val definingSheet =
                Evaluator.definedNameScope(wb, dn).map(_.name).getOrElse(currentSheet)
              extractQualifiedDependencies(
                target,
                definingSheet,
                cellsFor,
                workbook,
                visitingNames + key
              )
            ).getOrElse(Set.empty)

        // GH-394: a sheet-qualified name resolves like NameRef, but the lookup runs as seen
        // from its QUALIFIER (sheet-scoped names on that sheet shadow workbook-scoped ones)
        case TExpr.SheetNameRef(qualifier, name) =>
          val key = name.toUpperCase
          if visitingNames.contains(key) then Set.empty
          else
            (for
              wb <- workbook
              dn <- Evaluator.lookupDefinedName(wb, qualifier, name)
              target <- FormulaParser.parse(dn.formula).toOption
            yield
              val definingSheet =
                Evaluator.definedNameScope(wb, dn).map(_.name).getOrElse(currentSheet)
              extractQualifiedDependencies(
                target,
                definingSheet,
                cellsFor,
                workbook,
                visitingNames + key
              )
            ).getOrElse(Set.empty)

        // GH-306: runtime coercion wrapper — transparent for analysis
        case TExpr.Coerced(inner, _) => go(inner)

        // Literals and nullary functions (no dependencies)
        case TExpr.Lit(_) => Set.empty
        case TExpr.DateToSerial(dateExpr) => go(dateExpr)
        case TExpr.DateTimeToSerial(dtExpr) => go(dtExpr)

    go(expr)

  /**
   * Detect circular references across sheets using Tarjan's SCC algorithm.
   *
   * Similar to detectCycles but works with QualifiedRef to detect cycles that span multiple sheets.
   * A cross-sheet cycle occurs when cells across different sheets form a circular dependency (e.g.,
   * Sheet1!A1 = Sheet2!B1, Sheet2!B1 = Sheet1!A1).
   *
   * @param graph
   *   The cross-sheet dependency graph from fromWorkbook
   * @return
   *   Left(CircularRef) if cycle detected, Right(()) if acyclic
   *
   * Example:
   * {{{
   * val graph = DependencyGraph.fromWorkbook(workbook)
   * DependencyGraph.detectCrossSheetCycles(graph) match
   *   case Left(err) => println(s"Circular reference: $err")
   *   case Right(_) => println("No cycles")
   * }}}
   */
  @SuppressWarnings(
    Array(
      "org.wartremover.warts.Var",
      "org.wartremover.warts.IterableOps",
      "org.wartremover.warts.Return",
      "org.wartremover.warts.IsInstanceOf",
      "org.wartremover.warts.AsInstanceOf"
    )
  )
  def detectCrossSheetCycles(
    graph: Map[QualifiedRef, Set[QualifiedRef]]
  ): Either[EvalError.CircularRef, Unit] =
    // Tarjan's SCC algorithm adapted for QualifiedRef
    var index = 0
    var stack = List.empty[QualifiedRef]
    var indices = Map.empty[QualifiedRef, Int]
    var lowLinks = Map.empty[QualifiedRef, Int]
    var onStack = Set.empty[QualifiedRef]

    def strongConnect(v: QualifiedRef): Option[List[ARef]] =
      indices = indices.updated(v, index)
      lowLinks = lowLinks.updated(v, index)
      index += 1
      stack = v :: stack
      onStack = onStack + v

      val successors = graph.getOrElse(v, Set.empty)
      val cycleFound = successors.foldLeft(Option.empty[List[ARef]]) { (acc, w) =>
        acc match
          case Some(cycle) => Some(cycle)
          case None =>
            if !indices.contains(w) then
              strongConnect(w) match
                case Some(cycle) => Some(cycle)
                case None =>
                  lowLinks = lowLinks.updated(v, math.min(lowLinks(v), lowLinks(w)))
                  None
            else if onStack.contains(w) then
              lowLinks = lowLinks.updated(v, math.min(lowLinks(v), indices(w)))
              // Found cycle - reconstruct from stack
              val cycleNodes = (stack.takeWhile(_ != w) :+ w).reverse
              // Convert to List[ARef] with sheet prefix for error message
              Some(cycleNodes.map(_.ref) :+ cycleNodes.head.ref)
            else None
      }

      cycleFound match
        case Some(cycle) => Some(cycle)
        case None =>
          if lowLinks(v) == indices(v) then
            val (scc, remaining) = stack.span(_ != v)
            stack = remaining.tail
            onStack = onStack -- (scc :+ v)

            if scc.nonEmpty then
              // Multiple nodes in SCC = cycle
              val cycleNodes = (scc :+ v).reverse
              Some(cycleNodes.map(_.ref) :+ cycleNodes.head.ref)
            else if graph.get(v).exists(_.contains(v)) then Some(List(v.ref, v.ref)) // Self-loop
            else None
          else None

    val allNodes = graph.keySet
    val cycleFound = allNodes.foldLeft(Option.empty[List[ARef]]) { (acc, node) =>
      acc match
        case Some(cycle) => Some(cycle)
        case None =>
          if !indices.contains(node) then strongConnect(node)
          else None
    }

    cycleFound match
      case Some(cycle) => scala.util.Left(EvalError.CircularRef(cycle))
      case None => scala.util.Right(())
