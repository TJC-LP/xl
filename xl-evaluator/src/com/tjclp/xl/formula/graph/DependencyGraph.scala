package com.tjclp.xl.formula.graph

import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.functions.{FunctionSpecs, ArgValue}
import com.tjclp.xl.formula.parser.FormulaParser
import com.tjclp.xl.formula.eval.EvalError

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
      case TExpr.DateToSerial(e) => containsCellReferences(e)
      case TExpr.DateTimeToSerial(e) => containsCellReferences(e)

      // Literals and constants
      case TExpr.Lit(_) => false

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
      case TExpr.Aggregate(_, TExpr.RangeLocation.CrossSheet(_, _)) => false

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
      case TExpr.DateToSerial(e) => containsUnqualifiedCellReferences(e)
      case TExpr.DateTimeToSerial(e) => containsUnqualifiedCellReferences(e)

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
      case TExpr.Aggregate(_, location) => location.localCells

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
      case TExpr.Aggregate(_, location) => location.localCellsBounded(bounds)

      case call: TExpr.Call[?] =>
        depsFromArgValues(
          call.spec.argSpec.toValues(call.args),
          expr => extractDependenciesBounded(expr, bounds),
          loc => loc.localCellsBounded(bounds),
          range => boundedCells(range, bounds)
        )

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
    transitiveDependentsImpl(graph, refs, refs, Set.empty)

  @scala.annotation.tailrec
  private def transitiveDependentsImpl(
    graph: DependencyGraph,
    originalRefs: Set[ARef],
    frontier: Set[ARef],
    visited: Set[ARef]
  ): Set[ARef] =
    val toVisit = frontier -- visited
    if toVisit.isEmpty then visited -- originalRefs // Exclude starting refs
    else
      val directDeps = toVisit.flatMap(ref => graph.dependents.getOrElse(ref, Set.empty))
      transitiveDependentsImpl(graph, originalRefs, directDeps, visited ++ toVisit)

  /**
   * Detect circular references using Tarjan's strongly connected components algorithm.
   *
   * A circular reference occurs when a cell's formula depends (directly or transitively) on its own
   * value. For example: A1="=B1", B1="=C1", C1="=A1" forms a cycle.
   *
   * This uses Tarjan's SCC algorithm which runs in O(V + E) time with a single DFS traversal. The
   * algorithm maintains a stack and low-link values to detect strongly connected components
   * (cycles).
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
  @SuppressWarnings(
    Array(
      "org.wartremover.warts.Var",
      "org.wartremover.warts.IterableOps",
      "org.wartremover.warts.Return",
      "org.wartremover.warts.IsInstanceOf",
      "org.wartremover.warts.AsInstanceOf"
    )
  )
  def detectCycles(graph: DependencyGraph): Either[EvalError.CircularRef, Unit] =
    // Tarjan's SCC algorithm: Intentional imperative implementation
    // Rationale: Classic algorithm uses mutable state for O(V+E) performance.
    // Functional version sacrifices clarity without benefit. Compile-time only.
    var index = 0
    var stack = List.empty[ARef]
    var indices = Map.empty[ARef, Int]
    var lowLinks = Map.empty[ARef, Int]
    var onStack = Set.empty[ARef]

    def strongConnect(v: ARef): Option[List[ARef]] =
      // Set the depth index for v
      indices = indices.updated(v, index)
      lowLinks = lowLinks.updated(v, index)
      index += 1
      stack = v :: stack
      onStack = onStack + v

      // Consider successors of v (cells that v depends on)
      val successors = graph.dependencies.getOrElse(v, Set.empty)
      val cycleFound = successors.foldLeft(Option.empty[List[ARef]]) { (acc, w) =>
        acc match
          case Some(cycle) => Some(cycle) // Already found cycle, propagate
          case None =>
            if !indices.contains(w) then
              // Successor w has not yet been visited; recurse on it
              strongConnect(w) match
                case Some(cycle) => Some(cycle)
                case None =>
                  lowLinks = lowLinks.updated(v, math.min(lowLinks(v), lowLinks(w)))
                  None
            else if onStack.contains(w) then
              // Successor w is on stack and hence in the current SCC
              lowLinks = lowLinks.updated(v, math.min(lowLinks(v), indices(w)))
              // Found a cycle! Reconstruct it from stack (w to top of stack forms the cycle)
              val cycleNodes = (stack.takeWhile(_ != w) :+ w).reverse
              Some(cycleNodes :+ cycleNodes.head) // Add first node again to show cycle closes
            else None // w is not on stack, already processed
      }

      cycleFound match
        case Some(cycle) => Some(cycle)
        case None =>
          // If v is a root node, pop the stack and check for SCC
          if lowLinks(v) == indices(v) then
            // Pop nodes from stack until v
            val (scc, remaining) = stack.span(_ != v)
            stack = remaining.tail // Remove v from stack
            onStack = onStack -- (scc :+ v)

            // Check if SCC has more than one node (cycle)
            if scc.nonEmpty then
              // Multiple nodes in SCC means cycle
              val cycleNodes = (scc :+ v).reverse
              Some(cycleNodes :+ cycleNodes.head) // Add first node again to show cycle closes
            else
              // Single node - only a cycle if it has self-loop
              if graph.dependencies.get(v).exists(_.contains(v)) then
                Some(List(v, v)) // Self-loop: v -> v
              else None
          else None

    // Run Tarjan's algorithm on all unvisited nodes
    val allNodes = graph.dependencies.keySet
    val cycleFound = allNodes.foldLeft(Option.empty[List[ARef]]) { (acc, node) =>
      acc match
        case Some(cycle) => Some(cycle) // Already found cycle
        case None =>
          if !indices.contains(node) then strongConnect(node)
          else None
    }

    cycleFound match
      case Some(cycle) => scala.util.Left(EvalError.CircularRef(cycle))
      case None => scala.util.Right(())

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
    import scala.util.boundary, boundary.break

    boundary:
      // All formula cells (only process formulas, not constants)
      val allNodes = graph.dependencies.keySet

      // If no formula cells, early exit
      if allNodes.isEmpty then break(scala.util.Right(List.empty[ARef]))

      // Calculate in-degree for each node (number of formula cells it depends on)
      // Only count dependencies on other formula cells, not constants
      val inDegree = allNodes.map { node =>
        val deps = graph.dependencies.getOrElse(node, Set.empty)
        val formulaDeps = deps.filter(allNodes.contains)
        node -> formulaDeps.size
      }.toMap

      // Start with nodes that have in-degree 0 (no dependencies)
      val initialQueue = allNodes.filter(node => inDegree(node) == 0).toList

      @tailrec
      def process(
        queue: List[ARef],
        processedInDegree: Map[ARef, Int],
        acc: List[ARef]
      ): (List[ARef], Map[ARef, Int]) =
        queue match
          case Nil => (acc, processedInDegree)
          case node :: rest =>
            val deps = graph.dependents.getOrElse(node, Set.empty).filter(allNodes.contains)
            val (nextInDegree, newlyZero) = deps.foldLeft((processedInDegree, List.empty[ARef])) {
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
      else
        // Cycle detected: find one cycle for error reporting
        val remainingNodes = allNodes -- result.toSet
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
    override def toString: String = s"${sheet.value}!${ref.toA1}"

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
              case scala.util.Right(expr) => extractQualifiedDependencies(expr, sheet.name)
              case scala.util.Left(_) => Set.empty[QualifiedRef]
            Some(QualifiedRef(sheet.name, cellRef) -> deps)
          case _ => None
      }
    }.toMap

  /**
   * Convert a RangeLocation to qualified cell references.
   *
   * For Local ranges, uses the current sheet. For CrossSheet ranges, uses the specified sheet.
   */
  private def locationToQualifiedRefs(
    location: TExpr.RangeLocation,
    currentSheet: SheetName
  ): Set[QualifiedRef] =
    location match
      case TExpr.RangeLocation.Local(range) =>
        range.cells.map(ref => QualifiedRef(currentSheet, ref)).toSet
      case TExpr.RangeLocation.CrossSheet(sheet, range) =>
        range.cells.map(ref => QualifiedRef(sheet, ref)).toSet

  /**
   * Extract all qualified cell references from TExpr.
   *
   * Similar to extractDependencies but returns QualifiedRef to track cross-sheet references.
   * Same-sheet references are qualified with the current sheet name.
   *
   * @param expr
   *   The expression to analyze
   * @param currentSheet
   *   The sheet containing the formula (used for same-sheet ref qualification)
   * @return
   *   Set of qualified cell references used in the expression
   */
  @nowarn("msg=Unreachable case")
  private def extractQualifiedDependencies[A](
    expr: TExpr[A],
    currentSheet: SheetName
  ): Set[QualifiedRef] =
    expr match
      // Same-sheet references - qualify with current sheet
      case TExpr.Ref(at, _, _) => Set(QualifiedRef(currentSheet, at))
      case TExpr.PolyRef(at, _) => Set(QualifiedRef(currentSheet, at))
      case TExpr.RangeRef(range) =>
        range.cells.map(ref => QualifiedRef(currentSheet, ref)).toSet

      // Cross-sheet references - use target sheet
      case TExpr.SheetRef(sheet, at, _, _) => Set(QualifiedRef(sheet, at))
      case TExpr.SheetPolyRef(sheet, at, _) => Set(QualifiedRef(sheet, at))
      case TExpr.SheetRange(sheet, range) =>
        range.cells.map(ref => QualifiedRef(sheet, ref)).toSet
      case TExpr.Add(l, r) =>
        extractQualifiedDependencies(l, currentSheet) ++ extractQualifiedDependencies(
          r,
          currentSheet
        )
      case TExpr.Sub(l, r) =>
        extractQualifiedDependencies(l, currentSheet) ++ extractQualifiedDependencies(
          r,
          currentSheet
        )
      case TExpr.Mul(l, r) =>
        extractQualifiedDependencies(l, currentSheet) ++ extractQualifiedDependencies(
          r,
          currentSheet
        )
      case TExpr.Div(l, r) =>
        extractQualifiedDependencies(l, currentSheet) ++ extractQualifiedDependencies(
          r,
          currentSheet
        )
      case TExpr.Pow(l, r) =>
        extractQualifiedDependencies(l, currentSheet) ++ extractQualifiedDependencies(
          r,
          currentSheet
        )
      case TExpr.Concat(l, r) =>
        extractQualifiedDependencies(l, currentSheet) ++ extractQualifiedDependencies(
          r,
          currentSheet
        )
      case TExpr.Eq(l, r) =>
        extractQualifiedDependencies(l, currentSheet) ++ extractQualifiedDependencies(
          r,
          currentSheet
        )
      case TExpr.Neq(l, r) =>
        extractQualifiedDependencies(l, currentSheet) ++ extractQualifiedDependencies(
          r,
          currentSheet
        )
      case TExpr.Lt(l, r) =>
        extractQualifiedDependencies(l, currentSheet) ++ extractQualifiedDependencies(
          r,
          currentSheet
        )
      case TExpr.Lte(l, r) =>
        extractQualifiedDependencies(l, currentSheet) ++ extractQualifiedDependencies(
          r,
          currentSheet
        )
      case TExpr.Gt(l, r) =>
        extractQualifiedDependencies(l, currentSheet) ++ extractQualifiedDependencies(
          r,
          currentSheet
        )
      case TExpr.Gte(l, r) =>
        extractQualifiedDependencies(l, currentSheet) ++ extractQualifiedDependencies(
          r,
          currentSheet
        )
      // Unary operators
      case TExpr.ToInt(x) => extractQualifiedDependencies(x, currentSheet)

      // Reference functions
      case TExpr.Aggregate(_, location) => locationToQualifiedRefs(location, currentSheet)

      case call: TExpr.Call[?] =>
        val values = call.spec.argSpec.toValues(call.args)
        values.foldLeft(Set.empty[QualifiedRef]) { (acc, value) =>
          value match
            case ArgValue.Expr(expr) => acc ++ extractQualifiedDependencies(expr, currentSheet)
            case ArgValue.Range(range) => acc ++ locationToQualifiedRefs(range, currentSheet)
            case ArgValue.Cells(range) =>
              acc ++ range.cells.map(ref => QualifiedRef(currentSheet, ref)).toSet
        }

      // Literals and nullary functions (no dependencies)
      case TExpr.Lit(_) => Set.empty
      case TExpr.DateToSerial(dateExpr) => extractQualifiedDependencies(dateExpr, currentSheet)
      case TExpr.DateTimeToSerial(dtExpr) => extractQualifiedDependencies(dtExpr, currentSheet)

      // Financial functions
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
