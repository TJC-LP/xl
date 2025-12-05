package com.tjclp.xl.formula

import com.tjclp.xl.addressing.ARef
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
    val formulaCells = sheet.cells.flatMap { case (ref, cell) =>
      cell.value match
        case CellValue.Formula(expression, _) => Some(ref -> expression)
        case _ => None
    }

    // Build forward edges (dependencies)
    val dependencies = formulaCells.map { case (ref, formulaStr) =>
      val deps = FormulaParser.parse(formulaStr) match
        case scala.util.Right(expr) => extractDependencies(expr)
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
   * Extract all cell references from TExpr.
   *
   * Recursively traverses the expression AST and collects all cell references, including:
   *   - Single cell references (Ref)
   *   - Range references (FoldRange) expanded to all cells in range
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

      // Range reference (expand to all cells)
      case TExpr.FoldRange(range, _, _, _) =>
        range.cells.toSet

      // Recursive cases (binary operators)
      case TExpr.Add(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.Sub(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.Mul(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.Div(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.Eq(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.Neq(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.Lt(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.Lte(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.Gt(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.Gte(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.And(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.Or(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.ToInt(expr) =>
        extractDependencies(expr) // Type conversion - extract from wrapped expr
      case TExpr.Concatenate(xs) => xs.flatMap(extractDependencies).toSet
      case TExpr.Left(text, n) => extractDependencies(text) ++ extractDependencies(n)
      case TExpr.Right(text, n) => extractDependencies(text) ++ extractDependencies(n)
      case TExpr.Date(y, m, d) =>
        extractDependencies(y) ++ extractDependencies(m) ++ extractDependencies(d)
      case TExpr.Year(date) => extractDependencies(date)
      case TExpr.Month(date) => extractDependencies(date)
      case TExpr.Day(date) => extractDependencies(date)

      // Financial functions
      case TExpr.Npv(rate, values) =>
        extractDependencies(rate) ++ values.cells.toSet
      case TExpr.Irr(values, guessOpt) =>
        values.cells.toSet ++ guessOpt.map(extractDependencies).getOrElse(Set.empty)
      case TExpr.VLookup(lookup, table, colIndex, rangeLookup) =>
        extractDependencies(lookup) ++
          table.cells.toSet ++
          extractDependencies(colIndex) ++
          extractDependencies(rangeLookup)

      // Conditional aggregation functions
      case TExpr.SumIf(range, criteria, sumRangeOpt) =>
        range.cells.toSet ++
          extractDependencies(criteria) ++
          sumRangeOpt.map(_.cells.toSet).getOrElse(Set.empty)
      case TExpr.CountIf(range, criteria) =>
        range.cells.toSet ++ extractDependencies(criteria)
      case TExpr.SumIfs(sumRange, conditions) =>
        sumRange.cells.toSet ++
          conditions.flatMap { case (range, criteria) =>
            range.cells.toSet ++ extractDependencies(criteria)
          }.toSet
      case TExpr.CountIfs(conditions) =>
        conditions.flatMap { case (range, criteria) =>
          range.cells.toSet ++ extractDependencies(criteria)
        }.toSet

      // Array and advanced lookup functions
      case TExpr.SumProduct(arrays) =>
        arrays.flatMap(_.cells.toSet).toSet

      case TExpr.XLookup(
            lookupValue,
            lookupArray,
            returnArray,
            ifNotFound,
            matchMode,
            searchMode
          ) =>
        extractDependencies(lookupValue) ++
          lookupArray.cells.toSet ++
          returnArray.cells.toSet ++
          ifNotFound.map(extractDependencies).getOrElse(Set.empty) ++
          extractDependencies(matchMode) ++
          extractDependencies(searchMode)

      // Ternary operator
      case TExpr.If(cond, thenBranch, elseBranch) =>
        extractDependencies(cond) ++ extractDependencies(thenBranch) ++ extractDependencies(
          elseBranch
        )

      // Unary operators
      case TExpr.Not(x) => extractDependencies(x)
      case TExpr.Len(x) => extractDependencies(x)
      case TExpr.Upper(x) => extractDependencies(x)
      case TExpr.Lower(x) => extractDependencies(x)

      // Range aggregate functions (direct enum cases)
      case TExpr.Min(range) => range.cells.toSet
      case TExpr.Max(range) => range.cells.toSet
      case TExpr.Average(range) => range.cells.toSet

      // Literals and nullary functions (no dependencies)
      case TExpr.Lit(_) => Set.empty
      case TExpr.Today() => Set.empty
      case TExpr.Now() => Set.empty

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
