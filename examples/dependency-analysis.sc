//> using scala 3.7.3
//> using dep com.tjclp::xl-core:0.1.0-SNAPSHOT
//> using dep com.tjclp::xl-evaluator:0.1.0-SNAPSHOT

/**
 * XL Formula System - Dependency Analysis Example
 *
 * This example demonstrates advanced dependency graph features:
 * - Build dependency graphs from formula cells
 * - Detect circular references with precise cycle paths
 * - Compute topological evaluation order
 * - Query precedents and dependents for impact analysis
 * - Visualize formula relationships
 *
 * Perfect for meta-programming, formula auditing, and impact analysis.
 *
 * Run with: scala-cli examples/dependency-analysis.sc
 */

import com.tjclp.xl.*
import com.tjclp.xl.conversions.given  // Enables put(ref, primitiveValue) syntax
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.*
import com.tjclp.xl.formula.SheetEvaluator.* // Extension methods
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.addressing.SheetName

// ============================================================================
// Scenario 1: Complex Dependency Chain
// ============================================================================

println("=" * 80)
println("SCENARIO 1: Complex Dependency Chain")
println("=" * 80)
println()

// Build a sheet with complex dependencies
val complexSheet = Sheet(name = SheetName.unsafe("Complex"))
  // Constants
  .put(ref"A1", 100)
  .put(ref"A2", 50)

  // First level (depend on constants)
  .put(ref"B1", fx"=A1*2")        // 200
  .put(ref"B2", fx"=A2+10")       // 60

  // Second level (depend on first level)
  .put(ref"C1", fx"=B1+B2")       // 260
  .put(ref"C2", fx"=B1-A1")       // 100

  // Third level (depend on second level)
  .put(ref"D1", fx"=C1*C2")       // 26000

  // Range aggregate (depends on range)
  .put(ref"E1", fx"=SUM(A1:D1)")  // Depends on A1, B1, C1, D1

println("Formula structure:")
println("  Constants: A1=100, A2=50")
println("  Level 1: B1=A1*2, B2=A2+10")
println("  Level 2: C1=B1+B2, C2=B1-A1")
println("  Level 3: D1=C1*C2")
println("  Aggregate: E1=SUM(A1:D1)")
println()

// Build dependency graph
val graph1 = DependencyGraph.fromSheet(complexSheet)

println("Dependency Analysis:")
println(s"  Total formula cells: ${graph1.dependencies.size}")
println(s"  Total dependency edges: ${graph1.dependencies.values.map(_.size).sum}")
println()

// Topological sort
DependencyGraph.topologicalSort(graph1) match
  case Right(order) =>
    println(s"✓ Evaluation order (${order.size} formulas):")
    println(s"  ${order.map(_.toA1).mkString(" → ")}")
  case Left(error) =>
    println(s"✗ Cycle detected: ${error.cycle.map(_.toA1).mkString(" → ")}")

println()

// Impact analysis: What happens if A1 changes?
val a1Impact = DependencyGraph.dependents(graph1, ref"A1")
println(s"Impact of changing A1:")
println(s"  Direct dependents: ${a1Impact.map(_.toA1).mkString(", ")}")

// Find all transitive dependents using iterative BFS (avoids scala-cli opaque type issues)
val transitiveImpact: Set[String] = {
  var visited = Set.empty[String]
  var queue = a1Impact.map(_.toA1).toList
  while queue.nonEmpty do
    val current = queue.head
    queue = queue.tail
    if !visited.contains(current) then
      visited = visited + current
      // Get next level of dependents
      ARef.parse(current).toOption.foreach { ref =>
        val nextLevel = DependencyGraph.dependents(graph1, ref)
        queue = queue ++ nextLevel.map(_.toA1).filterNot(visited.contains)
      }
  visited
}
println(s"  Transitive dependents: ${transitiveImpact.mkString(", ")}")
println(s"  Total cells affected: ${transitiveImpact.size}")

println()

// ============================================================================
// Scenario 2: Circular Reference Detection
// ============================================================================

println("=" * 80)
println("SCENARIO 2: Circular Reference Detection")
println("=" * 80)
println()

// Test Case 1: Simple 2-cycle
println("Test Case 1: Simple 2-cycle (A1 → B1 → A1)")
val cycle2 = Sheet(name = SheetName.unsafe("Cycle2"))
  .put(ref"A1", fx"=B1+10")
  .put(ref"B1", fx"=A1*2")

val graph2 = DependencyGraph.fromSheet(cycle2)
DependencyGraph.detectCycles(graph2) match
  case Right(_) =>
    println("  ✗ No cycle detected (unexpected)")
  case Left(error) =>
    println(s"  ✓ Circular reference detected!")
    println(s"    Cycle path: ${error.cycle.map(_.toA1).mkString(" → ")}")

println()

// Test Case 2: Longer cycle (A1 → B1 → C1 → D1 → A1)
println("Test Case 2: 4-node cycle (A1 → B1 → C1 → D1 → A1)")
val cycle4 = Sheet(name = SheetName.unsafe("Cycle4"))
  .put(ref"A1", fx"=B1+1")
  .put(ref"B1", fx"=C1+1")
  .put(ref"C1", fx"=D1+1")
  .put(ref"D1", fx"=A1+1")

val graph3 = DependencyGraph.fromSheet(cycle4)
DependencyGraph.detectCycles(graph3) match
  case Right(_) =>
    println("  ✗ No cycle detected (unexpected)")
  case Left(error) =>
    println(s"  ✓ Circular reference detected!")
    println(s"    Cycle path: ${error.cycle.map(_.toA1).mkString(" → ")}")
    println(s"    Cycle length: ${error.cycle.size - 1} nodes")

println()

// Test Case 3: Self-loop (A1 → A1)
println("Test Case 3: Self-loop (A1 → A1)")
val selfLoop = Sheet(name = SheetName.unsafe("SelfLoop"))
  .put(ref"A1", fx"=A1+1")

val graph4 = DependencyGraph.fromSheet(selfLoop)
DependencyGraph.detectCycles(graph4) match
  case Right(_) =>
    println("  ✗ No cycle detected (unexpected)")
  case Left(error) =>
    println(s"  ✓ Self-reference detected!")
    println(s"    Cycle: ${error.cycle.map(_.toA1).mkString(" → ")}")

println()

// Test Case 4: Cycle through range (A1 → SUM(B1:B10) where B5 → A1)
println("Test Case 4: Cycle through range (A1 → SUM(B1:B10) where B5 → A1)")
val rangeCycle = Sheet(name = SheetName.unsafe("RangeCycle"))
  .put(ref"A1", fx"=SUM(B1:B10)")
  .put(ref"B1", 10)
  .put(ref"B5", fx"=A1*2")  // Creates cycle!
  .put(ref"B10", 20)

val graph5 = DependencyGraph.fromSheet(rangeCycle)
DependencyGraph.detectCycles(graph5) match
  case Right(_) =>
    println("  ✗ No cycle detected (unexpected)")
  case Left(error) =>
    println(s"  ✓ Cycle through range detected!")
    println(s"    Cycle path: ${error.cycle.map(_.toA1).mkString(" → ")}")

println()

// ============================================================================
// Scenario 3: Precedent/Dependent Queries
// ============================================================================

println("=" * 80)
println("SCENARIO 3: Precedent/Dependent Queries (Impact Analysis)")
println("=" * 80)
println()

// Use the complex sheet from Scenario 1
println("Analyzing precedents (cells this cell depends on):")
println(s"  D1 precedents: ${DependencyGraph.precedents(graph1, ref"D1").map(_.toA1)}")
println(s"  C1 precedents: ${DependencyGraph.precedents(graph1, ref"C1").map(_.toA1)}")
println(s"  B1 precedents: ${DependencyGraph.precedents(graph1, ref"B1").map(_.toA1)}")
println()

println("Analyzing dependents (cells that depend on this cell):")
println(s"  B1 dependents: ${DependencyGraph.dependents(graph1, ref"B1").map(_.toA1)}")
println(s"  C1 dependents: ${DependencyGraph.dependents(graph1, ref"C1").map(_.toA1)}")
println(s"  D1 dependents: ${DependencyGraph.dependents(graph1, ref"D1").map(_.toA1)}")
println()

// Find cells with most dependents (potential bottlenecks)
val dependentCounts = graph1.dependents.map { case (ref, deps) => ref -> deps.size }
val mostImpactful = dependentCounts.toSeq.sortBy(-_._2).take(3)
println("Most impactful cells (most dependents):")
mostImpactful.foreach { case (ref, count) =>
  println(s"  $ref: $count dependents")
}

println()

// ============================================================================
// Scenario 4: Real-World Use Case - What-If Analysis
// ============================================================================

println("=" * 80)
println("SCENARIO 4: What-If Analysis (Change Propagation)")
println("=" * 80)
println()

// Simulate changing A1 from 100 to 150
println("Original A1 value: 100")
println("What if A1 = 150?")
println()

// Create modified sheet
val modifiedSheet = complexSheet.put(ref"A1", 150)

// Evaluate both versions
val originalResults = complexSheet.evaluateWithDependencyCheck().toOption.get
val modifiedResults = modifiedSheet.evaluateWithDependencyCheck().toOption.get

// Show changes
println("Impact of changing A1 from 100 to 150:")
transitiveImpact.toSeq.sorted.foreach { refStr =>
  // Parse string back to ARef for map lookup
  ARef.parse(refStr).toOption.foreach { ref =>
    val oldVal = originalResults.get(ref)
    val newVal = modifiedResults.get(ref)
    (oldVal, newVal) match
      case (Some(CellValue.Number(o)), Some(CellValue.Number(n))) =>
        val delta = n - o
        println(f"  $refStr: ${o.toDouble}%.0f → ${n.toDouble}%.0f (Δ ${delta.toDouble}%+.0f)")
      case _ => ()
  }
}

println()

// ============================================================================
// Scenario 5: Formula Extraction & Introspection
// ============================================================================

println("=" * 80)
println("SCENARIO 5: Formula Extraction & Introspection")
println("=" * 80)
println()

// Extract all dependencies from a complex formula
val complexExpr = FormulaParser.parse("=IF(SUM(A1:A10)>100, MAX(B1:B10), MIN(C1:C10))").toOption.get
val deps = DependencyGraph.extractDependencies(complexExpr)

println("Formula: =IF(SUM(A1:A10)>100, MAX(B1:B10), MIN(C1:C10))")
println(s"  Total cell dependencies: ${deps.size}")
println(s"  Cells referenced: ${deps.toSeq.sortBy(_.toString).take(10).map(_.toA1).mkString(", ")}${if deps.size > 10 then ", ..." else ""}")
println()

// Show all available functions
println(s"Available functions (${FunctionParser.allFunctions.size} total):")
println(s"  ${FunctionParser.allFunctions.sorted.mkString(", ")}")
println()

// Function lookup
println("Function introspection:")
FunctionParser.lookup("SUM") match
  case Some(parser) =>
    println(s"  SUM: arity = ${parser.arity}")
  case None =>
    println("  SUM: not found")

FunctionParser.lookup("AVERAGE") match
  case Some(parser) =>
    println(s"  AVERAGE: arity = ${parser.arity}")
  case None =>
    println("  AVERAGE: not found")

println()

// ============================================================================
// Summary
// ============================================================================

println("=" * 80)
println("SUMMARY: Dependency Analysis Capabilities")
println("=" * 80)
println("""
✓ Build dependency graphs from formula cells (O(V+E))
✓ Detect circular references with precise cycle paths (Tarjan's SCC)
✓ Compute topological evaluation order (Kahn's algorithm)
✓ Query precedents (cells this cell depends on) - O(1) lookups
✓ Query dependents (cells that depend on this cell) - O(1) lookups
✓ Extract dependencies from TExpr AST programmatically
✓ Perform impact analysis (what-if scenarios)
✓ Find bottleneck cells (most dependents)
✓ Introspect function registry

Use Cases:
- Formula auditing (find circular references, validate models)
- Impact analysis (what cells are affected by changes?)
- Performance optimization (identify bottleneck formulas)
- Documentation generation (auto-generate dependency diagrams)
- Testing (verify formula relationships match expectations)
- Incremental calculation (only re-evaluate affected cells)

Technical Details:
- Tarjan's SCC: O(V+E) with early exit on first cycle
- Kahn's algorithm: O(V+E) topological sort
- Adjacency lists: O(1) precedent/dependent queries
- Handles 10k formula cells in <10ms

🎯 Perfect for: DevOps, testing, meta-programming, formula validation
""")
