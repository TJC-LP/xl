#!/usr/bin/env -S scala-cli shebang
//> using file project.scala


/**
 * XL Formula System - Dependency Analysis Example
 *
 * This example demonstrates advanced dependency graph features:
 * - Build the workbook dependency graph (QualifiedGraph.of)
 * - Detect circular references (strongly connected components)
 * - List the evaluation order (dependency-first components)
 * - Query precedents and dependents for impact analysis
 * - Visualize formula relationships
 *
 * Perfect for meta-programming, formula auditing, and impact analysis.
 *
 * Run with: scala-cli examples/dependency_analysis.sc
 */

import com.tjclp.xl.{*, given}
// Note: No unsafe import needed - all refs are compile-time validated literals

// ============================================================================
// Scenario 1: Complex Dependency Chain
// ============================================================================

println("=" * 80)
println("SCENARIO 1: Complex Dependency Chain")
println("=" * 80)
println()

// Build a sheet with complex dependencies
val complexSheet = Sheet("Complex")
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

// Build the dependency graph of the workbook (cross-sheet, bounded: a range expands only to its
// occupied cells)
val book1 = Workbook(complexSheet)
val graph1 = QualifiedGraph.of(book1)
def at(sheet: Sheet, a1: ARef): QualifiedRef = QualifiedRef(sheet.name, a1)
def cells(refs: Iterable[QualifiedRef]): String =
  QualifiedGraph.sorted(refs).map(_.ref.toA1).mkString(", ")

println("Dependency Analysis:")
println(s"  Total formula cells: ${graph1.formulas.size}")
println(s"  Total dependency edges: ${graph1.dependencies.values.map(_.size).sum}")
println()

// Evaluation order: the strongly connected components, dependency-first
val components = graph1.sccs
if components.exists(_.cyclic) then
  println(s"✗ Cycle detected: ${cells(components.filter(_.cyclic).flatMap(_.members))}")
else
  println(s"✓ Evaluation order (${components.size} formulas):")
  println(s"  ${components.flatMap(_.members).map(_.ref.toA1).mkString(" → ")}")

println()

// Impact analysis: What happens if A1 changes?
val a1Impact = graph1.dependentsOf(at(complexSheet, ref"A1"))
println(s"Impact of changing A1:")
println(s"  Direct dependents: ${cells(a1Impact)}")

// All transitive dependents, layer by layer (depth 0 = unbounded)
val impactLayers = graph1.dependents(at(complexSheet, ref"A1"), 0)
impactLayers.zipWithIndex.foreach { (layer, i) =>
  println(s"  ${i + 1} hop(s) away: ${cells(layer)}")
}
val transitiveImpact: Set[ARef] = impactLayers.flatten.map(_.ref).toSet
println(s"  Transitive dependents: ${cells(impactLayers.flatten)}")
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
val cycle2 = Sheet("Cycle2")
  .put(ref"A1", fx"=B1+10")
  .put(ref"B1", fx"=A1*2")

val graph2 = QualifiedGraph.of(Workbook(cycle2))
graph2.sccs.filter(_.cyclic) match
  case Vector() =>
    println("  ✗ No cycle detected (unexpected)")
  case cycles =>
    println(s"  ✓ Circular reference detected!")
    cycles.foreach(cycle => println(s"    Cycle members: ${cells(cycle.members)}"))

println()

// Test Case 2: Longer cycle (A1 → B1 → C1 → D1 → A1)
println("Test Case 2: 4-node cycle (A1 → B1 → C1 → D1 → A1)")
val cycle4 = Sheet("Cycle4")
  .put(ref"A1", fx"=B1+1")
  .put(ref"B1", fx"=C1+1")
  .put(ref"C1", fx"=D1+1")
  .put(ref"D1", fx"=A1+1")

val graph3 = QualifiedGraph.of(Workbook(cycle4))
graph3.sccs.filter(_.cyclic) match
  case Vector() =>
    println("  ✗ No cycle detected (unexpected)")
  case cycles =>
    println(s"  ✓ Circular reference detected!")
    cycles.foreach(cycle => println(s"    Cycle members: ${cells(cycle.members)}"))
    println(s"    Cycle length: ${cycles.map(_.members.size).sum} nodes")

println()

// Test Case 3: Self-loop (A1 → A1)
println("Test Case 3: Self-loop (A1 → A1)")
val selfLoop = Sheet("SelfLoop")
  .put(ref"A1", fx"=A1+1")

val graph4 = QualifiedGraph.of(Workbook(selfLoop))
graph4.sccs.filter(_.cyclic) match
  case Vector() =>
    println("  ✗ No cycle detected (unexpected)")
  case cycles =>
    println(s"  ✓ Self-reference detected!")
    cycles.foreach(cycle => println(s"    Cycle: ${cells(cycle.members)}"))

println()

// Test Case 4: Cycle through range (A1 → SUM(B1:B10) where B5 → A1)
println("Test Case 4: Cycle through range (A1 → SUM(B1:B10) where B5 → A1)")
val rangeCycle = Sheet("RangeCycle")
  .put(ref"A1", fx"=SUM(B1:B10)")
  .put(ref"B1", 10)
  .put(ref"B5", fx"=A1*2")  // Creates cycle!
  .put(ref"B10", 20)

val graph5 = QualifiedGraph.of(Workbook(rangeCycle))
graph5.sccs.filter(_.cyclic) match
  case Vector() =>
    println("  ✗ No cycle detected (unexpected)")
  case cycles =>
    println(s"  ✓ Cycle through range detected!")
    cycles.foreach(cycle => println(s"    Cycle members: ${cells(cycle.members)}"))

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
println(s"  D1 precedents: ${cells(graph1.precedentsOf(at(complexSheet, ref"D1")))}")
println(s"  C1 precedents: ${cells(graph1.precedentsOf(at(complexSheet, ref"C1")))}")
println(s"  B1 precedents: ${cells(graph1.precedentsOf(at(complexSheet, ref"B1")))}")
println()

println("Analyzing dependents (cells that depend on this cell):")
println(s"  B1 dependents: ${cells(graph1.dependentsOf(at(complexSheet, ref"B1")))}")
println(s"  C1 dependents: ${cells(graph1.dependentsOf(at(complexSheet, ref"C1")))}")
println(s"  D1 dependents: ${cells(graph1.dependentsOf(at(complexSheet, ref"D1")))}")
println()

// Find cells with most (transitive) dependents (potential bottlenecks)
val dependentCounts = complexSheet.cells.refs.toSeq.map { ref =>
  ref.toA1 -> graph1.dependents(at(complexSheet, ref), 0).flatten.size
}
val mostImpactful = dependentCounts.sortBy((a1, count) => (-count, a1)).take(3)
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

// Recalculate both versions (cycles would be reported per cell in `.errors`, never thrown)
val originalResults = Workbook(complexSheet).recalculate().evaluated(complexSheet.name)
val modifiedResults = Workbook(modifiedSheet).recalculate().evaluated(modifiedSheet.name)

// Show changes
println("Impact of changing A1 from 100 to 150:")
transitiveImpact.toSeq.sortBy(r => (r.row.index0, r.col.index0)).foreach { ref =>
  (originalResults.get(ref), modifiedResults.get(ref)) match
    case (Some(CellValue.Number(o)), Some(CellValue.Number(n))) =>
      val delta = n - o
      println(f"  ${ref.toA1}: ${o.toDouble}%.0f → ${n.toDouble}%.0f (Δ ${delta.toDouble}%+.0f)")
    case _ => ()
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
val allFunctions = FunctionRegistry.allNames
println(s"Available functions (${allFunctions.size} total):")
println(s"  ${allFunctions.mkString(", ")}")
println()

// Function lookup
println("Function introspection:")
FunctionRegistry.lookup("SUM") match
  case Some(spec) =>
    println(s"  SUM: arity = ${spec.arity}")
  case None =>
    println("  SUM: not found")

FunctionRegistry.lookup("AVERAGE") match
  case Some(spec) =>
    println(s"  AVERAGE: arity = ${spec.arity}")
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
✓ Build the cross-sheet dependency graph (QualifiedGraph.of)
✓ Detect circular references (strongly connected components, Tarjan's SCC)
✓ List the dependency-first evaluation order (the SCC condensation)
✓ Query precedents (cells this cell depends on), directly or in layers
✓ Query dependents (cells that depend on this cell), directly or in layers
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
- Tarjan's SCC: O(V+E), iterative (no stack overflow on deep chains)
- Ranges expand only to their occupied cells (SUM(A:A) never enumerates 1M cells)
- An empty cell inside a summed range still reports the sum as its dependent

🎯 Perfect for: DevOps, testing, meta-programming, formula validation
""")
