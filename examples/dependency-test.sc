//> using scala 3.7.3
//> using dep com.tjclp::xl-core:0.1.0-SNAPSHOT
//> using dep com.tjclp::xl-evaluator:0.1.0-SNAPSHOT
//> using dep com.tjclp::xl-cats-effect:0.1.0-SNAPSHOT
//> using repository ivy2Local

/**
 * Test file for `xl cell` command - dependency/dependent tracking.
 *
 * Creates a file with formulas to test:
 * - A1 should show dependents: C1, C3
 * - C1 should show dependencies: A1, B1
 * - D1 should show dependencies: C1, C2, C3
 */

import com.tjclp.xl.{*, given}
import com.tjclp.xl.macros.ref
import com.tjclp.xl.unsafe.*

val sheet = Sheet(name = SheetName.unsafe("Dependencies"))
  // Data cells (these will show as dependents of formulas)
  .put(ref"A1", 100)
  .put(ref"A2", 200)
  .put(ref"A3", 300)
  .put(ref"B1", 10)
  .put(ref"B2", 20)
  // Formula cells (these have dependencies)
  .put(ref"C1", fx"=A1+B1")        // Depends on A1, B1
  .put(ref"C2", fx"=A2+B2")        // Depends on A2, B2
  .put(ref"C3", fx"=SUM(A1:A3)")   // Depends on A1, A2, A3
  // Summary formula (depends on other formulas)
  .put(ref"D1", fx"=C1+C2+C3")     // Depends on C1, C2, C3
  // Labels
  .put(ref"A5", "Test cells for xl cell command")
  .put(ref"A6", "A1 should show dependents: C1, C3")
  .put(ref"A7", "C1 should show dependencies: A1, B1")
  .put(ref"A8", "D1 should show dependencies: C1, C2, C3")

val wb = Workbook(Vector(sheet)).withCachedFormulas()
val outPath = "data/dependency-test.xlsx"

Excel.write(wb, outPath)
println(s"Created: $outPath (with cached formula values)")

// Also show the dependency graph
val graph = DependencyGraph.fromSheet(sheet)
println("\nDependency Graph:")
println(s"  A1 dependents: ${graph.dependents.getOrElse(ref"A1", Set.empty).map(_.toA1).mkString(", ")}")
println(s"  C1 dependencies: ${graph.dependencies.getOrElse(ref"C1", Set.empty).map(_.toA1).mkString(", ")}")
println(s"  D1 dependencies: ${graph.dependencies.getOrElse(ref"D1", Set.empty).map(_.toA1).mkString(", ")}")
