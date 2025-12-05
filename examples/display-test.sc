//> using scala 3.7.3
//> using dep com.tjclp::xl-core:0.1.5-SNAPSHOT
//> using dep com.tjclp::xl-evaluator:0.1.5-SNAPSHOT
//> using repository ivy2Local

/**
 * Display Formatting Test
 *
 * Manual test to verify implicit display formatting works correctly.
 */

import com.tjclp.xl.{*, given}
import com.tjclp.xl.unsafe.*
// All display functionality now available from com.tjclp.xl.{*, given}

// Test 1: Core-only display (no evaluation)
println("=" * 70)
println("TEST 1: Core-only Display (formulas show as raw text)")
println("=" * 70)

val sheet1 = Sheet("Test1")
  .put(
    ref"A1" -> money"$$1000000",
    ref"A2" -> percent"60%",
    ref"A3" -> fx"=A1+A2"
  )
  // .style(ref"A1", CellStyle.default.withNumFmt(NumFmt.Currency))
  // .flatMap(_.style(ref"A2", CellStyle.default.withNumFmt(NumFmt.Percent)))

locally {
  given Sheet = sheet1

  // Test excel interpolator
  println(excel"A1 (currency): ${ref"A1"}")
  println(excel"A2 (percent): ${ref"A2"}")
  println(excel"A3 (formula): ${ref"A3"}")  // Should show "=A1+A2"
}
println()

// Test 2: With evaluator (formulas evaluate)
println("=" * 70)
println("TEST 2: With Evaluator (formulas evaluate automatically)")
println("=" * 70)

// EvaluatingFormulaDisplay is now available from com.tjclp.xl.{*, given}
val sheet2 = Sheet("Test2")
  .put(ref"A1", 100)
  .put(ref"A2", 200)
  .put(ref"B1", fx"=SUM(A1:A2)")
  .put(ref"B2", fx"=A1/A2")
  .style(ref"B2", CellStyle.default.withNumFmt(NumFmt.Percent))

locally {
  given Sheet = sheet2
  given FormulaDisplayStrategy = EvaluatingFormulaDisplay.evaluating  // Explicit strategy selection

  println(excel"A1: ${ref"A1"}")
  println(excel"A2: ${ref"A2"}")
  println(excel"B1 (formula sum): ${ref"B1"}")      // Should show "300" (evaluated!)
  println(excel"B2 (formula percent): ${ref"B2"}")  // Should show "50.0%" (evaluated + formatted!)
}
println()

// Test 3: Explicit display methods
println("=" * 70)
println("TEST 3: Explicit Display Methods")
println("=" * 70)

println(s"B1 display: ${sheet2.displayCell(ref"B1")}")
println(s"B1 formula: ${sheet2.displayFormula(ref"B1")}")
println()

println("✓ All tests completed!")
