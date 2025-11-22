//> using scala 3.7.3
//> using dep com.tjclp::xl-core:0.1.0-SNAPSHOT
//> using dep com.tjclp::xl-evaluator:0.1.0-SNAPSHOT

/**
 * Display Formatting Test
 *
 * Manual test to verify implicit display formatting works correctly.
 */

import com.tjclp.xl.*
import com.tjclp.xl.conversions.given
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.styles.{CellStyle, numfmt}
import numfmt.NumFmt
import com.tjclp.xl.unsafe.*

// Import display functionality
import com.tjclp.xl.display.{*, given}
import com.tjclp.xl.display.syntax.*
import com.tjclp.xl.display.DisplayConversions.given  // Import the conversions
import com.tjclp.xl.display.ExcelInterpolator.*  // Import excel"..." interpolator

// Test 1: Core-only display (no evaluation)
println("=" * 70)
println("TEST 1: Core-only Display (formulas show as raw text)")
println("=" * 70)

val sheet1 = Sheet(name = SheetName.unsafe("Test1"))
  .put(ref"A1", 1000000)
  .put(ref"A2", 0.6)
  .put(ref"A3", fx"=A1+A2")
  .style(ref"A1", CellStyle.default.withNumFmt(NumFmt.Currency))
  .flatMap(_.style(ref"A2", CellStyle.default.withNumFmt(NumFmt.Percent)))
  .unsafe

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

import com.tjclp.xl.formula.display.EvaluatingFormulaDisplay

val sheet2 = Sheet(name = SheetName.unsafe("Test2"))
  .put(ref"A1", 100)
  .put(ref"A2", 200)
  .put(ref"B1", fx"=SUM(A1:A2)")
  .put(ref"B2", fx"=A1/A2")
  .style(ref"B2", CellStyle.default.withNumFmt(NumFmt.Percent))
  .unsafe

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

println(s"B1 display: ${sheet2.display(ref"B1")}")
println(s"B1 formula: ${sheet2.displayFormula(ref"B1")}")
println()

println("✓ All tests completed!")
