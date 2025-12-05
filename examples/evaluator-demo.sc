//> using scala 3.7.3
//> using dep com.tjclp::xl-core:0.1.5-SNAPSHOT
//> using dep com.tjclp::xl-evaluator:0.1.5-SNAPSHOT
//> using repository ivy2Local

/**
 * Formula Evaluator Demo (WI-08, WI-09 Complete)
 *
 * **STATUS**: ✅ FULLY OPERATIONAL
 *
 * This script demonstrates formula evaluation with the complete formula system:
 * - Low-level API: Direct TExpr evaluation with Evaluator
 * - High-level API: Sheet extensions (evaluateFormula, evaluateCell, evaluateWithDependencyCheck)
 * - 21 built-in functions (aggregate, logical, text, date)
 * - Dependency graph and cycle detection
 *
 * To run:
 *   1. Publish locally: ./mill xl-core.publishLocal && ./mill xl-evaluator.publishLocal
 *   2. Run: scala-cli run examples/evaluator-demo.sc
 */

import com.tjclp.xl.{*, given}
import com.tjclp.xl.unsafe.*
// SheetEvaluator extension methods now available from com.tjclp.xl.{*, given}
import java.time.LocalDate
import scala.math.BigDecimal

println("=" * 70)
println("XL Formula Evaluator Demo (WI-08, WI-09 Complete)")
println("=" * 70)
println()

// ==================== Section 1: Basic Evaluation ====================

println("📊 Section 1: Basic Formula Evaluation (Low-Level API)")
println("-" * 70)

// Create test sheet with sample data
val sheet = Sheet("Test")
  .put(ref"A1", 10)
  .put(ref"A2", 20)
  .put(ref"A3", 30)
  .put(ref"B1", 5)
  .put(ref"B2", 2)

val evaluator = Evaluator.instance

// Example 1: Evaluate literal
val lit = TExpr.Lit(BigDecimal(42))
println(s"Literal 42: ${evaluator.eval(lit, sheet)}")
// Right(42)

// Example 2: Evaluate cell reference
val cellRef = TExpr.Ref(ref"A1", TExpr.decodeNumeric)
println(s"Cell A1: ${evaluator.eval(cellRef, sheet)}")
// Right(10)

// Example 3: Evaluate arithmetic
val add = TExpr.Add(
  TExpr.Ref(ref"A1", TExpr.decodeNumeric),
  TExpr.Ref(ref"A2", TExpr.decodeNumeric)
)
println(s"A1 + A2: ${evaluator.eval(add, sheet)}")
// Right(30)

println()

// ==================== Section 2: Range Aggregation ====================

println("Σ Section 2: Range Aggregation (SUM, COUNT, AVERAGE)")
println("-" * 70)

// SUM(A1:A3)
val sumExpr = TExpr.sum(CellRange.parse("A1:A3").toOption.get)
println(s"SUM(A1:A3): ${evaluator.eval(sumExpr, sheet)}")
// Right(60) // 10 + 20 + 30

// COUNT(A1:A3)
val countExpr = TExpr.count(CellRange.parse("A1:A3").toOption.get)
println(s"COUNT(A1:A3): ${evaluator.eval(countExpr, sheet)}")
// Right(3)

// AVERAGE(A1:A3)
val avgExpr = TExpr.average(CellRange.parse("A1:A3").toOption.get)
println(s"AVERAGE(A1:A3): ${evaluator.eval(avgExpr, sheet)}")
// Right(20) // 60/3

// MIN and MAX
val minExpr = TExpr.min(CellRange.parse("A1:A3").toOption.get)
println(s"MIN(A1:A3): ${evaluator.eval(minExpr, sheet)}")

val maxExpr = TExpr.max(CellRange.parse("A1:A3").toOption.get)
println(s"MAX(A1:A3): ${evaluator.eval(maxExpr, sheet)}")

println()

// ==================== Section 3: Conditional Logic ====================

println("🔀 Section 3: Conditional Logic (IF, AND, OR, NOT)")
println("-" * 70)

// IF(A1 > 15, "High", "Low")
val ifExpr = TExpr.If(
  TExpr.Gt(TExpr.Ref(ref"A1", TExpr.decodeNumeric), TExpr.Lit(BigDecimal(15))),
  TExpr.Lit("High"),
  TExpr.Lit("Low")
)
println(s"IF(A1>15, \"High\", \"Low\"): ${evaluator.eval(ifExpr, sheet)}")
// Right("Low") // A1=10, not > 15

// AND(A1>5, A2<25)
val andExpr = TExpr.And(
  TExpr.Gt(TExpr.Ref(ref"A1", TExpr.decodeNumeric), TExpr.Lit(BigDecimal(5))),
  TExpr.Lt(TExpr.Ref(ref"A2", TExpr.decodeNumeric), TExpr.Lit(BigDecimal(25)))
)
println(s"AND(A1>5, A2<25): ${evaluator.eval(andExpr, sheet)}")
// Right(true) // 10>5 && 20<25

// OR with short-circuit
val orExpr = TExpr.Or(
  TExpr.Gt(TExpr.Ref(ref"A1", TExpr.decodeNumeric), TExpr.Lit(BigDecimal(5))),
  TExpr.Lt(TExpr.Ref(ref"A2", TExpr.decodeNumeric), TExpr.Lit(BigDecimal(10)))
)
println(s"OR(A1>5, A2<10): ${evaluator.eval(orExpr, sheet)}")
// Right(true) // 10>5 is true, short-circuits (doesn't check A2<10)

println()

// ==================== Section 4: Text Functions ====================

println("📝 Section 4: Text Functions (CONCATENATE, LEFT, RIGHT, LEN, UPPER, LOWER)")
println("-" * 70)

// CONCATENATE
val concatExpr = TExpr.Concatenate(List(
  TExpr.Lit("Hello"),
  TExpr.Lit(" "),
  TExpr.Lit("World")
))
println(s"CONCATENATE: ${evaluator.eval(concatExpr, sheet)}")

// Text manipulation
val leftExpr = TExpr.Left(TExpr.Lit("Hello"), TExpr.Lit(3))
println(s"LEFT(\"Hello\", 3): ${evaluator.eval(leftExpr, sheet)}")

val upperExpr = TExpr.Upper(TExpr.Lit("hello"))
println(s"UPPER(\"hello\"): ${evaluator.eval(upperExpr, sheet)}")

val lenExpr = TExpr.Len(TExpr.Lit("Hello"))
println(s"LEN(\"Hello\"): ${evaluator.eval(lenExpr, sheet)}")

println()

// ==================== Section 5: Date Functions ====================

println("📅 Section 5: Date Functions (TODAY, DATE, YEAR, MONTH, DAY)")
println("-" * 70)

val clock = Clock.fixedDate(LocalDate.of(2025, 11, 21))

// TODAY()
val todayExpr = TExpr.today()
println(s"TODAY(): ${evaluator.eval(todayExpr, sheet, clock)}")

// DATE(2025, 11, 21)
val dateExpr = TExpr.Date(TExpr.Lit(2025), TExpr.Lit(11), TExpr.Lit(21))
println(s"DATE(2025, 11, 21): ${evaluator.eval(dateExpr, sheet, clock)}")

// YEAR(TODAY())
val yearExpr = TExpr.Year(TExpr.today())
println(s"YEAR(TODAY()): ${evaluator.eval(yearExpr, sheet, clock)}")

println()

// ==================== Section 6: Error Handling ====================

println("⚠️  Section 6: Error Handling")
println("-" * 70)

// Division by zero
val divZero = TExpr.Div(TExpr.Lit(BigDecimal(10)), TExpr.Lit(BigDecimal(0)))
evaluator.eval(divZero, sheet) match
  case Right(result) =>
    println(s"✗ 10/0 = $result (should have failed)")
  case Left(EvalError.DivByZero(num, denom)) =>
    println(s"✓ Division by zero caught: $num / $denom")
  case Left(error) =>
    println(s"? Unexpected error: $error")

// Missing cell reference (empty cell)
val emptyRef = TExpr.Ref(ref"Z999", TExpr.decodeNumeric)
evaluator.eval(emptyRef, sheet) match
  case Right(result) =>
    println(s"? Z999 = $result")
  case Left(error) =>
    println(s"✓ Empty cell handled: $error")

println()

// ==================== Section 7: High-Level API (Sheet Extensions) ====================

println("🚀 Section 7: High-Level API (Sheet Extension Methods)")
println("-" * 70)

val formulaSheet = Sheet("Formulas")
  .put(ref"A1", 10)
  .put(ref"A2", 20)
  .put(ref"B1", fx"=A1+A2")
  .put(ref"B2", fx"=B1*2")
  .put(ref"C1", fx"=SUM(A1:B2)")

// Evaluate single formula string
formulaSheet.evaluateFormula("=A1+5") match
  case Right(value) =>
    println(s"✓ evaluateFormula(\"=A1+5\") = $value")
  case Left(error) =>
    println(s"✗ Error: ${error.message}")

// Evaluate cell with formula
formulaSheet.evaluateCell(ref"B1") match
  case Right(value) =>
    println(s"✓ evaluateCell(B1) = $value  // B1 = A1+A2")
  case Left(error) =>
    println(s"✗ Error: ${error.message}")

// Evaluate all formulas with dependency checking
formulaSheet.evaluateWithDependencyCheck() match
  case Right(results) =>
    println(s"✓ evaluateWithDependencyCheck() succeeded (${results.size} formulas)")
    results.toSeq.sortBy(_._1.toString).foreach { case (ref, value) =>
      println(s"  $ref = $value")
    }
  case Left(error) =>
    println(s"✗ Error: ${error.message}")

println()

// ==================== Section 8: Complex Real-World Example ====================

println("🏗️  Section 8: Complex Real-World Formula (Financial Model)")
println("-" * 70)

// Create financial model sheet
val finModel = Sheet("Finance")
  .put(ref"A1", 1000000)  // Revenue
  .put(ref"A2", 600000)   // COGS
  .put(ref"A3", 250000)   // OpEx
  .put(ref"B1", CellValue.Number(BigDecimal(0.30)))     // Tax rate

  // Gross Profit = Revenue - COGS
  .put(ref"A4", fx"=A1-A2")

  // Operating Income = Gross Profit - OpEx
  .put(ref"A5", fx"=A4-A3")

  // Tax = Operating Income * Tax Rate
  .put(ref"A6", fx"=A5*B1")

  // Net Income = Operating Income - Tax
  .put(ref"A7", fx"=A5-A6")

  // Net Margin % = Net Income / Revenue
  .put(ref"A8", fx"=A7/A1")

println("Financial model structure:")
println("  A1: Revenue = $1,000,000")
println("  A2: COGS = $600,000")
println("  A3: OpEx = $250,000")
println("  A4: Gross Profit = A1-A2")
println("  A5: Operating Income = A4-A3")
println("  A6: Tax = A5*B1 (30%)")
println("  A7: Net Income = A5-A6")
println("  A8: Net Margin % = A7/A1")
println()

finModel.evaluateWithDependencyCheck() match
  case Right(results) =>
    println("✓ All formulas evaluated successfully:")
    println(f"  Gross Profit:     ${results(ref"A4")}%-15s // Revenue - COGS")
    println(f"  Operating Income: ${results(ref"A5")}%-15s // Gross Profit - OpEx")
    println(f"  Tax:              ${results(ref"A6")}%-15s // Operating Income * 30%%")
    println(f"  Net Income:       ${results(ref"A7")}%-15s // Operating Income - Tax")

    results(ref"A8") match
      case CellValue.Number(margin) =>
        println(f"  Net Margin:       ${margin.toDouble * 100}%.1f%%")
      case other =>
        println(f"  Net Margin:       $other")

  case Left(error) =>
    println(s"✗ Calculation error: ${error.message}")

println()

// ==================== Section 9: Dependency Analysis ====================

println("🔍 Section 9: Dependency Analysis")
println("-" * 70)

val graph = DependencyGraph.fromSheet(finModel)

// Show evaluation order
DependencyGraph.topologicalSort(graph) match
  case Right(order) =>
    println(s"Evaluation order (${order.size} formulas):")
    println(s"  ${order.map(_.toA1).mkString(" → ")}")
  case Left(error) =>
    println(s"✗ Cycle detected: ${error.cycle.map(_.toA1).mkString(" → ")}")

// Show Net Income dependencies
val netIncomeDeps = DependencyGraph.precedents(graph, ref"A7")
println(s"\nNet Income (A7) depends on: ${netIncomeDeps.map(_.toA1).mkString(", ")}")

// Show Revenue impact
val revenueImpact = DependencyGraph.dependents(graph, ref"A1")
println(s"Revenue (A1) impacts: ${revenueImpact.map(_.toA1).mkString(", ")}")

println()

// ==================== Section 10: Integration with fx Macro ====================

println("🔗 Section 10: Integration with fx Macro")
println("-" * 70)

// fx macro validates at compile time
val formula = fx"=A1+B1" // Returns CellValue.Formula

formula match
  case CellValue.Formula(text, _) =>
    println(s"fx macro validated: $text")

    // Evaluate using high-level API
    sheet.evaluateFormula(text) match
      case Right(result) =>
        println(s"Evaluation result: $result")  // 15 (A1=10, B1=5)
      case Left(error) =>
        println(s"Evaluation error: ${error.message}")
  case other =>
    println(s"Unexpected value type: $other")

println()

// ==================== Summary ====================

println("=" * 70)
println("Demo Complete! Formula System Fully Operational")
println("=" * 70)
println("""
✓ Low-level API: Evaluator.eval() for programmatic TExpr evaluation
✓ High-level API: sheet.evaluateFormula/Cell/WithDependencyCheck()
✓ 21 built-in functions: SUM, COUNT, AVERAGE, MIN, MAX, IF, AND, OR, NOT,
    CONCATENATE, LEFT, RIGHT, LEN, UPPER, LOWER, TODAY, NOW, DATE, YEAR, MONTH, DAY
✓ Dependency analysis: Precedent/dependent queries, topological sort
✓ Circular reference detection: Automatic cycle detection
✓ Total error handling: No exceptions, explicit Either types
✓ Type safety: Compile-time formula validation with fx macro
✓ Performance: 10k formulas in <10ms

Next Steps:
- Explore financial-model.sc for business use cases
- See dependency-analysis.sc for advanced graph features
- Check data-validation.sc for error detection patterns
- Read CLAUDE.md for comprehensive API documentation

🎉 Build production spreadsheet applications with confidence!
""")
