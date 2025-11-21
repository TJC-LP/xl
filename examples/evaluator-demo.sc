//> using scala 3.7.3
//> using dep com.tjclp::xl-core:0.1.0-SNAPSHOT
//> using dep com.tjclp::xl-evaluator:0.1.0-SNAPSHOT
//> using repository ivy2Local

/**
 * Formula Evaluator Demo (WI-08)
 *
 * **STATUS**: ⏳ NOT YET IMPLEMENTED
 *
 * This script will demonstrate formula evaluation once WI-08 is complete.
 * Currently, only the parser works (see formula-demo.sc).
 *
 * To run (after WI-08 implementation):
 *   1. Publish locally: ./mill xl-core.publishLocal && ./mill xl-evaluator.publishLocal
 *   2. Uncomment code below
 *   3. Run: scala-cli run examples/evaluator-demo.sc
 */

println("Formula Evaluator Demo - WI-08 Not Yet Implemented")
println("See formula-demo.sc for working formula parser examples")

// ==================== Uncomment After WI-08 Completion ====================

/*
import com.tjclp.xl.*
import com.tjclp.xl.formula.{FormulaParser, Evaluator, EvalError, TExpr}
import scala.math.BigDecimal

println("=" * 70)
println("XL Formula Evaluator Demo (WI-08)")
println("=" * 70)
println()

// ==================== Section 1: Basic Evaluation ====================

println("📊 Section 1: Basic Formula Evaluation")
println("-" * 70)

// Create test sheet with sample data
val sheet = Sheet.empty("Test")
  .put(ref"A1", BigDecimal(10))
  .put(ref"A2", BigDecimal(20))
  .put(ref"A3", BigDecimal(30))
  .put(ref"B1", BigDecimal(5))
  .put(ref"B2", BigDecimal(2))
  .toOption.get

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

println()

// ==================== Section 3: Conditional Logic ====================

println("🔀 Section 3: Conditional Logic (IF, AND, OR)")
println("-" * 70)

// IF(A1 > 15, "High", "Low")
val ifExpr = TExpr.If(
  TExpr.Gt(TExpr.Ref(ref"A1", TExpr.decodeNumeric), TExpr.Lit(BigDecimal(15))),
  TExpr.Lit("High"),
  TExpr.Lit("Low")
)
println(s"IF(A1>15, \"High\", \"Low\"): ${evaluator.eval(ifExpr, sheet)}")
// Right("Low") // A1=10, not > 15

// Nested IF
val nestedIf = TExpr.If(
  TExpr.Gt(TExpr.Ref(ref"A1", TExpr.decodeNumeric), TExpr.Lit(BigDecimal(50))),
  TExpr.Lit("Very High"),
  TExpr.If(
    TExpr.Gt(TExpr.Ref(ref"A1", TExpr.decodeNumeric), TExpr.Lit(BigDecimal(25))),
    TExpr.Lit("High"),
    TExpr.Lit("Low")
  )
)
println(s"Nested IF: ${evaluator.eval(nestedIf, sheet)}")
// Right("Low") // A1=10, neither > 50 nor > 25

println()

// ==================== Section 4: Error Handling ====================

println("⚠️  Section 4: Error Handling")
println("-" * 70)

// Division by zero
val divZero = TExpr.Div(TExpr.Lit(BigDecimal(10)), TExpr.Lit(BigDecimal(0)))
println(s"10/0: ${evaluator.eval(divZero, sheet)}")
// Left(EvalError.DivByZero("10", "0"))

// Missing cell reference
val missingRef = TExpr.Ref(ref"Z999", TExpr.decodeNumeric)
println(s"Missing cell Z999: ${evaluator.eval(missingRef, sheet)}")
// Left(EvalError.RefError(Z999, "cell not found"))

println()

// ==================== Section 5: Integration with fx Macro ====================

println("🔗 Section 5: Integration with fx Macro")
println("-" * 70)

// fx macro validates at compile time
val formula = fx"=A1+B1"  // Returns CellValue.Formula

formula match
  case CellValue.Formula(text) =>
    println(s"fx macro validated: $text")

    // Parse and evaluate
    FormulaParser.parse(text).flatMap { expr =>
      evaluator.eval(expr, sheet)
    } match
      case Right(result) =>
        println(s"Evaluation result: $result")  // 15 (A1=10, B1=5)
      case Left(error) =>
        println(s"Evaluation error: $error")

println()

// ==================== Section 6: Complex Real-World Example ====================

println("🏗️  Section 6: Complex Real-World Formula")
println("-" * 70)

// Create financial model sheet
val finModel = Sheet.empty("Finance")
  .put(
    ref"A1" -> BigDecimal(1000000),  // Revenue
    ref"A2" -> BigDecimal(600000),   // COGS
    ref"A3" -> BigDecimal(250000),   // OpEx
    ref"B1" -> BigDecimal(0.30)      // Tax rate
  )
  .toOption.get

// Formula: Net Income = (Revenue - COGS - OpEx) * (1 - TaxRate)
val netIncome = TExpr.Mul(
  TExpr.Sub(
    TExpr.Sub(
      TExpr.Ref(ref"A1", TExpr.decodeNumeric),  // Revenue
      TExpr.Ref(ref"A2", TExpr.decodeNumeric)   // COGS
    ),
    TExpr.Ref(ref"A3", TExpr.decodeNumeric)     // OpEx
  ),
  TExpr.Sub(
    TExpr.Lit(BigDecimal(1)),
    TExpr.Ref(ref"B1", TExpr.decodeNumeric)     // Tax rate
  )
)

evaluator.eval(netIncome, finModel) match
  case Right(result) =>
    println(s"Net Income: $$${result.setScale(2)}")
    // $105,000.00 = (1,000,000 - 600,000 - 250,000) * (1 - 0.30)
  case Left(error) =>
    println(s"Calculation error: $error")

println()
println("=" * 70)
println("Demo Complete! (Once WI-08 is implemented)")
println("=" * 70)
*/
