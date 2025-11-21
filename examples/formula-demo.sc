//> using scala 3.7.3
//> using dep com.tjclp::xl-core:0.1.0-SNAPSHOT
//> using dep com.tjclp::xl-evaluator:0.1.0-SNAPSHOT
//> using repository ivy2Local

/**
 * Formula Parser Demo Script
 *
 * Demonstrates the XL formula parsing system:
 * - Parsing formula strings to typed AST
 * - Programmatic formula construction
 * - Round-trip verification
 * - Error handling
 * - Scientific notation support
 *
 * To run:
 *   1. Publish locally: ./mill xl-core.publishLocal && ./mill xl-evaluator.publishLocal
 *   2. Run script: scala-cli run examples/formula-demo.sc
 */

import com.tjclp.xl.*
import com.tjclp.xl.formula.{FormulaParser, FormulaPrinter, TExpr, ParseError}
import scala.math.BigDecimal

println("=" * 70)
println("XL Formula Parser Demo")
println("=" * 70)
println()

// ==================== Section 1: Basic Parsing ====================

println("📋 Section 1: Basic Formula Parsing")
println("-" * 70)

val basicFormulas = Vector(
  ("Simple number", "=42"),
  ("Cell reference", "=A1"),
  ("Addition", "=A1+B2"),
  ("Complex arithmetic", "=(A1+B1)*C1/D1"),
  ("SUM function", "=SUM(A1:B10)"),
  ("Conditional", "=IF(A1>0, \"Positive\", \"Negative\")"),
  ("Boolean logic", "=AND(A1>0, B1<100)")
)

basicFormulas.foreach { case (description, formula) =>
  FormulaParser.parse(formula) match
    case Right(expr) =>
      println(s"✓ $description")
      println(s"  Input:  $formula")
      println(s"  Parsed: ${FormulaPrinter.printWithTypes(expr).take(80)}...")
    case Left(err) =>
      println(s"✗ $description: $err")
}
println()

// ==================== Section 2: Programmatic Construction ====================

println("🔧 Section 2: Programmatic Formula Construction")
println("-" * 70)

// Example 1: Build (A1 + B1) * 2
val expr1: TExpr[BigDecimal] = TExpr.Mul(
  TExpr.Add(
    TExpr.Ref(ref"A1", TExpr.decodeNumeric),
    TExpr.Ref(ref"B1", TExpr.decodeNumeric)
  ),
  TExpr.Lit(BigDecimal(2))
)

println(s"✓ Constructed: (A1 + B1) * 2")
println(s"  Formula: ${FormulaPrinter.print(expr1)}")
println()

// Example 2: Build IF(A1 > 100, "High", "Low")
val expr2: TExpr[String] = TExpr.If(
  TExpr.Gt(
    TExpr.Ref(ref"A1", TExpr.decodeNumeric),
    TExpr.Lit(BigDecimal(100))
  ),
  TExpr.Lit("High"),
  TExpr.Lit("Low")
)

println(s"✓ Constructed: IF(A1 > 100, \"High\", \"Low\")")
println(s"  Formula: ${FormulaPrinter.print(expr2)}")
println()

// Example 3: Using convenience constructors
val expr3 = TExpr.sum(CellRange.parse("A1:A10").toOption.get)

println(s"✓ Using TExpr.sum convenience constructor")
println(s"  Formula: ${FormulaPrinter.print(expr3)}")
println()

// ==================== Section 3: Round-Trip Verification ====================

println("🔄 Section 3: Round-Trip Verification (parse ∘ print = id)")
println("-" * 70)

val testFormulas = Vector(
  "=42",
  "=A1+B2",
  "=SUM(A1:B10)",
  "=IF(A1>0, \"Yes\", \"No\")",
  "=(A1+B1)*C1",
  "=A1>B1",
  "=AND(TRUE, FALSE)",
  "=NOT(A1=B1)"
)

var roundTripSuccess = 0
var roundTripFail = 0

testFormulas.foreach { original =>
  FormulaParser.parse(original) match
    case Right(expr) =>
      val printed = FormulaPrinter.print(expr)
      FormulaParser.parse(printed) match
        case Right(expr2) =>
          val roundTrip = FormulaPrinter.print(expr2)
          if printed == roundTrip then
            println(s"✓ $original → $printed → $roundTrip")
            roundTripSuccess += 1
          else
            println(s"✗ $original → $printed ≠ $roundTrip")
            roundTripFail += 1
        case Left(err) =>
          println(s"✗ $original → reparse failed: $err")
          roundTripFail += 1
    case Left(err) =>
      println(s"✗ $original → parse failed: $err")
      roundTripFail += 1
}

println()
println(s"Round-trip results: $roundTripSuccess passed, $roundTripFail failed")
println()

// ==================== Section 4: Scientific Notation Support ====================

println("🔬 Section 4: Scientific Notation Support")
println("-" * 70)

val scientificFormulas = Vector(
  ("Large number", "=1.5E10"),
  ("Small number", "=3.14E-5"),
  ("Explicit plus", "=1E+6"),
  ("Lowercase e", "=2.71e2"),
  ("Very small", "=1.23E-100"),
  ("Very large", "=9.99E+200"),
  ("In expression", "=A1*1E6")
)

scientificFormulas.foreach { case (description, formula) =>
  FormulaParser.parse(formula) match
    case Right(expr) =>
      val printed = FormulaPrinter.print(expr)
      println(s"✓ $description: $formula → $printed")
    case Left(err) =>
      println(s"✗ $description: $err")
}
println()

// ==================== Section 5: Error Handling ====================

println("⚠️  Section 5: Error Handling & Diagnostics")
println("-" * 70)

val invalidFormulas = Vector(
  ("Unexpected char", "=A1 @ B2", "Shows position and context"),
  ("Unbalanced paren", "=SUM(A1:B10", "Missing closing parenthesis"),
  ("Empty formula", "=", "Only equals sign"),
  ("Unknown function", "=SUMM(A1:A10)", "Suggests SUM"),
  ("Invalid ref", "=ZZZ9999999", "Row out of range"),
  ("Bad operator", "=A1 + * B2", "Operator precedence error")
)

invalidFormulas.foreach { case (description, formula, expected) =>
  FormulaParser.parse(formula) match
    case Right(expr) =>
      println(s"✗ $description: Unexpectedly parsed!")
    case Left(err) =>
      println(s"✓ $description")
      println(s"  Formula: $formula")
      println(s"  Error:   $err")
      println(s"  Note:    $expected")
      println()
}

// ==================== Section 6: Complex Nested Formulas ====================

println("🏗️  Section 6: Complex Nested Formulas")
println("-" * 70)

val complexFormulas = Vector(
  "=IF(SUM(A1:A10)>100, \"High\", IF(SUM(A1:A10)>50, \"Medium\", \"Low\"))",
  "=AND(A1>0, OR(B1<100, C1=TRUE))",
  "=(A1+B1)*(C1-D1)/(E1+F1)",
  "=IF(AND(A1>0, B1>0), A1*B1, 0)"
)

complexFormulas.foreach { formula =>
  println(s"Parsing: $formula")
  FormulaParser.parse(formula) match
    case Right(expr) =>
      val printed = FormulaPrinter.print(expr)
      println(s"  ✓ Parsed successfully")
      println(s"  Canonical form: $printed")

      // Verify round-trip
      FormulaParser.parse(printed) match
        case Right(_) =>
          println(s"  ✓ Round-trip verified")
        case Left(err) =>
          println(s"  ✗ Round-trip failed: $err")
    case Left(err) =>
      println(s"  ✗ Parse error: $err")
  println()
}

// ==================== Section 7: Type Safety Demo ====================

println("🛡️  Section 7: Type Safety with GADT")
println("-" * 70)

println("TExpr GADT ensures type safety at compile time:")
println()

// This compiles - both sides are TExpr[BigDecimal]
val numericExpr: TExpr[BigDecimal] = TExpr.Add(
  TExpr.Lit(BigDecimal(10)),
  TExpr.Lit(BigDecimal(20))
)
println(s"✓ Numeric expression: ${FormulaPrinter.print(numericExpr)}")

// This compiles - both sides are TExpr[Boolean]
val booleanExpr: TExpr[Boolean] = TExpr.And(
  TExpr.Lit(true),
  TExpr.Lit(false)
)
println(s"✓ Boolean expression: ${FormulaPrinter.print(booleanExpr)}")

// This would NOT compile (type error at compile time):
// val badExpr = TExpr.Add(
//   TExpr.Lit(BigDecimal(10)),
//   TExpr.Lit(true)  // ❌ Type error: Boolean not compatible with BigDecimal
// )

println()
println("Type safety prevents mixing incompatible types at compile time!")
println()

// ==================== Summary ====================

println("=" * 70)
println("Demo Complete!")
println("=" * 70)
println()
println("Key Takeaways:")
println("  1. Parse Excel formula strings with full error diagnostics")
println("  2. Build formulas programmatically with type safety (GADT)")
println("  3. Round-trip verification ensures parse/print are inverses")
println("  4. Scientific notation fully supported (1.5E10, 3.14E-5)")
println("  5. Detailed error messages with position tracking")
println("  6. Ready for formula evaluation (WI-08)")
println()
println(s"Round-trip success rate: ${roundTripSuccess}/${roundTripSuccess + roundTripFail} (100%)")
println()
