#!/usr/bin/env -S scala-cli shebang
//> using file project.scala


/**
 * XL Formula System - Quick Start Guide
 *
 * This 5-minute example demonstrates the core formula system capabilities:
 * - Parse Excel formulas to typed AST
 * - Evaluate formulas against sheets
 * - Handle errors gracefully
 * - Detect circular references automatically
 *
 * Run with: scala-cli examples/quick_start.sc
 */

import com.tjclp.xl.{*, given}
// Note: No unsafe import needed - all refs are compile-time validated literals

// ============================================================================
// STEP 1: Parse Formulas
// ============================================================================

println("=" * 70)
println("STEP 1: Formula Parsing")
println("=" * 70)

// Parse a simple arithmetic formula
val simpleFormula = "=10+20"
FormulaParser.parse(simpleFormula) match
  case Right(expr) =>
    println(s"✓ Parsed: $simpleFormula")
    println(s"  AST: $expr")
  case Left(error) =>
    println(s"✗ Parse error: $error")

// Parse a SUM formula with range
val sumFormula = "=SUM(A1:A3)"
FormulaParser.parse(sumFormula) match
  case Right(expr) =>
    println(s"✓ Parsed: $sumFormula")
  case Left(error) =>
    println(s"✗ Parse error: $error")

// Parse a complex nested formula
val complexFormula = "=IF(A1>100, \"High\", \"Low\")"
FormulaParser.parse(complexFormula) match
  case Right(expr) =>
    println(s"✓ Parsed: $complexFormula")
  case Left(error) =>
    println(s"✗ Parse error: $error")

println()

// ============================================================================
// STEP 2: Create a Sheet with Data
// ============================================================================

println("=" * 70)
println("STEP 2: Create Sheet with Data")
println("=" * 70)

val sheet = Sheet("Demo")
  .put(ref"A1", 150)
  .put(ref"A2", 200)
  .put(ref"A3", 75)
  .put(ref"B1", fx"=SUM(A1:A3)")
  .put(ref"B2", fx"=AVERAGE(A1:A3)")
  .put(ref"C1", fx"=IF(B1>400, \"High\", \"Low\")")

println("Sheet contents:")
println("  A1: 150")
println("  A2: 200")
println("  A3: 75")
println("  B1: =SUM(A1:A3)")
println("  B2: =AVERAGE(A1:A3)")
println("  C1: =IF(B1>400, \"High\", \"Low\")")
println()

// ============================================================================
// STEP 3: Apply Number Formatting
// ============================================================================

println("=" * 70)
println("STEP 3: Apply Number Formatting")
println("=" * 70)

// Style formula cells as decimal numbers
val decimalStyle = CellStyle.default.withNumFmt(NumFmt.Decimal)
val styledSheet = sheet
  .style(ref"B1:B2", decimalStyle)  // Format SUM and AVERAGE as decimals

println("✓ Applied decimal formatting to formula range B1:B2")
println()

// ============================================================================
// STEP 4: Evaluate Individual Formulas
// ============================================================================

println("=" * 70)
println("STEP 4: Evaluate Individual Formulas")
println("=" * 70)

// Evaluate B1 (SUM formula)
styledSheet.evaluateCell(ref"B1") match
  case Right(value) =>
    println(s"✓ B1 = $value  // Sum of A1:A3")
  case Left(error) =>
    println(s"✗ B1 evaluation error: ${error.message}")

// Evaluate B2 (AVERAGE formula)
styledSheet.evaluateCell(ref"B2") match
  case Right(value) =>
    println(s"✓ B2 = $value  // Average of A1:A3")
  case Left(error) =>
    println(s"✗ B2 evaluation error: ${error.message}")

// Evaluate C1 (IF formula) - references B1 which is still a Formula
// Note: evaluateCell() doesn't evaluate dependencies - B1 is still unevaluated
// This demonstrates the need for bulk evaluation with dependency checking (Step 5)
styledSheet.evaluateCell(ref"C1") match
  case Right(value) =>
    println(s"✓ C1 = $value  // Conditional based on B1")
  case Left(error) =>
    println(s"⚠️  C1 needs dependency evaluation (expected): ${error.message.take(80)}...")
    println(s"    Use Workbook(sheet).recalculate() for formulas with formula dependencies")

println()

// ============================================================================
// STEP 5: Evaluate All Formulas (with Dependency Checking)
// ============================================================================

println("=" * 70)
println("STEP 5: Bulk Evaluation with Dependency Checking")
println("=" * 70)

val recalculated = Workbook(styledSheet).recalculate()
if recalculated.isClean then
  val results = recalculated.evaluated.getOrElse(styledSheet.name, Map.empty)
  println("✓ All formulas evaluated successfully!")
  println(s"  Total formulas: ${results.size}")
  println("\nResults:")
  results.toSeq.sortBy(_._1.toString).foreach { case (ref, value) =>
    println(s"  $ref = $value")
  }
else recalculated.errors.foreach(error => println(s"✗ Evaluation error: ${error.render}"))

println()

// ============================================================================
// STEP 5: Circular Reference Detection
// ============================================================================

println("=" * 70)
println("STEP 5: Circular Reference Detection")
println("=" * 70)

// Create a sheet with a circular reference
val cyclicSheet = Sheet("Cyclic")
  .put(ref"A1", fx"=B1+10")
  .put(ref"B1", fx"=C1*2")
  .put(ref"C1", fx"=A1+5")  // Cycle: A1 → B1 → C1 → A1

println("Sheet with circular reference:")
println("  A1: =B1+10")
println("  B1: =C1*2")
println("  C1: =A1+5  // Creates cycle: A1 → B1 → C1 → A1")
println()

// A cycle never fails the recalculation: each member is reported as a per-cell error
val cyclicResult = Workbook(cyclicSheet).recalculate()
if cyclicResult.isClean then println("✗ Unexpected success (should have detected cycle)")
else
  println("✓ Circular reference detected!")
  cyclicResult.errors.foreach(error => println(s"  Error: ${error.render}"))

println()

// ============================================================================
// STEP 6: Dependency Analysis
// ============================================================================

println("=" * 70)
println("STEP 6: Dependency Analysis")
println("=" * 70)

// Build the workbook's dependency graph
val graph = QualifiedGraph.of(Workbook(sheet))
def at(a1: ARef): QualifiedRef = QualifiedRef(sheet.name, a1)
def cells(refs: Iterable[QualifiedRef]): String =
  QualifiedGraph.sorted(refs).map(_.ref.toA1).mkString(", ")

// Check for cycles (should be none): a cyclic strongly connected component is a circular reference
val cycles = graph.sccs.filter(_.cyclic)
if cycles.isEmpty then println("✓ No circular references detected")
else cycles.foreach(cycle => println(s"✗ Circular reference: ${cells(cycle.members)}"))

// Evaluation order: the components, dependency-first
if cycles.isEmpty then
  println(s"✓ Evaluation order: ${graph.sccs.flatMap(_.members).map(_.ref.toA1).mkString(" → ")}")

// Query precedents (cells B1 depends on)
val b1Precedents = graph.precedentsOf(at(ref"B1"))
println(s"✓ B1 depends on: ${cells(b1Precedents)}")

// Query dependents (cells that depend on A1)
val a1Dependents = graph.dependentsOf(at(ref"A1"))
println(s"✓ Cells depending on A1: ${cells(a1Dependents)}")

println()

// ============================================================================
// Summary
// ============================================================================

println("=" * 70)
println("SUMMARY: What You've Learned")
println("=" * 70)
println("""
✓ Parse Excel formulas to typed AST with FormulaParser.parse()
✓ Evaluate formulas against sheets with sheet.evaluateCell()
✓ Handle errors gracefully with Either types (no exceptions!)
✓ Detect circular references automatically
✓ Evaluate all formulas safely with dependency checking
✓ Analyze dependencies (precedents/dependents) with QualifiedGraph

Next Steps:
- Explore financial_model.sc for business use cases
- See dependency_analysis.sc for advanced graph features
- Check data_validation.sc for error detection patterns
- Read CLAUDE.md for comprehensive API documentation

🎉 You're ready to build spreadsheet applications with XL!
""")
