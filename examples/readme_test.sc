#!/usr/bin/env -S scala-cli shebang
//> using file project.scala


/**
 * README Examples Test Script
 *
 * This script validates all code examples from README.md
 * Run with: scala-cli run examples/readme-test.sc
 */

// Unified import - everything from core + formula + IO + display
import com.tjclp.xl.{*, given}
import com.tjclp.xl.unsafe.*
// SheetEvaluator extension methods now available from com.tjclp.xl.{*, given}

var passed = 0
var failed = 0

def test(name: String)(block: => Unit): Unit =
  print(s"Testing: $name... ")
  try
    block
    println("✓ PASSED")
    passed += 1
  catch
    case e: Throwable =>
      println(s"✗ FAILED: ${e.getMessage}")
      e.printStackTrace()
      failed += 1

println("=" * 70)
println("README.md Examples Test Suite")
println("=" * 70)
println()

// ============================================================================
// Hero Example (lines 8-23)
// ============================================================================
test("Hero Example - Financial Report") {
  val report = Sheet("Q1 Report")
    .put("A1", "Revenue")      .put("B1", 1250000, CellStyle.default.currency)
    .put("A2", "Expenses")     .put("B2", 875000, CellStyle.default.currency)
    .put("A3", "Net Income")   .put("B3", fx"=B1-B2")
    .put("A4", "Margin")       .put("B4", fx"=B3/B1")
    .style("A1:A4", CellStyle.default.bold)
    .style("B4", CellStyle.default.percent)

  Excel.write(Workbook.empty.put(report), "/tmp/readme-hero.xlsx")
  assert(report.cells.size >= 8, s"Expected at least 8 cells, got ${report.cells.size}")
  assert(report.cell("A1").isDefined, "A1 should exist")
}

// ============================================================================
// Easy Mode Example (lines 50-68)
// ============================================================================
test("Easy Mode - Read/Modify/Write") {
  // First create a workbook to test with
  val sheet = Sheet("Sheet1")
    .put("A1", "Test")
    .put("B1", 42)
  val workbook = Workbook.empty.put(sheet)
  Excel.write(workbook, "/tmp/readme-easymode-input.xlsx")

  // Now test read/modify/write cycle (the README example)
  val loaded = Excel.read("/tmp/readme-easymode-input.xlsx")

  val updated = loaded.update("Sheet1", sheet =>
    sheet
      .put("A1", "Updated!")
      .put("B1", 42)
      .style("A1:B1", CellStyle.default.bold)
  )

  Excel.write(updated, "/tmp/readme-easymode-output.xlsx")
  assert(updated.map(_.sheets.nonEmpty).getOrElse(false), "Should have at least one sheet")
}

// ============================================================================
// Patch DSL Example (lines 70-85)
// ============================================================================
test("Patch DSL - Declarative Sheet") {
  val sheet = Sheet("Sales")
    .put(
      (ref"A1" := "Product")   ++ (ref"B1" := "Price")   ++ (ref"C1" := "Qty") ++
      (ref"A2" := "Widget")    ++ (ref"B2" := 19.99)     ++ (ref"C2" := 100) ++
      (ref"A3" := "Gadget")    ++ (ref"B3" := 29.99)     ++ (ref"C3" := 50) ++
      (ref"D1" := "Total")     ++ (ref"D2" := fx"=B2*C2") ++ (ref"D3" := fx"=B3*C3") ++
      ref"A1:D1".styled(CellStyle.default.bold)
    )

  assert(sheet.cells.size >= 11, s"Expected at least 11 cells, got ${sheet.cells.size}")
  assert(sheet.cell("D2").map(_.value).exists(_.isInstanceOf[CellValue.Formula]), "D2 should be a formula")
}

// ============================================================================
// Compile-Time Validated References (lines 91-99)
// ============================================================================
test("Compile-Time Validated References") {
  // Note: Use addressing types directly for compatibility with scala-cli
  import com.tjclp.xl.addressing.{ARef => ARefType, CellRange => CellRangeType}
  import com.tjclp.xl.addressing.ARef.toA1

  val cell: ARefType = ref"A1"          // Single cell
  val range: CellRangeType = ref"A1:B10"    // Range
  val qualified = ref"Sheet1!A1:C100"   // With sheet name

  assert(toA1(cell) == "A1", "cell should be A1")
  assert(range.start == ref"A1", "Range start should be A1")
  assert(range.end == ref"B10", "Range end should be B10")

  // Runtime interpolation (returns Either)
  val col = "A"
  val row = "1"
  val dynamic = ref"$col$row"           // Either[XLError, RefType]
  assert(dynamic.isRight, "Dynamic ref should parse successfully")
}

// ============================================================================
// Formatted Literals (lines 103-108)
// ============================================================================
test("Formatted Literals") {
  val price = money"$$1,234.56"         // Currency format (escaped $)
  val growth = percent"12.5%"           // Percent format
  val date = date"2025-11-24"           // ISO date format
  val loss = accounting"($$500.00)"     // Accounting (negatives in parens)

  assert(price.isInstanceOf[Formatted], "money literal should be Formatted")
  assert(growth.isInstanceOf[Formatted], "percent literal should be Formatted")
  assert(date.isInstanceOf[Formatted], "date literal should be Formatted")
  assert(loss.isInstanceOf[Formatted], "accounting literal should be Formatted")
}

// ============================================================================
// Fluent Style DSL (lines 114-125)
// ============================================================================
test("Fluent Style DSL") {
  val header = CellStyle.default
    .bold
    .size(14.0)
    .bgBlue
    .white
    .center
    .bordered

  val currency = CellStyle.default.currency
  val percentStyle = CellStyle.default.percent

  // Verify the styles are constructed (no need to check internal details)
  assert(header != CellStyle.default, "Header should be modified from default")
}

// ============================================================================
// Rich Text (lines 129-132)
// ============================================================================
test("Rich Text - Multi-Format Cell") {
  val sheet = Sheet("RichText")
  val text = "Error: ".bold.red + "Fix this!".underline
  val richSheet = sheet.put("A1", text)

  assert(richSheet.cell("A1").isDefined, "A1 should exist")
}

// ============================================================================
// Patch Composition (lines 136-145)
// ============================================================================
test("Patch Composition") {
  val headerStyle = CellStyle.default.bold.size(14.0)
  val sheet = Sheet("Patch Test")

  val patch =
    (ref"A1" := "Title") ++
    ref"A1".styled(headerStyle) ++
    ref"A1:C1".merge

  val result = sheet.put(patch)

  assert(result.cell("A1").isDefined, "A1 should exist")
  assert(result.mergedRanges.nonEmpty, "Should have merged range")
}

// ============================================================================
// Formula System (lines 151-167)
// ============================================================================
test("Formula System - Parsing") {
  // Parse formulas
  val sum = FormulaParser.parse("=SUM(A1:B10)")
  val ifExpr = FormulaParser.parse("=IF(A1>0, B1, C1)")

  assert(sum.isRight, s"SUM formula should parse: $sum")
  assert(ifExpr.isRight, s"IF formula should parse: $ifExpr")
}

test("Formula System - Evaluation with Dependency Check") {
  val sheet = Sheet("FormulaTest")
    .put("A1", 100)
    .put("B1", 200)
    .put("C1", fx"=A1+B1")

  // Evaluate with cycle detection
  sheet.evaluateWithDependencyCheck() match
    case Right(results) =>
      assert(results.nonEmpty, "Should have results")
    case Left(error) =>
      throw new AssertionError(s"Unexpected error: $error")
}

test("Formula System - Dependency Analysis") {
  val sheet = Sheet("DepTest")
    .put("A1", 100)
    .put("B1", fx"=A1*2")
    .put("C1", fx"=B1+A1")

  val graph = DependencyGraph.fromSheet(sheet)
  val precedents = DependencyGraph.precedents(graph, ref"B1")

  assert(precedents.contains(ref"A1"), "B1 should depend on A1")
}

// ============================================================================
// Formula Roundtrip Test
// ============================================================================
test("Formula Roundtrip - Write and Read Back") {
  // Create a sheet with formulas and their values
  val sheet = Sheet("FormulaRoundtrip")
    .put("A1", 100)
    .put("A2", 200)
    .put("A3", 300)
    .put("B1", CellValue.Formula("A1*2", Some(CellValue.Number(BigDecimal(200)))))
    .put("B2", CellValue.Formula("A2*2", Some(CellValue.Number(BigDecimal(400)))))
    .put("B3", CellValue.Formula("SUM(A1:A3)", Some(CellValue.Number(BigDecimal(600)))))

  val workbook = Workbook.empty.put(sheet).remove("Sheet1").unsafe

  // Write to file
  Excel.write(workbook, "/tmp/readme-formula-roundtrip.xlsx")

  // Read back and verify formulas are preserved
  val readBack = Excel.read("/tmp/readme-formula-roundtrip.xlsx")

  readBack.sheets.find(_.name.value == "FormulaRoundtrip") match
    case Some(readSheet) =>
      // Verify B1 formula
      readSheet.cell("B1").map(_.value) match
        case Some(CellValue.Formula(expr, cached)) =>
          assert(expr == "A1*2", s"B1 formula should be 'A1*2', got '$expr'")
          cached match
            case Some(CellValue.Number(n)) =>
              assert(n == BigDecimal(200), s"B1 cached value should be 200, got $n")
            case other =>
              throw new AssertionError(s"B1 cached value should be Number(200), got $other")
        case other =>
          throw new AssertionError(s"B1 should be Formula, got $other")

      // Verify B3 SUM formula
      readSheet.cell("B3").map(_.value) match
        case Some(CellValue.Formula(expr, cached)) =>
          assert(expr == "SUM(A1:A3)", s"B3 formula should be 'SUM(A1:A3)', got '$expr'")
          cached match
            case Some(CellValue.Number(n)) =>
              assert(n == BigDecimal(600), s"B3 cached value should be 600, got $n")
            case other =>
              throw new AssertionError(s"B3 cached value should be Number(600), got $other")
        case other =>
          throw new AssertionError(s"B3 should be Formula, got $other")

      println("  → Formula expressions and cached values roundtripped successfully!")
    case None =>
      throw new AssertionError("FormulaRoundtrip sheet not found in read-back")
}

// ============================================================================
// Final Integration Test
// ============================================================================
test("Final Workbook Write") {
  // Create comprehensive workbook with all examples
  val heroSheet = Sheet("Q1 Report")
    .put("A1", "Revenue")      .put("B1", 1250000, CellStyle.default.currency)
    .put("A2", "Expenses")     .put("B2", 875000, CellStyle.default.currency)
    .put("A3", "Net Income")   .put("B3", fx"=B1-B2")
    .put("A4", "Margin")       .put("B4", fx"=B3/B1")
    .style("A1:A4", CellStyle.default.bold)
    .style("B4", CellStyle.default.percent)

  val salesSheet = Sheet("Sales")
    .put(
      (ref"A1" := "Product")   ++ (ref"B1" := "Price")   ++ (ref"C1" := "Qty") ++
      (ref"A2" := "Widget")    ++ (ref"B2" := 19.99)     ++ (ref"C2" := 100) ++
      (ref"A3" := "Gadget")    ++ (ref"B3" := 29.99)     ++ (ref"C3" := 50) ++
      (ref"D1" := "Total")     ++ (ref"D2" := fx"=B2*C2") ++ (ref"D3" := fx"=B3*C3")
    )

  val workbook = Workbook.empty
    .put(heroSheet)
    .put(salesSheet)
    .remove("Sheet1")
    .unsafe

  Excel.write(workbook, "/tmp/readme-examples.xlsx")

  // Verify file was created (read-back of formula-only cells requires evaluation)
  assert(java.nio.file.Files.exists(java.nio.file.Paths.get("/tmp/readme-examples.xlsx")),
    "Output file should exist")
}

// ============================================================================
// Summary
// ============================================================================
println()
println("=" * 70)
println(s"Test Results: $passed passed, $failed failed")
println("=" * 70)

if failed > 0 then
  println("\n❌ Some tests failed!")
  sys.exit(1)
else
  println("\n✅ All README examples work correctly!")
  println(s"   Output files written to /tmp/readme-*.xlsx")
