package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.{FunSuite, ScalaCheckSuite}
import org.scalacheck.Prop.*
import org.scalacheck.{Arbitrary, Gen}

/**
 * Tests for cross-sheet formula support (TJC-351).
 *
 * Tests parsing, printing, and evaluation of cross-sheet references:
 *   - Simple references: `=Sales!A1`
 *   - Range references: `=SUM(Sales!A1:A10)`
 *   - Quoted sheet names: `='Q1 Report'!A1`
 *   - Special character escaping: `='O''Brien''s'!A1`
 */
class CrossSheetFormulaSpec extends ScalaCheckSuite:

  // ===== Test Helpers =====

  def sheetWith(name: String, cells: (ARef, CellValue)*): Sheet =
    cells.foldLeft(Sheet(SheetName.unsafe(name))) { case (s, (ref, value)) =>
      s.put(ref, value)
    }

  def workbookWith(sheets: Sheet*): Workbook =
    Workbook(sheets.toVector)

  // ===== Parser Tests =====

  test("parses simple cross-sheet reference: Sales!A1") {
    val result = FormulaParser.parse("=Sales!A1")
    assert(result.isRight, s"Parse failed: $result")
    val printed = result.map(FormulaPrinter.print(_))
    assertEquals(printed, Right("=Sales!A1"))
  }

  test("parses cross-sheet reference with lowercase sheet name") {
    val result = FormulaParser.parse("=sales!B2")
    assert(result.isRight, s"Parse failed: $result")
  }

  test("parses cross-sheet range: Sales!A1:B10") {
    val result = FormulaParser.parse("=SUM(Sales!A1:B10)")
    assert(result.isRight, s"Parse failed: $result")
    val printed = result.map(FormulaPrinter.print(_))
    assertEquals(printed, Right("=SUM(Sales!A1:B10)"))
  }

  // Note: Quoted sheet names ('Q1 Report'!A1) are NOT yet supported by the parser.
  // The FormulaPrinter.formatSheetName correctly outputs them, but parsing is a TODO.
  // Tests below document current behavior (parse failure).
  // TODO: TJC-360 - Add parser support for quoted sheet names in formulas

  test("quoted sheet names are not yet supported (parse limitation)".ignore) {
    // When implemented, these should parse:
    // - ='Q1 Report'!A1
    // - ='Sales&Marketing'!A1
    // - ='O''Brien''s Data'!A1
    // - ='2024Q1'!A1
    // - ='Jan-Mar'!A1
    val result = FormulaParser.parse("='Q1 Report'!A1")
    // Currently fails - this test documents the limitation
    assert(result.isLeft, "Quoted sheet names not yet supported")
  }

  // ===== Round-Trip Property Tests =====

  // Generator for valid unquoted sheet names
  val genSimpleSheetName: Gen[String] = for
    first <- Gen.alphaChar
    rest <- Gen.listOfN(5, Gen.alphaNumChar)
  yield (first :: rest).mkString

  // Generator for sheet names requiring quotes (spaces, special chars)
  val genQuotedSheetName: Gen[String] = for
    name <- Gen.oneOf(
      Gen.const("Q1 Report"),
      Gen.const("Sales&Marketing"),
      Gen.const("2024 Data"),
      Gen.const("Jan-Mar"),
      Gen.const("O'Brien's")
    )
  yield name

  property("round-trip: parse . print = id for simple cross-sheet refs") {
    forAll(genSimpleSheetName, Gen.choose(1, 100), Gen.choose(1, 26)) { (sheetName, row, col) =>
      val colLetter = ('A' + col - 1).toChar
      val formula = s"=$sheetName!$colLetter$row"
      val parsed = FormulaParser.parse(formula)
      val reprinted = parsed.map(FormulaPrinter.print(_))
      reprinted == Right(formula)
    }
  }

  property("round-trip: parse . print = id for cross-sheet SUM") {
    forAll(genSimpleSheetName, Gen.choose(1, 50), Gen.choose(2, 100)) { (sheetName, startRow, endRow) =>
      val actualEnd = Math.max(startRow + 1, endRow)
      val formula = s"=SUM($sheetName!A$startRow:A$actualEnd)"
      val parsed = FormulaParser.parse(formula)
      val reprinted = parsed.map(FormulaPrinter.print(_))
      reprinted == Right(formula)
    }
  }

  // ===== Evaluator Tests =====

  test("SheetPolyRef: evaluates number from target sheet") {
    val sales = sheetWith("Sales", ref"A1" -> CellValue.Number(BigDecimal(100)))
    val main = sheetWith("Main")
    val wb = workbookWith(main, sales)

    val result = main.evaluateFormula("=Sales!A1", workbook = Some(wb))
    assertEquals(result, Right(CellValue.Number(BigDecimal(100))))
  }

  test("SheetPolyRef: evaluates text from target sheet") {
    val data = sheetWith("Data", ref"B2" -> CellValue.Text("Hello"))
    val main = sheetWith("Main")
    val wb = workbookWith(main, data)

    val result = main.evaluateFormula("=Data!B2", workbook = Some(wb))
    assertEquals(result, Right(CellValue.Text("Hello")))
  }

  test("SheetPolyRef: evaluates formula with cached value") {
    val sales = sheetWith(
      "Sales",
      ref"A1" -> CellValue.Formula("=10*5", Some(CellValue.Number(BigDecimal(50))))
    )
    val main = sheetWith("Main")
    val wb = workbookWith(main, sales)

    val result = main.evaluateFormula("=Sales!A1", workbook = Some(wb))
    assertEquals(result, Right(CellValue.Number(BigDecimal(50))))
  }

  // ===== GH-161: Cross-sheet reference to formula cell without cached value =====

  test("GH-161: evaluates formula WITHOUT cached value") {
    // This is the bug case: formula with no cache should evaluate, not return 0
    val sales = sheetWith(
      "Sales",
      ref"A1" -> CellValue.Formula("10*5", None) // No cached value!
    )
    val main = sheetWith("Main")
    val wb = workbookWith(main, sales)

    val result = main.evaluateFormula("=Sales!A1", workbook = Some(wb))
    assertEquals(result, Right(CellValue.Number(BigDecimal(50))))
  }

  test("GH-161: evaluates nested formula chain without cache") {
    // A1 = B1 + 10, B1 = 20, both without cache
    val data = sheetWith(
      "Data",
      ref"A1" -> CellValue.Formula("B1+10", None),
      ref"B1" -> CellValue.Number(BigDecimal(20))
    )
    val main = sheetWith("Main")
    val wb = workbookWith(main, data)

    val result = main.evaluateFormula("=Data!A1", workbook = Some(wb))
    assertEquals(result, Right(CellValue.Number(BigDecimal(30))))
  }

  test("GH-161: cross-sheet ref to formula with SUM") {
    // Sales!A1 = SUM(B1:B3), B1:B3 contain values
    val sales = sheetWith(
      "Sales",
      ref"A1" -> CellValue.Formula("SUM(B1:B3)", None),
      ref"B1" -> CellValue.Number(BigDecimal(10)),
      ref"B2" -> CellValue.Number(BigDecimal(20)),
      ref"B3" -> CellValue.Number(BigDecimal(30))
    )
    val main = sheetWith("Main")
    val wb = workbookWith(main, sales)

    val result = main.evaluateFormula("=Sales!A1", workbook = Some(wb))
    assertEquals(result, Right(CellValue.Number(BigDecimal(60))))
  }

  test("GH-161: cross-sheet ref to formula with arithmetic") {
    // Data!A1 = A2*A3 where A2=5, A3=7
    val data = sheetWith(
      "Data",
      ref"A1" -> CellValue.Formula("A2*A3", None),
      ref"A2" -> CellValue.Number(BigDecimal(5)),
      ref"A3" -> CellValue.Number(BigDecimal(7))
    )
    val main = sheetWith("Main")
    val wb = workbookWith(main, data)

    val result = main.evaluateFormula("=Data!A1", workbook = Some(wb))
    assertEquals(result, Right(CellValue.Number(BigDecimal(35))))
  }

  test("SheetPolyRef: empty cell returns 0") {
    val empty = sheetWith("Empty")
    val main = sheetWith("Main")
    val wb = workbookWith(main, empty)

    val result = main.evaluateFormula("=Empty!Z99", workbook = Some(wb))
    assertEquals(result, Right(CellValue.Number(BigDecimal(0))))
  }

  test("SheetPolyRef: boolean value from target sheet") {
    val flags = sheetWith("Flags", ref"A1" -> CellValue.Bool(true))
    val main = sheetWith("Main")
    val wb = workbookWith(main, flags)

    val result = main.evaluateFormula("=Flags!A1", workbook = Some(wb))
    assertEquals(result, Right(CellValue.Bool(true)))
  }

  test("SUM: cross-sheet range") {
    val sales = sheetWith(
      "Sales",
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"A2" -> CellValue.Number(BigDecimal(20)),
      ref"A3" -> CellValue.Number(BigDecimal(30))
    )
    val main = sheetWith("Main")
    val wb = workbookWith(main, sales)

    val result = main.evaluateFormula("=SUM(Sales!A1:A3)", workbook = Some(wb))
    assertEquals(result, Right(CellValue.Number(BigDecimal(60))))
  }

  // Note: COUNT, AVERAGE with cross-sheet ranges aren't implemented yet.
  // Cross-sheet aggregate support now uses TExpr.Aggregate with RangeLocation

  test("Aggregate: COUNT across cross-sheet range") {
    val data = sheetWith(
      "Data",
      ref"B1" -> CellValue.Number(BigDecimal(1)),
      ref"B2" -> CellValue.Number(BigDecimal(2)),
      ref"B3" -> CellValue.Text("skip"),
      ref"B4" -> CellValue.Number(BigDecimal(4))
    )
    val main = sheetWith("Main")
    val wb = workbookWith(main, data)

    val result = main.evaluateFormula("=COUNT(Data!B1:B4)", workbook = Some(wb))
    assertEquals(result, Right(CellValue.Number(BigDecimal(3))))
  }

  test("Aggregate: AVERAGE across cross-sheet range") {
    val nums = sheetWith(
      "Numbers",
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"A2" -> CellValue.Number(BigDecimal(20)),
      ref"A3" -> CellValue.Number(BigDecimal(30))
    )
    val main = sheetWith("Main")
    val wb = workbookWith(main, nums)

    val result = main.evaluateFormula("=AVERAGE(Numbers!A1:A3)", workbook = Some(wb))
    assertEquals(result, Right(CellValue.Number(BigDecimal(20))))
  }

  test("SUM: cross-sheet range handles empty cells") {
    val sparse = sheetWith(
      "Sparse",
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      // A2 is empty
      ref"A3" -> CellValue.Number(BigDecimal(30))
    )
    val main = sheetWith("Main")
    val wb = workbookWith(main, sparse)

    val result = main.evaluateFormula("=SUM(Sparse!A1:A3)", workbook = Some(wb))
    assertEquals(result, Right(CellValue.Number(BigDecimal(40)))) // Empty cells treated as 0
  }

  // MIN, MAX with cross-sheet ranges now use TExpr.Aggregate

  test("Aggregate: MIN across cross-sheet range") {
    val vals = sheetWith(
      "Values",
      ref"A1" -> CellValue.Number(BigDecimal(50)),
      ref"A2" -> CellValue.Number(BigDecimal(10)),
      ref"A3" -> CellValue.Number(BigDecimal(30))
    )
    val main = sheetWith("Main")
    val wb = workbookWith(main, vals)

    val result = main.evaluateFormula("=MIN(Values!A1:A3)", workbook = Some(wb))
    assertEquals(result, Right(CellValue.Number(BigDecimal(10))))
  }

  test("Aggregate: MAX across cross-sheet range") {
    val vals = sheetWith(
      "Values",
      ref"A1" -> CellValue.Number(BigDecimal(50)),
      ref"A2" -> CellValue.Number(BigDecimal(10)),
      ref"A3" -> CellValue.Number(BigDecimal(30))
    )
    val main = sheetWith("Main")
    val wb = workbookWith(main, vals)

    val result = main.evaluateFormula("=MAX(Values!A1:A3)", workbook = Some(wb))
    assertEquals(result, Right(CellValue.Number(BigDecimal(50))))
  }

  // ===== Error Case Tests =====

  test("cross-sheet ref without workbook context returns error") {
    val main = sheetWith("Main")

    val result = main.evaluateFormula("=Sales!A1") // No workbook provided
    assert(result.isLeft, s"Expected error but got: $result")
    result match
      case Left(err) =>
        assert(
          err.message.contains("workbook") || err.message.contains("context"),
          s"Error should mention workbook context: ${err.message}"
        )
      case _ => fail("Expected Left")
  }

  test("cross-sheet ref to nonexistent sheet returns error") {
    val main = sheetWith("Main")
    val wb = workbookWith(main)

    val result = main.evaluateFormula("=NonExistent!A1", workbook = Some(wb))
    assert(result.isLeft, s"Expected error but got: $result")
    result match
      case Left(err) =>
        assert(
          err.message.contains("not found") || err.message.contains("NonExistent"),
          s"Error should mention sheet not found: ${err.message}"
        )
      case _ => fail("Expected Left")
  }

  test("cross-sheet range without workbook context returns error") {
    val main = sheetWith("Main")

    val result = main.evaluateFormula("=SUM(Sales!A1:A10)") // No workbook provided
    assert(result.isLeft, s"Expected error but got: $result")
  }

  test("cross-sheet range to nonexistent sheet returns error") {
    val main = sheetWith("Main")
    val wb = workbookWith(main)

    val result = main.evaluateFormula("=SUM(Missing!A1:A10)", workbook = Some(wb))
    assert(result.isLeft, s"Expected error but got: $result")
  }

  // ===== Integration Tests =====

  test("cross-sheet formula in arithmetic: Sales!A1 + Sales!A2") {
    val sales = sheetWith(
      "Sales",
      ref"A1" -> CellValue.Number(BigDecimal(100)),
      ref"A2" -> CellValue.Number(BigDecimal(50))
    )
    val main = sheetWith("Main")
    val wb = workbookWith(main, sales)

    val result = main.evaluateFormula("=Sales!A1+Sales!A2", workbook = Some(wb))
    assertEquals(result, Right(CellValue.Number(BigDecimal(150))))
  }

  test("cross-sheet formula with multiplication: Sales!A1 * Revenue!B1") {
    val sales = sheetWith("Sales", ref"A1" -> CellValue.Number(BigDecimal(10)))
    val revenue = sheetWith("Revenue", ref"B1" -> CellValue.Number(BigDecimal(5)))
    val main = sheetWith("Main")
    val wb = workbookWith(main, sales, revenue)

    val result = main.evaluateFormula("=Sales!A1*Revenue!B1", workbook = Some(wb))
    assertEquals(result, Right(CellValue.Number(BigDecimal(50))))
  }

  test("cross-sheet formula in condition: IF(Sales!A1 > 50, ...)") {
    val sales = sheetWith("Sales", ref"A1" -> CellValue.Number(BigDecimal(100)))
    val main = sheetWith("Main")
    val wb = workbookWith(main, sales)

    val result = main.evaluateFormula("=IF(Sales!A1>50,\"High\",\"Low\")", workbook = Some(wb))
    assertEquals(result, Right(CellValue.Text("High")))
  }

  test("cross-sheet SUM in complex formula: SUM(Sales!A1:A10) * 2") {
    val sales = sheetWith(
      "Sales",
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"A2" -> CellValue.Number(BigDecimal(20)),
      ref"A3" -> CellValue.Number(BigDecimal(30))
    )
    val main = sheetWith("Main")
    val wb = workbookWith(main, sales)

    val result = main.evaluateFormula("=SUM(Sales!A1:A3)*2", workbook = Some(wb))
    assertEquals(result, Right(CellValue.Number(BigDecimal(120))))
  }

  test("mixed cross-sheet and local references") {
    val external = sheetWith("External", ref"A1" -> CellValue.Number(BigDecimal(100)))
    val main = sheetWith(
      "Main",
      ref"B1" -> CellValue.Number(BigDecimal(50))
    )
    val wb = workbookWith(main, external)

    val result = main.evaluateFormula("=External!A1+B1", workbook = Some(wb))
    assertEquals(result, Right(CellValue.Number(BigDecimal(150))))
  }

  // Note: Tests for quoted sheet names are skipped because parsing is not yet implemented.
  // See the ignored test above documenting this limitation.
  // TODO: TJC-360 - Add parser support for quoted sheet names in formulas

  test("cross-sheet reference with quoted sheet name (not yet supported)".ignore) {
    // When quoted sheet name parsing is implemented, this should work:
    val quarterly = sheetWith("Q1 Report", ref"A1" -> CellValue.Number(BigDecimal(1000)))
    val main = sheetWith("Main")
    val wb = workbookWith(main, quarterly)

    val result = main.evaluateFormula("='Q1 Report'!A1", workbook = Some(wb))
    assertEquals(result, Right(CellValue.Number(BigDecimal(1000))))
  }

  test("cross-sheet SUM with quoted sheet name (not yet supported)".ignore) {
    // When quoted sheet name parsing is implemented, this should work:
    val quarterly = sheetWith(
      "Q1 Report",
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"A2" -> CellValue.Number(BigDecimal(20))
    )
    val main = sheetWith("Main")
    val wb = workbookWith(main, quarterly)

    val result = main.evaluateFormula("=SUM('Q1 Report'!A1:A2)", workbook = Some(wb))
    assertEquals(result, Right(CellValue.Number(BigDecimal(30))))
  }

  // ===== evaluateWithDependencyCheck Cross-Sheet Tests =====

  test("evaluateWithDependencyCheck handles same-sheet formulas with workbook context") {
    // This tests that existing dependency checking still works when workbook is provided
    val sheet = sheetWith(
      "Main",
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"B1" -> CellValue.Formula("=A1+5"),
      ref"C1" -> CellValue.Formula("=B1*2")
    )
    val wb = workbookWith(sheet)

    val result = sheet.evaluateWithDependencyCheck(workbook = Some(wb))
    assert(result.isRight, s"Expected Right but got: $result")
    result.foreach { results =>
      assertEquals(results.get(ref"B1"), Some(CellValue.Number(BigDecimal(15))))
      assertEquals(results.get(ref"C1"), Some(CellValue.Number(BigDecimal(30))))
    }
  }

  // ===== Cross-Sheet Cycle Detection Tests =====

  test("DependencyGraph.fromWorkbook: extracts cross-sheet dependencies") {
    val sheet1 = sheetWith(
      "Sheet1",
      ref"A1" -> CellValue.Formula("=Sheet2!B1+10")
    )
    val sheet2 = sheetWith(
      "Sheet2",
      ref"B1" -> CellValue.Number(BigDecimal(5))
    )
    val wb = workbookWith(sheet1, sheet2)

    val graph = DependencyGraph.fromWorkbook(wb)

    // Sheet1!A1 should depend on Sheet2!B1
    val sheet1A1 = DependencyGraph.QualifiedRef(SheetName.unsafe("Sheet1"), ref"A1")
    val sheet2B1 = DependencyGraph.QualifiedRef(SheetName.unsafe("Sheet2"), ref"B1")

    assert(graph.contains(sheet1A1), "Graph should contain Sheet1!A1")
    assert(graph(sheet1A1).contains(sheet2B1), "Sheet1!A1 should depend on Sheet2!B1")
  }

  test("DependencyGraph.detectCrossSheetCycles: no cycle in valid workbook") {
    val sheet1 = sheetWith(
      "Sheet1",
      ref"A1" -> CellValue.Formula("=Sheet2!B1+10")
    )
    val sheet2 = sheetWith(
      "Sheet2",
      ref"B1" -> CellValue.Number(BigDecimal(5))
    )
    val wb = workbookWith(sheet1, sheet2)

    val graph = DependencyGraph.fromWorkbook(wb)
    val result = DependencyGraph.detectCrossSheetCycles(graph)

    assert(result.isRight, s"Expected no cycle but got: $result")
  }

  test("DependencyGraph.detectCrossSheetCycles: detects cross-sheet cycle") {
    // Sheet1!A1 = Sheet2!B1, Sheet2!B1 = Sheet1!A1 → cycle
    val sheet1 = sheetWith(
      "Sheet1",
      ref"A1" -> CellValue.Formula("=Sheet2!B1")
    )
    val sheet2 = sheetWith(
      "Sheet2",
      ref"B1" -> CellValue.Formula("=Sheet1!A1")
    )
    val wb = workbookWith(sheet1, sheet2)

    val graph = DependencyGraph.fromWorkbook(wb)
    val result = DependencyGraph.detectCrossSheetCycles(graph)

    assert(result.isLeft, s"Expected cycle error but got: $result")
  }

  test("DependencyGraph.detectCrossSheetCycles: detects longer cross-sheet cycle") {
    // Sheet1!A1 → Sheet2!B1 → Sheet3!C1 → Sheet1!A1
    val sheet1 = sheetWith(
      "Sheet1",
      ref"A1" -> CellValue.Formula("=Sheet2!B1")
    )
    val sheet2 = sheetWith(
      "Sheet2",
      ref"B1" -> CellValue.Formula("=Sheet3!C1")
    )
    val sheet3 = sheetWith(
      "Sheet3",
      ref"C1" -> CellValue.Formula("=Sheet1!A1")
    )
    val wb = workbookWith(sheet1, sheet2, sheet3)

    val graph = DependencyGraph.fromWorkbook(wb)
    val result = DependencyGraph.detectCrossSheetCycles(graph)

    assert(result.isLeft, s"Expected cycle error but got: $result")
  }

  test("DependencyGraph.detectCrossSheetCycles: detects self-reference") {
    val sheet1 = sheetWith(
      "Sheet1",
      ref"A1" -> CellValue.Formula("=Sheet1!A1")
    )
    val wb = workbookWith(sheet1)

    val graph = DependencyGraph.fromWorkbook(wb)
    val result = DependencyGraph.detectCrossSheetCycles(graph)

    assert(result.isLeft, s"Expected cycle error but got: $result")
  }

  test("DependencyGraph.fromWorkbook: handles mixed same-sheet and cross-sheet deps") {
    val sheet1 = sheetWith(
      "Sheet1",
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"A2" -> CellValue.Formula("=A1+Sheet2!B1") // Depends on both A1 and Sheet2!B1
    )
    val sheet2 = sheetWith(
      "Sheet2",
      ref"B1" -> CellValue.Number(BigDecimal(5))
    )
    val wb = workbookWith(sheet1, sheet2)

    val graph = DependencyGraph.fromWorkbook(wb)

    val sheet1A2 = DependencyGraph.QualifiedRef(SheetName.unsafe("Sheet1"), ref"A2")
    val sheet1A1 = DependencyGraph.QualifiedRef(SheetName.unsafe("Sheet1"), ref"A1")
    val sheet2B1 = DependencyGraph.QualifiedRef(SheetName.unsafe("Sheet2"), ref"B1")

    assert(graph.contains(sheet1A2), "Graph should contain Sheet1!A2")
    val deps = graph(sheet1A2)
    assert(deps.contains(sheet1A1), "Sheet1!A2 should depend on Sheet1!A1")
    assert(deps.contains(sheet2B1), "Sheet1!A2 should depend on Sheet2!B1")
  }

  // ===== VLOOKUP Cross-Sheet Tests (TJC-352) =====

  test("VLOOKUP: cross-sheet reference with exact match") {
    val lookup = sheetWith(
      "Lookup",
      ref"A1" -> CellValue.Number(BigDecimal(100)),
      ref"B1" -> CellValue.Text("Result1"),
      ref"A2" -> CellValue.Number(BigDecimal(200)),
      ref"B2" -> CellValue.Text("Result2"),
      ref"A3" -> CellValue.Number(BigDecimal(300)),
      ref"B3" -> CellValue.Text("Result3")
    )
    val calc = sheetWith(
      "Calc",
      ref"A1" -> CellValue.Number(BigDecimal(200)) // Lookup value
    )
    val wb = workbookWith(calc, lookup)

    val result = calc.evaluateFormula("=VLOOKUP(A1,Lookup!A1:B3,2,FALSE)", workbook = Some(wb))
    assertEquals(result, Right(CellValue.Text("Result2")))
  }

  test("VLOOKUP: cross-sheet reference with approximate match") {
    val lookup = sheetWith(
      "Lookup",
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"B1" -> CellValue.Text("Small"),
      ref"A2" -> CellValue.Number(BigDecimal(50)),
      ref"B2" -> CellValue.Text("Medium"),
      ref"A3" -> CellValue.Number(BigDecimal(100)),
      ref"B3" -> CellValue.Text("Large")
    )
    val calc = sheetWith(
      "Calc",
      ref"A1" -> CellValue.Number(BigDecimal(75)) // Between 50 and 100
    )
    val wb = workbookWith(calc, lookup)

    // Approximate match (TRUE) should find 50 (largest <= 75)
    val result = calc.evaluateFormula("=VLOOKUP(A1,Lookup!A1:B3,2,TRUE)", workbook = Some(wb))
    assertEquals(result, Right(CellValue.Text("Medium")))
  }

  test("VLOOKUP: cross-sheet reference with text lookup") {
    val lookup = sheetWith(
      "Products",
      ref"A1" -> CellValue.Text("Apple"),
      ref"B1" -> CellValue.Number(BigDecimal("1.50")),
      ref"A2" -> CellValue.Text("Banana"),
      ref"B2" -> CellValue.Number(BigDecimal("0.75")),
      ref"A3" -> CellValue.Text("Cherry"),
      ref"B3" -> CellValue.Number(BigDecimal("3.00"))
    )
    val order = sheetWith(
      "Order",
      ref"A1" -> CellValue.Text("Banana")
    )
    val wb = workbookWith(order, lookup)

    val result = order.evaluateFormula("=VLOOKUP(A1,Products!A1:B3,2,FALSE)", workbook = Some(wb))
    assertEquals(result, Right(CellValue.Number(BigDecimal("0.75"))))
  }

  test("VLOOKUP: cross-sheet reference not found returns error") {
    val lookup = sheetWith(
      "Lookup",
      ref"A1" -> CellValue.Number(BigDecimal(100)),
      ref"B1" -> CellValue.Text("Found")
    )
    val calc = sheetWith(
      "Calc",
      ref"A1" -> CellValue.Number(BigDecimal(999)) // Not in lookup table
    )
    val wb = workbookWith(calc, lookup)

    val result = calc.evaluateFormula("=VLOOKUP(A1,Lookup!A1:B1,2,FALSE)", workbook = Some(wb))
    // VLOOKUP returns error when value not found (either Left or Right(Error))
    result match
      case Left(err) =>
        assert(
          err.message.contains("not found") || err.message.contains("VLOOKUP"),
          s"Error should mention lookup failure: ${err.message}"
        )
      case Right(CellValue.Error(_)) =>
        () // Also acceptable
      case other =>
        fail(s"Expected error for not-found lookup, got: $other")
  }

  // ===== GH-161: Negative Test Cases =====

  test("GH-161: detects circular cross-sheet reference at runtime") {
    // Create circular reference: Sheet1!A1 → Sheet2!B1 → Sheet1!A1
    // Both cells have formulas WITHOUT cached values, so they trigger recursive evaluation
    val sheet1 = sheetWith(
      "Sheet1",
      ref"A1" -> CellValue.Formula("Sheet2!B1", None) // No cache - triggers recursive eval
    )
    val sheet2 = sheetWith(
      "Sheet2",
      ref"B1" -> CellValue.Formula("Sheet1!A1", None) // No cache - triggers recursive eval
    )
    val wb = workbookWith(sheet1, sheet2)

    val result = sheet1.evaluateFormula("=Sheet2!B1", workbook = Some(wb))
    assert(result.isLeft, s"Expected error for circular reference but got: $result")
    result match
      case Left(err) =>
        assert(
          err.message.contains("recursion") || err.message.contains("circular"),
          s"Error should mention recursion or circular reference: ${err.message}"
        )
      case _ => fail("Expected Left with circular reference error")
  }

  test("GH-161: handles parse error in referenced formula") {
    // Cross-sheet reference points to a cell with an invalid formula
    val data = sheetWith(
      "Data",
      ref"A1" -> CellValue.Formula("INVALID((", None) // Malformed formula, no cache
    )
    val main = sheetWith("Main")
    val wb = workbookWith(main, data)

    val result = main.evaluateFormula("=Data!A1", workbook = Some(wb))
    assert(result.isLeft, s"Expected error for malformed formula but got: $result")
    result match
      case Left(err) =>
        assert(
          err.message.contains("parse") || err.message.contains("Failed"),
          s"Error should mention parse failure: ${err.message}"
        )
      case _ => fail("Expected Left with parse error")
  }

  // ===== GH-187: SUM/Aggregate of uncached formula cells =====

  test("GH-187: SUM of uncached formula cells") {
    val sales = sheetWith(
      "Sales",
      ref"A1" -> CellValue.Formula("10", None),
      ref"A2" -> CellValue.Formula("20", None),
      ref"A3" -> CellValue.Formula("30", None)
    )
    val main = sheetWith("Main")
    val wb = workbookWith(main, sales)

    val result = main.evaluateFormula("=SUM(Sales!A1:A3)", workbook = Some(wb))
    assertEquals(result, Right(CellValue.Number(BigDecimal(60))))
  }

  test("GH-187: SUM of mixed cached and uncached formula cells") {
    val data = sheetWith(
      "Data",
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"A2" -> CellValue.Formula("20", None), // Uncached
      ref"A3" -> CellValue.Formula("=30", Some(CellValue.Number(BigDecimal(30)))), // Cached
      ref"A4" -> CellValue.Formula("40", None) // Uncached
    )
    val main = sheetWith("Main")
    val wb = workbookWith(main, data)

    val result = main.evaluateFormula("=SUM(Data!A1:A4)", workbook = Some(wb))
    assertEquals(result, Right(CellValue.Number(BigDecimal(100))))
  }

  test("GH-187: AVERAGE of uncached formula cells") {
    val vals = sheetWith(
      "Values",
      ref"A1" -> CellValue.Formula("10", None),
      ref"A2" -> CellValue.Formula("20", None),
      ref"A3" -> CellValue.Formula("30", None)
    )
    val main = sheetWith("Main")
    val wb = workbookWith(main, vals)

    val result = main.evaluateFormula("=AVERAGE(Values!A1:A3)", workbook = Some(wb))
    assertEquals(result, Right(CellValue.Number(BigDecimal(20))))
  }

  test("GH-187: cross-sheet SUM with nested formula references") {
    // A1 contains formula referencing B1, B1 contains uncached formula
    val data = sheetWith(
      "Data",
      ref"A1" -> CellValue.Formula("B1*2", None), // References B1
      ref"A2" -> CellValue.Formula("B2*2", None), // References B2
      ref"B1" -> CellValue.Number(BigDecimal(5)),
      ref"B2" -> CellValue.Number(BigDecimal(10))
    )
    val main = sheetWith("Main")
    val wb = workbookWith(main, data)

    // SUM(A1:A2) = (5*2) + (10*2) = 10 + 20 = 30
    val result = main.evaluateFormula("=SUM(Data!A1:A2)", workbook = Some(wb))
    assertEquals(result, Right(CellValue.Number(BigDecimal(30))))
  }

  test("GH-187: same-sheet SUM of uncached formula cells") {
    // Also verify the fix works for same-sheet aggregates
    val sheet = sheetWith(
      "Main",
      ref"A1" -> CellValue.Formula("10", None),
      ref"A2" -> CellValue.Formula("20", None),
      ref"A3" -> CellValue.Formula("30", None)
    )
    val wb = workbookWith(sheet)

    val result = sheet.evaluateFormula("=SUM(A1:A3)", workbook = Some(wb))
    assertEquals(result, Right(CellValue.Number(BigDecimal(60))))
  }

  // ===== GH-187: Conditional Aggregate Functions with Uncached Formulas =====

  test("GH-187: SUMIF with uncached formula in sum range") {
    val sheet = sheetWith(
      "Data",
      ref"A1" -> CellValue.Number(BigDecimal(100)),
      ref"A2" -> CellValue.Number(BigDecimal(50)),
      ref"B1" -> CellValue.Formula("10", None),
      ref"B2" -> CellValue.Formula("20", None)
    )
    val wb = workbookWith(sheet)

    // Sum B column where A > 60 (only A1=100 matches, so sum B1=10)
    val result = sheet.evaluateFormula("=SUMIF(A1:A2, \">60\", B1:B2)", workbook = Some(wb))
    assertEquals(result, Right(CellValue.Number(BigDecimal(10))))
  }

  test("GH-187: SUMIF with uncached formula in criteria range") {
    val sheet = sheetWith(
      "Data",
      ref"A1" -> CellValue.Formula("100", None),
      ref"A2" -> CellValue.Formula("50", None),
      ref"B1" -> CellValue.Number(BigDecimal(10)),
      ref"B2" -> CellValue.Number(BigDecimal(20))
    )
    val wb = workbookWith(sheet)

    // Sum B column where A > 60 (A1=100 matches, so sum B1=10)
    val result = sheet.evaluateFormula("=SUMIF(A1:A2, \">60\", B1:B2)", workbook = Some(wb))
    assertEquals(result, Right(CellValue.Number(BigDecimal(10))))
  }

  test("GH-187: SUMIFS with uncached formula in sum range") {
    val sheet = sheetWith(
      "Data",
      ref"A1" -> CellValue.Number(BigDecimal(100)),
      ref"A2" -> CellValue.Number(BigDecimal(200)),
      ref"B1" -> CellValue.Text("East"),
      ref"B2" -> CellValue.Text("West"),
      ref"C1" -> CellValue.Formula("10", None),
      ref"C2" -> CellValue.Formula("20", None)
    )
    val wb = workbookWith(sheet)

    // Sum C where A > 50 and B = "East" (only row 1 matches)
    val result =
      sheet.evaluateFormula("=SUMIFS(C1:C2, A1:A2, \">50\", B1:B2, \"East\")", workbook = Some(wb))
    assertEquals(result, Right(CellValue.Number(BigDecimal(10))))
  }

  test("GH-187: AVERAGEIF with uncached formula in average range") {
    val sheet = sheetWith(
      "Data",
      ref"A1" -> CellValue.Number(BigDecimal(100)),
      ref"A2" -> CellValue.Number(BigDecimal(50)),
      ref"A3" -> CellValue.Number(BigDecimal(200)),
      ref"B1" -> CellValue.Formula("10", None),
      ref"B2" -> CellValue.Formula("20", None),
      ref"B3" -> CellValue.Formula("30", None)
    )
    val wb = workbookWith(sheet)

    // Average B where A > 60 (A1=100, A3=200 match, so average B1=10, B3=30 = 20)
    val result = sheet.evaluateFormula("=AVERAGEIF(A1:A3, \">60\", B1:B3)", workbook = Some(wb))
    assertEquals(result, Right(CellValue.Number(BigDecimal(20))))
  }

  test("GH-187: AVERAGEIFS with uncached formula in average range") {
    val sheet = sheetWith(
      "Data",
      ref"A1" -> CellValue.Number(BigDecimal(100)),
      ref"A2" -> CellValue.Number(BigDecimal(200)),
      ref"A3" -> CellValue.Number(BigDecimal(100)),
      ref"B1" -> CellValue.Text("East"),
      ref"B2" -> CellValue.Text("West"),
      ref"B3" -> CellValue.Text("East"),
      ref"C1" -> CellValue.Formula("10", None),
      ref"C2" -> CellValue.Formula("20", None),
      ref"C3" -> CellValue.Formula("30", None)
    )
    val wb = workbookWith(sheet)

    // Average C where A >= 100 and B = "East" (rows 1 and 3 match, avg 10+30 / 2 = 20)
    val result = sheet.evaluateFormula(
      "=AVERAGEIFS(C1:C3, A1:A3, \">=100\", B1:B3, \"East\")",
      workbook = Some(wb)
    )
    assertEquals(result, Right(CellValue.Number(BigDecimal(20))))
  }

  test("GH-187: SUMPRODUCT with uncached formula cells") {
    val sheet = sheetWith(
      "Data",
      ref"A1" -> CellValue.Formula("2", None),
      ref"A2" -> CellValue.Formula("3", None),
      ref"B1" -> CellValue.Number(BigDecimal(10)),
      ref"B2" -> CellValue.Number(BigDecimal(20))
    )
    val wb = workbookWith(sheet)

    // 2*10 + 3*20 = 20 + 60 = 80
    val result = sheet.evaluateFormula("=SUMPRODUCT(A1:A2, B1:B2)", workbook = Some(wb))
    assertEquals(result, Right(CellValue.Number(BigDecimal(80))))
  }

  test("GH-187: SUMPRODUCT with all uncached formulas") {
    val sheet = sheetWith(
      "Data",
      ref"A1" -> CellValue.Formula("2", None),
      ref"A2" -> CellValue.Formula("3", None),
      ref"B1" -> CellValue.Formula("10", None),
      ref"B2" -> CellValue.Formula("20", None)
    )
    val wb = workbookWith(sheet)

    // 2*10 + 3*20 = 20 + 60 = 80
    val result = sheet.evaluateFormula("=SUMPRODUCT(A1:A2, B1:B2)", workbook = Some(wb))
    assertEquals(result, Right(CellValue.Number(BigDecimal(80))))
  }

  test("GH-187: SUMPRODUCT with expression formulas") {
    val sheet = sheetWith(
      "Data",
      ref"A1" -> CellValue.Formula("1+1", None), // = 2
      ref"A2" -> CellValue.Formula("1+2", None), // = 3
      ref"B1" -> CellValue.Number(BigDecimal(10)),
      ref"B2" -> CellValue.Number(BigDecimal(20))
    )
    val wb = workbookWith(sheet)

    // 2*10 + 3*20 = 20 + 60 = 80
    val result = sheet.evaluateFormula("=SUMPRODUCT(A1:A2, B1:B2)", workbook = Some(wb))
    assertEquals(result, Right(CellValue.Number(BigDecimal(80))))
  }

  test("GH-187: COUNTIF with uncached formula in criteria range") {
    val sheet = sheetWith(
      "Data",
      ref"A1" -> CellValue.Formula("100", None),
      ref"A2" -> CellValue.Formula("50", None),
      ref"A3" -> CellValue.Formula("200", None)
    )
    val wb = workbookWith(sheet)

    // Count where A > 60 (A1=100, A3=200 match, so count = 2)
    val result = sheet.evaluateFormula("=COUNTIF(A1:A3, \">60\")", workbook = Some(wb))
    assertEquals(result, Right(CellValue.Number(BigDecimal(2))))
  }

  test("GH-187: COUNTIFS with uncached formula in criteria ranges") {
    val sheet = sheetWith(
      "Data",
      ref"A1" -> CellValue.Formula("100", None),
      ref"A2" -> CellValue.Formula("200", None),
      ref"A3" -> CellValue.Formula("100", None),
      ref"B1" -> CellValue.Text("East"),
      ref"B2" -> CellValue.Text("West"),
      ref"B3" -> CellValue.Text("East")
    )
    val wb = workbookWith(sheet)

    // Count where A >= 100 and B = "East" (rows 1 and 3 match, so count = 2)
    val result = sheet.evaluateFormula(
      "=COUNTIFS(A1:A3, \">=100\", B1:B3, \"East\")",
      workbook = Some(wb)
    )
    assertEquals(result, Right(CellValue.Number(BigDecimal(2))))
  }

  // ===== GH-192: Cross-Sheet Full-Column References in Conditional Aggregates =====

  test("GH-192: SUMIFS with cross-sheet ranges") {
    val data = sheetWith(
      "Data",
      ref"A1" -> CellValue.Text("1001-A"),
      ref"B1" -> CellValue.Text("Screw"),
      ref"C1" -> CellValue.Number(BigDecimal(20)),
      ref"A2" -> CellValue.Text("1001-A"),
      ref"B2" -> CellValue.Text("Screw"),
      ref"C2" -> CellValue.Number(BigDecimal(50)),
      ref"A3" -> CellValue.Text("1002-B"),
      ref"B3" -> CellValue.Text("Bolt"),
      ref"C3" -> CellValue.Number(BigDecimal(100))
    )
    val main = sheetWith("Main")
    val wb = workbookWith(main, data)

    // Sum C where A = "1001-A" and B = "Screw" -> 20 + 50 = 70
    val result = main.evaluateFormula(
      "=SUMIFS(Data!C1:C3, Data!A1:A3, \"1001-A\", Data!B1:B3, \"Screw\")",
      workbook = Some(wb)
    )
    assertEquals(result, Right(CellValue.Number(BigDecimal(70))))
  }

  test("GH-192: SUMIFS with cross-sheet full-column references") {
    val data = sheetWith(
      "Data",
      ref"A1" -> CellValue.Text("1001-A"),
      ref"B1" -> CellValue.Text("Screw"),
      ref"C1" -> CellValue.Number(BigDecimal(20)),
      ref"A2" -> CellValue.Text("1001-A"),
      ref"B2" -> CellValue.Text("Screw"),
      ref"C2" -> CellValue.Number(BigDecimal(50)),
      ref"A3" -> CellValue.Text("1002-B"),
      ref"B3" -> CellValue.Text("Bolt"),
      ref"C3" -> CellValue.Number(BigDecimal(100))
    )
    val main = sheetWith("Main")
    val wb = workbookWith(main, data)

    // Sum C:C where A:A = "1001-A" and B:B = "Screw" -> 20 + 50 = 70
    val result = main.evaluateFormula(
      "=SUMIFS(Data!C:C, Data!A:A, \"1001-A\", Data!B:B, \"Screw\")",
      workbook = Some(wb)
    )
    assertEquals(result, Right(CellValue.Number(BigDecimal(70))))
  }

  test("GH-192: COUNTIFS with cross-sheet ranges") {
    val data = sheetWith(
      "Data",
      ref"A1" -> CellValue.Text("Apple"),
      ref"B1" -> CellValue.Text("Red"),
      ref"A2" -> CellValue.Text("Apple"),
      ref"B2" -> CellValue.Text("Green"),
      ref"A3" -> CellValue.Text("Banana"),
      ref"B3" -> CellValue.Text("Yellow")
    )
    val main = sheetWith("Main")
    val wb = workbookWith(main, data)

    // Count where A = "Apple" -> 2
    val result = main.evaluateFormula(
      "=COUNTIFS(Data!A1:A3, \"Apple\")",
      workbook = Some(wb)
    )
    assertEquals(result, Right(CellValue.Number(BigDecimal(2))))
  }

  test("GH-192: SUMIF with cross-sheet ranges") {
    val data = sheetWith(
      "Data",
      ref"A1" -> CellValue.Text("Apple"),
      ref"B1" -> CellValue.Number(BigDecimal(10)),
      ref"A2" -> CellValue.Text("Banana"),
      ref"B2" -> CellValue.Number(BigDecimal(20)),
      ref"A3" -> CellValue.Text("Apple"),
      ref"B3" -> CellValue.Number(BigDecimal(30))
    )
    val main = sheetWith("Main")
    val wb = workbookWith(main, data)

    // Sum B where A = "Apple" -> 10 + 30 = 40
    val result = main.evaluateFormula(
      "=SUMIF(Data!A1:A3, \"Apple\", Data!B1:B3)",
      workbook = Some(wb)
    )
    assertEquals(result, Right(CellValue.Number(BigDecimal(40))))
  }

  test("GH-192: COUNTIF with cross-sheet range") {
    val data = sheetWith(
      "Data",
      ref"A1" -> CellValue.Number(BigDecimal(100)),
      ref"A2" -> CellValue.Number(BigDecimal(50)),
      ref"A3" -> CellValue.Number(BigDecimal(200))
    )
    val main = sheetWith("Main")
    val wb = workbookWith(main, data)

    // Count where A > 60 -> 2
    val result = main.evaluateFormula(
      "=COUNTIF(Data!A1:A3, \">60\")",
      workbook = Some(wb)
    )
    assertEquals(result, Right(CellValue.Number(BigDecimal(2))))
  }

  test("GH-192: AVERAGEIF with cross-sheet ranges") {
    val data = sheetWith(
      "Data",
      ref"A1" -> CellValue.Text("Apple"),
      ref"B1" -> CellValue.Number(BigDecimal(10)),
      ref"A2" -> CellValue.Text("Banana"),
      ref"B2" -> CellValue.Number(BigDecimal(20)),
      ref"A3" -> CellValue.Text("Apple"),
      ref"B3" -> CellValue.Number(BigDecimal(30))
    )
    val main = sheetWith("Main")
    val wb = workbookWith(main, data)

    // Average B where A = "Apple" -> (10 + 30) / 2 = 20
    val result = main.evaluateFormula(
      "=AVERAGEIF(Data!A1:A3, \"Apple\", Data!B1:B3)",
      workbook = Some(wb)
    )
    assertEquals(result, Right(CellValue.Number(BigDecimal(20))))
  }

  test("GH-192: AVERAGEIFS with cross-sheet ranges") {
    val data = sheetWith(
      "Data",
      ref"A1" -> CellValue.Number(BigDecimal(100)),
      ref"B1" -> CellValue.Text("East"),
      ref"C1" -> CellValue.Number(BigDecimal(10)),
      ref"A2" -> CellValue.Number(BigDecimal(200)),
      ref"B2" -> CellValue.Text("West"),
      ref"C2" -> CellValue.Number(BigDecimal(20)),
      ref"A3" -> CellValue.Number(BigDecimal(100)),
      ref"B3" -> CellValue.Text("East"),
      ref"C3" -> CellValue.Number(BigDecimal(30))
    )
    val main = sheetWith("Main")
    val wb = workbookWith(main, data)

    // Average C where A >= 100 and B = "East" -> (10 + 30) / 2 = 20
    val result = main.evaluateFormula(
      "=AVERAGEIFS(Data!C1:C3, Data!A1:A3, \">=100\", Data!B1:B3, \"East\")",
      workbook = Some(wb)
    )
    assertEquals(result, Right(CellValue.Number(BigDecimal(20))))
  }

  test("GH-192: SUMPRODUCT with cross-sheet ranges") {
    val data = sheetWith(
      "Data",
      ref"A1" -> CellValue.Number(BigDecimal(2)),
      ref"B1" -> CellValue.Number(BigDecimal(10)),
      ref"A2" -> CellValue.Number(BigDecimal(3)),
      ref"B2" -> CellValue.Number(BigDecimal(20))
    )
    val main = sheetWith("Main")
    val wb = workbookWith(main, data)

    // 2*10 + 3*20 = 20 + 60 = 80
    val result = main.evaluateFormula(
      "=SUMPRODUCT(Data!A1:A2, Data!B1:B2)",
      workbook = Some(wb)
    )
    assertEquals(result, Right(CellValue.Number(BigDecimal(80))))
  }

  test("GH-192: mixed local and cross-sheet references in SUMIFS") {
    val data = sheetWith(
      "Data",
      ref"A1" -> CellValue.Text("1001-A"),
      ref"A2" -> CellValue.Text("1001-A"),
      ref"A3" -> CellValue.Text("1002-B")
    )
    val main = sheetWith(
      "Main",
      ref"B1" -> CellValue.Number(BigDecimal(20)),
      ref"B2" -> CellValue.Number(BigDecimal(50)),
      ref"B3" -> CellValue.Number(BigDecimal(100))
    )
    val wb = workbookWith(main, data)

    // Sum local B where cross-sheet A = "1001-A" -> 20 + 50 = 70
    val result = main.evaluateFormula(
      "=SUMIFS(B1:B3, Data!A1:A3, \"1001-A\")",
      workbook = Some(wb)
    )
    assertEquals(result, Right(CellValue.Number(BigDecimal(70))))
  }

  // ===== GH-192: Full-Column Optimization Tests =====

  test("GH-192: SUMIF with full-column references (same sheet)") {
    val sheet = sheetWith(
      "Data",
      ref"A1" -> CellValue.Text("Apple"),
      ref"A2" -> CellValue.Text("Banana"),
      ref"A3" -> CellValue.Text("Apple"),
      ref"B1" -> CellValue.Number(BigDecimal(10)),
      ref"B2" -> CellValue.Number(BigDecimal(20)),
      ref"B3" -> CellValue.Number(BigDecimal(30))
    )
    val wb = workbookWith(sheet)

    // Sum B:B where A:A = "Apple" -> 10 + 30 = 40
    val result = sheet.evaluateFormula(
      "=SUMIF(A:A, \"Apple\", B:B)",
      workbook = Some(wb)
    )
    assertEquals(result, Right(CellValue.Number(BigDecimal(40))))
  }

  test("GH-192: COUNTIF with full-column reference (same sheet)") {
    val sheet = sheetWith(
      "Data",
      ref"A1" -> CellValue.Number(BigDecimal(100)),
      ref"A2" -> CellValue.Number(BigDecimal(50)),
      ref"A3" -> CellValue.Number(BigDecimal(200))
    )
    val wb = workbookWith(sheet)

    // Count A:A where > 60 -> 2
    val result = sheet.evaluateFormula(
      "=COUNTIF(A:A, \">60\")",
      workbook = Some(wb)
    )
    assertEquals(result, Right(CellValue.Number(BigDecimal(2))))
  }

  test("GH-192: COUNTIFS with cross-sheet full-column references") {
    val data = sheetWith(
      "Data",
      ref"A1" -> CellValue.Text("Apple"),
      ref"B1" -> CellValue.Text("Red"),
      ref"A2" -> CellValue.Text("Apple"),
      ref"B2" -> CellValue.Text("Green"),
      ref"A3" -> CellValue.Text("Banana"),
      ref"B3" -> CellValue.Text("Yellow")
    )
    val main = sheetWith("Main")
    val wb = workbookWith(main, data)

    // Count where A:A = "Apple" and B:B = "Red" -> 1
    val result = main.evaluateFormula(
      "=COUNTIFS(Data!A:A, \"Apple\", Data!B:B, \"Red\")",
      workbook = Some(wb)
    )
    assertEquals(result, Right(CellValue.Number(BigDecimal(1))))
  }

  test("GH-192: AVERAGEIF with full-column references (same sheet)") {
    val sheet = sheetWith(
      "Data",
      ref"A1" -> CellValue.Text("Apple"),
      ref"A2" -> CellValue.Text("Banana"),
      ref"A3" -> CellValue.Text("Apple"),
      ref"B1" -> CellValue.Number(BigDecimal(10)),
      ref"B2" -> CellValue.Number(BigDecimal(20)),
      ref"B3" -> CellValue.Number(BigDecimal(30))
    )
    val wb = workbookWith(sheet)

    // Average B:B where A:A = "Apple" -> (10 + 30) / 2 = 20
    val result = sheet.evaluateFormula(
      "=AVERAGEIF(A:A, \"Apple\", B:B)",
      workbook = Some(wb)
    )
    assertEquals(result, Right(CellValue.Number(BigDecimal(20))))
  }

  test("GH-192: AVERAGEIFS with cross-sheet full-column references") {
    val data = sheetWith(
      "Data",
      ref"A1" -> CellValue.Number(BigDecimal(100)),
      ref"B1" -> CellValue.Text("East"),
      ref"C1" -> CellValue.Number(BigDecimal(10)),
      ref"A2" -> CellValue.Number(BigDecimal(200)),
      ref"B2" -> CellValue.Text("West"),
      ref"C2" -> CellValue.Number(BigDecimal(20)),
      ref"A3" -> CellValue.Number(BigDecimal(100)),
      ref"B3" -> CellValue.Text("East"),
      ref"C3" -> CellValue.Number(BigDecimal(30))
    )
    val main = sheetWith("Main")
    val wb = workbookWith(main, data)

    // Average C:C where A:A >= 100 and B:B = "East" -> (10 + 30) / 2 = 20
    val result = main.evaluateFormula(
      "=AVERAGEIFS(Data!C:C, Data!A:A, \">=100\", Data!B:B, \"East\")",
      workbook = Some(wb)
    )
    assertEquals(result, Right(CellValue.Number(BigDecimal(20))))
  }

  test("GH-192: SUMPRODUCT with full-column references") {
    val sheet = sheetWith(
      "Data",
      ref"A1" -> CellValue.Number(BigDecimal(2)),
      ref"B1" -> CellValue.Number(BigDecimal(10)),
      ref"A2" -> CellValue.Number(BigDecimal(3)),
      ref"B2" -> CellValue.Number(BigDecimal(20))
    )
    val wb = workbookWith(sheet)

    // 2*10 + 3*20 = 20 + 60 = 80
    val result = sheet.evaluateFormula(
      "=SUMPRODUCT(A:A, B:B)",
      workbook = Some(wb)
    )
    assertEquals(result, Right(CellValue.Number(BigDecimal(80))))
  }

  test("GH-192: SUM with full-column reference") {
    val sheet = sheetWith(
      "Data",
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"A2" -> CellValue.Number(BigDecimal(20)),
      ref"A3" -> CellValue.Number(BigDecimal(30))
    )
    val wb = workbookWith(sheet)

    // SUM(A:A) -> 60
    val result = sheet.evaluateFormula("=SUM(A:A)", workbook = Some(wb))
    assertEquals(result, Right(CellValue.Number(BigDecimal(60))))
  }

  test("GH-192: AVERAGE with cross-sheet full-column reference") {
    val data = sheetWith(
      "Data",
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"A2" -> CellValue.Number(BigDecimal(20)),
      ref"A3" -> CellValue.Number(BigDecimal(30))
    )
    val main = sheetWith("Main")
    val wb = workbookWith(main, data)

    // AVERAGE(Data!A:A) -> 20
    val result = main.evaluateFormula("=AVERAGE(Data!A:A)", workbook = Some(wb))
    assertEquals(result, Right(CellValue.Number(BigDecimal(20))))
  }

  test("GH-192: COUNT with full-column reference") {
    val sheet = sheetWith(
      "Data",
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"A2" -> CellValue.Text("text"),
      ref"A3" -> CellValue.Number(BigDecimal(30))
    )
    val wb = workbookWith(sheet)

    // COUNT(A:A) -> 2 (only numbers)
    val result = sheet.evaluateFormula("=COUNT(A:A)", workbook = Some(wb))
    assertEquals(result, Right(CellValue.Number(BigDecimal(2))))
  }

  test("GH-192: empty sheet with full-column conditional returns 0") {
    val data = sheetWith("Data") // Empty sheet
    val main = sheetWith("Main")
    val wb = workbookWith(main, data)

    // SUMIF on empty sheet -> 0
    val result = main.evaluateFormula(
      "=SUMIF(Data!A:A, \"anything\", Data!B:B)",
      workbook = Some(wb)
    )
    assertEquals(result, Right(CellValue.Number(BigDecimal(0))))
  }

  test("GH-192: COUNTIF on empty sheet with full-column reference returns 0") {
    val data = sheetWith("Data") // Empty sheet
    val main = sheetWith("Main")
    val wb = workbookWith(main, data)

    val result = main.evaluateFormula(
      "=COUNTIF(Data!A:A, \">0\")",
      workbook = Some(wb)
    )
    assertEquals(result, Right(CellValue.Number(BigDecimal(0))))
  }
