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

  test("SheetFoldRange: SUM across cross-sheet range") {
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

  test("SheetFoldRange: handles empty cells in range") {
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
