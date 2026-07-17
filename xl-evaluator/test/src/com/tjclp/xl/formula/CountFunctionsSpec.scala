package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.addressing.SheetName
import munit.FunSuite

/**
 * Comprehensive tests for COUNTBLANK function.
 *
 * COUNTBLANK counts empty cells in a range. Cells with formulas that return "" are NOT counted as
 * blank (Excel behavior).
 */
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial"))
class CountFunctionsSpec extends FunSuite:
  val emptySheet = new Sheet(name = SheetName.unsafe("Test"))
  val evaluator = Evaluator.instance

  /** Helper to create sheet with cells */
  def sheetWith(cells: (ARef, CellValue)*): Sheet =
    cells.foldLeft(emptySheet) { case (s, (ref, value)) =>
      s.put(ref, value)
    }

  // ===== COUNTBLANK Tests =====

  test("COUNTBLANK: all cells empty") {
    val sheet = emptySheet
    val range = CellRange.parse("A1:A5").toOption.get
    val expr = TExpr.Aggregate("COUNTBLANK", TExpr.RangeLocation.Local(range))
    val result = evaluator.eval(expr, sheet)
    assertEquals(result, Right(BigDecimal(5)))
  }

  test("COUNTBLANK: no cells empty") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(1),
      ref"A2" -> CellValue.Number(2),
      ref"A3" -> CellValue.Number(3)
    )
    val range = CellRange.parse("A1:A3").toOption.get
    val expr = TExpr.Aggregate("COUNTBLANK", TExpr.RangeLocation.Local(range))
    val result = evaluator.eval(expr, sheet)
    assertEquals(result, Right(BigDecimal(0)))
  }

  test("COUNTBLANK: mixed cells - some empty") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(1),
      ref"A3" -> CellValue.Number(3),
      ref"A5" -> CellValue.Number(5)
    )
    val range = CellRange.parse("A1:A5").toOption.get
    val expr = TExpr.Aggregate("COUNTBLANK", TExpr.RangeLocation.Local(range))
    val result = evaluator.eval(expr, sheet)
    assertEquals(result, Right(BigDecimal(2))) // A2 and A4 are empty
  }

  test("COUNTBLANK: text cells are not blank") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Text(""),
      ref"A2" -> CellValue.Text("Hello"),
      ref"A3" -> CellValue.Number(0)
    )
    val range = CellRange.parse("A1:A4").toOption.get
    val expr = TExpr.Aggregate("COUNTBLANK", TExpr.RangeLocation.Local(range))
    val result = evaluator.eval(expr, sheet)
    // Note: Excel considers Text("") as blank, but we count only truly empty cells (no CellValue)
    // A4 is truly empty (no cell)
    assertEquals(result, Right(BigDecimal(1))) // Only A4 is truly empty
  }

  test("COUNTBLANK: formula cells are not blank") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Formula("=1+1", Some(CellValue.Number(2))),
      ref"A2" -> CellValue.Number(5)
    )
    val range = CellRange.parse("A1:A3").toOption.get
    val expr = TExpr.Aggregate("COUNTBLANK", TExpr.RangeLocation.Local(range))
    val result = evaluator.eval(expr, sheet)
    assertEquals(result, Right(BigDecimal(1))) // Only A3 is empty
  }

  test("COUNTBLANK: single cell empty") {
    val sheet = emptySheet
    val range = CellRange.parse("A1:A1").toOption.get
    val expr = TExpr.Aggregate("COUNTBLANK", TExpr.RangeLocation.Local(range))
    val result = evaluator.eval(expr, sheet)
    assertEquals(result, Right(BigDecimal(1)))
  }

  test("COUNTBLANK: single cell not empty") {
    val sheet = sheetWith(ref"A1" -> CellValue.Number(42))
    val range = CellRange.parse("A1:A1").toOption.get
    val expr = TExpr.Aggregate("COUNTBLANK", TExpr.RangeLocation.Local(range))
    val result = evaluator.eval(expr, sheet)
    assertEquals(result, Right(BigDecimal(0)))
  }

  test("COUNTBLANK: 2D range") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(1),
      ref"B2" -> CellValue.Number(2)
    )
    val range = CellRange.parse("A1:B2").toOption.get
    val expr = TExpr.Aggregate("COUNTBLANK", TExpr.RangeLocation.Local(range))
    val result = evaluator.eval(expr, sheet)
    // 4 cells total (A1, A2, B1, B2), 2 filled (A1, B2), 2 empty (A2, B1)
    assertEquals(result, Right(BigDecimal(2)))
  }

  // ===== GH-395: direct single-cell args triage like 1×1 ranges =====

  test("GH-395: =COUNT(A1, A2) with blank A2 is 1 (issue repro — Excel ignores the blank)") {
    val sheet = sheetWith(ref"A1" -> CellValue.Number(10))
    assertEquals(
      sheet.evaluateFormula("=COUNT(A1, A2)"),
      Right(CellValue.Number(BigDecimal(1)))
    )
  }

  test("GH-395: =COUNT over a text direct ref skips it like the range fold (not an error)") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(10),
      ref"A2" -> CellValue.Text("abc")
    )
    assertEquals(
      sheet.evaluateFormula("=COUNT(A1, A2)"),
      Right(CellValue.Number(BigDecimal(1)))
    )
    // range form agrees (the direct/range parity law)
    assertEquals(
      sheet.evaluateFormula("=COUNT(A1:A2)"),
      Right(CellValue.Number(BigDecimal(1)))
    )
  }

  test("GH-395: =COUNTA(A1, A2) with blank A2 counts only the filled cell") {
    val sheet = sheetWith(ref"A1" -> CellValue.Text("x"))
    assertEquals(
      sheet.evaluateFormula("=COUNTA(A1, A2)"),
      Right(CellValue.Number(BigDecimal(1)))
    )
  }

  test("GH-395: =COUNTBLANK(A1) with blank A1 stays 1 (the countsEmpty trap)") {
    val sheet = sheetWith(ref"Z9" -> CellValue.Text("unrelated"))
    assertEquals(
      sheet.evaluateFormula("=COUNTBLANK(A1)"),
      Right(CellValue.Number(BigDecimal(1)))
    )
    // and a FILLED direct ref is not blank
    val filled = sheetWith(ref"A1" -> CellValue.Number(5))
    assertEquals(
      filled.evaluateFormula("=COUNTBLANK(A1)"),
      Right(CellValue.Number(BigDecimal(0)))
    )
  }

  test("GH-395: =COUNTBLANK(5) is 0 — a literal is never blank") {
    assertEquals(
      emptySheet.evaluateFormula("=COUNTBLANK(5)"),
      Right(CellValue.Number(BigDecimal(0)))
    )
  }

  test("GH-395: direct-ref triage keeps cached and uncached formula cells counted") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Formula("=1+1", Some(CellValue.Number(2))),
      ref"A2" -> CellValue.Formula("=2+3", None)
    )
    assertEquals(
      sheet.evaluateFormula("=COUNT(A1, A2)"),
      Right(CellValue.Number(BigDecimal(2)))
    )
    assertEquals(
      sheet.evaluateFormula("=SUM(A1, A2)"),
      Right(CellValue.Number(BigDecimal(7)))
    )
  }

  test("GH-395: cross-sheet direct single-cell ref triages against the target sheet") {
    import com.tjclp.xl.workbooks.Workbook
    val data = Sheet(SheetName.unsafe("Data")).put(ref"B1", CellValue.Number(BigDecimal(7)))
    val main = Sheet(SheetName.unsafe("Main"))
    val wb = Workbook(Vector(main, data))
    // Data!B2 is blank: COUNT ignores it, SUM folds only the filled cell
    assertEquals(
      wb.evaluateFormula("=COUNT(Data!B1, Data!B2)", SheetName.unsafe("Main")),
      Right(CellValue.Number(BigDecimal(1)))
    )
    assertEquals(
      wb.evaluateFormula("=SUM(Data!B1, Data!B2)", SheetName.unsafe("Main")),
      Right(CellValue.Number(BigDecimal(7)))
    )
  }

  test("GH-395: error-valued direct ref still propagates through SUM and skips in COUNT") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(10),
      ref"A2" -> CellValue.Error(com.tjclp.xl.cells.CellError.Div0)
    )
    // SUM propagates the element's error VALUE (GH-344 policy)
    assertEquals(
      sheet.evaluateFormula("=SUM(A1, A2)"),
      Right(CellValue.Error(com.tjclp.xl.cells.CellError.Div0))
    )
    // COUNT: errors are not numbers — skipped, not counted
    assertEquals(
      sheet.evaluateFormula("=COUNT(A1, A2)"),
      Right(CellValue.Number(BigDecimal(1)))
    )
  }

  // ===== COUNTBLANK Parser Test =====

  test("COUNTBLANK: parse from string") {
    val result = FormulaParser.parse("=COUNTBLANK(A1:A5)")
    assert(result.isRight, s"Failed to parse COUNTBLANK: $result")
    result match
      case Right(expr) =>
        val printed = FormulaPrinter.print(expr)
        assertEquals(printed, "=COUNTBLANK(A1:A5)")
      case Left(err) =>
        fail(s"Parse failed: $err")
  }
