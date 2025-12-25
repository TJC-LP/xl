package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.addressing.SheetName
import munit.FunSuite

/**
 * Comprehensive tests for COUNTBLANK function.
 *
 * COUNTBLANK counts empty cells in a range.
 * Cells with formulas that return "" are NOT counted as blank (Excel behavior).
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
