package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.addressing.SheetName
import munit.FunSuite

/**
 * Comprehensive tests for reference functions: ROW, COLUMN, ROWS, COLUMNS, ADDRESS.
 *
 * These functions provide information about cell references and ranges.
 */
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial"))
class ReferenceFunctionsSpec extends FunSuite:
  val emptySheet = new Sheet(name = SheetName.unsafe("Test"))
  val evaluator = Evaluator.instance

  /** Helper to create sheet with cells */
  def sheetWith(cells: (ARef, CellValue)*): Sheet =
    cells.foldLeft(emptySheet) { case (s, (ref, value)) =>
      s.put(ref, value)
    }

  // ===== ROW Tests =====

  test("ROW: returns row number of reference A1 = 1") {
    val expr = TExpr.row(TExpr.PolyRef(ref"A1", Anchor.Relative))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(1)))
  }

  test("ROW: returns row number of reference B5 = 5") {
    val expr = TExpr.row(TExpr.PolyRef(ref"B5", Anchor.Relative))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(5)))
  }

  test("ROW: returns row number of reference Z100 = 100") {
    val expr = TExpr.row(TExpr.PolyRef(ref"Z100", Anchor.Relative))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(100)))
  }

  test("ROW: works with absolute reference") {
    val expr = TExpr.row(TExpr.PolyRef(ref"C10", Anchor.Absolute))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(10)))
  }

  test("ROW: zero-argument returns row of current cell") {
    val expr = TExpr.row()
    // Evaluate with currentCell = B5, so ROW() should return 5
    val result = evaluator.eval(expr, emptySheet, Clock.system, None, Some(ref"B5"))
    assertEquals(result, Right(BigDecimal(5)))
  }

  test("ROW: zero-argument returns row 1 when currentCell is A1") {
    val expr = TExpr.row()
    val result = evaluator.eval(expr, emptySheet, Clock.system, None, Some(ref"A1"))
    assertEquals(result, Right(BigDecimal(1)))
  }

  test("ROW: zero-argument returns row 100 when currentCell is Z100") {
    val expr = TExpr.row()
    val result = evaluator.eval(expr, emptySheet, Clock.system, None, Some(ref"Z100"))
    assertEquals(result, Right(BigDecimal(100)))
  }

  test("ROW: zero-argument fails when no current cell context") {
    val expr = TExpr.row()
    val result = evaluator.eval(expr, emptySheet, Clock.system, None, None)
    assert(result.isLeft, "ROW() without current cell should fail")
  }

  // ===== COLUMN Tests =====

  test("COLUMN: returns column number of reference A1 = 1") {
    val expr = TExpr.column(TExpr.PolyRef(ref"A1", Anchor.Relative))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(1)))
  }

  test("COLUMN: returns column number of reference B5 = 2") {
    val expr = TExpr.column(TExpr.PolyRef(ref"B5", Anchor.Relative))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(2)))
  }

  test("COLUMN: returns column number of reference Z1 = 26") {
    val expr = TExpr.column(TExpr.PolyRef(ref"Z1", Anchor.Relative))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(26)))
  }

  test("COLUMN: returns column number of reference AA1 = 27") {
    val expr = TExpr.column(TExpr.PolyRef(ref"AA1", Anchor.Relative))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(27)))
  }

  test("COLUMN: zero-argument returns column of current cell") {
    val expr = TExpr.column()
    // Evaluate with currentCell = C5, so COLUMN() should return 3
    val result = evaluator.eval(expr, emptySheet, Clock.system, None, Some(ref"C5"))
    assertEquals(result, Right(BigDecimal(3)))
  }

  test("COLUMN: zero-argument returns column 1 when currentCell is A1") {
    val expr = TExpr.column()
    val result = evaluator.eval(expr, emptySheet, Clock.system, None, Some(ref"A1"))
    assertEquals(result, Right(BigDecimal(1)))
  }

  test("COLUMN: zero-argument returns column 26 when currentCell is Z1") {
    val expr = TExpr.column()
    val result = evaluator.eval(expr, emptySheet, Clock.system, None, Some(ref"Z1"))
    assertEquals(result, Right(BigDecimal(26)))
  }

  test("COLUMN: zero-argument fails when no current cell context") {
    val expr = TExpr.column()
    val result = evaluator.eval(expr, emptySheet, Clock.system, None, None)
    assert(result.isLeft, "COLUMN() without current cell should fail")
  }

  // ===== ROWS Tests =====

  test("ROWS: count rows in single cell range = 1") {
    val range = CellRange.parse("A1:A1").toOption.get
    val expr = TExpr.rows(TExpr.RangeRef(range))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(1)))
  }

  test("ROWS: count rows in vertical range A1:A10 = 10") {
    val range = CellRange.parse("A1:A10").toOption.get
    val expr = TExpr.rows(TExpr.RangeRef(range))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(10)))
  }

  test("ROWS: count rows in 2D range B2:D5 = 4") {
    val range = CellRange.parse("B2:D5").toOption.get
    val expr = TExpr.rows(TExpr.RangeRef(range))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(4)))
  }

  // ===== COLUMNS Tests =====

  test("COLUMNS: count columns in single cell range = 1") {
    val range = CellRange.parse("A1:A1").toOption.get
    val expr = TExpr.columns(TExpr.RangeRef(range))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(1)))
  }

  test("COLUMNS: count columns in horizontal range A1:J1 = 10") {
    val range = CellRange.parse("A1:J1").toOption.get
    val expr = TExpr.columns(TExpr.RangeRef(range))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(10)))
  }

  test("COLUMNS: count columns in 2D range B2:D5 = 3") {
    val range = CellRange.parse("B2:D5").toOption.get
    val expr = TExpr.columns(TExpr.RangeRef(range))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(3)))
  }

  // ===== ADDRESS Tests =====

  test("ADDRESS: basic absolute reference $A$1") {
    val expr = TExpr.address(
      TExpr.Lit(BigDecimal(1)), // row
      TExpr.Lit(BigDecimal(1)), // col
      TExpr.Lit(BigDecimal(1)), // abs_num = 1 ($A$1)
      TExpr.Lit(true),          // a1_style = true
      None                       // no sheet name
    )
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right("$A$1"))
  }

  test("ADDRESS: row absolute A$1") {
    val expr = TExpr.address(
      TExpr.Lit(BigDecimal(1)),
      TExpr.Lit(BigDecimal(1)),
      TExpr.Lit(BigDecimal(2)), // abs_num = 2 (A$1)
      TExpr.Lit(true),
      None
    )
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right("A$1"))
  }

  test("ADDRESS: col absolute $A1") {
    val expr = TExpr.address(
      TExpr.Lit(BigDecimal(1)),
      TExpr.Lit(BigDecimal(1)),
      TExpr.Lit(BigDecimal(3)), // abs_num = 3 ($A1)
      TExpr.Lit(true),
      None
    )
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right("$A1"))
  }

  test("ADDRESS: relative A1") {
    val expr = TExpr.address(
      TExpr.Lit(BigDecimal(1)),
      TExpr.Lit(BigDecimal(1)),
      TExpr.Lit(BigDecimal(4)), // abs_num = 4 (A1)
      TExpr.Lit(true),
      None
    )
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right("A1"))
  }

  test("ADDRESS: column 26 = Z") {
    val expr = TExpr.address(
      TExpr.Lit(BigDecimal(1)),
      TExpr.Lit(BigDecimal(26)), // column 26 = Z
      TExpr.Lit(BigDecimal(4)),
      TExpr.Lit(true),
      None
    )
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right("Z1"))
  }

  test("ADDRESS: column 27 = AA") {
    val expr = TExpr.address(
      TExpr.Lit(BigDecimal(1)),
      TExpr.Lit(BigDecimal(27)), // column 27 = AA
      TExpr.Lit(BigDecimal(4)),
      TExpr.Lit(true),
      None
    )
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right("AA1"))
  }

  test("ADDRESS: row 100, column 3 = C100") {
    val expr = TExpr.address(
      TExpr.Lit(BigDecimal(100)),
      TExpr.Lit(BigDecimal(3)), // column 3 = C
      TExpr.Lit(BigDecimal(1)),
      TExpr.Lit(true),
      None
    )
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right("$C$100"))
  }

  test("ADDRESS: with sheet name") {
    val expr = TExpr.address(
      TExpr.Lit(BigDecimal(5)),
      TExpr.Lit(BigDecimal(2)),
      TExpr.Lit(BigDecimal(1)),
      TExpr.Lit(true),
      Some(TExpr.Lit("Sales"))
    )
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right("Sales!$B$5"))
  }

  test("ADDRESS: R1C1 notation") {
    val expr = TExpr.address(
      TExpr.Lit(BigDecimal(5)),
      TExpr.Lit(BigDecimal(3)),
      TExpr.Lit(BigDecimal(1)), // abs = 1 for R5C3 (absolute)
      TExpr.Lit(false),         // a1_style = false (R1C1)
      None
    )
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right("R5C3"))
  }

  test("ADDRESS: invalid row returns #VALUE!") {
    val expr = TExpr.address(
      TExpr.Lit(BigDecimal(0)), // invalid row
      TExpr.Lit(BigDecimal(1)),
      TExpr.Lit(BigDecimal(1)),
      TExpr.Lit(true),
      None
    )
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right("#VALUE!"))
  }

  test("ADDRESS: invalid column returns #VALUE!") {
    val expr = TExpr.address(
      TExpr.Lit(BigDecimal(1)),
      TExpr.Lit(BigDecimal(0)), // invalid column
      TExpr.Lit(BigDecimal(1)),
      TExpr.Lit(true),
      None
    )
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right("#VALUE!"))
  }

  // ===== Parser Tests =====

  test("ROW: parse from string") {
    val result = FormulaParser.parse("=ROW(A5)")
    assert(result.isRight, s"Failed to parse ROW: $result")
    result.foreach { expr =>
      val printed = FormulaPrinter.print(expr)
      assertEquals(printed, "=ROW(A5)")
    }
  }

  test("ROW: parse zero-argument form") {
    val result = FormulaParser.parse("=ROW()")
    assert(result.isRight, s"Failed to parse ROW(): $result")
    result.foreach { expr =>
      val printed = FormulaPrinter.print(expr)
      assertEquals(printed, "=ROW()")
    }
  }

  test("COLUMN: parse from string") {
    val result = FormulaParser.parse("=COLUMN(C1)")
    assert(result.isRight, s"Failed to parse COLUMN: $result")
    result.foreach { expr =>
      val printed = FormulaPrinter.print(expr)
      assertEquals(printed, "=COLUMN(C1)")
    }
  }

  test("COLUMN: parse zero-argument form") {
    val result = FormulaParser.parse("=COLUMN()")
    assert(result.isRight, s"Failed to parse COLUMN(): $result")
    result.foreach { expr =>
      val printed = FormulaPrinter.print(expr)
      assertEquals(printed, "=COLUMN()")
    }
  }

  test("ROWS: parse from string") {
    val result = FormulaParser.parse("=ROWS(A1:A10)")
    assert(result.isRight, s"Failed to parse ROWS: $result")
    result.foreach { expr =>
      val printed = FormulaPrinter.print(expr)
      assertEquals(printed, "=ROWS(A1:A10)")
    }
  }

  test("COLUMNS: parse from string") {
    val result = FormulaParser.parse("=COLUMNS(A1:D1)")
    assert(result.isRight, s"Failed to parse COLUMNS: $result")
    result.foreach { expr =>
      val printed = FormulaPrinter.print(expr)
      assertEquals(printed, "=COLUMNS(A1:D1)")
    }
  }

  test("ADDRESS: parse from string (4 args)") {
    val result = FormulaParser.parse("=ADDRESS(1, 1, 1, TRUE)")
    assert(result.isRight, s"Failed to parse ADDRESS: $result")
    result.foreach { expr =>
      val printed = FormulaPrinter.print(expr)
      assertEquals(printed, "=ADDRESS(1, 1, 1, TRUE)")
    }
  }

  test("ADDRESS: parse from string (5 args with sheet)") {
    val result = FormulaParser.parse("=ADDRESS(1, 1, 1, TRUE, \"Sheet1\")")
    assert(result.isRight, s"Failed to parse ADDRESS with sheet: $result")
    result.foreach { expr =>
      val printed = FormulaPrinter.print(expr)
      assertEquals(printed, "=ADDRESS(1, 1, 1, TRUE, \"Sheet1\")")
    }
  }

  // ===== SheetEvaluator Integration Tests =====

  import com.tjclp.xl.formula.eval.SheetEvaluator.*

  test("evaluateCell: ROW() formula returns row of containing cell") {
    // Put =ROW() formula in cell B5 - should evaluate to 5
    val sheet = emptySheet.put(ref"B5", CellValue.Formula("ROW()", None))
    val result = sheet.evaluateCell(ref"B5")
    assertEquals(result, Right(CellValue.Number(BigDecimal(5))))
  }

  test("evaluateCell: COLUMN() formula returns column of containing cell") {
    // Put =COLUMN() formula in cell D3 - should evaluate to 4 (D is 4th column)
    val sheet = emptySheet.put(ref"D3", CellValue.Formula("COLUMN()", None))
    val result = sheet.evaluateCell(ref"D3")
    assertEquals(result, Right(CellValue.Number(BigDecimal(4))))
  }

  test("evaluateCell: ROW()+COLUMN() formula in C7 returns 10") {
    // C7: row=7, column=3, sum=10
    val sheet = emptySheet.put(ref"C7", CellValue.Formula("ROW()+COLUMN()", None))
    val result = sheet.evaluateCell(ref"C7")
    assertEquals(result, Right(CellValue.Number(BigDecimal(10))))
  }

  test("evaluateFormula: ROW() without cell context fails gracefully") {
    // When evaluateFormula is called without cell context, ROW() should fail
    val result = emptySheet.evaluateFormula("=ROW()")
    assert(result.isLeft, "ROW() without cell context should fail")
  }
