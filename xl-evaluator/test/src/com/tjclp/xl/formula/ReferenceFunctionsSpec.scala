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
    val expr = TExpr.Row_(TExpr.PolyRef(ref"A1", Anchor.Relative))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(1)))
  }

  test("ROW: returns row number of reference B5 = 5") {
    val expr = TExpr.Row_(TExpr.PolyRef(ref"B5", Anchor.Relative))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(5)))
  }

  test("ROW: returns row number of reference Z100 = 100") {
    val expr = TExpr.Row_(TExpr.PolyRef(ref"Z100", Anchor.Relative))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(100)))
  }

  test("ROW: works with absolute reference") {
    val expr = TExpr.Row_(TExpr.PolyRef(ref"C10", Anchor.Absolute))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(10)))
  }

  // ===== COLUMN Tests =====

  test("COLUMN: returns column number of reference A1 = 1") {
    val expr = TExpr.Column_(TExpr.PolyRef(ref"A1", Anchor.Relative))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(1)))
  }

  test("COLUMN: returns column number of reference B5 = 2") {
    val expr = TExpr.Column_(TExpr.PolyRef(ref"B5", Anchor.Relative))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(2)))
  }

  test("COLUMN: returns column number of reference Z1 = 26") {
    val expr = TExpr.Column_(TExpr.PolyRef(ref"Z1", Anchor.Relative))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(26)))
  }

  test("COLUMN: returns column number of reference AA1 = 27") {
    val expr = TExpr.Column_(TExpr.PolyRef(ref"AA1", Anchor.Relative))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(27)))
  }

  // ===== ROWS Tests =====

  test("ROWS: count rows in single cell range = 1") {
    val range = CellRange.parse("A1:A1").toOption.get
    val foldRange = TExpr.FoldRange(
      range,
      BigDecimal(0),
      (_: BigDecimal, _: BigDecimal) => BigDecimal(0),
      TExpr.decodeNumeric
    )
    val expr = TExpr.Rows(foldRange)
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(1)))
  }

  test("ROWS: count rows in vertical range A1:A10 = 10") {
    val range = CellRange.parse("A1:A10").toOption.get
    val foldRange = TExpr.FoldRange(
      range,
      BigDecimal(0),
      (_: BigDecimal, _: BigDecimal) => BigDecimal(0),
      TExpr.decodeNumeric
    )
    val expr = TExpr.Rows(foldRange)
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(10)))
  }

  test("ROWS: count rows in 2D range B2:D5 = 4") {
    val range = CellRange.parse("B2:D5").toOption.get
    val foldRange = TExpr.FoldRange(
      range,
      BigDecimal(0),
      (_: BigDecimal, _: BigDecimal) => BigDecimal(0),
      TExpr.decodeNumeric
    )
    val expr = TExpr.Rows(foldRange)
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(4)))
  }

  // ===== COLUMNS Tests =====

  test("COLUMNS: count columns in single cell range = 1") {
    val range = CellRange.parse("A1:A1").toOption.get
    val foldRange = TExpr.FoldRange(
      range,
      BigDecimal(0),
      (_: BigDecimal, _: BigDecimal) => BigDecimal(0),
      TExpr.decodeNumeric
    )
    val expr = TExpr.Columns(foldRange)
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(1)))
  }

  test("COLUMNS: count columns in horizontal range A1:J1 = 10") {
    val range = CellRange.parse("A1:J1").toOption.get
    val foldRange = TExpr.FoldRange(
      range,
      BigDecimal(0),
      (_: BigDecimal, _: BigDecimal) => BigDecimal(0),
      TExpr.decodeNumeric
    )
    val expr = TExpr.Columns(foldRange)
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(10)))
  }

  test("COLUMNS: count columns in 2D range B2:D5 = 3") {
    val range = CellRange.parse("B2:D5").toOption.get
    val foldRange = TExpr.FoldRange(
      range,
      BigDecimal(0),
      (_: BigDecimal, _: BigDecimal) => BigDecimal(0),
      TExpr.decodeNumeric
    )
    val expr = TExpr.Columns(foldRange)
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(3)))
  }

  // ===== ADDRESS Tests =====

  test("ADDRESS: basic absolute reference $A$1") {
    val expr = TExpr.Address(
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
    val expr = TExpr.Address(
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
    val expr = TExpr.Address(
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
    val expr = TExpr.Address(
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
    val expr = TExpr.Address(
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
    val expr = TExpr.Address(
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
    val expr = TExpr.Address(
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
    val expr = TExpr.Address(
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
    val expr = TExpr.Address(
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
    val expr = TExpr.Address(
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
    val expr = TExpr.Address(
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

  test("COLUMN: parse from string") {
    val result = FormulaParser.parse("=COLUMN(C1)")
    assert(result.isRight, s"Failed to parse COLUMN: $result")
    result.foreach { expr =>
      val printed = FormulaPrinter.print(expr)
      assertEquals(printed, "=COLUMN(C1)")
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
