package com.tjclp.xl.formula

import java.nio.file.Path

import scala.collection.immutable.ArraySeq

import com.tjclp.xl.{*, given}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.addressing.{Column, SheetName}
import com.tjclp.xl.context.{SourceContext, SourceFingerprint}
import com.tjclp.xl.ooxml.PartManifest
import com.tjclp.xl.sheets.ColumnProperties
import munit.FunSuite

/**
 * Comprehensive tests for reference functions: ROW, COLUMN, ROWS, COLUMNS, ADDRESS, CELL.
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
      TExpr.Lit(true), // a1_style = true
      None // no sheet name
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
      TExpr.Lit(false), // a1_style = false (R1C1)
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

  // ===== CELL Tests (GH-424) =====

  /** Parse a formula or fail the test — for evaluating with a non-default Evaluator instance. */
  private def parseOrFail(formula: String): TExpr[?] =
    FormulaParser.parse(formula) match
      case Right(expr) => expr
      case Left(err) => fail(s"parse failed for $formula: $err")

  test("CELL: filename is empty string pre-save (no workbook path) — Excel-faithful") {
    // The house self-label idiom: =CELL("Filename",$B$2) on every print tab
    assertEquals(
      emptySheet.evaluateFormula("=CELL(\"Filename\",$B$2)"),
      Right(CellValue.Text(""))
    )
  }

  test("CELL: filename is directory + [file] + sheet once a workbook path is known") {
    val expr = parseOrFail("=CELL(\"Filename\",$B$2)")
    Evaluator.instance(workbookPath = Some("/models/lbo.xlsx")).eval(expr, emptySheet) match
      case Right(CellValue.Text(s)) => assertEquals(s, "/models/[lbo.xlsx]Test")
      case other => fail(s"expected filename text, got $other")
  }

  test("CELL: filename splits on the last backslash for Windows-shaped paths") {
    // No-reference form: the ambient sheet is the subject, no cell context required
    val expr = parseOrFail("=CELL(\"filename\")")
    Evaluator.instance(workbookPath = Some("C:\\models\\lbo.xlsx")).eval(expr, emptySheet) match
      case Right(CellValue.Text(s)) => assertEquals(s, "C:\\models\\[lbo.xlsx]Test")
      case other => fail(s"expected filename text, got $other")
  }

  test("CELL: filename reports the QUALIFIER sheet for a cross-sheet reference") {
    val expr = parseOrFail("=CELL(\"filename\", Other!B2)")
    Evaluator.instance(workbookPath = Some("/m/x.xlsx")).eval(expr, emptySheet) match
      case Right(CellValue.Text(s)) => assertEquals(s, "/m/[x.xlsx]Other")
      case other => fail(s"expected filename text, got $other")
  }

  test("CELL: address returns the absolute A1 form of the reference's top-left cell") {
    assertEquals(
      emptySheet.evaluateFormula("=CELL(\"address\", B5)"),
      Right(CellValue.Text("$B$5"))
    )
    assertEquals(
      emptySheet.evaluateFormula("=CELL(\"address\", C3:D9)"),
      Right(CellValue.Text("$C$3"))
    )
  }

  test("CELL: row and col arms return 1-based coordinates of the top-left cell") {
    assertEquals(
      emptySheet.evaluateFormula("=CELL(\"row\", B5)"),
      Right(CellValue.Number(BigDecimal(5)))
    )
    assertEquals(
      emptySheet.evaluateFormula("=CELL(\"col\", B5)"),
      Right(CellValue.Number(BigDecimal(2)))
    )
    assertEquals(
      emptySheet.evaluateFormula("=CELL(\"row\", C3:D9)"),
      Right(CellValue.Number(BigDecimal(3)))
    )
    // info_type is case-insensitive like Excel
    assertEquals(
      emptySheet.evaluateFormula("=CELL(\"COL\", AA1)"),
      Right(CellValue.Number(BigDecimal(27)))
    )
  }

  test("CELL: no-reference form uses the current cell (ROW()/COLUMN() convention)") {
    val sheet = emptySheet.put(ref"B5", CellValue.Formula("CELL(\"address\")", None))
    assertEquals(sheet.evaluateCell(ref"B5"), Right(CellValue.Text("$B$5")))
  }

  test("CELL: address/row/col with no reference and no cell context is a clean error") {
    assert(
      emptySheet.evaluateFormula("=CELL(\"row\")").isLeft,
      "CELL(\"row\") without a cell context should fail cleanly"
    )
  }

  test("CELL: unsupported or unknown info_type is a clean #VALUE! error value, never a throw") {
    assertEquals(
      emptySheet.evaluateFormula("=CELL(\"format\", A1)"),
      Right(CellValue.Error(CellError.Value))
    )
    assertEquals(
      emptySheet.evaluateFormula("=CELL(\"no_such_arm\", A1)"),
      Right(CellValue.Error(CellError.Value))
    )
  }

  // ----- #959: width, contents, type; filename from the workbook's own file -----

  /** A workbook holding `sheets`, read from `path` the way XlsxReader records a file source. */
  private def bookReadFrom(path: String, sheets: Sheet*): Workbook =
    Workbook(sheets.toVector).copy(sourceContext =
      Some(
        SourceContext.fromFile(
          Path.of(path),
          PartManifest.empty,
          SourceFingerprint(0L, ArraySeq.empty[Byte])
        )
      )
    )

  test("#959 CELL: filename comes from the file the workbook was read from") {
    val book = bookReadFrom("/models/lbo.xlsx", emptySheet)
    assertEquals(
      emptySheet.evaluateFormula("=CELL(\"filename\",A1)", workbook = Some(book)),
      Right(CellValue.Text("/models/[lbo.xlsx]Test"))
    )
  }

  test("#959 CELL: filename of a workbook never read from a file is still empty") {
    assertEquals(
      emptySheet.evaluateFormula("=CELL(\"filename\",A1)", workbook = Some(Workbook(emptySheet))),
      Right(CellValue.Text(""))
    )
  }

  test("#959 CELL: an explicit workbook path wins over the file the workbook was read from") {
    val book = bookReadFrom("/old/in.xlsx", emptySheet)
    val expr = parseOrFail("=CELL(\"filename\",A1)")
    Evaluator
      .instance(workbookPath = Some("/new/out.xlsx"))
      .eval(expr, emptySheet, workbook = Some(book)) match
      case Right(CellValue.Text(s)) => assertEquals(s, "/new/[out.xlsx]Test")
      case other => fail(s"expected filename text, got $other")
  }

  test("#959 CELL: width of a default-width column is 8 (Excel's 8.43 characters, truncated)") {
    assertEquals(
      emptySheet.evaluateFormula("=CELL(\"width\",A1)"),
      Right(CellValue.Number(BigDecimal(8)))
    )
  }

  test("#959 CELL: width reads the column's stored width as whole characters, padding removed") {
    // <col width> carries 5px of padding at the default font's 7px digit: Excel shows a column
    // stored as 11.42578125 as 10.71 characters and one stored as 15 as 14.29
    val sheet = emptySheet
      .setColumnProperties(Column.from0(1), ColumnProperties(width = Some(11.42578125)))
      .setColumnProperties(Column.from0(2), ColumnProperties(width = Some(15.0)))
    assertEquals(sheet.evaluateFormula("=CELL(\"width\",B1)"), Right(CellValue.Number(10)))
    assertEquals(sheet.evaluateFormula("=CELL(\"width\",C7)"), Right(CellValue.Number(14)))
    // a range reports its top-left cell's column
    assertEquals(sheet.evaluateFormula("=CELL(\"width\",C1:D9)"), Right(CellValue.Number(14)))
  }

  test("#959 CELL: width of a hidden column is 0") {
    val sheet = emptySheet.setColumnProperties(
      Column.from0(1),
      ColumnProperties(width = Some(20.0), hidden = true)
    )
    assertEquals(sheet.evaluateFormula("=CELL(\"width\",B1)"), Right(CellValue.Number(0)))
  }

  test("#959 CELL: width of an unsized column follows the sheet's default column width") {
    val sheet = emptySheet.copy(defaultColumnWidth = Some(12.0))
    assertEquals(sheet.evaluateFormula("=CELL(\"width\",E1)"), Right(CellValue.Number(11)))
  }

  test("#959 CELL: width of a cross-sheet reference reads that sheet's columns") {
    val other = Sheet("Other")
      .setColumnProperties(Column.from0(0), ColumnProperties(width = Some(30.0)))
    val book = Workbook(Vector(emptySheet, other))
    assertEquals(
      emptySheet.evaluateFormula("=CELL(\"width\",Other!A1)", workbook = Some(book)),
      Right(CellValue.Number(29))
    )
  }

  test("#959 CELL: width with no reference uses the formula's own column") {
    val sheet = emptySheet
      .setColumnProperties(Column.from0(1), ColumnProperties(width = Some(15.0)))
      .put(ref"B5", CellValue.Formula("CELL(\"width\")", None))
    assertEquals(sheet.evaluateCell(ref"B5"), Right(CellValue.Number(14)))
  }

  test("#959 CELL: contents is the value of the top-left cell, not its formula") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(42)),
      ref"A2" -> CellValue.Text("Revenue"),
      ref"A3" -> CellValue.Formula("1+2", None),
      ref"A4" -> CellValue.Bool(true)
    )
    assertEquals(sheet.evaluateFormula("=CELL(\"contents\",A1)"), Right(CellValue.Number(42)))
    assertEquals(
      sheet.evaluateFormula("=CELL(\"contents\",A2:B9)"),
      Right(CellValue.Text("Revenue"))
    )
    assertEquals(sheet.evaluateFormula("=CELL(\"contents\",A3)"), Right(CellValue.Number(3)))
    assertEquals(sheet.evaluateFormula("=CELL(\"contents\",A4)"), Right(CellValue.Bool(true)))
  }

  test("#959 CELL: contents of an error cell is that error") {
    val sheet = sheetWith(ref"A1" -> CellValue.Error(CellError.Div0))
    assertEquals(
      sheet.evaluateFormula("=CELL(\"contents\",A1)"),
      Right(CellValue.Error(CellError.Div0))
    )
  }

  test("#959 CELL: contents of a blank cell is what a plain reference to it gives") {
    assertEquals(
      emptySheet.evaluateFormula("=CELL(\"contents\",A1)"),
      emptySheet.evaluateFormula("=A1")
    )
  }

  test("#959 CELL: type is b for blank, l for text, v for anything else") {
    val sheet = sheetWith(
      ref"A2" -> CellValue.Text("Revenue"),
      ref"A3" -> CellValue.Number(BigDecimal(1)),
      ref"A4" -> CellValue.Formula("\"x\"&\"y\"", None),
      ref"A5" -> CellValue.Formula("1/0", None),
      ref"A6" -> CellValue.Bool(false),
      ref"A7" -> CellValue.Formula("\"\"", None)
    )
    val types = (1 to 7).map(r => sheet.evaluateFormula(s"=CELL(\"type\",A$r)"))
    assertEquals(
      types.toList,
      List("b", "l", "v", "l", "v", "v", "l").map(t => Right(CellValue.Text(t)))
    )
  }

  test(
    "#959 CELL: contents and type with no reference are #VALUE!, never a read of their own cell"
  ) {
    // Excel reads the selected cell; the current cell would be a circular read of itself
    val sheet = emptySheet
      .put(ref"B5", CellValue.Formula("CELL(\"contents\")", None))
      .put(ref"B6", CellValue.Formula("CELL(\"type\")", None))
    assertEquals(sheet.evaluateCell(ref"B5"), Right(CellValue.Error(CellError.Value)))
    assertEquals(sheet.evaluateCell(ref"B6"), Right(CellValue.Error(CellError.Value)))
  }

  test("#959 CELL: every info type xl does not support is #VALUE!") {
    List("format", "prefix", "protect", "color", "parentheses", "no_such_arm").foreach { info =>
      assertEquals(
        emptySheet.evaluateFormula(s"=CELL(\"$info\",A1)"),
        Right(CellValue.Error(CellError.Value)),
        info
      )
    }
  }

  test("CELL: non-reference second argument is a clean error, never a throw") {
    assert(
      emptySheet.evaluateFormula("=CELL(\"row\", 1+1)").isLeft,
      "CELL over a non-reference should fail cleanly"
    )
  }

  test("CELL: parses and round-trips through print") {
    List("=CELL(\"Filename\", $B$2)", "=CELL(\"row\", B5)", "=CELL(\"address\")").foreach { f =>
      FormulaParser.parse(f) match
        case Right(expr) => assertEquals(FormulaPrinter.print(expr), f)
        case Left(err) => fail(s"parse failed for $f: $err")
    }
  }
