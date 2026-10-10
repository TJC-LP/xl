package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.eval.ArrayResult
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * GH-713: XLOOKUP returns a reference — the matched row (or column) of `return_array` — so a 2-D
 * return array yields a whole row, a reference position (SUM, ROWS, the `:` operator) reads it
 * whole, and a plain cell reads it through the implicit intersection, as Excel 365's `=@XLOOKUP(…)`
 * does.
 *
 * LibreOffice 24.2 has no XLOOKUP: the values are Excel's documented semantics. The INDEX lifting
 * pins are LibreOffice 24.2.7's.
 */
@SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
class XLookupReferenceSpec extends FunSuite:

  private def N(n: String): CellValue = CellValue.Number(BigDecimal(n))
  private def E(code: String): CellValue =
    CellValue.Error(CellError.parse(code).fold(msg => fail(msg), identity))

  /** Sheet1: A1:A3 = 1,2,3, B1:B3 = 10,20,30, C1:C3 = 100,200,300, D1 blank, row 5 = keys. */
  private val book: Workbook =
    val grid = (1 to 3).foldLeft(Sheet(SheetName.unsafe("Sheet1"))) { (s, i) =>
      s.put(ARef.from1(1, i), N(i.toString))
        .put(ARef.from1(2, i), N((i * 10).toString))
        .put(ARef.from1(3, i), N((i * 100).toString))
    }
    // a horizontal key row: A5:C5 = 1,2,3
    val keyed = (1 to 3).foldLeft(grid)((s, i) => s.put(ARef.from1(i, 5), N(i.toString)))
    Workbook(Vector(keyed))

  private def sheet1: Sheet = book.sheets.headOption.getOrElse(fail("no sheet"))

  private def plain(at: String, formula: String): Either[String, CellValue] =
    val ref = ARef.parse(at).fold(msg => fail(msg), identity)
    val placed = sheet1.put(ref, CellValue.Formula(formula, None))
    placed.evaluateCell(ref, Clock.system, Some(book.put(placed))).left.map(_.message)

  private def array(formula: String): Either[EvalError, Any] =
    Evaluator.arrayInstance.eval(
      FormulaParser
        .parse(formula)
        .fold(err => fail(s"$formula: $err"), identity)
        .asInstanceOf[TExpr[Any]],
      sheet1,
      workbook = Some(book)
    )

  private def row(values: String*): ArrayResult = ArrayResult(Vector(values.map(N).toVector))

  test("GH-713: a 2-D return array yields the whole matched row") {
    assertEquals(array("=XLOOKUP(2,A1:A3,B1:C3)"), Right(row("20", "200")))
    assertEquals(plain("H1", "ROWS(XLOOKUP(2,A1:A3,B1:C3))"), Right(N("1")))
    assertEquals(plain("H1", "COLUMNS(XLOOKUP(2,A1:A3,B1:C3))"), Right(N("2")))
    assertEquals(plain("H1", "SUM(XLOOKUP(2,A1:A3,B1:C3))"), Right(N("220")))
    assertEquals(array("=SUM(XLOOKUP(2,A1:A3,B1:C3))"), Right(BigDecimal(220)))
  }

  test("GH-713: a plain cell reads the matched row in its own column (=@XLOOKUP)") {
    assertEquals(plain("B9", "XLOOKUP(2,A1:A3,B1:C3)"), Right(N("20")))
    assertEquals(plain("C9", "XLOOKUP(2,A1:A3,B1:C3)"), Right(N("200")))
    assertEquals(plain("E9", "XLOOKUP(2,A1:A3,B1:C3)"), Right(E("#VALUE!")))
    // a one-cell result is the same value in every column
    for at <- List("B9", "E9", "Z42") do
      assertEquals(plain(at, "XLOOKUP(2,A1:A3,B1:B3)"), Right(N("20")), at)
      assertEquals(plain(at, "XLOOKUP(2,A1:A3,B1:B3)*2"), Right(N("40")), at)
  }

  test("GH-713: a horizontal lookup returns a column; other shapes are #VALUE!") {
    // keys across A5:C5, return rows 1..3 of A:C: the matched column, top to bottom
    assertEquals(
      array("=XLOOKUP(2,A5:C5,A1:C3)"),
      Right(ArrayResult(Vector(Vector(N("10")), Vector(N("20")), Vector(N("30")))))
    )
    assertEquals(plain("H1", "SUM(XLOOKUP(3,A5:C5,A1:C3))"), Right(N("600")))
    // a 2-D lookup array, and a return array of another length
    plain("H1", "XLOOKUP(2,A1:B3,B1:C3)") match
      case Right(CellValue.Error(CellError.Value)) => ()
      case other => fail(s"2-D lookup_array: $other")
    val message = array("=XLOOKUP(2,A1:B3,B1:C3)") match
      case Left(EvalError.ErrorValue(CellError.Value, Some(msg))) => msg
      case other => fail(s"2-D lookup_array: $other")
    assert(message.contains("dimension"), message)
    assertEquals(plain("H1", "XLOOKUP(2,A1:A3,B1:B2)"), Right(E("#VALUE!")))
    // a 1×1 lookup array with a one-row return array: the row (the vertical rule wins)
    assertEquals(array("=XLOOKUP(1,A1,A1:C1)"), Right(row("1", "10", "100")))
  }

  test("GH-713: a miss is #N/A without if_not_found, its value with one") {
    assertEquals(plain("H1", "XLOOKUP(9,A1:A3,B1:C3)"), Right(E("#N/A")))
    assertEquals(plain("H1", "ISNA(XLOOKUP(9,A1:A3,B1:C3))"), Right(CellValue.Bool(true)))
    assertEquals(plain("H1", "IFNA(XLOOKUP(9,A1:A3,B1:B3),-1)"), Right(N("-1")))
    assertEquals(plain("H1", "XLOOKUP(9,A1:A3,B1:B3,\"none\")"), Right(CellValue.Text("none")))
    assertEquals(plain("H1", "SUM(XLOOKUP(9,A1:A3,B1:C3,0))"), Right(N("0")))
    // a matched blank reads as INDEX over a blank does
    val blankRow = sheet1.put(ref"A7", N("7"))
    val placed = blankRow.put(ref"H1", CellValue.Formula("XLOOKUP(7,A1:A7,D1:D7)", None))
    val index = blankRow.put(ref"H1", CellValue.Formula("INDEX(D1:D7,7)", None))
    assertEquals(
      placed.evaluateCell(ref"H1", Clock.system, Some(book.put(placed))),
      index.evaluateCell(ref"H1", Clock.system, Some(book.put(index)))
    )
  }

  test("GH-713: XLOOKUP is an operand of ':'") {
    assertEquals(plain("H1", "SUM(B1:XLOOKUP(2,A1:A3,B1:B3))"), Right(N("30")))
    assertEquals(
      plain("H1", "SUM(XLOOKUP(1,A1:A3,B1:B3):XLOOKUP(3,A1:A3,C1:C3))"),
      Right(N("660"))
    )
    assertEquals(plain("H1", "SUM(B1:XLOOKUP(9,A1:A3,B1:B3))"), Right(E("#N/A")))
  }

  test("GH-713: a year's row — SUM(XLOOKUP(key,ids,B:M)) — costs the data") {
    val wide = (1 to 3).foldLeft(sheet1) { (s, r) =>
      (2 to 13).foldLeft(s)((acc, c) => acc.put(ARef.from1(c, r), N((r * c).toString)))
    }
    val placed = wide.put(ref"O1", CellValue.Formula("SUM(XLOOKUP(2,A:A,B:M))", None))
    // row 2: 2*(2+…+13) = 2*90
    assertEquals(
      placed.evaluateCell(ref"O1", Clock.system, Some(book.put(placed))),
      Right(N("180"))
    )
  }

  test("GH-713: an array key still lifts under an aggregate (no reference collapse)") {
    assertEquals(plain("H1", "SUM(XLOOKUP({1;2},A1:A3,B1:B3))"), Right(N("30")))
    assertEquals(plain("H1", "SUMPRODUCT(XLOOKUP({1;2},A1:A3,B1:B3))"), Right(N("30")))
    // LibreOffice 24.2.7: INDEX lifts over an array position under an aggregate
    assertEquals(plain("H1", "SUM(INDEX(B1:C3,{1;2},1))"), Right(N("30")))
    assertEquals(plain("H1", "SUM(INDEX(B1:B3,{1;2}))"), Right(N("30")))
  }

  test("GH-713: the result keeps round-tripping as _xlfn.XLOOKUP") {
    val expr =
      FormulaParser.parse("=SUM(B1:XLOOKUP(2,A1:A3,B1:B3))").fold(e => fail(e.toString), identity)
    assertEquals(FormulaPrinter.printFileForm(expr), "SUM(B1:XLOOKUP(2,A1:A3,B1:B3))")
    assertEquals(
      com.tjclp.xl.ooxml.FormulaStorage.toStored(FormulaPrinter.printFileForm(expr)),
      "SUM(B1:_xlfn.XLOOKUP(2,A1:A3,B1:B3))"
    )
  }
