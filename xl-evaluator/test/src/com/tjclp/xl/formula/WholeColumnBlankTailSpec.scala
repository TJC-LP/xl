package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.functions.{ArgSpec, SumProductTail}
import com.tjclp.xl.formula.parser.FormulaParser
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * GH-711 / GH-192: SUMPRODUCT, COUNTIF and COUNTIFS read a whole column within the sheet's used
 * extent, and add back the rows trimmed past it, every one blank, as Excel counts them. Each
 * expected value is LibreOffice 24.2's, except where a comment says otherwise.
 */
class WholeColumnBlankTailSpec extends FunSuite:

  private def num(n: Double): CellValue = CellValue.Number(BigDecimal(n))

  /** The review's book: A1=4, A2=5, B1=1; every other cell blank. */
  private val sheet: Sheet =
    Sheet(SheetName.unsafe("T"))
      .put(ref"A1", num(4))
      .put(ref"A2", num(5))
      .put(ref"B1", num(1))

  private def eval(formula: String, on: Sheet = sheet): XLResult[CellValue] =
    on.evaluateFormula(formula, Clock.system, Some(Workbook(Vector(on))), None)

  private def evalIn(wb: Workbook, on: Sheet)(formula: String): XLResult[CellValue] =
    on.evaluateFormula(formula, Clock.system, Some(wb), None)

  test("a selector's whole-column condition counts the blank rows past the used range") {
    // 1048576 rows, two filled: the trimmed condition must not lose the 1048574 blank ones
    assertEquals(eval("=SUMPRODUCT(IF(A:A=\"\",1,0))"), Right(num(1048574)))
    assertEquals(eval("=SUMPRODUCT(IF(ISBLANK(A:A),1,0))"), Right(num(1048574)))
    assertEquals(eval("=SUMPRODUCT(CHOOSE(1,--(A:A=\"\")))"), Right(num(1048574)))
    // the trimmed branches still line up with the condition (GH-711)
    assertEquals(eval("=SUMPRODUCT(IF(A:A>4,B:B,A:A))"), Right(num(4)))
    assertEquals(eval("=SUMPRODUCT(IF(A:A>4,B:B,A:A),A:A)"), Right(num(16)))
  }

  test("SUMPRODUCT counts every blank row of a whole column, as Excel does (GH-192)") {
    assertEquals(eval("=SUMPRODUCT(--(A:A=\"\"))"), Right(num(1048574)))
    assertEquals(eval("=SUMPRODUCT(--ISBLANK(A:A))"), Right(num(1048574)))
    assertEquals(eval("=SUMPRODUCT((A:A=\"\")*(B:B=\"\"))"), Right(num(1048574)))
    assertEquals(eval("=SUMPRODUCT(--(A:A=\"\"),--(B:B=\"\"))"), Right(num(1048574)))
    assertEquals(eval("=SUMPRODUCT((A:A=\"\")*2)"), Right(num(2097148)))
    // two columns: 2 * 1048576 cells, three filled
    assertEquals(eval("=SUMPRODUCT(--(A:B=\"\"))"), Right(num(2097149)))
    // a whole row: 16384 cells, two filled
    assertEquals(eval("=SUMPRODUCT(--(1:1=\"\"))"), Right(num(16382)))
  }

  test("a blank row that contributes nothing leaves the trimmed sum as it was") {
    assertEquals(eval("=SUMPRODUCT(--(A:A>0))"), Right(num(2)))
    assertEquals(eval("=SUMPRODUCT(A:A,B:B)"), Right(num(4)))
    assertEquals(eval("=SUMPRODUCT(A:A*2)"), Right(num(18)))
    assertEquals(eval("=SUMPRODUCT(IF(TRUE,B:B,A:A))"), Right(num(1)))
  }

  test("the rows above a used range that starts below row 1 count too") {
    val lower = Sheet(SheetName.unsafe("U"))
      .put(ref"A5", num(1))
      .put(ref"A6", num(2))
      .put(ref"A7", num(3))
    assertEquals(eval("=SUMPRODUCT(--(A:A=\"\"))", lower), Right(num(1048573)))
    assertEquals(eval("=SUMPRODUCT(A:A)", lower), Right(num(6)))
    assertEquals(eval("=COUNTIF(A:A,\"\")", lower), Right(num(1048573)))
  }

  test("an empty sheet is blank throughout") {
    val empty = Sheet(SheetName.unsafe("E"))
    assertEquals(eval("=SUMPRODUCT(--(A:A=\"\"))", empty), Right(num(1048576)))
    assertEquals(eval("=SUMPRODUCT(A:A)", empty), Right(num(0)))
    assertEquals(eval("=COUNTIF(A:A,\"\")", empty), Right(num(1048576)))
  }

  test("an error at a blank row is the result, as over the whole column in Excel") {
    // 1/blank is #DIV/0! in every trimmed row
    assertEquals(eval("=SUMPRODUCT(1/A:A)"), Right(CellValue.Error(CellError.Div0)))
    // ...and stays one when no row is trimmed off
    val bounded = eval("=SUMPRODUCT(1/A1:A3)")
    assertEquals(bounded, Right(CellValue.Error(CellError.Div0)))
  }

  test("the distinct-count idiom over whole columns stays exact: blanks match the criterion") {
    // A1:A3 = x, y, x. Each blank row adds 0/COUNTIF(A:A,"") = 0, COUNTIF counting every blank
    val texts = Sheet(SheetName.unsafe("W"))
      .put(ref"A1", CellValue.Text("x"))
      .put(ref"A2", CellValue.Text("y"))
      .put(ref"A3", CellValue.Text("x"))
    assertEquals(eval("=SUMPRODUCT((A:A<>\"\")/COUNTIF(A:A,A:A&\"\"))", texts), Right(num(2)))
    // Excel's whole-column 1/COUNTIF(A:A,A:A) divides by COUNTIF(A:A,<blank>) at every blank
    // row: no zeros, so #DIV/0! (the bounded range is the distinct count)
    assertEquals(
      eval("=SUMPRODUCT(1/COUNTIF(A:A,A:A))", texts),
      Right(CellValue.Error(CellError.Div0))
    )
    assertEquals(eval("=SUMPRODUCT(1/COUNTIF(A1:A3,A1:A3))", texts), Right(num(2)))
  }

  test("COUNTIF and COUNTIFS count the blank cells past the used range for a blank criterion") {
    assertEquals(eval("=COUNTIF(A:A,\"\")"), Right(num(1048574)))
    assertEquals(eval("=COUNTIF(A:A,\"=\")"), Right(num(1048574)))
    assertEquals(eval("=COUNTIF(A:A,\"<>\")"), Right(num(2)))
    assertEquals(eval("=COUNTIF(A:A,0)"), Right(num(0)))
    assertEquals(eval("=COUNTIF(A:A,\"<5\")"), Right(num(1)))
    assertEquals(eval("=COUNTIF(A:B,\"\")"), Right(num(2097149)))
    assertEquals(eval("=COUNTIFS(A:A,\"\",B:B,\"\")"), Right(num(1048574)))
    assertEquals(eval("=COUNTIFS(A:A,\"\",B:B,\"<>\")"), Right(num(0)))
    // a whole row (Excel's rule; LibreOffice 24.2 answers 0 here, its whole-row COUNTIF reads no
    // blank cell, while it counts the whole-column ones above)
    assertEquals(eval("=COUNTIF(1:1,\"\")"), Right(num(16382)))
  }

  test("the trimmed blank rows are blank on every sheet the arguments read") {
    val other = Sheet(SheetName.unsafe("O"))
      .put(ref"A1", num(1))
      .put(ref"A2", num(2))
      .put(ref"A3", num(3))
      .put(ref"A4", num(4))
    val wb = Workbook(Vector(sheet, other))
    val e = evalIn(wb, sheet)
    // shared extent rows 1-4: A3, A4 here are blank inside it, rows 5 on are blank in both
    assertEquals(e("=SUMPRODUCT(--(A:A=\"\"),--(O!A:A>0))"), Right(num(2)))
    assertEquals(e("=SUMPRODUCT(--(A:A=\"\"),--(O!A:A=\"\"))"), Right(num(1048572)))
  }

  test("a zero or blank-cell criterion adds no trimmed cell; a not-equal-to-value one adds all") {
    // "<>x": a blank is not x, so every blank cell counts (LibreOffice: 1048576). The trimmed
    // cells follow the matcher's verdict on a blank, the same as a blank inside the used range
    assertEquals(eval("=COUNTIF(A:A,\"<>x\")"), Right(num(1048576)))
    assertEquals(eval("=COUNTIF(A1:A3,\"<>x\")"), Right(num(3)))
    // a blank criterion cell counts zeros, never blanks (LibreOffice: 0 and 0)
    assertEquals(eval("=COUNTIF(A:A,C99)"), Right(num(0)))
    assertEquals(eval("=COUNTIF(A:A,0)"), Right(num(0)))
    val zero = sheet.put(ref"A3", num(0))
    assertEquals(eval("=COUNTIF(A:A,C99)", zero), Right(num(1)))
    assertEquals(eval("=COUNTIF(A:A,0)", zero), Right(num(1)))
  }

  test("COUNTIFS takes the trimmed count from a first range on another sheet") {
    // O!A1:A4 filled, T!A1:A2 filled: the shared bound is rows 1-4 (LibreOffice: 1048572 both ways)
    val other = (1 to 4).foldLeft(Sheet(SheetName.unsafe("O"))) { (s, i) =>
      s.put(com.tjclp.xl.addressing.ARef.from0(0, i - 1), num(i.toDouble))
    }
    val e = evalIn(Workbook(Vector(sheet, other)), sheet)
    assertEquals(e("=COUNTIFS(O!A:A,\"\",A:A,\"\")"), Right(num(1048572)))
    assertEquals(e("=COUNTIFS(A:A,\"\",O!A:A,\"\")"), Right(num(1048572)))
  }

  test("shapes the tail does not apply to keep their value") {
    // a whole column only under an aggregate is no element source: SUM(A:A) once, not per row
    // (LibreOffice: 9)
    assertEquals(eval("=SUMPRODUCT(SUM(A:A))"), Right(num(9)))
    // ...while beside one that is, SUM's range slot is one fixed reference and the line repeats
    // it on every blank row: 1048574 * SUM(B:B) (LibreOffice: 1048574)
    assertEquals(eval("=SUMPRODUCT((A:A=\"\")*SUM(B:B))"), Right(num(1048574)))
    // ROW of a range is its first row in xl (documented in LIMITATIONS; Excel and LibreOffice
    // fold every row: 1048571), so ROW(A:A) is no element source and nothing is read twice
    assertEquals(eval("=SUMPRODUCT(--(ROW(A:A)>5))"), Right(num(0)))
    // whole columns mixed with whole rows: no tail, the used-range blanks only (B2), where
    // Excel counts every blank cell of the sheet (documented in LIMITATIONS)
    assertEquals(eval("=SUMPRODUCT(--(A:XFD=\"\"))"), Right(num(1)))
    // the same column under ROW stays whole, so no tail: the trimmed sum, A1 and A2 not blank.
    // Excel and LibreOffice sum the blank rows' numbers (549756338173); xl's ROW of a range is its
    // first row, and the plan rejects the shape on its own whatever ROW returns
    assertEquals(eval("=SUMPRODUCT((A:A=\"\")*ROW(A:A))"), Right(num(0)))
    assertEquals(eval("=SUMPRODUCT((A:A=\"\")*ROW(B:B))"), Right(num(0)))
  }

  test("a volatile argument is read once: no tail, so RAND draws once") {
    // RAND()>=0 is TRUE for every row, so Excel counts the 1048574 blank rows; the tail would read
    // the argument a second time, so it is not read and the trimmed sum stands
    assertEquals(eval("=SUMPRODUCT((A:A=\"\")*(RAND()>=0))"), Right(num(0)))
    assertEquals(eval("=SUMPRODUCT((A:A=\"\")*(RAND()>=0),--(B:B=\"\"))"), Right(num(0)))
  }

  test("COUNTIFS counts the trimmed cells only when every range trimmed the same count") {
    // every range is one shape under one bound, so the counts agree and the blanks count
    assertEquals(eval("=COUNTIFS(A:A,\"\",B:B,\"\",C:C,\"\")"), Right(num(1048574)))
    assertEquals(eval("=COUNTIFS(A:B,\"\",B:C,\"\")"), Right(num(2097149)))
  }

  test("the tail's plan rebuilds no argument that already read every row (structural)") {
    @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
    def args(inner: String): List[ArgSpec.SumProductArg] =
      FormulaParser.parse(s"=SUMPRODUCT($inner)") match
        case Right(call: TExpr.Call[?]) => call.args.asInstanceOf[List[ArgSpec.SumProductArg]]
        case other => fail(s"not a SUMPRODUCT call: $other")
    val column = CellRange.parse("A:A").fold(e => fail(e), identity)
    val tail = SumProductTail
      .of(List(column), (Some((Row.from0(0), Row.from0(1))), None))
      .getOrElse(fail("rows 1-2 kept, so a tail"))
    assertEquals((tail.count, tail.kept), (1048574L, 2L))
    def plan(inner: String, extent: Int) = SumProductTail.plan(tail, args(inner), extent, _ => true)
    // an argument whose whole column no transform moved is never read over the blank row: ROW
    // keeps A:A, and at any extent the plan is empty
    assertEquals(plan("--(ROW(A:A)>5)", 1), None)
    assertEquals(plan("--(ROW(A:A)>5)", 2), None)
    // an argument that came back at another extent than the kept rows (a whole column a
    // shape-sensitive call read whole) is never read again either
    assertEquals(plan("--(A:A=\"\")", 1048576), None)
    // every whole column read as a value must move onto the line: one left whole rules the tail
    // out on its own, whatever its call returns (ROW(A:A) spreads over every row once ROW
    // returns an array). A range slot (SUM's, COUNTIF's range) is one reference and may stay
    assertEquals(plan("(A:A=\"\")*ROW(A:A)", 2), None)
    assertEquals(plan("(A:A=\"\")*ROW(B:B)", 2), None)
    assert(plan("(A:A=\"\")*SUM(B:B)", 2).isDefined)
    assert(plan("1/COUNTIF(A:A,A:A)", 2).isDefined)
    assertEquals(plan("SUM(A:A)", 1), None)
    // a volatile call anywhere in an argument rules it out: the line would draw again
    assertEquals(plan("(A:A=\"\")*(RAND()>=0)", 2), None)
    assertEquals(plan("--(A:A=\"\"),(B:B=\"\")*RANDBETWEEN(1,1)", 2), None)
    assert(plan("--(A:A=\"\")", 2).isDefined)
    assert(plan("(A:A=\"\")*(B:B<5),C:C", 2).isDefined)
    // the planned line reads row 3, the first one past the kept rows
    assertEquals(
      plan("--(A:A=\"\")", 2).map(_.map(_.fold(_ => Nil, TExpr.collectRanges(_).map(_._2.toA1)))),
      Some(List(List("A3:A3")))
    )
  }

  test("only a random draw rules the tail out: TODAY and NOW read one clock") {
    // 2026-10-12 is serial 46307: 4, 5 and every blank (0) are before it, as with a literal
    val clock = Clock.fixedDate(java.time.LocalDate.of(2026, 10, 12))
    def at(formula: String) =
      sheet.evaluateFormula(formula, clock, Some(Workbook(Vector(sheet))), None)
    assertEquals(at("=SUMPRODUCT(--(A:A<TODAY()))"), Right(num(1048576)))
    assertEquals(at("=SUMPRODUCT(--(A:A<46307))"), Right(num(1048576)))
    assertEquals(at("=SUMPRODUCT(--(A:A<NOW()))"), Right(num(1048576)))
    // RAND draws once, so no tail (Excel: 1048574)
    assertEquals(at("=SUMPRODUCT((A:A=\"\")*(RAND()>=0))"), Right(num(0)))
  }

  test("drift: every volatile spec is classified as a draw or a clock read") {
    val volatile = com.tjclp.xl.formula.functions.FunctionRegistry.volatileFunctionNames.toSet
    assert(SumProductTail.ClockReads.subsetOf(volatile), volatile)
    // a new volatile spec lands here, a draw until it is classified
    assertEquals(volatile -- SumProductTail.ClockReads, Set("RAND", "RANDBETWEEN"))
  }

  test("a defined name that draws rules the tail out; one that does not keeps it") {
    val wb = Workbook(Vector(sheet))
      .withDefinedName("draw", "RAND()")
      .withDefinedName("drawtwice", "draw*2")
      .withDefinedName("five", "5")
      .withDefinedName("asof", "TODAY()")
    val clock = Clock.fixedDate(java.time.LocalDate.of(2026, 10, 12))
    def e(formula: String) = sheet.evaluateFormula(formula, clock, Some(wb), None)
    // a draw through a name, and through a name's name: the trimmed sum (Excel: 1048574)
    assertEquals(e("=SUMPRODUCT((A:A=\"\")*(draw>=0))"), Right(num(0)))
    assertEquals(e("=SUMPRODUCT((A:A=\"\")*(drawtwice>=0))"), Right(num(0)))
    // a constant and a clock read through a name keep the tail
    assertEquals(e("=SUMPRODUCT(--(A:A<five))"), Right(num(1048575)))
    assertEquals(e("=SUMPRODUCT(--(A:A<asof))"), Right(num(1048576)))
  }

  test("whole rows: the blank line can sit before the used columns") {
    // C1:D1 filled, row 2 empty: the bound keeps columns C:D, so the blank line is column A
    val right = Sheet(SheetName.unsafe("R")).put(ref"C1", num(1)).put(ref"D1", num(2))
    assertEquals(eval("=SUMPRODUCT(--(1:1=\"\"))", right), Right(num(16382)))
    assertEquals(eval("=SUMPRODUCT(1:1)", right), Right(num(3)))
    val row = CellRange.parse("1:1").fold(e => fail(e), identity)
    val tail = SumProductTail
      .of(List(row), (None, Some((Column.from0(2), Column.from0(3)))))
      .getOrElse(fail("columns C:D kept, so a tail"))
    assertEquals((tail.count, tail.kept, tail.alongRows), (16382L, 2L, false))
    assertEquals(tail.line(row), CellRange(ARef.from0(0, 0), ARef.from0(0, 0)))
    // COUNTIFS over two whole rows: one shape, one bound, so the trimmed counts agree (Excel's
    // rule; LibreOffice 24.2's whole-row COUNTIF reads no blank cell)
    assertEquals(eval("=COUNTIFS(1:1,\"\",2:2,\"\")", right), Right(num(16382)))
  }

  test("a selector that may pick a bounded range beside a whole column has no tail") {
    // A1 = 0, B1:B5 = 1..5, C1:C5 = 10..50: the used rows are 1-5, the bounded branch is 5 rows,
    // so the extent alone would let the tail through; the plan declines it on its own
    val base = (1 to 5).foldLeft(Sheet(SheetName.unsafe("S")).put(ref"A1", num(0))) { (s, i) =>
      s.put(ARef.from0(1, i - 1), num(i.toDouble)).put(ARef.from0(2, i - 1), num(10.0 * i))
    }
    assertEquals(eval("=SUMPRODUCT(IF(A1>0,B:B,C1:C5))", base), Right(num(150)))
    assertEquals(
      eval("=SUMPRODUCT(IF(A1>0,B:B,C1:C5))", base.put(ref"A1", num(1))),
      Right(num(15))
    )
    @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
    val args = FormulaParser.parse("=SUMPRODUCT(IF(A1>0,B:B,C1:C5))") match
      case Right(call: TExpr.Call[?]) => call.args.asInstanceOf[List[ArgSpec.SumProductArg]]
      case other => fail(s"not a SUMPRODUCT call: $other")
    val column = CellRange.parse("B:B").fold(e => fail(e), identity)
    val tail = SumProductTail
      .of(List(column), (Some((Row.from0(0), Row.from0(4))), None))
      .getOrElse(fail("rows 1-5 kept, so a tail"))
    assertEquals(SumProductTail.plan(tail, args, 5, _ => true), None)
  }
