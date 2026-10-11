package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.parser.FormulaParser
import com.tjclp.xl.formula.printer.FormulaPrinter
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * GH-710: a dynamic named range — a name whose formula computes a reference (OFFSET, INDIRECT,
 * INDEX, IF or CHOOSE over references) — and a LET name bound to a reference are references in
 * every slot that takes a range: the criteria and lookup functions, INDEX, MATCH, LARGE, RANK, NPV,
 * TRANSPOSE. They used to be `#VALUE!` ("does not refer to a range") or a parse error.
 *
 * The first table is LibreOffice 24.2's recalculation of the same plain cells (Empty!A1, A2, … in
 * order). LibreOffice 24.2 has neither LET, MAXIFS, MINIFS, XMATCH nor XLOOKUP: those cases are
 * pinned to the value of the same formula written with the reference itself, which is what Excel's
 * LET (it binds the reference) and those functions compute.
 */
class DynamicNameRangeSlotSpec extends FunSuite:

  private def N(n: String): CellValue = CellValue.Number(BigDecimal(n))
  private def T(s: String): CellValue = CellValue.Text(s)
  private def E(error: CellError): CellValue = CellValue.Error(error)

  private val a = Vector("1", "-2", "3.5", "0", "10", "7", "-4.25", "100", "2", "5")
  private val c = Vector(
    "apple",
    "banana",
    "cherry",
    "date",
    "elder",
    "fig",
    "grape",
    "honeydew",
    "iceberg",
    "jam"
  )
  private val d = Vector(1, 2, 3, 1, 2, 3, 1, 2, 3, 1)

  private val sheet1Name = SheetName.unsafe("Sheet1")
  private val emptyName = SheetName.unsafe("Empty")

  /** Sheet1: A1:A10 numbers, B1:B10 = 2, C1:C10 text, D1:D10 a key column, A12:E12 = 1..5. */
  private val sheet1: Sheet =
    val columns = (0 until 10).foldLeft(Sheet(sheet1Name)) { (s, i) =>
      s.put(ARef.from0(0, i), N(a(i)))
        .put(ARef.from0(1, i), N("2"))
        .put(ARef.from0(2, i), T(c(i)))
        .put(ARef.from0(3, i), CellValue.Number(BigDecimal(d(i))))
    }
    (0 until 5).foldLeft(columns) { (s, j) =>
      s.put(ARef.from0(j, 11), CellValue.Number(BigDecimal(j + 1)))
    }

  private val other: Sheet = (0 until 10).foldLeft(Sheet(SheetName.unsafe("Other"))) { (s, i) =>
    s.put(ARef.from0(0, i), CellValue.Number(BigDecimal((i + 1) * 10)))
  }

  private def withNames(wb: Workbook): Workbook =
    wb.withDefinedName("dyn", "OFFSET(Sheet1!$A$1,0,0,10,1)")
      .withDefinedName("dynTab", "OFFSET(Sheet1!$A$1,0,0,10,4)")
      .withDefinedName("dynRow", "OFFSET(Sheet1!$A$12,0,0,1,5)")
      .withDefinedName("dynKey", "OFFSET(Sheet1!$D$1,0,0,10,1)")
      .withDefinedName("dynTop", "OFFSET(Sheet1!$A$1,0,0,1,1)")
      .withDefinedName("dynOther", "OFFSET(Other!$A$1,0,0,10,1)")
      .withDefinedName("dynInd", "INDIRECT(\"Sheet1!A1:A10\")")
      .withDefinedName("dynIdx", "INDEX(Sheet1!$A$1:$D$10,0,1)")
      .withDefinedName("dynIf", "IF(TRUE,Sheet1!$A$1:$A$10,Sheet1!$B$1:$B$10)")
      .withDefinedName("dynCh", "CHOOSE(2,Sheet1!$B$1:$B$10,Sheet1!$A$1:$A$10)")
      .withDefinedName("dynBad", "INDIRECT(\"nope!A1:A3\")")
      .withDefinedName("dynRef", "OFFSET(Sheet1!$A$1,-1,0,10,1)")
      .withDefinedName("alias", "dyn")
      .withDefinedName("konst", "5")

  private val book: Workbook = withNames(Workbook(Vector(sheet1, other, Sheet(emptyName))))

  /** `formula` as a plain formula cell at `sheet!at`, evaluated at its own position. */
  private def plain(sheet: String, at: String, formula: String): Either[String, CellValue] =
    val ref = ARef.parse(at).fold(msg => fail(msg), identity)
    book.sheets.find(_.name.value == sheet) match
      case None => Left(s"no sheet $sheet")
      case Some(target) =>
        val placed = target.put(ref, CellValue.Formula(formula, None))
        placed.evaluateCell(ref, Clock.system, Some(book.put(placed))).left.map(_.message)

  /** `formula` evaluated without a cell position — an array context, as `xl eval` evaluates. */
  private def unplaced(formula: String): Either[String, CellValue] =
    book.evaluateFormula(s"=$formula", sheet1Name).left.map(_.message)

  private def agree(obtained: Either[String, CellValue], expected: CellValue): Unit =
    (obtained, expected) match
      case (Right(CellValue.Number(got)), CellValue.Number(want)) =>
        assert((got - want).abs < BigDecimal("1e-9"), s"$got is not $want")
      case _ => assertEquals(obtained, Right(expected))

  private val libreOffice: List[(String, CellValue)] = List(
    "COUNTIF(dyn,\">2\")" -> N("5"),
    "INDEX(dyn,3)" -> N("3.5"),
    "MATCH(7,dyn,0)" -> N("6"),
    "VLOOKUP(7,dyn,1,FALSE)" -> N("7"),
    "VLOOKUP(7,dynTab,3,FALSE)" -> T("fig"),
    "HLOOKUP(3,dynRow,1,FALSE)" -> N("3"),
    "INDEX(dynTab,6,3)" -> T("fig"),
    "SUMIF(dyn,\">2\")" -> N("125.5"),
    "SUMIF(dynKey,2,dyn)" -> N("108"),
    "AVERAGEIF(dyn,\">2\")" -> N("25.1"),
    "AVERAGEIF(dynKey,1,dyn)" -> N("0.4375"),
    "COUNTIFS(dyn,\">0\",dynKey,1)" -> N("2"),
    "SUMIFS(dyn,dynKey,1)" -> N("1.75"),
    "AVERAGEIFS(dyn,dynKey,1)" -> N("0.4375"),
    "LARGE(dyn,2)" -> N("10"),
    "SMALL(dyn,1)" -> N("-4.25"),
    "PERCENTILE(dyn,0.5)" -> N("2.75"),
    "QUARTILE(dyn,1)" -> N("0.25"),
    "RANK(7,dyn)" -> N("3"),
    "RANK(-4.25,dyn,1)" -> N("1"),
    "NPV(0.1,dyn)" -> N("59.29205859357761"),
    "SUM(TRANSPOSE(dyn))" -> N("122.25"),
    "COUNTIF(dynInd,\">2\")" -> N("5"),
    "COUNTIF(dynIdx,\">2\")" -> N("5"),
    "COUNTIF(dynIf,\">2\")" -> N("5"),
    "COUNTIF(dynCh,\">2\")" -> N("5"),
    "MATCH(7,dynInd,0)" -> N("6"),
    "INDEX(dynIf,3)" -> N("3.5"),
    "COUNTIF(dynOther,\">20\")" -> N("8"),
    "MATCH(30,dynOther,0)" -> N("3"),
    "MATCH(3,dynRow,0)" -> N("3"),
    "SUMIF(Sheet1!A1:A10,\">2\",dyn)" -> N("125.5"),
    // the sum range is resized to the criteria range's shape, from the reference's top-left cell
    "SUMIF(Sheet1!D1:D10,1,dynTop)" -> N("1.75"),
    "COUNTIF(dynTab,\">2\")" -> N("8"),
    "COUNTIF(dynKey,1)" -> N("4"),
    "HLOOKUP(1,dyn,3,FALSE)" -> N("3.5"),
    "COUNTIF(alias,\">2\")" -> N("5"),
    "NPV(0.1,alias)" -> N("59.29205859357761"),
    // a name that computes no reference is still #VALUE!
    "COUNTIF(konst,\">2\")" -> E(CellError.Value),
    // the error the name's formula evaluates to is the slot's error
    "COUNTIF(dynBad,1)" -> E(CellError.Ref)
  )

  libreOffice.zipWithIndex.foreach { case ((formula, expected), i) =>
    val at = s"A${i + 1}"
    test(s"=$formula at Empty!$at is $expected") {
      agree(plain("Empty", at, formula), expected)
    }
  }

  test("INDEX's column of a dynamic name in a plain cell's value position is intersected") {
    // LibreOffice: Empty!A8 reads row 8 of the column INDEX returns
    agree(plain("Empty", "A8", "INDEX(dyn,0,1)"), N("100"))
    agree(plain("Empty", "A9", "SUM(INDEX(dyn,0,1))"), N("122.25"))
  }

  test("a name computing an off-grid reference is #REF! in a range slot, as OFFSET is") {
    // Excel's OFFSET off the grid is #REF! (LibreOffice says #VALUE! for OFFSET itself, and so
    // for COUNTIF over it): the slot carries the error the reference evaluated to
    agree(plain("Empty", "A1", "COUNTIF(dynRef,\">2\")"), E(CellError.Ref))
    agree(plain("Empty", "A2", "INDEX(dynRef,3)"), E(CellError.Ref))
  }

  test("MAXIFS, MINIFS, XMATCH and XLOOKUP take a dynamic name like a written range") {
    agree(plain("Empty", "A1", "MAXIFS(dyn,dynKey,1)"), N("5"))
    agree(plain("Empty", "A2", "MINIFS(dyn,dynKey,1)"), N("-4.25"))
    agree(plain("Empty", "A3", "MAXIFS(Sheet1!A1:A10,Sheet1!D1:D10,1)"), N("5"))
    agree(plain("Empty", "A4", "XMATCH(7,dyn)"), N("6"))
    agree(plain("Empty", "A5", "XLOOKUP(7,dyn,dynKey)"), N("3"))
  }

  test("a LET name bound to a dynamic name is a reference in a range slot") {
    agree(plain("Empty", "A1", "LET(r,dyn,COUNTIF(r,\">2\"))"), N("5"))
    agree(plain("Empty", "A2", "LET(r,dyn,INDEX(r,3))"), N("3.5"))
    agree(plain("Empty", "A3", "LET(r,dyn,MATCH(7,r,0))"), N("6"))
    agree(plain("Empty", "A4", "LET(t,dynTab,VLOOKUP(7,t,3,FALSE))"), T("fig"))
    agree(plain("Empty", "A5", "LET(r,dyn,k,dynKey,SUMIFS(r,k,1))"), N("1.75"))
  }

  test("a LET name bound to OFFSET, INDIRECT, INDEX or a written range is a reference too") {
    agree(plain("Sheet1", "H1", "LET(r,OFFSET(A1,0,0,10,1),COUNTIF(r,\">2\"))"), N("5"))
    agree(plain("Sheet1", "H2", "LET(r,INDIRECT(\"A1:A10\"),LARGE(r,2))"), N("10"))
    agree(plain("Sheet1", "H3", "LET(r,INDEX(A1:D10,0,1),SUMIF(r,\">2\"))"), N("125.5"))
    agree(plain("Sheet1", "H4", "LET(r,A1:A10,COUNTIF(r,\">2\"))"), N("5"))
    agree(plain("Sheet1", "H5", "LET(x,A1,COUNTIF(x,1))"), N("1"))
    // nested LETs: the inner name is bound to the outer one
    agree(plain("Sheet1", "H6", "LET(r,dyn,LET(s,r,COUNTIF(s,\">2\")))"), N("5"))
  }

  test("INDEX over a LET name bound to an array indexes the array") {
    agree(plain("Empty", "A1", "LET(x,SEQUENCE(3),INDEX(x,2))"), N("2"))
    agree(unplaced("LET(x,SEQUENCE(3,2),INDEX(x,3,2))"), N("6"))
    agree(unplaced("LET(r,dyn,INDEX(r,3))"), N("3.5"))
    agree(unplaced("LET(r,dyn,AREAS(r))"), N("1"))
  }

  test("a LET name bound to a value is #VALUE! in a range slot, as a name bound to one is") {
    agree(plain("Empty", "A1", "LET(r,5,COUNTIF(r,\">2\"))"), E(CellError.Value))
    agree(plain("Empty", "A2", "LET(r,konst,COUNTIF(r,\">2\"))"), E(CellError.Value))
  }

  test("without a cell position (an array context) dynamic names and LET names resolve too") {
    agree(unplaced("COUNTIF(dyn,\">2\")"), N("5"))
    agree(unplaced("INDEX(dyn,3)"), N("3.5"))
    agree(unplaced("VLOOKUP(7,dynTab,3,FALSE)"), T("fig"))
    agree(unplaced("LET(r,dyn,COUNTIF(r,\">2\"))"), N("5"))
    agree(unplaced("LET(r,OFFSET(A1,0,0,10,1),COUNTIF(r,\">2\")+SUM(r))"), N("127.25"))
    agree(unplaced("SUMPRODUCT(LET(r,dyn,COUNTIF(r,r)))"), N("10"))
  }

  test("a LET name in a range slot prints and re-parses as written") {
    val text = "=LET(r,dyn,COUNTIF(r,\">2\"))"
    val parsed = FormulaParser.parse(text).fold(err => fail(err.toString), identity)
    val printed = FormulaPrinter.printFileForm(parsed)
    assertEquals(printed, "LET(r,dyn,COUNTIF(r,\">2\"))")
    assertEquals(FormulaParser.parse(s"=$printed"), Right(parsed))
  }

  test("a name whose reference reads itself through a range slot is a clean error") {
    val looped = book.withDefinedName("loop", "OFFSET(Sheet1!$A$1,0,0,COUNTIF(loop,\">2\"),1)")
    val placed = Sheet(emptyName).put(ref"A1", CellValue.Formula("COUNTIF(loop,\">2\")", None))
    val result = placed.evaluateCell(ref"A1", Clock.system, Some(looped.put(placed)))
    assert(result.isLeft || result.exists(_.isInstanceOf[CellValue.Error]), result.toString)
  }

  test("recalculation reads a dynamic range's precedent formulas before the range slot") {
    // A2 is an uncached formula inside the dynamic range: the readers must see its value (20)
    val withFormula = sheet1
      .put(ref"A2", CellValue.Formula("B1*10", None))
      .put(ref"H1", CellValue.Formula("COUNTIF(dyn,\">2\")", None))
      .put(ref"H2", CellValue.Formula("LET(r,OFFSET(A1,0,0,10,1),COUNTIF(r,\">2\"))", None))
      .put(ref"H3", CellValue.Formula("MATCH(20,dyn,0)", None))
    val result = withNames(Workbook(Vector(withFormula, other, Sheet(emptyName)))).recalculate()
    assert(result.isClean, result.errors.map(_.render).mkString("; "))
    assertEquals(result.evaluated(sheet1Name)(ref"H1"), N("6"))
    assertEquals(result.evaluated(sheet1Name)(ref"H2"), N("6"))
    assertEquals(result.evaluated(sheet1Name)(ref"H3"), N("2"))
  }
