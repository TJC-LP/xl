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
 * #725: ROWS and COLUMNS of an argument that evaluates to an error value are that error, and INDEX
 * over a scalar treats it as a 1x1 array. #723 (a): IF takes two arguments, and a false test with
 * no value_if_false is the logical FALSE.
 *
 * Every expected value is desktop Excel's (Excel for Mac, formulas entered as dynamic-array
 * formulas, each value's type read with TYPE) over the same synthetic data: Data!A1:A3 = 1, 2, 3,
 * A4 blank, A5 `=""`, A6 `=NA()`, A7 TRUE, A8 "txt", B1:B3 = 10, 20, 30, K1:K3 = 1.5, 2, 3; names
 * `konst = 5`, `ktext = "t"`, `dynRef` and `dynCol` (OFFSET off the top and left edges) and `dyn =
 * OFFSET(Data!$K$1,0,0,3,1)`. A second Excel run added `errCell = Data!$A$6`, `arrNA = {#N/A}`,
 * `arrThree = {1,2,3}`, `arrComp = Data!$A$1:$A$3*2` and `seqName = SEQUENCE(3)`.
 */
class ScalarIndexAndIfSpec extends FunSuite:

  private def N(n: String): CellValue = CellValue.Number(BigDecimal(n))
  private def T(s: String): CellValue = CellValue.Text(s)
  private def B(b: Boolean): CellValue = CellValue.Bool(b)
  private def E(error: CellError): CellValue = CellValue.Error(error)

  private val dataName = SheetName.unsafe("Data")
  private val calcName = SheetName.unsafe("Calc")

  private val data: Sheet = Sheet(dataName)
    .put(ref"A1", N("1"))
    .put(ref"A2", N("2"))
    .put(ref"A3", N("3"))
    .put(ref"A5", CellValue.Formula("\"\"", Some(T(""))))
    .put(ref"A6", CellValue.Formula("NA()", Some(E(CellError.NA))))
    .put(ref"A7", B(true))
    .put(ref"A8", T("txt"))
    .put(ref"B1", N("10"))
    .put(ref"B2", N("20"))
    .put(ref"B3", N("30"))
    .put(ref"K1", N("1.5"))
    .put(ref"K2", N("2"))
    .put(ref"K3", N("3"))

  private val book: Workbook = Workbook(Vector(data, Sheet(calcName)))
    .withDefinedName("konst", "5")
    .withDefinedName("ktext", "\"t\"")
    .withDefinedName("dynRef", "OFFSET(Data!$A$1,-1,0,10,1)")
    .withDefinedName("dynCol", "OFFSET(Data!$A$1,0,-1,1,5)")
    .withDefinedName("dyn", "OFFSET(Data!$K$1,0,0,3,1)")
    .withDefinedName("errCell", "Data!$A$6")
    .withDefinedName("arrNA", "{#N/A}")
    .withDefinedName("arrThree", "{1,2,3}")
    .withDefinedName("arrComp", "Data!$A$1:$A$3*2")
    .withDefinedName("seqName", "SEQUENCE(3)")

  /** `formula` as a plain formula cell at Calc!G1, evaluated at its own position. */
  private def plain(formula: String): Either[String, CellValue] =
    val placed = Sheet(calcName).put(ref"G1", CellValue.Formula(formula, None))
    placed.evaluateCell(ref"G1", Clock.system, Some(book.put(placed))).left.map(_.message)

  /** `formula` evaluated without a cell position: an array context, as `xl eval` evaluates. */
  private def unplaced(formula: String): Either[String, CellValue] =
    book.evaluateFormula(s"=$formula", calcName).left.map(_.message)

  private def agree(obtained: Either[String, CellValue], expected: CellValue): Unit =
    (obtained, expected) match
      case (Right(CellValue.Number(got)), CellValue.Number(want)) =>
        assert((got - want).abs < BigDecimal("1e-9"), s"$got is not $want")
      case _ => assertEquals(obtained, Right(expected))

  /** Cases whose value is the same in a plain cell and in an array context. */
  private val bothModes: List[(String, CellValue)] = List(
    // ----- #725 (1): ROWS and COLUMNS of an error -----
    "ROWS(OFFSET(Data!A1,-1,0,10,1))" -> E(CellError.Ref),
    "COLUMNS(OFFSET(Data!A1,0,-1,1,5))" -> E(CellError.Ref),
    "ROWS(dynRef)" -> E(CellError.Ref),
    "COLUMNS(dynCol)" -> E(CellError.Ref),
    "ROWS(INDIRECT(\"nosuch!A1\"))" -> E(CellError.Ref),
    "ROWS(#REF!)" -> E(CellError.Ref),
    "ROWS(#N/A)" -> E(CellError.NA),
    "ROWS(NA())" -> E(CellError.NA),
    "COLUMNS(1/0)" -> E(CellError.Div0),
    // a reference to a cell holding an error is still a reference, and an array keeps its shape
    "ROWS(Data!A6)" -> N("1"),
    "ROWS(Data!A1:A9)" -> N("9"),
    "ROWS({1,#N/A;2,3})" -> N("2"),
    // a 1x1 array holding an error keeps its shape, written, named or LET-bound
    "ROWS({#N/A})" -> N("1"),
    "COLUMNS({#N/A})" -> N("1"),
    "ROWS(arrNA)" -> N("1"),
    "LET(x,{#N/A},ROWS(x))" -> N("1"),
    "ROWS(LET(x,{#N/A},x))" -> N("1"),
    // a LET whose body is a reference to an error cell is that reference
    "ROWS(LET(x,Data!A6,x))" -> N("1"),
    "COLUMNS(LET(x,Data!A6,x))" -> N("1"),
    "SUM(ROWS(LET(x,Data!A6,x)))" -> N("1"),
    "ROWS(LET(x,errCell,x))" -> N("1"),
    "LET(x,Data!A6,ROWS(x))" -> N("1"),
    "ROWS(IF(TRUE,Data!A6))" -> N("1"),
    "ROWS(errCell)" -> N("1"),
    "SUM(LET(x,Data!A1,x))" -> N("1"),
    // an error value a selector or INDEX gives is the error
    "ROWS(IF(TRUE,NA()))" -> E(CellError.NA),
    "ROWS(INDEX({#N/A},1))" -> E(CellError.NA),
    "COLUMNS({1,#N/A;2,3})" -> N("2"),
    "ROWS(5)" -> N("1"),
    "COLUMNS(\"\")" -> N("1"),
    "ROWS(LET(x,OFFSET(Data!A1,-1,0,2,1),x))" -> E(CellError.Ref),
    "IFERROR(ROWS(OFFSET(Data!A1,-1,0,10,1)),\"e\")" -> T("e"),
    "ROWS(OFFSET(Data!A1,0,0,10,1))" -> N("10"),
    "ROWS(Data!A1:A3 Data!B5:B7)" -> E(CellError.Null),
    "ROWS(OFFSET(Data!A1,1048575,0,2,1))" -> E(CellError.Ref),
    // ----- #725 (2): INDEX over a scalar is INDEX over a 1x1 array -----
    "INDEX(5,1)" -> N("5"),
    "INDEX(5,1,1)" -> N("5"),
    "INDEX(5,2)" -> E(CellError.Ref),
    "INDEX(5,1,2)" -> E(CellError.Ref),
    "INDEX(5,0)" -> N("5"),
    "INDEX(5,0,0)" -> N("5"),
    "INDEX(5,-1)" -> E(CellError.Value),
    "INDEX(5,\"a\")" -> E(CellError.Value),
    "INDEX(5,1.9)" -> N("5"),
    "INDEX(\"t\",1)" -> T("t"),
    "INDEX(\"\",1)" -> T(""),
    "INDEX(TRUE,1,1)" -> B(true),
    "INDEX(#N/A,1)" -> E(CellError.NA),
    "INDEX(1+2,1)" -> N("3"),
    "INDEX(konst,1)" -> N("5"),
    "INDEX(ktext,1,1)" -> T("t"),
    "LET(x,5,INDEX(x,1))" -> N("5"),
    "LET(x,5,INDEX(x,1,1))" -> N("5"),
    "LET(x,5,INDEX(x,2))" -> E(CellError.Ref),
    "LET(x,\"t\",INDEX(x,1)&\"!\")" -> T("t!"),
    "INDEX(5,1,1,1)" -> N("5"),
    "INDEX(5,1,1,2)" -> E(CellError.Ref),
    "INDEX(5,,1)" -> N("5"),
    "INDEX(SUM(Data!B1:B3),1)" -> N("60"),
    "INDEX({7},1)" -> N("7"),
    "INDEX(5,2,1)" -> E(CellError.Ref),
    "ROWS(INDEX(5,1))" -> N("1"),
    "INDEX(5,1,1,0)" -> E(CellError.Value),
    "COLUMNS(INDEX(5,0,0))" -> N("1"),
    "INDEX(5,TRUE)" -> N("5"),
    "INDEX(5,1,0)" -> N("5"),
    "INDEX(5,0,1)" -> N("5"),
    "INDEX(5,\"\")" -> E(CellError.Value),
    // a defined name bound to an array, constant or computed
    "INDEX(arrThree,2)" -> N("2"),
    "INDEX(arrThree,1,2)" -> N("2"),
    "INDEX(arrComp,2)" -> N("4"),
    "INDEX(seqName,2)" -> N("2"),
    // ----- #723 (a): two-argument IF -----
    "IF(TRUE,5)" -> N("5"),
    "IF(FALSE,5)" -> B(false),
    "IF(0,5)" -> B(false),
    "IF(Data!A4,5)" -> B(false),
    "IF(Data!A5,5)" -> E(CellError.Value),
    "IF(\"\",5)" -> E(CellError.Value),
    "IF(Data!A6,5)" -> E(CellError.NA),
    "IF(FALSE,5)&\"\"" -> T("FALSE"),
    "IF(FALSE,5)+1" -> N("1"),
    "IF(TRUE,)" -> N("0"),
    "IF(FALSE,)" -> B(false),
    "SUMPRODUCT(--IF({TRUE,FALSE,TRUE},{1,2,4}))" -> N("5"),
    "IF(\"TRUE\",5)" -> N("5"),
    "IF(Data!A8,5)" -> E(CellError.Value),
    "IF(1,Data!A4)" -> N("0"),
    "IF(TRUE,Data!A5)&\"|\"" -> T("|"),
    "SUMPRODUCT((Data!A1:A3>0)*IF(TRUE,ROW(Data!A1:A3)))" -> N("6"),
    "ROWS(IF(TRUE,OFFSET(Data!A1,-1,0,2,1)))" -> E(CellError.Ref),
    "IF(TRUE,5,)" -> N("5"),
    "IF(FALSE,5,)" -> N("0"),
    // ----- #753: a range-slot read of one LET name leaves another binding's array value -----
    "LET(c,IF(TRUE,Data!A1:A3,0),r,dyn,SUMPRODUCT(c*1)+COUNTIF(r,\">2\"))" -> N("7")
  )

  bothModes.foreach { case (formula, expected) =>
    test(s"=$formula is $expected in a plain cell and in an array context") {
      agree(plain(formula), expected)
      agree(unplaced(formula), expected)
    }
  }

  /**
   * Array formulas: an array or range in IF's test spreads over its elements, and an array row_num
   * spreads INDEX. A plain cell is Excel's legacy formula, which intersects a range and reads IF's
   * array test by its first element (`SUM(IF({TRUE,FALSE,TRUE},{1,2,4}))` is 7 there).
   */
  private val arrayContext: List[(String, CellValue)] = List(
    // SUM and COUNT skip the FALSE an unmatched element gives
    "SUM(IF({TRUE,FALSE,TRUE},{1,2,4}))" -> N("5"),
    "COUNT(IF({TRUE,FALSE,TRUE},{1,2,4}))" -> N("2"),
    "SUM(IF(Data!A1:A3>1,Data!B1:B3))" -> N("50"),
    "MAX(IF(Data!A1:A3<3,Data!B1:B3))" -> N("20"),
    // COUNTA counts the FALSE an unmatched row gives
    "COUNTA(IF(Data!A1:A3>1,Data!B1:B3))" -> N("3"),
    "SUM(INDEX(5,{1,1}))" -> N("10"),
    "SUMPRODUCT(INDEX(5,{1,1}))" -> N("10"),
    "SUM(INDEX(5,{1,2}))" -> E(CellError.Ref),
    "SUM(INDEX({5},{1,1}))" -> N("10")
  )

  arrayContext.foreach { case (formula, expected) =>
    test(s"=$formula is $expected in an array context") {
      agree(unplaced(formula), expected)
    }
  }

  test("two-argument IF and INDEX over a scalar print and re-parse as written") {
    List("IF(A1>0,5)", "IF(FALSE,)", "INDEX(5,1)", "INDEX(1+2,1,1)", "INDEX(\"t\",1)").foreach {
      formula =>
        val parsed = FormulaParser.parse(s"=$formula").fold(e => fail(s"$formula: $e"), identity)
        assertEquals(FormulaPrinter.printFileForm(parsed), formula)
    }
  }

  test("INDEX's first slot still refuses a reference-returning call (#723 part b)") {
    assert(FormulaParser.parse("=INDEX(OFFSET(A1,0,0,3,1),2)").isLeft)
  }
