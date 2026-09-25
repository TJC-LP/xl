package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.eval.ArrayResult
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * GH-669: formulas Excel and LibreOffice open that the parser refused — `TRUE()`/`FALSE()` as
 * calls, a bare `NOT` name, a range whose two ends repeat the sheet (`Sheet1!A1:Sheet1!B2`) and
 * array constants (`{1,2;3,4}`). Each parses, prints back as written and evaluates; every value
 * pinned against LibreOffice is its recalculation of the same plain `<f>` cell (the cell each
 * formula sits in is part of the case). Where LibreOffice departs from Excel the row says so.
 */
@SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
class GrammarGapsSpec extends FunSuite:

  private def N(n: String): CellValue = CellValue.Number(BigDecimal(n))
  private def T(s: String): CellValue = CellValue.Text(s)
  private def B(b: Boolean): CellValue = CellValue.Bool(b)
  private def E(code: String): CellValue =
    CellValue.Error(CellError.parse(code).fold(msg => fail(msg), identity))

  /** Sheet1: A1:A3 = 1,2,3; B1:B2 = 10,20. Other: A1 = 100. The name NOT refers to Sheet1!$A$3. */
  private val book: Workbook =
    val sheet1 = Sheet(SheetName.unsafe("Sheet1"))
      .put(ARef.from1(1, 1), N("1"))
      .put(ARef.from1(1, 2), N("2"))
      .put(ARef.from1(1, 3), N("3"))
      .put(ARef.from1(2, 1), N("10"))
      .put(ARef.from1(2, 2), N("20"))
    val other = Sheet(SheetName.unsafe("Other")).put(ARef.from1(1, 1), N("100"))
    Workbook(Vector(sheet1, other)).withDefinedName("NOT", "Sheet1!$A$3")

  /** `formula` as a plain formula cell at `Sheet1!at`, evaluated at its own position. */
  private def plain(at: String, formula: String): Either[String, CellValue] =
    val ref = ARef.parse(at).fold(msg => fail(msg), identity)
    val target = book.sheets.headOption.getOrElse(fail("no sheet"))
    val placed = target.put(ref, CellValue.Formula(formula, None))
    placed.evaluateCell(ref, Clock.system, Some(book.put(placed))).left.map(_.message)

  private def parsed(formula: String): TExpr[?] =
    FormulaParser.parse(formula).fold(err => fail(s"$formula should parse: $err"), identity)

  /** Parses, prints back byte-for-byte, and the printed text re-parses to the same tree. */
  private def assertPreserved(formula: String): TExpr[?] =
    val expr = parsed(formula)
    assertEquals(FormulaPrinter.print(expr), formula)
    assertEquals(FormulaParser.parse(FormulaPrinter.printFileForm(expr)), Right(expr))
    expr

  private def assertRefused(formula: String): ParseError =
    FormulaParser.parse(formula).swap.getOrElse(fail(s"$formula should not parse"))

  // ===== TRUE() / FALSE() =====

  test("GH-669: TRUE() and FALSE() are zero-argument calls that print back as written") {
    assertPreserved("=TRUE()")
    assertPreserved("=FALSE()")
    assertPreserved("=IF(FALSE(), 1, 2)")
    assertPreserved("=TRUE()+1")
    // the bare constants keep their literal form
    assertEquals(parsed("=TRUE"), TExpr.Lit(true))
    assertPreserved("=TRUE")
    assert(FunctionRegistry.lookup("TRUE").isDefined && FunctionRegistry.lookup("FALSE").isDefined)
    assertRefused("=TRUE(1)") match
      case ParseError.InvalidArguments("TRUE", _, _, _) => ()
      case other => fail(s"TRUE(1) should be an arity error, got $other")
  }

  // ===== NOT as a name =====

  test("GH-669: a bare NOT at the end or before a binary operator is a name, not the operator") {
    for formula <- List("=NOT", "=not", "=NOT+1", "=NOT*2", "=NOT&\"x\"", "=NOT=3")
    do assertPreserved(formula)
    assertEquals(parsed("=NOT"), TExpr.NameRef("NOT"))
    assertEquals(parsed("=(NOT)"), TExpr.NameRef("NOT"))
    assertPreserved("=IF(NOT, 1, 2)")
    // the prefix operator and the function are unchanged
    parsed("=NOT A1") match
      case TExpr.Call(spec, _) => assertEquals(spec.name, "NOT")
      case other => fail(s"NOT A1 should stay the operator, got $other")
    parsed("=NOT(A1)") match
      case TExpr.Call(spec, _) => assertEquals(spec.name, "NOT")
      case other => fail(s"NOT(A1) should stay the call, got $other")
    // a defined name spelled NOT resolves
    assertEquals(plain("H1", "NOT+1"), Right(N("4")))
    assertEquals(plain("H2", "NOT"), Right(N("3")))
  }

  // ===== Sheet1!A1:Sheet1!B2 =====

  test("GH-669: a range whose end repeats the sheet is the same range, printed canonically") {
    val expected = parsed("=Sheet1!A1:B2")
    assertEquals(parsed("=Sheet1!A1:Sheet1!B2"), expected)
    assertEquals(parsed("=Sheet1!A1:sheet1!B2"), expected)
    assertEquals(FormulaPrinter.print(parsed("=Sheet1!A1:Sheet1!B2")), "=Sheet1!A1:B2")
    assertEquals(parsed("='My Sheet'!A1:'My Sheet'!B2"), parsed("='My Sheet'!A1:B2"))
    assertEquals(parsed("=SUM(Sheet1!$A$1:Sheet1!B2)"), parsed("=SUM(Sheet1!$A$1:B2)"))
    // two different sheets is a 3-D reference, which xl does not implement: refused, never a
    // silently narrowed range
    assertRefused("=Sheet1!A1:Other!B2") match
      case ParseError.InvalidCellRef(_, _, reason) => assert(reason.contains("same sheet"), reason)
      case other => fail(s"expected InvalidCellRef, got $other")
  }

  // ===== Array constants =====

  test("GH-669: array constants parse and print back verbatim") {
    for formula <- List(
        "={1,2;3,4}",
        "={\"a\",\"b\"}",
        "={TRUE,FALSE}",
        "={1,-2,3.5}",
        "={#N/A,1;\"x\",FALSE}",
        "={1;2;3}",
        "={1.50}",
        "=SUM({1,2;3,4})",
        "=INDEX({1,2;3,4}, 2, 1)",
        "=SUM({1,2}*2)",
        "=A1:A3={1;2;3}"
      )
    do assertPreserved(formula)
    assertEquals(
      parsed("={1,2;3,4}"),
      TExpr.Lit(ArrayResult(Vector(Vector(N("1"), N("2")), Vector(N("3"), N("4")))))
    )
    // the elements are constants — Excel's own rule — and the rows must agree in width
    for formula <- List("={1,2;3}", "={}", "={1,,2}", "={A1}", "={1+1}", "={1,2", "={(1)}")
    do assertRefused(formula)
  }

  test("GH-669: array constants evaluate as arrays") {
    val grid = ArrayResult(Vector(Vector(N("1"), N("2")), Vector(N("3"), N("4"))))
    assertEquals(book.sheets.headOption.map(_.evaluateFormula("={1,2;3,4}")), Some(Right(N("1"))))
    def array(formula: String): Either[EvalError, Any] =
      Evaluator.arrayInstance.eval(parsed(formula).asInstanceOf[TExpr[Any]], book.sheets(0))
    assertEquals(array("={1,2;3,4}"), Right(grid))
    assertEquals(array("={1,2}*10"), Right(ArrayResult(Vector(Vector(N("10"), N("20"))))))
  }

  // ===== LibreOffice oracle =====

  private val libreOffice: List[(String, String, CellValue)] = List(
    ("H1", "TRUE()", B(true)),
    ("H2", "FALSE()", B(false)),
    ("H3", "TRUE()+1", N("2")),
    ("H4", "IF(FALSE(),1,2)", N("2")),
    ("H5", "SUM({1,2;3,4})", N("10")),
    ("H6", "INDEX({1,2;3,4},2,1)", N("3")),
    ("H7", "{1,2}", N("1")),
    ("H8", "{\"a\",\"b\"}", T("a")),
    ("H9", "SUM({1,-2,3.5})", N("2.5")),
    ("H11", "ROWS({1,2;3,4;5,6})", N("3")),
    ("H12", "COLUMNS({1,2,3})", N("3")),
    ("H13", "SUM({1,2}*2)", N("6")),
    ("H15", "Sheet1!A1:Sheet1!B2", E("#VALUE!")),
    ("H16", "SUM(Sheet1!A1:Sheet1!B2)", N("33")),
    ("H18", "MAX({1,5,3})", N("5")),
    ("H19", "AND({TRUE,FALSE})", B(false)),
    ("H20", "{1;2}+10", N("11")),
    ("H22", "SUM({1,2;3,4})+{10,20}", N("20"))
  )

  test("GH-669: the new shapes agree with LibreOffice in a plain cell") {
    libreOffice.foreach { (at, formula, expected) =>
      assertEquals(plain(at, formula), Right(expected), s"$at: =$formula")
    }
    // the repeated-sheet range is the plain range: intersected in a row it crosses
    assertEquals(plain("A2", "Sheet1!B1:Sheet1!B2"), Right(N("20")))
  }

  test("GH-669: error elements of an array constant are values, as in Excel") {
    // LibreOffice 25.8 answers #N/A for the whole of COUNTA({1,"x",TRUE,#N/A}) and
    // ISERROR({#N/A}); Excel's array constants carry error VALUES — COUNTA counts one, ISERROR
    // sees one — so these rows pin Excel's semantics
    assertEquals(plain("H10", "COUNTA({1,\"x\",TRUE,#N/A})"), Right(N("4")))
    assertEquals(plain("H21", "ISERROR({#N/A})"), Right(B(true)))
    assertEquals(plain("H23", "{#DIV/0!,1}"), Right(E("#DIV/0!")))
  }

  test("GH-669 review: INDEX's refused first argument reports the call's own argument count") {
    FormulaParser.parse("=INDEX(A1:A3+1+1,3)") match
      case Left(ParseError.InvalidArguments("INDEX", _, _, got)) => assertEquals(got, "2 arguments")
      case other => fail(s"expected InvalidArguments for INDEX, got $other")
  }
