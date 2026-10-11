package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.ast.RangeForm
import com.tjclp.xl.formula.eval.ArrayResult
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * GH-713: the range operator `:` between reference-valued operands — `A1:INDEX(A:A,COUNTA(A:A))`,
 * `INDEX(r,1):INDEX(r,2)`, `B1:XLOOKUP(…)`, `Start:Finish` — evaluating to the bounding range of
 * its operands on one sheet.
 *
 * Every value marked LibreOffice is LibreOffice 24.2.7's recalculation of the same plain `<f>` cell
 * on the same grid. The shapes LibreOffice cannot answer (it has no XLOOKUP, and its xlsx import
 * refuses `A1:IF(…)`) follow Excel, each with its reason.
 */
@SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
class RangeOperatorSpec extends FunSuite:

  private def N(n: String): CellValue = CellValue.Number(BigDecimal(n))
  private def E(code: String): CellValue =
    CellValue.Error(CellError.parse(code).fold(msg => fail(msg), identity))

  /**
   * Sheet1: A1:A3 = 1,2,3, B1:B3 = 10,20,30, C1:C3 = 100,200,300. Other: A1:A3 = 1000,2000,3000.
   * Names: Start = Sheet1!$A$1, Finish = Sheet1!$B$2, OtherCell = Other!$A$1.
   */
  private val book: Workbook =
    val sheet1 = (1 to 3).foldLeft(Sheet(SheetName.unsafe("Sheet1"))) { (s, i) =>
      s.put(ARef.from1(1, i), N(i.toString))
        .put(ARef.from1(2, i), N((i * 10).toString))
        .put(ARef.from1(3, i), N((i * 100).toString))
    }
    val other = (1 to 3).foldLeft(Sheet(SheetName.unsafe("Other"))) { (s, i) =>
      s.put(ARef.from1(1, i), N((i * 1000).toString))
    }
    Workbook(Vector(sheet1, other))
      .withDefinedName("Start", "Sheet1!$A$1")
      .withDefinedName("Finish", "Sheet1!$B$2")
      .withDefinedName("OtherCell", "Other!$A$1")

  private def sheet1: Sheet = book.sheets.headOption.getOrElse(fail("no sheet"))

  /** `formula` as a plain formula cell at `Sheet1!at`, evaluated at its own position. */
  private def plain(at: String, formula: String): Either[String, CellValue] =
    val ref = ARef.parse(at).fold(msg => fail(msg), identity)
    val placed = sheet1.put(ref, CellValue.Formula(formula, None))
    placed.evaluateCell(ref, Clock.system, Some(book.put(placed))).left.map(_.message)

  /** `formula` evaluated as an array (evala). */
  private def array(formula: String): Either[EvalError, Any] =
    Evaluator.arrayInstance.eval(
      parsed(formula).asInstanceOf[TExpr[Any]],
      sheet1,
      workbook = Some(book)
    )

  private def parsed(formula: String): TExpr[?] =
    FormulaParser.parse(formula).fold(err => fail(s"$formula should parse: $err"), identity)

  /** Accepted; both printed forms re-parse to the same tree. */
  private def assertRoundTrip(formula: String): TExpr[?] =
    val expr = parsed(formula)
    assertEquals(FormulaParser.parse(FormulaPrinter.printFileForm(expr)), Right(expr), formula)
    assertEquals(FormulaParser.parse(FormulaPrinter.print(expr)), Right(expr), formula)
    expr

  private def refused(formula: String): ParseError =
    FormulaParser.parse(formula).swap.getOrElse(fail(s"$formula should not parse"))

  // ===== Parser =====

  test("GH-713: computed ranges parse and round-trip in both printed forms") {
    List(
      "=SUM(A1:INDEX(A:A,COUNTA(A:A)))",
      "=SUM($A$1:INDEX($A:$A,3))",
      "=SUM(INDEX(B1:B3,1):INDEX(B1:B3,2))",
      "=SUM(B1:XLOOKUP(2,A1:A3,B1:B3))",
      "=SUM(B1:_xlfn.XLOOKUP(2,A1:A3,B1:B3))",
      "=SUM(Sheet1!A1:INDEX(Sheet1!A:A,3))",
      "=SUM('My Sheet'!A1:INDEX('My Sheet'!A:A,3))",
      "=SUM(Start:Finish)",
      "=SUM(A1:IF(TRUE,B2,C3))",
      "=SUM(A1:CHOOSE(2,B1,C3))",
      "=SUM(A1:OFFSET(A1,2,1))",
      "=SUM(A1:INDIRECT(\"B2\"))",
      "=SUM(A1:B1:C3)",
      "=SUM((A1:C3 B2:D4):E5)",
      "=A1:INDEX(B1:B9,3) B1:B9",
      "=-A1:INDEX(B1:B3,3)",
      "=@A1:INDEX(B1:B3,3)",
      "=LET(x,A1,SUM(x:INDEX(A:A,3)))",
      "=INDEX(A1:INDEX(C1:C3,3),2,3)",
      "=SUM(#REF!:INDEX(B1:B3,2))",
      "=SUM(A1:#REF!)",
      "=SUM(INDEX(B1:B3,1):B5:C6)",
      "=SUM((INDEX(B1:B3,1):B5):C6)",
      "=SUM(A1:(B2:C3))",
      "=SUM(INDEX(B1:B3,1):B5:INDEX(C1:C3,2))",
      "=SUM((INDEX(B1:B3,1):B5):INDEX(C1:C3,2))",
      "=SUM(A1:INDEX(B1:B3,1):INDEX(C1:C3,2))",
      "=SUM(A1:Start:Finish)",
      "=SUM(Sheet1!A1:B2:INDEX(C1:C3,2))",
      "=SUM(Sheet1!A1:Start:Finish)",
      "=ROWS(A1:INDEX(A:A,3))%",
      "=SUM((A1,C3):A1)"
    ).foreach(assertRoundTrip)
  }

  test("GH-713: every unparenthesized computed range prints back byte-for-byte") {
    List(
      "SUM(A1:INDEX(A:A,COUNTA(A:A)))",
      "SUM(INDEX(B1:B3,1):INDEX(B1:B3,2))",
      "SUM(Sheet1!A1:INDEX(Sheet1!A:A,3))",
      "SUM('My Sheet'!A1:INDEX('My Sheet'!A:A,3))",
      "SUM(Start:Finish)",
      "SUM(A1:B1:C3)",
      "SUM(INDEX(B1:B3,1):B5:C6)",
      "SUM(INDEX(B1:B3,1):B5:INDEX(C1:C3,2))",
      "SUM(A1:INDEX(B1:B3,1):INDEX(C1:C3,2))",
      "SUM(A1:Start:Finish)",
      "A1:INDEX(B1:B9,3) B1:B9",
      "-A1:INDEX(B1:B3,3)",
      "SUM(#REF!:INDEX(B1:B3,2))",
      "SUM(A1:#REF!)"
    ).foreach { text =>
      assertEquals(FormulaPrinter.printFileForm(parsed(s"=$text")), text, text)
    }
    // the grouping parentheses the printer must keep, and the one fold
    assertEquals(
      FormulaPrinter.printFileForm(parsed("=SUM((INDEX(B1:B3,1):B5):C6)")),
      "SUM((INDEX(B1:B3,1):B5):C6)"
    )
    assertEquals(
      FormulaPrinter.printFileForm(parsed("=SUM((INDEX(B1:B3,1):B5):INDEX(C1:C3,2))")),
      "SUM((INDEX(B1:B3,1):B5):INDEX(C1:C3,2))"
    )
    assertEquals(FormulaPrinter.printFileForm(parsed("=SUM(A1:(B2:C3))")), "SUM(A1:(B2:C3))")
    assertEquals(parsed("=SUM((A1):B2)"), parsed("=SUM(A1:B2)"))
    // `@` keeps its operand whole for the stored form's `_xlfn.SINGLE(…)`
    assertEquals(
      FormulaPrinter.printFileForm(parsed("=@A1:INDEX(B1:B3,3)")),
      "@(A1:INDEX(B1:B3,3))"
    )
  }

  test("GH-713: still-illegal texts keep their exact diagnostics") {
    // each of these was refused before the range operator, with this very error
    val before = Map(
      "=A1:B" -> FormulaParser.parse("=A1:B"),
      "=A1:1+2" -> FormulaParser.parse("=A1:1+2"),
      "=A1:\"x\"" -> FormulaParser.parse("=A1:\"x\""),
      "=A1:{1,2}" -> FormulaParser.parse("=A1:{1,2}")
    )
    before.foreach((formula, result) => assert(result.isLeft, formula))
    refused("=A1:B") match
      case ParseError.InvalidCellRef("A1:B", 0, _) => ()
      case other => fail(s"A1:B: $other")
    refused("=A1:\"x\"") match
      case ParseError.InvalidCellRef("A1:", 0, _) => ()
      case other => fail(s"A1:\"x\": $other")
    for formula <- List(
        "=1:A1",
        "=\"x\":A1",
        "=A1 : B2",
        "=A1: B2",
        "=A1 :B2",
        "=A1:Sheet2!B2",
        "=Sheet1:Sheet3!A1",
        "=A1:LOG10(5)",
        "=A:IF(TRUE,B2,C3)",
        "=1:INDEX(A:A,3)",
        "=A1:1+2",
        "=A1:{1,2}"
      )
    do refused(formula)
    refused("=A1:LOG10(5)") match
      case ParseError.UnexpectedChar('(', _, _) => ()
      case other => fail(s"A1:LOG10(5): $other")
    // a defect inside the right operand's call is that call's own diagnostic
    refused("=SUM(A1:INDEX(A:A,FOO(1)))") match
      case ParseError.UnknownFunction("FOO", _, _) => ()
      case other => fail(s"A1:INDEX(A:A,FOO(1)): $other")
    // a computed reference in a typed range slot names the construct (follow-up: #710)
    refused("=SUMIF(A1:INDEX(A:A,3),\">1\",B1:B3)") match
      case ParseError.InvalidArguments("SUMIF", _, _, actual) =>
        assert(actual.contains("computed reference"), actual)
      case other => fail(s"SUMIF: $other")
    refused("=VLOOKUP(1,OFFSET(A1,0,0,3,2),2,FALSE)") match
      case ParseError.InvalidArguments("VLOOKUP", _, _, actual) =>
        assert(actual.contains("computed reference"), actual)
      case other => fail(s"VLOOKUP: $other")
  }

  test("GH-713: lexical ranges keep their exact tree and form") {
    assertEquals(
      parsed("=A1:B2"),
      TExpr.RangeRef(CellRange.parse("A1:B2").toOption.get, RangeForm.Cells)
    )
    assertEquals(
      parsed("=A:A"),
      TExpr.RangeRef(CellRange.parse("A:A").toOption.get, RangeForm.Columns)
    )
    assertEquals(
      parsed("=1:5"),
      TExpr.RangeRef(CellRange.parse("1:5").toOption.get, RangeForm.Rows)
    )
    assertEquals(
      parsed("=Jan:Dec"),
      TExpr.RangeRef(CellRange.parse("JAN:DEC").toOption.get, RangeForm.Columns)
    )
    assertEquals(parsed("=Sheet1!A1:Sheet1!B2"), parsed("=Sheet1!A1:B2"))
    parsed("=Sheet1!A1:B2") match
      case TExpr.SheetRange(_, _, RangeForm.Cells) => ()
      case other => fail(s"Sheet1!A1:B2: $other")
  }

  test("GH-713: long ':' chains are bounded by the nesting budget, never the stack") {
    val chain = "=SUM(" + List.fill(100)("INDEX(B1:B3,1)").mkString(":") + ")"
    assertRoundTrip(chain)
    val long = "=SUM(" + List.fill(400)("INDEX(B1:B3,1)").mkString(":") + ")"
    FormulaParser.parse(long) match
      case Left(_: ParseError.NestingTooDeep) => ()
      case other => fail(s"expected NestingTooDeep, got ${other.map(_.getClass)}")
    val absorbed = "=SUM(" + List.fill(400)("A1").mkString(":INDEX(B1:B3,1):") + ")"
    FormulaParser.parse(absorbed) match
      case Left(_: ParseError.NestingTooDeep) | Right(_) => ()
      case other => fail(s"expected a parse or NestingTooDeep, got $other")
  }

  test("GH-713: the range operator is never callable by name nor listed") {
    refused("=RANGE(A1,B1)") match
      case _: ParseError.UnknownFunction => ()
      case other => fail(s"expected UnknownFunction, got $other")
    assert(!FunctionRegistry.allNames.exists(_.startsWith("(")))
  }

  // ===== Evaluation: the LibreOffice oracle =====

  // LibreOffice 24.2.7 on the same grid (cells in column E, F and G as named)
  private val libreOffice: List[(String, String, CellValue)] = List(
    ("E1", "SUM(INDEX(B1:B3,1):INDEX(B1:B3,2))", N("30")),
    ("E2", "SUM(B2:A1)", N("33")),
    ("E3", "SUM(INDEX(B1:B3,3):A1)", N("66")),
    ("E4", "SUM(OFFSET(A1:INDEX(A1:A3,2),1,1))", N("50")),
    ("E5", "SUM((A1,C3):A1)", N("666")),
    ("E6", "SUM(A1:INDEX(B1:B3,5))", E("#REF!")),
    ("E7", "SUM(A1:INDEX(B1:B3,NA()))", E("#N/A")),
    ("E8", "SUM(A1:CHOOSE(1,5,B2))", E("#VALUE!")),
    ("E9", "ROWS(A1:INDEX(A:A,3))", N("3")),
    ("E10", "SUM(A1:B1:C3)", N("666")),
    ("E11", "SUM(INDEX(B1:C3,{1;2},1))", N("30")),
    ("E12", "SUM(INDEX(B1:B3,{1;2}))", N("30")),
    ("E13", "SUM($A$1:INDEX($A:$A,COUNTA($A:$A)))", N("6")),
    ("E14", "COLUMNS(A1:INDEX(C1:C3,2))", N("3")),
    ("E15", "SUM(A1:INDEX(C1:C3,2))", N("333")),
    ("E16", "SUM(INDEX(A1:INDEX(C1:C3,3),2,3))", N("200")),
    ("E17", "SUM(A1:CHOOSE(2,B1,C3))", N("666")),
    ("E18", "SUM(A1:OFFSET(A1,2,1))", N("66")),
    ("E19", "SUM(A1:INDIRECT(\"B2\"))", N("33")),
    ("E20", "SUM((A1:C3 B2:D4):C3)", N("550")),
    ("E21", "AREAS(A1:INDEX(C1:C3,2))", N("1")),
    ("E22", "SUM(INDEX(A1:INDEX(C1:C3,3),0,2))", N("60")),
    ("E23", "SUMPRODUCT(INDEX(B1:B3,{1;2}))", N("30")),
    ("E24", "SUM(B1:B3:A2)", N("66")),
    ("E25", "SUM(INDEX(B1:B3,2):INDEX(B1:B3,2))", N("20")),
    ("E27", "SUM(A1:INDEX(B1:B3,0))", N("66")),
    ("E28", "SUM(A1:#REF!)", E("#REF!")),
    ("E29", "SUM(INDEX(B1:B3,2):B3:C1)", N("660")),
    // a plain cell reads the bounding range through the implicit intersection
    ("F2", "B1:INDEX(B1:B3,3)", N("20")),
    ("F5", "B1:INDEX(B1:B3,3)", E("#VALUE!")),
    ("G2", "A1:INDEX(B1:B3,3)", E("#VALUE!"))
  )

  test("GH-713: computed ranges agree with LibreOffice") {
    libreOffice.foreach((at, formula, expected) =>
      assertEquals(plain(at, formula), Right(expected), s"$at: =$formula")
    )
  }

  test("GH-713: a name computing a computed range, and the names the operator joins") {
    val dyn = book.withDefinedName("dyn", "Sheet1!$A$1:INDEX(Sheet1!$A:$A,COUNTA(Sheet1!$A:$A))")
    val placed = sheet1.put(ref"H1", CellValue.Formula("SUM(dyn)", None))
    assertEquals(
      placed.evaluateCell(ref"H1", Clock.system, Some(dyn.put(placed))).left.map(_.message),
      Right(N("6"))
    )
    // Start = A1, Finish = B2: A1:B2
    assertEquals(plain("H1", "SUM(Start:Finish)"), Right(N("33")))
    assertEquals(plain("H1", "SUM(Finish:Start)"), Right(N("33")))
    assertEquals(plain("H1", "SUM(A1:Start:Finish)"), Right(N("33")))
  }

  test("GH-713: array mode reads the whole bounding range") {
    assertEquals(
      array("=B1:INDEX(B1:B3,3)"),
      Right(ArrayResult(Vector(Vector(N("10")), Vector(N("20")), Vector(N("30")))))
    )
    assertEquals(array("=SUM(A1:INDEX(C1:C3,2))"), Right(BigDecimal(333)))
  }

  test("GH-713: shapes with no LibreOffice oracle follow Excel") {
    // LibreOffice's xlsx import refuses `A1:IF(…)` (#NAME?); Excel's IF returns a reference
    assertEquals(plain("H1", "SUM(A1:IF(TRUE,B2,C3))"), Right(N("33")))
    assertEquals(plain("H1", "SUM(A1:IF(FALSE,B2,C3))"), Right(N("666")))
    // an operand that evaluates to a value is not a reference
    assertEquals(plain("H1", "SUM(A1:IF(TRUE,5,B2))"), Right(E("#VALUE!")))
    // a range's two ends must be on one sheet: Excel's #VALUE! (LibreOffice spans the sheets)
    assertEquals(plain("H1", "SUM(A1:INDEX(Other!A1:A3,2))"), Right(E("#VALUE!")))
    assertEquals(plain("H1", "SUM(Start:OtherCell)"), Right(E("#VALUE!")))
    assertEquals(plain("H1", "SUM(Sheet1!A1:INDEX(Sheet1!B1:B3,2))"), Right(N("33")))
    // a LET name bound to a reference is an operand
    assertEquals(plain("H1", "LET(x,A1,SUM(x:INDEX(B1:B3,2)))"), Right(N("33")))
    // INDEX over a computed range, and the semilattice laws on a sample
    assertEquals(plain("H1", "INDEX(A1:INDEX(C1:C3,3),2,3)"), Right(N("200")))
    assertEquals(plain("H1", "SUM(INDEX(B1:B3,3):INDEX(B1:B3,1))"), Right(N("60")))
  }

  // ===== Dependencies, recalculation and structural edits =====

  test("GH-713: a covered computed range is static; an uncovered one is dynamic") {
    val covered = parsed("=SUM(A1:INDEX(A:A,COUNTA(A:A)))")
    assert(!DependencyGraph.containsDynamicReference(covered))
    assert(!DependencyGraph.containsDynamicReference(parsed("=SUM(INDEX(B1:B3,1):INDEX(B1:B3,2))")))
    assert(!DependencyGraph.containsDynamicReference(parsed("=SUM(B1:XLOOKUP(2,A1:A3,B1:B3))")))
    assert(!DependencyGraph.containsDynamicReference(parsed("=SUM(A1:B1:C3)")))
    assert(DependencyGraph.containsDynamicReference(parsed("=SUM(B1:INDEX(D:D,3))")))
    assert(DependencyGraph.containsDynamicReference(parsed("=SUM(A1:IF(TRUE,B2,C3))")))
    assert(DependencyGraph.containsDynamicReference(parsed("=SUM(Start:Finish)")))
    // the exact bounding box of static operands: B2 is read though neither operand names it
    assert(DependencyGraph.extractDependencies(parsed("=SUM(A1:B1:C3)")).contains(ref"B2"))
  }

  test("GH-713: B10 =SUM(A1:INDEX(A:A,COUNTA(A:A))) is static and follows an edit in A") {
    val placed = sheet1.put(ref"B10", CellValue.Formula("SUM(A1:INDEX(A:A,COUNTA(A:A)))", None))
    assert(!DependencyGraph.dynamicCells(placed).contains(ref"B10"))
    val graph = DependencyGraph.fromSheet(placed)
    assert(DependencyGraph.detectCycles(graph).isRight)
    assert(DependencyGraph.precedents(graph, ref"B10").contains(ref"A3"))
    assertEquals(
      placed.evaluateCell(ref"B10", Clock.system, Some(book.put(placed))),
      Right(N("6"))
    )
    val grown = placed.put(ref"A4", N("4"))
    assertEquals(grown.evaluateCell(ref"B10", Clock.system, Some(book.put(grown))), Right(N("10")))
  }

  test("GH-713: C5 =SUM(B1:INDEX(D:D,3)) is dynamic and not circular") {
    val placed = sheet1.put(ref"C5", CellValue.Formula("SUM(B1:INDEX(D:D,3))", None))
    assert(DependencyGraph.dynamicCells(placed).contains(ref"C5"))
    assert(DependencyGraph.detectCycles(DependencyGraph.fromSheet(placed)).isRight)
    // B1:D3: 60 + 600 (C1:C3) — D is blank
    assertEquals(
      placed.evaluateCell(ref"C5", Clock.system, Some(book.put(placed))),
      Right(N("660"))
    )
  }

  test("GH-713: the dynamic pre-filter admits every computed range the classifier flags") {
    for text <- List(
        "SUM(B1:INDEX(D:D,3))",
        "SUM(INDEX(D:D,3):B1)",
        "SUM(Start:Finish)",
        "SUM(Start:A5)",
        "LET(x,A1,SUM(x:A5))",
        "SUM(A1:IF(TRUE,B2,C3))",
        "SUM(Sheet1!A1:_xlfn.XLOOKUP(1,A1:A3,C1:C3))",
        "SUM((A1,C3):A1)"
      )
    do
      val expr = parsed(s"=$text")
      if DependencyGraph.containsDynamicReference(expr) then
        assert(
          functions.ReferenceOperators.mayContainComputedRange(text),
          s"pre-filter must admit $text"
        )
    assert(!functions.ReferenceOperators.mayContainComputedRange("SUM(A1:B2)+\"x:INDEX(\""))
    assert(!functions.ReferenceOperators.mayContainComputedRange("SUM(Sheet1!A1:B2,$A$1:$C3)"))
  }

  test("GH-713: shifting and renaming move both operands; a deleted operand is #REF!") {
    val shifted = FormulaShifter.shift(parsed("=SUM(A1:INDEX(A:A,COUNTA(A:A)))"), 1, 1)
    assertEquals(FormulaPrinter.printFileForm(shifted), "SUM(B2:INDEX(B:B,COUNTA(B:B)))")
    def structural(formula: String, at: Int, delta: Int, rows: Boolean = false): String =
      FormulaPrinter.printFileForm(
        FormulaShifter.shiftStructural(parsed(formula), true, "Sheet1", rows, at, delta)
      )
    assertEquals(
      structural("=SUM(B1:INDEX(C1:C3,2))", 1, 1),
      "SUM(C1:INDEX(D1:D3,2))"
    )
    assertEquals(
      structural("=SUM(B5:INDEX(C1:C9,2))", 2, 3, rows = true),
      "SUM(B8:INDEX(C1:C12,2))"
    )
    val deleted = structural("=SUM(A1:INDEX(B1:B3,2))", 0, -1)
    assertEquals(deleted, "SUM(#REF!:INDEX(A1:A3,2))")
    assertEquals(FormulaParser.parse(s"=$deleted").isRight, true)
    assertEquals(plain("H1", deleted), Right(E("#REF!")))
    val renamed = FormulaShifter.renameSheet(
      parsed("=SUM(Sheet1!$A$1:INDEX(Sheet1!$A:$A,5))"),
      SheetName.unsafe("Sheet1"),
      SheetName.unsafe("Data")
    )
    assertEquals(FormulaPrinter.printFileForm(renamed), "SUM(Data!$A$1:INDEX(Data!$A:$A,5))")
  }
