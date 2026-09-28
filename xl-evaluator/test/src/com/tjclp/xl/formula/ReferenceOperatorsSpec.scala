package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.eval.ArrayResult
import com.tjclp.xl.formula.functions.ReferenceOperators
import com.tjclp.xl.ooxml.{TestFixtures, XlsxReader, XlsxWriter}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * GH-669: Excel's union (`,` inside parentheses) and intersection (a space) reference operators.
 *
 * Every value in the LibreOffice table is LibreOffice 25.8's recalculation of the same plain `<f>`
 * cell on the same grid (the cell each formula sits in is part of the case). The shapes where xl
 * follows Excel rather than LibreOffice — or where LibreOffice's xlsx import or evaluation falls
 * short — are pinned after the table, each with its reason.
 */
@SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
class ReferenceOperatorsSpec extends FunSuite:

  private def N(n: String): CellValue = CellValue.Number(BigDecimal(n))
  private def T(s: String): CellValue = CellValue.Text(s)
  private def B(b: Boolean): CellValue = CellValue.Bool(b)
  private def E(code: String): CellValue =
    CellValue.Error(CellError.parse(code).fold(msg => fail(msg), identity))

  /**
   * Sheet1: A1:A10 = 1..10, B1:B5 = 10..50, C1 "x", C3 "y", C5 "z" (C2, C4 blank), D1:E2 =
   * 100,200;300,400. Other: A1 = 1000, B1 = 2000.
   */
  private val book: Workbook =
    val numbers = (1 to 10).foldLeft(Sheet(SheetName.unsafe("Sheet1"))) { (s, i) =>
      s.put(ARef.from1(1, i), N(i.toString))
    }
    val sheet1 = (1 to 5)
      .foldLeft(numbers)((s, i) => s.put(ARef.from1(2, i), N((i * 10).toString)))
      .put(ref"C1", T("x"))
      .put(ref"C3", T("y"))
      .put(ref"C5", T("z"))
      .put(ref"D1", N("100"))
      .put(ref"E1", N("200"))
      .put(ref"D2", N("300"))
      .put(ref"E2", N("400"))
    val other = Sheet(SheetName.unsafe("Other"))
      .put(ref"A1", N("1000"))
      .put(ref"B1", N("2000"))
    Workbook(Vector(sheet1, other))

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

  /** Accepted, and the file form re-parses to the same tree. */
  private def assertRoundTrip(formula: String): TExpr[?] =
    val expr = parsed(formula)
    assertEquals(FormulaParser.parse(FormulaPrinter.printFileForm(expr)), Right(expr), formula)
    assertEquals(FormulaParser.parse(FormulaPrinter.print(expr)), Right(expr), formula)
    expr

  private def refused(formula: String): ParseError =
    FormulaParser.parse(formula).swap.getOrElse(fail(s"$formula should not parse"))

  /** LibreOffice prints 15 significant digits: numbers agree to 1e-12. */
  private def agree(obtained: Either[String, CellValue], expected: CellValue, clue: String): Unit =
    (obtained, expected) match
      case (Right(CellValue.Number(got)), CellValue.Number(want)) =>
        assert((got - want).abs < BigDecimal("1e-12"), s"$clue: $got is not $want")
      case _ => assertEquals(obtained, Right(expected), clue)

  // ===== Parser =====

  test("GH-669: union and intersection formulas parse and round-trip") {
    List(
      "=SUM((A1,A2))",
      "=SUM((A1:A2,B1:B2))",
      "=SUM((A1,(B1,C1)))",
      "=INDEX((A1:B2,A1:C2),1,1,2)",
      "=AREAS((A1,B1))",
      "=SUM((A1:B2 B1:C2))",
      "=(A1:B2 B1:C2)",
      "=A1:C1 B1:B5",
      "=A:A 3:3",
      "=Jan Sales",
      "=A1 (B1:C2)",
      "=(A1,B1) A1:B2",
      "=A1:C3 (B1:C3 C1:C3)",
      "=-A1:B2 B1:C2",
      "=A1:B2 B1:C2%",
      "=(A1:B2 B1:C2,D1)",
      "=@(A1:B2 B1:C2)",
      "=A1:B2 #REF!",
      "=(#REF!,A1)",
      "=OFFSET(A1,0,0,3,3) B1:B5",
      "=SUM(Sheet1!A1:B2 Sheet1!B1:C2)",
      "=SUM('My Sheet'!A1:B2 'My Sheet'!B2)"
    ).foreach(assertRoundTrip)
  }

  test("GH-669: the printed forms: bare ',' in every form, one space, '~' read as ','") {
    def printsAs(source: String, expected: String): Unit =
      assertEquals(FormulaPrinter.print(parsed(source)), expected, source)
      assertEquals("=" + FormulaPrinter.printFileForm(parsed(source)), expected, source)
    printsAs("=SUM((A1,A2))", "=SUM((A1,A2))")
    assertEquals(
      FormulaPrinter.printFileForm(parsed("=INDEX((A1:B2,D1:E2),1,1,2)")),
      "INDEX((A1:B2,D1:E2),1,1,2)"
    )
    assertEquals(
      FormulaPrinter.print(parsed("=INDEX((A1:B2,D1:E2),1,1,2)")),
      "=INDEX((A1:B2,D1:E2), 1, 1, 2)"
    )
    printsAs("=SUM((A1~A2))", "=SUM((A1,A2))")
    printsAs("=SUM((A1 , A2))", "=SUM((A1,A2))")
    printsAs("=A1:B2   B1:C2", "=A1:B2 B1:C2")
    printsAs("=A1:C3 (B1:C3 C1:C3)", "=A1:C3 (B1:C3 C1:C3)")
    printsAs("=(A1:C3 B1:C3) C1:C3", "=A1:C3 B1:C3 C1:C3")
    printsAs("=@(A1:B2 B1:C2)", "=@(A1:B2 B1:C2)")
    printsAs("=@(A1,B1)", "=@(A1,B1)")
    // precedence: the intersection binds tighter than negation, % and ^
    assertEquals(parsed("=-A1:B2 B1:C2"), parsed("=-(A1:B2 B1:C2)"))
    assertEquals(parsed("=A1:B2 B1:C2%"), parsed("=(A1:B2 B1:C2)%"))
    assertEquals(parsed("=A1:B2 B1:C2^2"), parsed("=(A1:B2 B1:C2)^2"))
  }

  test("GH-669: non-reference operands, empty areas and stray closers are refused") {
    for formula <- List("=(1,2)", "=(A1,{1,2})", "=(A1,A1#)", "=A1 5", "=A1 SUM(B1)")
    do
      refused(formula) match
        case _: ParseError.InvalidOperator => ()
        case other => fail(s"$formula: expected InvalidOperator, got $other")
    refused("=(A1,)") match
      case _: ParseError.UnexpectedChar | _: ParseError.UnexpectedEOF => ()
      case other => fail(s"(A1,): expected UnexpectedChar/EOF, got $other")
    assertEquals(refused("=(A1,A2]"), ParseError.UnbalancedDelimiter(6, ']', "expected ')'"))
    // a top-level bare ',' stays an error (defined-name unions are a follow-up)
    refused("=A1,B1")
  }

  test("GH-669: spaces that are not an intersection are unchanged") {
    def same(a: String, b: String): Unit = assertEquals(parsed(a), parsed(b), a)
    for (formula, name) <- List(("=A1 AND B1", "AND"), ("=A1 OR B1", "OR")) do
      parsed(formula) match
        case TExpr.Call(spec, _) => assertEquals(spec.name, name)
        case other => fail(s"$formula should stay the $name keyword, got $other")
    same("=SUM(A1 , B1)", "=SUM(A1,B1)")
    same("=A1 &B1", "=A1&B1")
    same("=IF(A1 =1,2,3)", "=IF(A1=1,2,3)")
    // a function name before a spaced '(' stays a call: LOG (a registered name) and LOG10 (an
    // Excel function shaped like a cell address, which xl does not implement)
    parsed("=LOG (A1)") match
      case TExpr.Call(spec, _) => assertEquals(spec.name, "LOG")
      case other => fail(s"LOG (A1) should stay a call, got $other")
    refused("=LOG10 (A1)") match
      case ParseError.UnknownFunction("LOG10", _, _) => ()
      case other => fail(s"LOG10 (A1) should stay an unknown function, got $other")
    // a tab or a line feed is never the operator
    assert(FormulaParser.parse("=A1:B2\tB1:C2").isLeft)
    assert(FormulaParser.parse("=A1:B2\nB1:C2").isLeft)
  }

  test("GH-669: long unions and intersection chains hit the depth guard, never the stack") {
    // each area of a union re-enters the expression grammar: 200 areas is fine…
    assertRoundTrip("=SUM((" + List.fill(200)("A1").mkString(",") + "))")
    // …and deep nesting of unions is bounded by the nesting budget
    val deep = "=SUM(" + ("(A1," * 200) + "A1" + (")" * 200) + ")"
    FormulaParser.parse(deep) match
      case Left(_: ParseError.NestingTooDeep) => ()
      case other => fail(s"expected NestingTooDeep, got ${other.map(_.getClass)}")
    val chain = "=" + List.fill(300)("A1:C3").mkString(" ")
    FormulaParser.parse(chain) match
      case Left(_: ParseError.NestingTooDeep) | Right(_) => ()
      case other => fail(s"expected a parse or NestingTooDeep, got $other")
  }

  test("GH-669: the operators are never callable by name nor listed") {
    assertEquals(FunctionRegistry.lookup(ReferenceOperators.UnionName), None)
    assertEquals(FunctionRegistry.lookup(ReferenceOperators.IntersectionName), None)
    refused("=UNION(A1,B1)") match
      case _: ParseError.UnknownFunction => ()
      case other => fail(s"expected UnknownFunction, got $other")
    assert(!FunctionRegistry.allNames.exists(_.startsWith("(")))
    assert(FunctionRegistry.lookup("AREAS").isDefined)
  }

  // ===== Evaluation =====

  test("GH-669: aggregates fold every area — overlaps count twice, nested unions flatten") {
    assertEquals(plain("H1", "SUM((A1,A2))"), Right(N("3")))
    assertEquals(plain("H1", "SUM((A1:A3,A2:A4))"), Right(N("15")))
    assertEquals(plain("H1", "SUM((A1,A1))"), Right(N("2")))
    assertEquals(plain("H1", "SUM((A1,(B1,B2)))"), Right(N("31")))
    assertEquals(array("=SUM((A1,A2))"), Right(BigDecimal(3)))
    assertEquals(array("=SUM((A1:A5,C1:C5) A3:C3)"), Right(BigDecimal(3)))
    assertEquals(plain("H1", "COUNTBLANK(C1:C4 C2:C3)"), Right(N("1")))
    assertEquals(plain("H1", "COUNTBLANK((C1:C2,C3:C4))"), Right(E("#VALUE!")))
  }

  test("GH-669: a union in a value position is #VALUE!") {
    for formula <- List("(A1,A2)", "(A1,A2)+1", "ABS((A1,A2))", "@(A1,B1)", "ROWS((A1,A2))")
    do assertEquals(plain("H1", formula), Right(E("#VALUE!")), formula)
    assertEquals(plain("H1", "SUM(IF(TRUE,(A1,A2),0))"), Right(E("#VALUE!")))
    assertEquals(plain("H1", "OFFSET((A1,A2),0,1)"), Right(E("#VALUE!")))
    assertEquals(plain("H1", "LET(x,(A1,A2),SUM(x))"), Right(E("#VALUE!")))
    // a parenthesized union name parses; SUM over it is #VALUE! (a top-level bare union name
    // stays unparseable — the defined-name follow-up)
    val named = book.withDefinedName("pair", "(Sheet1!$A$1,Sheet1!$A$2)")
    val placed = sheet1.put(ref"H1", CellValue.Formula("SUM(pair)", None))
    assertEquals(
      placed.evaluateCell(ref"H1", Clock.system, Some(named.put(placed))).left.map(_.message),
      Right(E("#VALUE!"))
    )
  }

  test("GH-669: an intersection is one reference; empty is #NULL!, cross-sheet #VALUE!") {
    assertEquals(plain("H5", "A1:C10 B1:B10"), Right(N("50")))
    assertEquals(plain("H12", "A1:C10 B1:B10"), Right(E("#VALUE!")))
    assertEquals(plain("H1", "A1:A3 C1:C3"), Right(E("#NULL!")))
    assertEquals(plain("H1", "ERROR.TYPE(A1:A3 C1:C3)"), Right(N("1")))
    assertEquals(plain("H1", "SUM(Sheet1!A1:B2 Other!A1:B2)"), Right(E("#VALUE!")))
    // (intersection)*2 in array mode equals the INDEX column it names
    assertEquals(array("=(A1:B3 B1:C3)*2"), array("=INDEX(A1:B3,0,2)*2"))
    assertEquals(
      array("=A1:B3 B1:C3"),
      Right(ArrayResult(Vector(Vector(N("10")), Vector(N("20")), Vector(N("30")))))
    )
  }

  test("GH-669: INDEX area_num picks an area; out of range is #REF!, below 1 #VALUE!") {
    assertEquals(plain("H1", "INDEX((A1:B2,D1:E2),1,1,2)"), Right(N("100")))
    assertEquals(plain("H1", "INDEX((A1:B2,D1:E2),2,2,2)"), Right(N("400")))
    assertEquals(plain("H1", "INDEX((A1:B2,D1:E2),1,1,3)"), Right(E("#REF!")))
    assertEquals(plain("H1", "INDEX((A1:B2,D1:E2),1,1,0)"), Right(E("#VALUE!")))
    assertEquals(plain("H1", "INDEX(A1:B2,1,1,2)"), Right(E("#REF!")))
    assertEquals(plain("H1", "INDEX(A1:B2,1,1,)"), Right(N("1")))
    assertEquals(plain("H1", "SUM(INDEX((A1:B2,D1:E2),0,0,2))"), Right(N("1000")))
  }

  test("GH-669: AREAS counts areas") {
    assertEquals(plain("H1", "AREAS((A1,B1))"), Right(N("2")))
    assertEquals(plain("H1", "AREAS(A1:B2)"), Right(N("1")))
    assertEquals(plain("H1", "AREAS(((A1,B1),C1))"), Right(N("3")))
    assertEquals(plain("H1", "AREAS((A1:A5,C1:C5) A2:C2)"), Right(N("2")))
    assertEquals(plain("H1", "AREAS(A1 C1)"), Right(E("#NULL!")))
    assertEquals(plain("H1", "AREAS(OFFSET(A1,0,0,2,2))"), Right(N("1")))
    for formula <- List("=AREAS(1)", "=AREAS(\"A1\")") do refused(formula)
  }

  // ===== Dependencies =====

  test(
    "GH-669: dependency edges — every area of a union, only the shared cells of an intersection"
  ) {
    assertEquals(
      DependencyGraph.extractDependencies(parsed("=SUM((A1,C1:C3))")),
      Set(ref"A1", ref"C1", ref"C2", ref"C3")
    )
    assertEquals(
      DependencyGraph.extractDependencies(parsed("=SUM(A1:C3 A1:A3)")),
      Set(ref"A1", ref"A2", ref"A3")
    )
    assertEquals(DependencyGraph.extractDependencies(parsed("=SUM(A1:A3 C1:C3)")), Set.empty[ARef])
    // an external leaf inside a union makes the formula external (its cache stays pinned)
    assert(TExpr.containsExternalRef(parsed("=SUM(([1]Book!A1,A2))")))
    assert(!TExpr.containsExternalRef(parsed("=SUM((A1,A2))")))
    // C3: =SUM(A1:C3 A1:A3) reads A1:A3 only — not circular
    val placed = sheet1.put(ref"C3", CellValue.Formula("SUM(A1:C3 A1:A3)", None))
    val graph = DependencyGraph.fromSheet(placed)
    assert(DependencyGraph.detectCycles(graph).isRight)
    assertEquals(
      placed.evaluateCell(ref"C3", Clock.system, Some(book.put(placed))),
      Right(N("6"))
    )
  }

  test("GH-669: shifting moves every area and both operands; a deleted area is #REF!") {
    val shifted = FormulaShifter.shift(parsed("=SUM((A1,C1:C3)) + SUM(A1:B2 B1:C2)"), 1, 1)
    assertEquals(FormulaPrinter.print(shifted), "=SUM((B2,D2:D4))+SUM(B2:C3 C2:D3)")
    val renamed = FormulaShifter.renameSheet(
      parsed("=SUM((Other!A1,Other!B1))"),
      SheetName.unsafe("Other"),
      SheetName.unsafe("Renamed")
    )
    assertEquals(FormulaPrinter.print(renamed), "=SUM((Renamed!A1,Renamed!B1))")
    assertEquals(plain("H1", "SUM((#REF!,A1))"), Right(E("#REF!")))
    assertEquals(plain("H1", "SUM(A1:B2 #REF!)"), Right(E("#REF!")))
    // structural edits: inserting a column before B moves every area and both operands; deleting
    // an area's column leaves #REF! where it was, which re-parses and evaluates to #REF!
    def structural(formula: String, at: Int, delta: Int): String =
      FormulaPrinter.print(
        FormulaShifter.shiftStructural(parsed(formula), true, "Sheet1", false, at, delta)
      )
    assertEquals(
      structural("=SUM((A1,C1:C3))+SUM(B1:C2 C1:D2)", 1, 1),
      "=SUM((A1,D1:D3))+SUM(C1:D2 D1:E2)"
    )
    val deleted = structural("=SUM((C1,A1))", 2, -1)
    assertEquals(deleted, "=SUM((#REF!,A1))")
    assertEquals(plain("H1", deleted.stripPrefix("=")), Right(E("#REF!")))
    val gone = structural("=SUM(A1:B2 C1:C2)", 2, -1)
    assertEquals(gone, "=SUM(A1:B2 #REF!)")
    assertEquals(plain("H1", gone.stripPrefix("=")), Right(E("#REF!")))
  }

  // ===== LibreOffice oracle =====

  private val libreOffice: List[(String, String, CellValue)] = List(
    ("H1", "SUM((A1,A2))", N("3")),
    ("H2", "SUM((A1:A3,A2:A4))", N("15")),
    ("H3", "SUM((A1,C1:C3))", N("1")),
    ("H5", "SUM(((A1,B1),C1))", N("11")),
    ("H6", "COUNT((A1,A2))", N("2")),
    ("H7", "COUNT((A1,C1:C3))", N("1")),
    ("H8", "COUNTA((A1,C1:C3))", N("3")),
    ("H9", "AVERAGE((A1:A3,A2:A4))", N("2.5")),
    ("H10", "MIN((A1:A3,B1:B2))", N("1")),
    ("H11", "MAX((A1:A3,B1:B2))", N("20")),
    ("H12", "MEDIAN((A1:A3,B1:B2))", N("3")),
    ("H13", "STDEV((A1:A3,B1:B2))", N("7.98122797569397")),
    ("H14", "VAR((A1:A3,B1:B2))", N("63.7")),
    ("H15", "STDEVP((A1:A3,B1:B2))", N("7.1386273190299")),
    ("H16", "VARP((A1:A3,B1:B2))", N("50.96")),
    ("H17", "SUM((A1,A1))", N("2")),
    ("H20", "SUM((A1:A2 A2:A3,B1))", N("12")),
    ("H21", "INDEX((A1:B2,D1:E2),1,1,1)", N("1")),
    ("H22", "INDEX((A1:B2,D1:E2),1,1,2)", N("100")),
    ("H23", "INDEX((A1:B2,D1:E2),2,2,2)", N("400")),
    ("H24", "INDEX((A1:B2,D1:E2),1,1,3)", E("#REF!")),
    ("H25", "INDEX((A1:B2,D1:E2),1,1,0)", E("#VALUE!")),
    ("H26", "SUM(INDEX((A1:B2,D1:E2),0,0,2))", N("1000")),
    ("H27", "INDEX(A1:B2,1,1,2)", E("#REF!")),
    ("H28", "INDEX((A1:B2,D1:E2),2,1)", N("2")),
    ("H29", "AREAS((A1,B1))", N("2")),
    ("H30", "AREAS(A1:B2)", N("1")),
    ("H31", "AREAS(((A1,B1),C1))", N("3")),
    ("H33", "AREAS(A1 C1)", E("#NULL!")),
    ("H34", "AREAS(OFFSET(A1,0,0,2,2))", N("1")),
    ("H35", "AREAS(A1)", N("1")),
    ("H36", "SUM(A1:B3 B2:C5)", N("50")),
    ("H37", "SUM(A:A 3:3)", N("3")),
    ("H38", "ROWS(A1:C10 B2:B4)", N("3")),
    ("H39", "COLUMNS(A1:C10 B2:D2)", N("2")),
    ("H43", "ISERROR(A1:A3 C1:C3)", B(true)),
    ("H44", "ERROR.TYPE(A1:A3 C1:C3)", N("1")),
    ("H45", "IFERROR(A1:A3 C1:C3,\"none\")", T("none")),
    ("H46", "COUNT(A1:A3 C1:C3)", N("0")),
    ("H47", "COUNTA(A1:A3 C1:C3)", N("1")),
    ("H48", "A1:A3 C1:C3", E("#NULL!")),
    ("H49", "SUM(A1:B2 B1:C2)", N("30")),
    ("H50", "(A1:B2 B1:C2)", E("#VALUE!")),
    ("H51", "A1:C1 B1:B5", N("10")),
    ("H54", "-A1:B2 B2:C3", N("-20")),
    ("H56", "INDEX(A1:C3 A2:C3,1,1)", N("2")),
    ("H57", "(A1,A2)", E("#VALUE!")),
    ("H58", "(A1,A2)+1", E("#VALUE!")),
    ("H59", "ABS((A1,A2))", E("#VALUE!")),
    ("H60", "ROWS((A1,A2))", E("#VALUE!")),
    ("H63", "COUNTBLANK(C1:C4 C2:C3)", N("1")),
    ("H66", "SUM((A1~A2))", N("3")),
    ("H67", "SUM(ABS({-1,-2}))", N("3")),
    ("H68", "SUM({1,2}*{3,4})", N("11")),
    ("H69", "SUMPRODUCT({1,2},{3,4})", N("11")),
    ("H70", "SUM(COUNTIF(A1:A10,{1,2,3}))", N("3")),
    ("H71", "SUM({1,2}*A1)", N("3")),
    ("H72", "SUM({1;2}*A1:A2)", E("#VALUE!")),
    ("H73", "MAX({1,2}+A3)", N("5")),
    ("H74", "{1,2}&\"x\"", T("1x")),
    ("H75", "SUM(-{1,2})", N("-3")),
    ("H76", "SUM({10,20}%)", N("0.3")),
    ("H77", "SUM(--({1,2}>1))", N("1")),
    ("H78", "AVERAGE({1,2,3}*2)", N("4")),
    ("H79", "SUM(({1,2}))", N("3")),
    ("H80", "INDEX({1,2;3,4},0,2)", N("2")),
    ("H81", "SUM(INDEX({1,2;3,4},0,2))", N("6")),
    ("H82", "INDEX({1,2;3,4},1)", N("1")),
    ("H83", "INDEX({10,20,30},2)", N("20")),
    ("H86", "ROWS(A1:A3 A2:A5)", N("2")),
    ("H87", "SUM(A1:A10 A5)", N("5")),
    ("H88", "SUM(A1:A10 A11)", E("#NULL!")),
    ("H89", "Sheet1!B1:Sheet1!B5", E("#VALUE!")),
    ("H90", "SUM(Sheet1!B1:Sheet1!B5)", N("150")),
    ("I5", "A1:C10 B1:B10", N("50")),
    ("I12", "A1:C10 B1:B10", E("#VALUE!")),
    ("J2", "A1:A5 A1:A5", N("2")),
    ("J3", "(A1:A5,B1:B5)", E("#VALUE!")),
    ("J4", "A1:A10 A:A", N("4"))
  )

  test("GH-669: union, intersection and array-constant cells agree with LibreOffice") {
    libreOffice.foreach((at, formula, expected) =>
      agree(plain(at, formula), expected, s"$at: =$formula")
    )
  }

  test("GH-669: shapes where xl follows Excel, each pinned with its reason") {
    // LibreOffice's xlsx import does not turn a union nested inside another union's parentheses
    // (or one holding an error literal) into its own union operator and answers #VALUE!; Excel
    // flattens nested unions and SUM propagates a #REF! area
    assertEquals(plain("H4", "SUM((A1,(B1,B2)))"), Right(N("31")))
    assertEquals(plain("H18", "SUM((#REF!,A1))"), Right(E("#REF!")))
    // the union is resolved as one argument and triaged by COUNT's rule for an error argument,
    // which COUNT skips (GH-630): 0, like COUNT(#REF!)
    assertEquals(plain("H19", "COUNT((#REF!,A1))"), Right(N("0")))
    // LibreOffice does not intersect a union, a parenthesized intersection, a reference a function
    // returns, or a cell before a parenthesized reference (#VALUE!); Excel does
    assertEquals(plain("H32", "AREAS((A1:A5,C1:C5) A2:C2)"), Right(N("2")))
    assertEquals(plain("H40", "SUM((A1:A5,C1:C5) A3:C3)"), Right(N("3")))
    assertEquals(plain("H41", "COUNT((A1:A5,B1:B5) A3:C3)"), Right(N("2")))
    assertEquals(plain("H52", "A1 (A1:C2)"), Right(N("1")))
    assertEquals(plain("H53", "SUM(A1:C3 (B1:C3 C1:C3))"), Right(N("0")))
    assertEquals(plain("H55", "SUM(OFFSET(A1,0,0,3,3) B1:B5)"), Right(N("60")))
    // LibreOffice has no `@` (#NAME?); Excel's implicit intersection of A1:A10 from row 42 fails
    assertEquals(plain("H42", "@(A1:A10 A1:B10)"), Right(E("#VALUE!")))
    assertEquals(plain("H5", "@(A1:A10 A1:B10)"), Right(N("5")))
    // COUNTBLANK takes one range: Excel's #VALUE! for a union, where LibreOffice counts (2)
    assertEquals(plain("H62", "COUNTBLANK((C1:C2,C3:C4))"), Right(E("#VALUE!")))
    // a union or intersection spans one sheet: xl's #VALUE! where LibreOffice sums across sheets
    // (1001) and answers #NULL! for the cross-sheet intersection
    assertEquals(plain("H64", "SUM((A1,Other!A1))"), Right(E("#VALUE!")))
    assertEquals(plain("H65", "SUM(Sheet1!A1:B2 Other!A1:B2)"), Right(E("#VALUE!")))
    // a union IF selects reaches SUM as a value (#VALUE!), where LibreOffice sums it (3) — a
    // documented residual: never a wrong number
    assertEquals(plain("H61", "SUM(IF(TRUE,(A1,A2),0))"), Right(E("#VALUE!")))
    // lookups over an array constant (MATCH 2, VLOOKUP "b" in LibreOffice) and a union in a range
    // slot stay parse errors — refused, never a wrong number
    refused("=MATCH(2,{1,2,3},0)")
    refused("=VLOOKUP(2,{1,\"a\";2,\"b\"},2,FALSE)")
    refused("=LARGE((A1:A3,C1:C3),1)")
    refused("=COUNTIF((A1,A2),1)")
  }

  test("GH-669: the audit no longer lists the operators as unparseable") {
    val formulas = List(
      "SUM((A1,A2))",
      "SUM((A1~A2))",
      "INDEX((A1:B2,D1:E2),1,1,2)",
      "AREAS((A1,B1))",
      "(A1:B2 B1:C2)",
      "TRUE()",
      "SUM({1,2;3,4})",
      "Sheet1!A1:Sheet1!B2"
    )
    val placed = formulas.zipWithIndex.foldLeft(sheet1) { case (acc, (f, i)) =>
      acc.put(ARef.from1(10, i + 1), CellValue.Formula(f, None))
    }
    assertEquals(eval.WorkbookAudit.of(book.put(placed)).unparseable, Vector.empty)
  }

  // ===== The LibreOffice fixture =====

  test("GH-669: the LibreOffice fixture parses, recalculates to its caches and keeps its text") {
    // reference-operators-lo.xlsx (scripts/generate-fixtures.py): LibreOffice's own <f> text —
    // `SUM((A1~A2))` for a union — and its cached values
    val path = TestFixtures.copyToTemp("reference-operators-lo.xlsx")
    val read = XlsxReader.read(path).fold(e => fail(e.message), identity)
    val sheet = read.sheets.headOption.getOrElse(fail("no sheet"))
    val formulas = sheet.cells.values.toList.flatMap { cell =>
      cell.value match
        case CellValue.Formula(text, Some(cached), _) => List((cell.ref, text, cached))
        case _ => Nil
    }
    assert(formulas.size >= 24, s"${formulas.size} formula cells")
    assert(formulas.exists((_, text, _) => text.contains("~")), "the union keeps LibreOffice's '~'")
    val recalculated = read.withCachedFormulas().sheets.headOption.getOrElse(fail("no sheet"))
    formulas.foreach { (at, text, cached) =>
      assert(FormulaParser.parse(s"=$text").isRight, s"${at.toA1}: '$text' should parse")
      recalculated(at).value match
        case CellValue.Formula(after, Some(value), _) =>
          assertEquals(after, text, s"${at.toA1}: the formula text is untouched")
          (value, cached) match
            case (CellValue.Number(got), CellValue.Number(want)) =>
              assert((got - want).abs < BigDecimal("1e-12"), s"${at.toA1} '$text': $got vs $want")
            case _ => assertEquals(value, cached, s"${at.toA1}: '$text'")
        case other => fail(s"${at.toA1}: '$text' recalculated to $other")
    }
    // an untouched <f> is written back as it was read
    val out = java.nio.file.Files.createTempFile("xl-refops-", ".xlsx")
    out.toFile.deleteOnExit()
    XlsxWriter.write(read, out).fold(e => fail(e.message), identity)
    val reread = XlsxReader.read(out).fold(e => fail(e.message), identity)
    formulas.foreach { (at, text, _) =>
      reread.sheets.headOption.map(_(at).value) match
        case Some(CellValue.Formula(after, _, _)) => assertEquals(after, text, at.toA1)
        case other => fail(s"${at.toA1}: $other")
    }
  }
