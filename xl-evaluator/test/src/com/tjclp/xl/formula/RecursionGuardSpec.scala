package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.parser.{FormulaParser, ParseError}
import munit.FunSuite

/**
 * GH-56: pathologically nested formulas must be rejected with a `ParseError.NestingTooDeep` (a
 * total `Left`), never a `StackOverflowError`. Capping parse depth also keeps the AST shallow
 * enough that evaluation cannot overflow, so this single guard protects both parser and evaluator.
 */
class RecursionGuardSpec extends FunSuite:

  private val s = new Sheet(name = SheetName.unsafe("S"))

  test("deeply nested parentheses → NestingTooDeep, not StackOverflowError") {
    val formula = "=" + ("(" * 300) + "1" + (")" * 300)
    FormulaParser.parse(formula) match
      case Left(_: ParseError.NestingTooDeep) => ()
      case other => fail(s"expected NestingTooDeep, got $other")
  }

  test("deeply nested unary minus → NestingTooDeep, not StackOverflowError") {
    val formula = "=" + ("-" * 300) + "1"
    FormulaParser.parse(formula) match
      case Left(_: ParseError.NestingTooDeep) => ()
      case other => fail(s"expected NestingTooDeep, got $other")
  }

  test("GH-271: deeply chained unary plus → NestingTooDeep, not StackOverflowError") {
    val formula = "=" + ("+" * 300) + "1"
    FormulaParser.parse(formula) match
      case Left(_: ParseError.NestingTooDeep) => ()
      case other => fail(s"expected NestingTooDeep, got $other")
  }

  test("GH-355: deeply chained percent → NestingTooDeep, not StackOverflowError") {
    val formula = "=1" + ("%" * 300)
    FormulaParser.parse(formula) match
      case Left(_: ParseError.NestingTooDeep) => ()
      case other => fail(s"expected NestingTooDeep, got $other")
  }

  test("deeply nested function calls → NestingTooDeep") {
    val formula = "=" + ("ABS(" * 300) + "1" + (")" * 300)
    FormulaParser.parse(formula) match
      case Left(_: ParseError.NestingTooDeep) => ()
      case other => fail(s"expected NestingTooDeep, got $other")
  }

  test("deeply nested formula via evaluateFormula returns Left, never throws") {
    val formula = "=" + ("(" * 300) + "1" + (")" * 300)
    assert(s.evaluateFormula(formula).isLeft)
  }

  test("normal moderately-nested formulas still parse and evaluate") {
    assertEquals(s.evaluateFormula("=((((1+2))))*3"), Right(CellValue.Number(BigDecimal(9))))
    assert(FormulaParser.parse("=SUM(1,ABS(-2),MAX(3,4))").isRight)
    // ~50 nested parens is well under the 256 cap and must still work
    val ok = "=" + ("(" * 50) + "7" + (")" * 50)
    assertEquals(s.evaluateFormula(ok), Right(CellValue.Number(BigDecimal(7))))
  }

  // The binary operators are right-recursive, so chained operators are a distinct overflow
  // vector from parens/unary (caught by the claude-review of PR #250).
  private def assertTooDeep(formula: String): Unit =
    FormulaParser.parse(formula) match
      case Left(_: ParseError.NestingTooDeep) => ()
      // never the parsed AST in the message: case-class toString recurses along a deep spine
      case Right(_) => fail(s"expected NestingTooDeep, ${formula.take(40)}… parsed")
      case Left(other) => fail(s"expected NestingTooDeep, got $other")

  test("the right-nested keyword chains still count every level") {
    assertTooDeep("=1" + (" AND 1" * 300)) // logical AND chain (a right-nested call spine)
    assertTooDeep("=1" + (" OR 1" * 300)) // logical OR chain
  }

  test("normal short operator chains still parse") {
    assert(FormulaParser.parse("=1=1").isRight)
    assert(FormulaParser.parse("=1&2&3").isRight)
    assert(FormulaParser.parse("=TRUE AND FALSE OR TRUE").isRight)
  }

  // GH-680: the budget bounds nesting — function arguments and parentheses — not the length of a
  // flat operator chain. Excel's limit is 64 levels of nesting; a chain of any length within the
  // 8192-character cap opens intact.
  private def assertRoundTrips(formula: String): TExpr[?] =
    FormulaParser.parse(formula) match
      case Right(expr) =>
        assertEquals(FormulaPrinter.print(expr), formula)
        expr
      case Left(err) => fail(s"${formula.take(40)}… should parse: $err")

  test("GH-680: a 200-term + chain parses, evaluates and prints back as written") {
    val sheet =
      (1 to 200).foldLeft(s)((acc, i) => acc.put(ARef.from1(2, i), CellValue.Number(i)))
    val formula = "=" + (1 to 200).map(i => s"B$i").mkString("+")
    assertRoundTrips(formula)
    assertEquals(sheet.evaluateFormula(formula), Right(CellValue.Number(BigDecimal(20100))))
  }

  test("GH-680: a 200-term & chain parses, evaluates and prints back as written") {
    val formula = "=" + (1 to 200).map(i => "\"" + (i % 10) + "\"").mkString("&")
    assertRoundTrips(formula)
    val expected = (1 to 200).map(i => (i % 10).toString).mkString
    assertEquals(s.evaluateFormula(formula), Right(CellValue.Text(expected)))
  }

  test("GH-680: a flat chain of every binary operator costs no nesting level") {
    for op <- List("+", "-", "*", "/", "^", "&", "=", "<", "<>", ">=")
    do assertRoundTrips("=1" + (op + "1") * 300)
  }

  test("GH-680: 64 levels of nesting parse; 129 fail NestingTooDeep") {
    assert(FormulaParser.parse("=" + ("SUM(" * 64) + "1" + (")" * 64)).isRight)
    assert(FormulaParser.parse("=" + ("(" * 64) + "1" + (")" * 64)).isRight)
    // a chain inside every level costs nothing: only the calls nest
    assertRoundTrips("=" + ("SUM(1+1+1+1+1+" * 64) + "1" + (")" * 64))
    assertTooDeep("=" + ("SUM(" * 129) + "1" + (")" * 129))
    assertTooDeep("=" + ("(" * 129) + "1" + (")" * 129))
  }

  /**
   * Run `body` on a fresh thread with a 1MB stack — the JVM default on Linux x64, where CI runs —
   * rather than the test runner's.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  private def onSmallStack[A](body: => A): A =
    var result: Either[Throwable, A] = Left(new IllegalStateException("did not run"))
    // catches StackOverflowError too (Try would not), so an overflow fails the test by name
    val run: Runnable = () =>
      result =
        try Right(body)
        catch case e: Throwable => Left(e)
    val t = Thread.ofPlatform().name("gh-680-small-stack").stackSize(1024L * 1024L).unstarted(run)
    t.start()
    t.join()
    result.fold(e => throw e, identity)

  private def assertTooManyOperators(formula: String): Unit =
    FormulaParser.parse(formula) match
      case Left(ParseError.TooManyOperators(1025, 1024, _)) => ()
      case other => fail(s"expected TooManyOperators(1025, 1024), got $other")

  test("GH-680: 1024 chained operators parse; the 1025th is TooManyOperators, a total Left") {
    val terms = List.fill(1025)("1")
    for op <- List("+", "&", "*", "=", "^")
    do
      assertRoundTrips("=" + terms.mkString(op))
      assertTooManyOperators("=" + (terms :+ "1").mkString(op))
    // the budget is per formula: operators in sibling arguments add up
    val half = List.fill(514)("1").mkString("+")
    assertTooManyOperators(s"=SUM($half,$half)")
    // the caret names the first operator over the limit
    val over = "=" + (terms :+ "1").mkString("+")
    FormulaParser.parse(over) match
      case Left(err @ ParseError.TooManyOperators(_, _, pos)) =>
        assertEquals(pos, over.lastIndexOf('+') - 1)
        assert(ParseError.describe(err).startsWith("Too many chained operators"), err.toString)
      case other => fail(s"expected TooManyOperators, got $other")
  }

  test("GH-680: the longest chain inside the deepest nest walks stack-safely on a 1MB stack") {
    // the worst shape both budgets admit: 127 nested calls around a 1024-operator chain, through
    // every walker that meets a formula — each takes the chain's left spine in one loop
    val terms = 1025
    val other = SheetName.unsafe("Other")
    val refChain = ("Other!A1" :: List.fill(terms - 1)("A1")).mkString("+")
    val nested = "=" + ("SUM(" * 127) + refChain + (")" * 127)
    val concat =
      "=" + ("SUM(" * 126) + "LEN(" + List.fill(terms)("A1").mkString("&") + ")" + (")" * 126)
    val compare =
      "=" + ("SUM(" * 126) + "N(" + List.fill(terms)("A1").mkString("=") + ")" + (")" * 126)
    val sheet = s.put(ARef.from1(1, 1), CellValue.Number(1))
    val local = nested.replace("Other!", "")
    onSmallStack {
      val expr = assertRoundTrips(nested)
      assertEquals(sheet.evaluateFormula(local), Right(CellValue.Number(BigDecimal(terms))))
      assertEquals(sheet.evaluateFormula(concat), Right(CellValue.Number(BigDecimal(terms))))
      assertEquals(sheet.evaluateFormula(compare), Right(CellValue.Number(BigDecimal(0))))
      assert(FormulaPrinter.printFileForm(expr).startsWith("SUM(SUM("))
      assertEquals(DependencyGraph.extractDependencies(expr), Set(ARef.from1(1, 1)))
      assert(DependencyGraph.containsCellReferences(expr))
      assert(!DependencyGraph.containsDynamicReference(expr))
      assert(!TExpr.containsExternalRef(expr) && !TExpr.containsDateFunction(expr))
      assert(FormulaShifter.referencesSheet(expr, "Other"))
      val renamed = FormulaShifter.renameSheet(expr, other, SheetName.unsafe("New"))
      assert(FormulaPrinter.print(renamed).contains("New!A1+A1"))
      val shifted = FormulaShifter.shift(expr, 1, 1)
      assertEquals(DependencyGraph.extractDependencies(shifted), Set(ARef.from1(2, 2)))
      assertEquals(FormulaParser.parse(FormulaPrinter.print(shifted)), Right(shifted))
      assertEquals(FormulaParser.parse(nested).map(_.hashCode), Right(expr.hashCode))
      // a recalculation walks the dependency graph and the evaluator over the stored formula
      val stored = sheet.put(ARef.from1(2, 2), CellValue.Formula(local.drop(1), None))
      Workbook(stored).withCachedFormulas().sheets.headOption.map(_(ARef.from1(2, 2)).value) match
        case Some(CellValue.Formula(_, Some(CellValue.Number(n)), _)) =>
          assertEquals(n, BigDecimal(terms))
        case other => fail(s"expected the recalculated chain, got ${other.map(_.getClass)}")
    }
  }

  // GH-669 review: an intersection chain `A1 A1 A1 …` builds a left spine of intersection calls that
  // the chain walkers do not unwind, so each space costs a nesting level like a postfix `%` does
  test(
    "GH-669: a 2700-term intersection chain is NestingTooDeep on a 1MB stack, never an overflow"
  ) {
    val formula = "=" + List.fill(2700)("A1").mkString(" ")
    assert(formula.length <= 8192, formula.length)
    val sheet = s.put(ARef.from1(1, 1), CellValue.Number(1))
    onSmallStack {
      assertTooDeep(formula)
      assert(sheet.evaluateFormula(formula).isLeft)
      val stored = sheet.put(ARef.from1(2, 2), CellValue.Formula(formula.drop(1), None))
      assert(stored.evaluateWithDependencyCheck().isLeft)
    }
  }

  test("GH-669: intersections share the nesting budget; the longest admitted chain walks safely") {
    assertTooDeep("=" + List.fill(129)("A1").mkString(" "))
    // GH-680 review: the levels are the operand's own spine; sibling operands of a flat chain do
    // not add up, so two 70-term chains joined by `+` parse
    val siblings = "=" + List.fill(2)(List.fill(70)("A1").mkString(" ")).mkString("+")
    val formula = "=" + List.fill(127)("A1").mkString(" ")
    val sheet = s.put(ARef.from1(1, 1), CellValue.Number(7))
    onSmallStack {
      assertRoundTrips(siblings)
      assertEquals(sheet.evaluateFormula(siblings), Right(CellValue.Number(BigDecimal(14))))
      val expr = assertRoundTrips(formula)
      assertEquals(sheet.evaluateFormula(formula), Right(CellValue.Number(BigDecimal(7))))
      assertEquals(DependencyGraph.extractDependencies(expr), Set(ARef.from1(1, 1)))
      val shifted = FormulaShifter.shift(expr, 1, 1)
      assertEquals(DependencyGraph.extractDependencies(shifted), Set(ARef.from1(2, 2)))
      assertEquals(FormulaParser.parse(formula).map(_.hashCode), Right(expr.hashCode))
      val stored = sheet.put(ARef.from1(2, 2), CellValue.Formula(formula.drop(1), None))
      Workbook(stored).withCachedFormulas().sheets.headOption.map(_(ARef.from1(2, 2)).value) match
        case Some(CellValue.Formula(_, Some(CellValue.Number(n)), _)) =>
          assertEquals(n, BigDecimal(7))
        case other => fail(s"expected the recalculated intersection, got ${other.map(_.getClass)}")
    }
  }

  // GH-680 review: a postfix `%` or an intersection spends a nesting level only within its own
  // operand — the spine it builds — never summed across the terms of a flat chain
  test("GH-680: a 200-term chain of percent operands parses, evaluates and prints back") {
    val formula = "=" + List.fill(200)("1%").mkString("+")
    onSmallStack {
      assertRoundTrips(formula)
      assertEquals(s.evaluateFormula(formula), Right(CellValue.Number(BigDecimal(2))))
    }
  }

  test("GH-680: a 200-term chain of A1*5% products parses, evaluates and prints back") {
    val formula = "=" + List.fill(200)("A1*5%").mkString("+")
    val sheet = s.put(ARef.from1(1, 1), CellValue.Number(2))
    onSmallStack {
      assertRoundTrips(formula)
      assertEquals(sheet.evaluateFormula(formula), Right(CellValue.Number(BigDecimal(20))))
    }
  }

  test("GH-680: a 200-term chain of intersection operands parses and evaluates") {
    val formula = "=" + List.fill(200)("A1 A1").mkString("+")
    val sheet = s.put(ARef.from1(1, 1), CellValue.Number(7))
    onSmallStack {
      assert(FormulaParser.parse(formula).isRight, formula.take(40))
      assertEquals(sheet.evaluateFormula(formula), Right(CellValue.Number(BigDecimal(1400))))
    }
  }

  test("GH-680: 129 percents on ONE operand still spend a level each") {
    onSmallStack {
      assertTooDeep("=1" + ("%" * 129))
      assertTooDeep("=1+1" + ("%" * 129))
      assert(FormulaParser.parse("=1" + ("%" * 127)).isRight)
    }
  }

  // GH-680 review: ROW/COLUMN/CELL quoted a non-reference argument through the case-class toString,
  // which recurses along a chain's spine — `=ROW(A1+…)` of 1025 terms exhausted a 1MB stack
  test("GH-680: ROW/COLUMN/CELL of a long chain name the argument as written, stack-safely") {
    val chain = List.fill(1025)("A1").mkString("+")
    val sheet = s.put(ARef.from1(1, 1), CellValue.Number(1))
    onSmallStack {
      for (fn, formula) <- List(
          "ROW" -> s"=ROW($chain)",
          "COLUMN" -> s"=COLUMN($chain)",
          "CELL" -> s"""=CELL("row",$chain)"""
        )
      do
        sheet.evaluateFormula(formula) match
          case Left(err) =>
            assert(err.message.contains(s"$fn requires a cell reference"), err.message.take(200))
            assert(!err.message.contains("Add("), err.message.take(200))
          case Right(v) => fail(s"$fn of a chain should refuse, got $v")
    }
  }
