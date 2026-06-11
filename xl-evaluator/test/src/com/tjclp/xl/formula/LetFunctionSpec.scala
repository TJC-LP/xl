package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import com.tjclp.xl.formula.parser.{FormulaParser, ParseError}
import munit.{FunSuite, ScalaCheckSuite}
import org.scalacheck.Prop.*
import org.scalacheck.Gen

/**
 * GH-193: LET(name1, value1, [name2, value2, ...], calculation).
 *
 * Excel 365 lexical bindings: each binding is visible to later bindings and the body. Names are
 * case-insensitive for reuse, must start with a letter or underscore, and must not be cell-ref
 * shaped. Bindings may hold scalars or range/array values.
 */
class LetFunctionSpec extends ScalaCheckSuite:

  private def sheetWith(cells: (String, CellValue)*): Sheet =
    cells.foldLeft(Sheet(SheetName.unsafe("S"))) { case (s, (refStr, value)) =>
      ARef.parse(refStr) match
        case Right(ref) => s.put(ref, value)
        case Left(err) => fail(s"bad ref $refStr: $err")
    }

  private val sheet = sheetWith(
    "A1" -> CellValue.Number(BigDecimal(10)),
    "A2" -> CellValue.Number(BigDecimal(20)),
    "A3" -> CellValue.Number(BigDecimal(30)),
    "B1" -> CellValue.Number(BigDecimal(2))
  )

  private def evalNum(formula: String, s: Sheet = sheet): BigDecimal =
    s.evaluateFormula(formula) match
      case Right(CellValue.Number(n)) => n
      case other => fail(s"expected Number for $formula, got $other")

  // ===== Basic evaluation =====

  test("basic: LET(x, 1, x+1) = 2") {
    assertEquals(evalNum("=LET(x, 1, x+1)"), BigDecimal(2))
  }

  test("body can be just the binding: LET(x, 5, x) = 5") {
    assertEquals(evalNum("=LET(x, 5, x)"), BigDecimal(5))
  }

  test("multi-binding: LET(x, 1, y, 2, x+y) = 3") {
    assertEquals(evalNum("=LET(x, 1, y, 2, x+y)"), BigDecimal(3))
  }

  test("binding referencing prior binding: LET(x, 2, y, x*3, y+x) = 8") {
    assertEquals(evalNum("=LET(x, 2, y, x*3, y+x)"), BigDecimal(8))
  }

  test("cell-ref binding value: LET(x, A1, x*2) = 20") {
    assertEquals(evalNum("=LET(x, A1, x*2)"), BigDecimal(20))
  }

  test("string binding: LET(s, \"ab\", s&\"c\") = abc") {
    assertEquals(
      sheet.evaluateFormula("""=LET(s, "ab", s&"c")"""),
      Right(CellValue.Text("abc"))
    )
  }

  test("boolean binding: LET(b, TRUE, IF(b, 1, 2)) = 1") {
    assertEquals(evalNum("=LET(b, TRUE, IF(b, 1, 2))"), BigDecimal(1))
  }

  test("binding used multiple times in body: LET(x, 3, x*x+x) = 12") {
    assertEquals(evalNum("=LET(x, 3, x*x+x)"), BigDecimal(12))
  }

  test("numeric binding coerces to text in concat: LET(x, 1, \"v: \"&x) = 'v: 1'") {
    assertEquals(
      sheet.evaluateFormula("""=LET(x, 1, "v: "&x)"""),
      Right(CellValue.Text("v: 1"))
    )
  }

  test("LET result coerces to text in outer concat: LET(x, 1, x)&\"a\" = '1a'") {
    assertEquals(
      sheet.evaluateFormula("""=LET(x, 1, x)&"a""""),
      Right(CellValue.Text("1a"))
    )
  }

  test("date binding supports serial arithmetic: LET(d, TODAY(), d+1-(d)) = 1") {
    assertEquals(evalNum("=LET(d, TODAY(), d+1-d)"), BigDecimal(1))
  }

  test("LET composes as a scalar function argument: ABS(LET(x, 5, -x)) = 5") {
    assertEquals(evalNum("=ABS(LET(x, 5, 0-x))"), BigDecimal(5))
  }

  test("multi-line LET (whitespace/newlines) parses and evaluates") {
    val formula =
      """=LET(
        |    current, A1,
        |    doubled, current*2,
        |    doubled+1
        |)""".stripMargin
    assertEquals(evalNum(formula), BigDecimal(21))
  }

  // ===== Shadowing and scoping =====

  test("inner LET shadows outer: LET(x, 1, LET(x, 2, x)+x) = 3") {
    assertEquals(evalNum("=LET(x, 1, LET(x, 2, x)+x)"), BigDecimal(3))
  }

  test("sequential rebinding (let* semantics): LET(x, 1, x, x+1, x) = 2") {
    assertEquals(evalNum("=LET(x, 1, x, x+1, x)"), BigDecimal(2))
  }

  test("case-insensitive name reuse: LET(MyVal, 7, myval*2) = 14") {
    assertEquals(evalNum("=LET(MyVal, 7, myval*2)"), BigDecimal(14))
  }

  test("binding name shadows a function name for value refs: LET(sum, 5, sum*2) = 10") {
    assertEquals(evalNum("=LET(sum, 5, sum*2)"), BigDecimal(10))
  }

  test("function calls still work when a binding shadows the name: LET(sum, 5, SUM(A1:A2)+sum)") {
    assertEquals(evalNum("=LET(sum, 5, SUM(A1:A2)+sum)"), BigDecimal(35))
  }

  test("binding names are not visible outside the LET") {
    val result = sheet.evaluateFormula("=LET(x, 1, x)+x")
    assert(result.isLeft, s"expected error for out-of-scope name, got $result")
  }

  // ===== Range-valued bindings =====

  test("range binding in aggregate: LET(r, A1:A3, SUM(r)) = 60") {
    assertEquals(evalNum("=LET(r, A1:A3, SUM(r))"), BigDecimal(60))
  }

  test("LET returning a range flattens into an outer aggregate: SUM(LET(r, A1:A3, r)) = 60") {
    assertEquals(evalNum("=SUM(LET(r, A1:A3, r))"), BigDecimal(60))
  }

  test("range binding in criteria aggregate: LET(r, A1:A3, SUMIF(r, \">15\")) = 50") {
    assertEquals(evalNum("""=LET(r, A1:A3, SUMIF(r, ">15"))"""), BigDecimal(50))
  }

  test("range binding used twice: LET(r, A1:A3, SUM(r)/COUNT(r)) = 20") {
    assertEquals(evalNum("=LET(r, A1:A3, SUM(r)/COUNT(r))"), BigDecimal(20))
  }

  test("range binding with arithmetic broadcast: LET(r, A1:A3, SUM(r*2)) = 120") {
    assertEquals(evalNum("=LET(r, A1:A3, SUM(r*2))"), BigDecimal(120))
  }

  test("scalar binding inside an aggregate: LET(x, 5, SUM(x, A1)) = 15") {
    assertEquals(evalNum("=LET(x, 5, SUM(x, A1))"), BigDecimal(15))
  }

  test("array-call binding: LET(t, TRANSPOSE(A1:A3), SUM(t)) = 60") {
    assertEquals(evalNum("=LET(t, TRANSPOSE(A1:A3), SUM(t))"), BigDecimal(60))
  }

  test("binding holding a string criteria: LET(c, \">15\", SUMIF(A1:A3, c)) = 50") {
    assertEquals(evalNum("""=LET(c, ">15", SUMIF(A1:A3, c))"""), BigDecimal(50))
  }

  test("cross-sheet range binding with SUMIFS (issue 193 shape)") {
    val data = sheetWith("E2" -> CellValue.Text("widget"))
    val stock = Seq(
      "C2" -> CellValue.Text("widget"),
      "C3" -> CellValue.Text("gadget"),
      "C4" -> CellValue.Text("widget"),
      "D2" -> CellValue.Number(BigDecimal(5)),
      "D3" -> CellValue.Number(BigDecimal(7)),
      "D4" -> CellValue.Number(BigDecimal(11))
    ).foldLeft(Sheet(SheetName.unsafe("Sheet2"))) { case (s, (refStr, value)) =>
      ARef.parse(refStr) match
        case Right(ref) => s.put(ref, value)
        case Left(err) => fail(s"bad ref $refStr: $err")
    }
    val wb = Workbook(Vector(data, stock))
    val formula =
      "=LET(item, E2, qtys, Sheet2!D2:D4, names, Sheet2!C2:C4, total, SUMIFS(qtys, names, item), total)"
    import com.tjclp.xl.formula.eval.SheetEvaluator.*
    val result = data.evaluateFormula(formula, Clock.system, Some(wb))
    assertEquals(result, Right(CellValue.Number(BigDecimal(16))))
  }

  // ===== Name validation =====

  private def assertParseError(formula: String): ParseError =
    FormulaParser.parse(formula) match
      case Left(err) => err
      case Right(expr) => fail(s"expected parse error for $formula, got $expr")

  test("cell-ref-shaped name rejected: LET(A1, 1, A1)") {
    assertParseError("=LET(A1, 1, A1)")
  }

  test("name starting with a digit rejected") {
    assertParseError("=LET(1x, 1, 1x)")
  }

  test("boolean literal as name rejected") {
    assertParseError("=LET(TRUE, 1, TRUE)")
  }

  test("operator keywords as names rejected") {
    assertParseError("=LET(AND, 1, AND)")
    assertParseError("=LET(OR, 1, OR)")
    assertParseError("=LET(NOT, 1, NOT)")
  }

  test("underscore-led names are valid") {
    assertEquals(evalNum("=LET(_tmp, 4, _tmp+1)"), BigDecimal(5))
  }

  test("names with digits and underscores are valid") {
    assertEquals(evalNum("=LET(total_stock2, 4, total_stock2*2)"), BigDecimal(8))
  }

  // ===== Arity / structure validation =====

  test("LET() with no args is a parse error") {
    assertParseError("=LET()")
  }

  test("LET(x) with only a name is a parse error") {
    assertParseError("=LET(x)")
  }

  test("LET(x, 1) missing the body is a parse error") {
    assertParseError("=LET(x, 1)")
  }

  test("LET(x, 1, y, 2) even arg count (trailing pair, no body) is a parse error") {
    assertParseError("=LET(x, 1, y, 2)")
  }

  test("unknown bare identifier in body is a parse error") {
    assertParseError("=LET(x, 1, y)")
  }

  // ===== Error propagation =====

  test("binding evaluation error short-circuits and names the failing binding") {
    val result = sheet.evaluateFormula("=LET(bad, 1/0, 42)")
    result match
      case Left(err) =>
        assert(err.message.contains("bad"), s"error should name the binding: ${err.message}")
      case Right(v) => fail(s"expected error, got $v")
  }

  // ===== Round-trip law =====

  test("parse . print = id for representative LET formulas") {
    val formulas = List(
      "=LET(x, 1, x+1)",
      "=LET(x, 1, y, 2, x+y)",
      "=LET(r, A1:A3, SUM(r))",
      "=LET(x, A1, x*2)",
      """=LET(s, "ab", s&"c")""",
      "=LET(x, 1, LET(y, 2, x+y))"
    )
    formulas.foreach { f =>
      val parsed = FormulaParser.parse(f)
      assert(parsed.isRight, s"parse failed for $f: $parsed")
      parsed.foreach { expr =>
        val printed = FormulaPrinter.print(expr)
        val reparsed = FormulaParser.parse(printed)
        assertEquals(reparsed, Right(expr), s"round-trip failed: $f -> $printed")
      }
    }
  }

  test("printer emits canonical LET form") {
    val parsed = FormulaParser.parse("=LET(x, 1, x+1)")
    assertEquals(parsed.map(FormulaPrinter.print(_)), Right("=LET(x, 1, x+1)"))
  }

  property("parse . print = id for generated scalar LET bindings") {
    val nameGen = for
      head <- Gen.alphaChar
      tail <- Gen.listOfN(3, Gen.alphaNumChar)
      name = (head :: tail).mkString
      if ARef.parse(name).isLeft
      if !Set("TRUE", "FALSE", "AND", "OR", "NOT").contains(name.toUpperCase)
    yield name
    forAll(nameGen, Gen.choose(-1000, 1000), Gen.choose(-1000, 1000)) {
      (name: String, a: Int, b: Int) =>
        val formula = s"=LET($name, $a, $name+$b)"
        FormulaParser.parse(formula) match
          case Right(expr) =>
            FormulaParser.parse(FormulaPrinter.print(expr)) == Right(expr)
          case Left(_) => false
    }
  }

  // ===== Depth guard (GH-56) =====

  test("deeply nested LET hits the recursion guard, never StackOverflowError") {
    val formula = "=" + ("LET(x, 1, " * 300) + "x" + (")" * 300)
    FormulaParser.parse(formula) match
      case Left(_: ParseError.NestingTooDeep) => ()
      case other => fail(s"expected NestingTooDeep, got $other")
  }

  test("moderately nested LET still works") {
    val formula = "=" + ("LET(x, 1, " * 50) + "x" + (")" * 50)
    assertEquals(evalNum(formula), BigDecimal(1))
  }

  // ===== Shifting (formula dragging) =====

  test("FormulaShifter shifts refs inside LET bindings and body") {
    val parsed = FormulaParser.parse("=LET(x, A1, x+B1)")
    assert(parsed.isRight, s"$parsed")
    parsed.foreach { expr =>
      val shifted = FormulaShifter.shift(expr, 0, 1)
      assertEquals(FormulaPrinter.print(shifted), "=LET(x, A2, x+B2)")
    }
  }

  test("anchored refs inside LET stay fixed when shifted") {
    val parsed = FormulaParser.parse("=LET(x, $A$1, x+B1)")
    assert(parsed.isRight, s"$parsed")
    parsed.foreach { expr =>
      val shifted = FormulaShifter.shift(expr, 1, 1)
      assertEquals(FormulaPrinter.print(shifted), "=LET(x, $A$1, x+C2)")
    }
  }

  // ===== Dependency extraction =====

  test("dependencies include refs from binding values and body") {
    val parsed = FormulaParser.parse("=LET(x, A1, r, B1:B2, x+SUM(r))")
    assert(parsed.isRight, s"$parsed")
    parsed.foreach { expr =>
      val deps = DependencyGraph.extractDependencies(expr).map(_.toA1)
      assertEquals(deps, Set("A1", "B1", "B2"))
    }
  }

  test("recalculate orders LET cells after their dependencies") {
    val s = Sheet(SheetName.unsafe("S"))
      .put(ref"A1", CellValue.Formula("=2+3", None))
      .put(ref"B1", CellValue.Formula("=LET(x, A1, x*10)", None))
    val result = Workbook(Vector(s)).recalculate()
    assert(result.isClean, s"errors: ${result.errors}")
    assertEquals(
      result.evaluated(SheetName.unsafe("S")).get(ref"B1"),
      Some(CellValue.Number(BigDecimal(50)))
    )
  }

  // ===== Typed argument positions (total coercion at the as*Expr boundary) =====
  // BindingRef is Any-typed like PolyRef; a binding used where a function expects
  // String/Int/Boolean/BigDecimal/LocalDate must coerce TOTALLY (clean Left when
  // uncoercible), never ClassCastException through the XLResult-returning APIs.

  test("Int position: LET(k, 2, LEFT(\"hey\", k)) = 'he'") {
    assertEquals(
      sheet.evaluateFormula("""=LET(k, 2, LEFT("hey", k))"""),
      Right(CellValue.Text("he"))
    )
  }

  test("Int position via COUNT binding: LET(n, COUNT(A1:A2), MID(\"hello\", n, 1)) = 'e'") {
    assertEquals(
      sheet.evaluateFormula("""=LET(n, COUNT(A1:A2), MID("hello", n, 1))"""),
      Right(CellValue.Text("e"))
    )
  }

  test("Int position coerces boolean binding like decodeAsInt: LET(b, TRUE, LEFT(\"hey\", b)) = 'h'") {
    assertEquals(
      sheet.evaluateFormula("""=LET(b, TRUE, LEFT("hey", b))"""),
      Right(CellValue.Text("h"))
    )
  }

  test("Int position TRUNCATES a non-integral binding: LET(x, 1.7, LEFT(\"hey\", x)) = \"h\"") {
    // GH-306/GH-307: Excel truncates fractional values toward zero in integer positions
    // (LEFT("hey", 1.7) → "h"); previously this was a clean Left. Updated with the shared
    // ScalarCoercion table — totality preserved, now also Excel-correct.
    assertEquals(
      sheet.evaluateFormula("""=LET(x, 1.7, LEFT("hey", x))"""),
      Right(CellValue.Text("h"))
    )
  }

  test("String position: LET(x, A1, UPPER(x)) = '10' (parity with UPPER(A1))") {
    assertEquals(sheet.evaluateFormula("=UPPER(A1)"), Right(CellValue.Text("10")))
    assertEquals(sheet.evaluateFormula("=LET(x, A1, UPPER(x))"), Right(CellValue.Text("10")))
  }

  test("String position: LET(x, 1, LEN(x)) = 1") {
    assertEquals(evalNum("=LET(x, 1, LEN(x))"), BigDecimal(1))
  }

  test("Boolean position: LET(x, A1, IF(x, 1, 2)) matches IF(A1, 1, 2) and never throws") {
    val direct = sheet.evaluateFormula("=IF(A1, 1, 2)")
    val bound = sheet.evaluateFormula("=LET(x, A1, IF(x, 1, 2))")
    // Load-bearing assertion: totality + parity with the direct form. GH-306: numeric
    // truthiness landed (Excel: 0 = FALSE, non-zero = TRUE) — both forms now follow,
    // exactly as this pin's original comment demanded. A1 = 10 → truthy → 1.
    assertEquals(bound, direct)
    assertEquals(direct, Right(CellValue.Number(BigDecimal(1))))
  }

  test("Numeric position gives clean Left for a text binding: LET(s, \"ab\", ABS(s))") {
    val result = sheet.evaluateFormula("""=LET(s, "ab", ABS(s))""")
    assert(result.isLeft, s"expected clean Left for text in numeric position, got $result")
  }

  test("Date position: LET(d, A1, YEAR(d)) works when A1 holds a date") {
    val s = sheetWith("A1" -> CellValue.DateTime(java.time.LocalDateTime.of(2024, 3, 15, 0, 0)))
    assertEquals(evalNum("=LET(d, A1, YEAR(d))", s), BigDecimal(2024))
  }

  test("Date position: LET(d, DATE(2024, 1, 5), MONTH(d)) = 1 (LocalDate binding)") {
    assertEquals(evalNum("=LET(d, DATE(2024, 1, 5), MONTH(d))"), BigDecimal(1))
  }

  test("array binding in a scalar text position collapses: LET(t, TRANSPOSE(A1:A2), UPPER(t))") {
    // GH-302: scalar argument positions collapse arrays to their top-left value (implicit
    // intersection) instead of rejecting them; previously this was a clean Left. The bound
    // TRANSPOSE(A1:A2) is [10, 20] → top-left 10 → UPPER renders "10". Array contexts are
    // unaffected (SUM(t) above still aggregates the whole array).
    assertEquals(
      sheet.evaluateFormula("=LET(t, TRANSPOSE(A1:A2), UPPER(t))"),
      Right(CellValue.Text("10"))
    )
  }

  test("parse.print = id through typed argument positions") {
    val formulas = List(
      """=LET(k, 2, LEFT("hey", k))""",
      """=LET(n, COUNT(A1:A2), MID("hello", n, 1))""",
      "=LET(x, A1, UPPER(x))",
      "=LET(x, A1, IF(x, 1, 2))",
      "=LET(d, A1, YEAR(d))"
    )
    formulas.foreach { f =>
      val parsed = FormulaParser.parse(f)
      assert(parsed.isRight, s"parse failed for $f: $parsed")
      parsed.foreach(expr => assertEquals(FormulaPrinter.print(expr), f))
    }
  }

  test("FormulaShifter shifts through typed argument positions") {
    val parsed = FormulaParser.parse("""=LET(k, A1, LEFT("hey", k))""")
    assert(parsed.isRight, s"$parsed")
    parsed.foreach { expr =>
      val shifted = FormulaShifter.shift(expr, 1, 1)
      assertEquals(FormulaPrinter.print(shifted), """=LET(k, B2, LEFT("hey", k))""")
    }
  }

  test("structural shift survives typed argument positions (row insert above)") {
    val parsed = FormulaParser.parse("""=LET(k, A5, LEFT("hey", k))""")
    assert(parsed.isRight, s"$parsed")
    parsed.foreach { expr =>
      val shifted =
        FormulaShifter.shiftStructural(expr, shiftLocal = true, "S", isRow = true, 0, 2)
      assertEquals(
        shifted.map(e => FormulaPrinter.print(e)),
        Some("""=LET(k, A7, LEFT("hey", k))""")
      )
    }
  }

  test("recalculate is total over typed-position LET formulas (per-cell errors, never throws)") {
    val s = Sheet(SheetName.unsafe("S"))
      .put(ref"A1", CellValue.Number(BigDecimal(10)))
      .put(ref"A2", CellValue.Number(BigDecimal(20)))
      .put(ref"B1", CellValue.Formula("""=LET(k, 2, LEFT("hey", k))""", None))
      .put(ref"B2", CellValue.Formula("=LET(x, A1, UPPER(x))", None))
      .put(ref"B3", CellValue.Formula("""=LET(s, "ab", ABS(s))""", None))
      .put(ref"B4", CellValue.Formula("=LET(x, A1, IF(x, 1, 2))", None))
    // Documented contract: total whole-workbook recalculation with per-cell error
    // reporting. This call must return a RecalcResult, never throw.
    val result = Workbook(Vector(s)).recalculate()
    val evaluated = result.evaluated(SheetName.unsafe("S"))
    assertEquals(evaluated.get(ref"B1"), Some(CellValue.Text("he")))
    assertEquals(evaluated.get(ref"B2"), Some(CellValue.Text("10")))
    assert(
      result.errors.exists(e => e.ref == ref"B3"),
      s"expected a per-cell error for B3, got: ${result.errors}"
    )
  }
