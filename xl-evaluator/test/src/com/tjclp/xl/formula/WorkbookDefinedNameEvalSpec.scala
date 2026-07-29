package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.graph.DependencyGraph
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.{DefinedName, Workbook}
import munit.FunSuite

/**
 * GH-384: defined-name references resolve inside formulas.
 *
 * Names are parsed as [[TExpr.NameRef]] and resolved at evaluation time against
 * `WorkbookMetadata.definedNames`: a sheet-scoped name SHADOWS a workbook-scoped one of the same
 * identifier (OOXML/Excel semantics), lookup is case-insensitive, and every failure mode (missing
 * name, unparseable refersTo, name cycle, missing workbook context) is a clean per-cell error —
 * never a throw.
 */
class WorkbookDefinedNameEvalSpec extends FunSuite:

  private def num(i: Int): CellValue = CellValue.Number(BigDecimal(i))
  private def formula(expr: String): CellValue = CellValue.Formula(expr, None)
  private val a1 = ARef.from0(0, 0)
  private val b2 = ARef.from0(1, 1)

  private def model: Sheet =
    Sheet(SheetName.unsafe("Model"))
      .put(b2, num(2)) // Model!B2 = 2 (the 'case' toggle)
      .put(ARef.from0(0, 0), num(10)) // Model!A1
      .put(ARef.from0(0, 1), num(20)) // Model!A2
      .put(ARef.from0(0, 2), num(30)) // Model!A3

  private def withName(wb: Workbook, dn: DefinedName): Workbook =
    wb.copy(metadata = wb.metadata.copy(definedNames = wb.metadata.definedNames :+ dn))

  private def cached(wb: Workbook, sheetName: String, ref: ARef): Option[CellValue] =
    wb.sheets
      .find(_.name.value == sheetName)
      .flatMap(_.cells.get(ref))
      .map(_.value)
      .collect { case CellValue.Formula(_, Some(v), _) => v }

  test("GH-384: single-cell name resolves to the target cell's value (=case)") {
    val wb = Workbook(model).withDefinedName("case", "Model!$B$2")
    assertEquals(wb.evaluateFormula("=case", "Model"), Right(num(2)))
  }

  test("GH-384: name lookup is case-insensitive (=CASE finds 'case')") {
    val wb = Workbook(model).withDefinedName("case", "Model!$B$2")
    assertEquals(wb.evaluateFormula("=CASE", "Model"), Right(num(2)))
  }

  test("GH-384: names compose in arithmetic (=entry_mult*ltm_ebitda)") {
    val wb = Workbook(model)
      .withDefinedName("entry_mult", "Model!$A$1") // 10
      .withDefinedName("ltm_ebitda", "Model!$A$2") // 20
    assertEquals(wb.evaluateFormula("=entry_mult*ltm_ebitda", "Model"), Right(num(200)))
  }

  test("GH-384: name in a comparison gates IF (=IF(case=2, ...)) — the LBO toggle shape") {
    val wb = Workbook(model).withDefinedName("case", "Model!$B$2")
    assertEquals(
      wb.evaluateFormula("""=IF(case=2, "Downside", "Base")""", "Model"),
      Right(CellValue.Text("Downside"))
    )
  }

  test("GH-384: constant name (refersTo a literal) resolves (=tax_rate*100)") {
    val wb = Workbook(model).withDefinedName("tax_rate", "0.25")
    assertEquals(
      wb.evaluateFormula("=tax_rate*100", "Model"),
      Right(CellValue.Number(BigDecimal(25)))
    )
  }

  test("GH-384: refersTo may carry a leading '=' (some producers write it)") {
    val wb = Workbook(model).withDefinedName("tax_rate", "=0.25")
    assertEquals(
      wb.evaluateFormula("=tax_rate*4", "Model"),
      Right(CellValue.Number(BigDecimal(1)))
    )
  }

  test("GH-384: range name aggregates (=SUM(rev_range))") {
    val wb = Workbook(model).withDefinedName("rev_range", "Model!$A$1:$A$3")
    assertEquals(wb.evaluateFormula("=SUM(rev_range)", "Model"), Right(num(60)))
  }

  test("GH-384: name in a typed argument position coerces (=EOMONTH(named_date, 0))") {
    val sheet = Sheet(SheetName.unsafe("Model"))
      .put(a1, CellValue.DateTime(java.time.LocalDateTime.of(2026, 1, 15, 0, 0)))
    val wb = Workbook(sheet).withDefinedName("named_date", "Model!$A$1")
    wb.evaluateFormula("=EOMONTH(named_date, 0)", "Model") match
      case Right(CellValue.DateTime(dt)) => assertEquals(dt.toLocalDate.getDayOfMonth, 31)
      case other => fail(s"Expected end-of-month date, got $other")
  }

  test("GH-384: sheet-scoped name shadows workbook-scoped name of the same identifier") {
    val other = Sheet(SheetName.unsafe("Other")).put(a1, num(99))
    val wb0 = Workbook(model, other)
    // Workbook-scoped: case -> Model!B2 (2). Sheet-scoped to Other (index 1): case -> Other!A1 (99).
    val wb = withName(
      withName(wb0, DefinedName("case", "Model!$B$2")),
      DefinedName("case", "Other!$A$1", localSheetId = Some(1))
    )
    assertEquals(wb.evaluateFormula("=case", "Other"), Right(num(99)), "sheet scope must shadow")
    assertEquals(wb.evaluateFormula("=case", "Model"), Right(num(2)), "global scope elsewhere")
  }

  test("GH-384: name-of-name chain resolves (alias -> case -> Model!B2)") {
    val wb = Workbook(model)
      .withDefinedName("case", "Model!$B$2")
      .withDefinedName("alias", "case")
    assertEquals(wb.evaluateFormula("=alias", "Model"), Right(num(2)))
  }

  test("GH-384: name cycle (A->B->A) is a clean error, never a throw or hang") {
    val wb = Workbook(model)
      .withDefinedName("aa", "bb")
      .withDefinedName("bb", "aa")
    wb.evaluateFormula("=aa", "Model") match
      case Left(err) => assert(err.message.toLowerCase.contains("cycle"), s"got: ${err.message}")
      case Right(v) => fail(s"name cycle must not evaluate, got $v")
  }

  test("GH-384: missing name is a clean per-cell error naming the identifier") {
    val wb = Workbook(model)
    wb.evaluateFormula("=no_such_name", "Model") match
      case Left(err) => assert(err.message.contains("no_such_name"), s"got: ${err.message}")
      case Right(v) => fail(s"missing name must not evaluate, got $v")
  }

  test("GH-384: unparseable refersTo is a clean per-cell error") {
    val wb = withName(Workbook(model), DefinedName("broken", "###not-a-formula###"))
    wb.evaluateFormula("=broken", "Model") match
      case Left(err) => assert(err.message.contains("broken"), s"got: ${err.message}")
      case Right(v) => fail(s"unparseable refersTo must not evaluate, got $v")
  }

  test("GH-384: single-sheet eval without workbook context is a clean error (SheetRef posture)") {
    import com.tjclp.xl.formula.eval.SheetEvaluator.*
    model.evaluateFormula("=case") match
      case Left(err) =>
        assert(err.message.contains("workbook"), s"got: ${err.message}")
      case Right(v) => fail(s"expected clean error without workbook context, got $v")
  }

  test("GH-384: LET binding shadows a defined name of the same identifier") {
    val wb = Workbook(model).withDefinedName("case", "Model!$B$2") // 2
    assertEquals(
      wb.evaluateFormula("=LET(case, 1, case+1)", "Model"),
      Right(CellValue.Number(BigDecimal(2))) // binding (1) + 1, NOT defined name (2) + 1
    )
  }

  test("GH-384: dependency graph carries an edge from the name user to the name's target") {
    val other = Sheet(SheetName.unsafe("Other")).put(a1, formula("=IF(case=2, 10, 20)"))
    val wb = Workbook(model.put(b2, formula("=1+1")), other).withDefinedName("case", "Model!$B$2")
    val graph = DependencyGraph.fromWorkbook(wb)
    val edges = graph.getOrElse(QualifiedRef(SheetName.unsafe("Other"), a1), Set.empty)
    assert(
      edges.contains(QualifiedRef(SheetName.unsafe("Model"), b2)),
      s"expected Other!A1 -> Model!B2 edge via defined name, got $edges"
    )
  }

  test("GH-384: recalculate() orders a computed toggle before its name-gated dependents") {
    // Model!B2 is COMPUTED (=1+1); Other!A1 gates on the name 'case' -> Model!$B$2.
    // Without the name edge, Kahn could order Other!A1 first and read a stale/uncached toggle.
    val other = Sheet(SheetName.unsafe("Other")).put(a1, formula("=IF(case=2, 10, 20)"))
    val wb = Workbook(model.put(b2, formula("=1+1")), other).withDefinedName("case", "Model!$B$2")
    val result = wb.recalculate()
    assert(result.isClean, s"expected clean recalc, got: ${result.errors.map(_.render)}")
    assertEquals(cached(result.workbook, "Other", a1), Some(num(10)))
  }

  test("GH-384: cell cycle THROUGH a name is detected as circular by recalculate()") {
    // Model!A9 = "=user" where user -> Model!$B$9, and Model!B9 = "=A9" — a cycle only visible
    // when the graph resolves the name to its target.
    val a9 = ARef.from0(0, 8)
    val b9 = ARef.from0(1, 8)
    val sheet = Sheet(SheetName.unsafe("Model"))
      .put(a9, formula("=user*1"))
      .put(b9, formula("=A9"))
    val wb = Workbook(sheet).withDefinedName("user", "Model!$B$9")
    val result = wb.recalculate()
    assert(!result.isClean)
    assert(
      result.errors.exists(_.error.message.contains("Circular")),
      s"expected a Circular reference error via the name edge, got: ${result.errors.map(_.render)}"
    )
  }

  test("GH-384: recalculate() with defined-name formulas caches values (probe shape)") {
    val other = Sheet(SheetName.unsafe("Calc"))
      .put(a1, formula("=entry_mult*ltm_ebitda"))
    val wb = Workbook(model, other)
      .withDefinedName("entry_mult", "Model!$A$1")
      .withDefinedName("ltm_ebitda", "Model!$A$2")
    val result = wb.recalculate()
    assert(result.isClean, s"expected clean, got: ${result.errors.map(_.render)}")
    assertEquals(cached(result.workbook, "Calc", a1), Some(num(200)))
  }

  test("GH-384: name cycle inside recalculate() is a per-cell error, never a throw") {
    val sheet = Sheet(SheetName.unsafe("Model")).put(a1, formula("=aa"))
    val wb = Workbook(sheet)
      .withDefinedName("aa", "bb")
      .withDefinedName("bb", "aa")
    val result = wb.recalculate() // must not throw or hang
    assert(!result.isClean)
    assertEquals(result.errors.map(_.ref), Vector(a1))
  }

  // ==========================================================================
  // GH-394: names + sheet-qualified refs in RANGE-TYPED argument positions
  // ==========================================================================

  test("GH-394: =VLOOKUP(x, named_table, 2) parses and evaluates (issue repro)") {
    // Lookup table on Model: A1:B3 = (10, 20), (20, 30), (30, 40)-ish — reuse model column A
    // as keys and add a value column
    val sheet = model
      .put(ARef.from0(1, 0), num(101))
      .put(ARef.from0(1, 2), num(103))
    val wb = Workbook(sheet).withDefinedName("named_table", "Model!$A$1:$B$3")
    assertEquals(
      wb.evaluateFormula("=VLOOKUP(10, named_table, 2, FALSE)", "Model"),
      Right(num(101))
    )
  }

  test("GH-394: =SUMIF(named_range, criteria) works") {
    val wb = Workbook(model).withDefinedName("rev_range", "Model!$A$1:$A$3")
    assertEquals(
      wb.evaluateFormula("=SUMIF(rev_range, \">15\")", "Model"),
      Right(num(50)) // 20 + 30
    )
  }

  test("GH-394: =INDEX(named_range, n) and =MATCH(x, named_range) work") {
    val wb = Workbook(model).withDefinedName("rev_range", "Model!$A$1:$A$3")
    assertEquals(wb.evaluateFormula("=INDEX(rev_range, 2)", "Model"), Right(num(20)))
    assertEquals(
      wb.evaluateFormula("=MATCH(30, rev_range, 0)", "Model"),
      Right(CellValue.Number(BigDecimal(3)))
    )
  }

  test("GH-394: XIRR with sheet-qualified literal ranges parses WITHOUT sheet context") {
    // Field repro: `xl eval "=XIRR(S!A1:B1,S!A2:B2)"` failed at parse with
    // InvalidArguments(XIRR,0,range,...) while the in-sheet form worked
    val serial2020 = CellValue.Number(
      BigDecimal(CellValue.dateTimeToExcelSerial(java.time.LocalDateTime.of(2020, 1, 1, 0, 0)))
    )
    val sheet = Sheet(SheetName.unsafe("S"))
      .put(ARef.from0(0, 0), num(-1000)) // S!A1
      .put(ARef.from0(1, 0), num(1100)) // S!B1
      .put(ARef.from0(0, 1), serial2020) // S!A2
      .put(
        ARef.from0(1, 1),
        CellValue.DateTime(java.time.LocalDateTime.of(2021, 1, 1, 0, 0))
      ) // S!B2
    val wb = Workbook(sheet)
    // Parse with NO sheet context (the ad-hoc eval path)
    val parsed = FormulaParser.parse("=XIRR(S!A1:B1,S!A2:B2)")
    assert(parsed.isRight, s"sheet-qualified ranges must parse in range slots, got $parsed")
    wb.evaluateFormula("=XIRR(S!A1:B1, S!A2:B2)", "S") match
      case Right(CellValue.Number(rate)) =>
        assert((rate.toDouble - 0.1).abs < 0.01, s"expected ~0.1, got $rate")
      case other => fail(s"expected XIRR rate, got $other")
  }

  test("GH-394: NPV/XNPV with sheet-qualified ranges evaluate") {
    val sheet = Sheet(SheetName.unsafe("S"))
      .put(ARef.from0(0, 0), num(100))
      .put(ARef.from0(0, 1), num(200))
    val wb = Workbook(sheet)
    wb.evaluateFormula("=NPV(0.1, S!A1:A2)", "S") match
      case Right(CellValue.Number(v)) =>
        val expected = 100.0 / 1.1 + 200.0 / 1.21
        assert((v.toDouble - expected).abs < 0.01, s"got $v, expected ~$expected")
      case other => fail(s"expected NPV value, got $other")
  }

  test("GH-394: =Model!case (sheet-qualified NAME) parses and evaluates") {
    val wb = Workbook(model).withDefinedName("case", "Model!$B$2")
    val parsed = FormulaParser.parse("=Model!case")
    assert(parsed.isRight, s"sheet-qualified name must parse, got $parsed")
    assertEquals(wb.evaluateFormula("=Model!case", "Model"), Right(num(2)))
    // and from ANOTHER sheet (the point of qualification)
    val other = Sheet(SheetName.unsafe("Other"))
    val wb2 = Workbook(model, other).withDefinedName("case", "Model!$B$2")
    assertEquals(wb2.evaluateFormula("=Model!case*10", "Other"), Right(num(20)))
  }

  test("GH-394: sheet-qualified name picks the SHEET-SCOPED binding of its qualifier") {
    val other = Sheet(SheetName.unsafe("Other")).put(a1, num(99))
    val wb0 = Workbook(model, other)
    val wb = withName(
      withName(wb0, DefinedName("case", "Model!$B$2")), // workbook-scoped -> 2
      DefinedName("case", "Other!$A$1", localSheetId = Some(1)) // Other-scoped -> 99
    )
    assertEquals(wb.evaluateFormula("=Other!case", "Model"), Right(num(99)))
  }

  test("GH-394: name bound to a CONSTANT in a range slot is a clean error VALUE, not a crash") {
    val wb = Workbook(model).withDefinedName("tax_rate", "0.25")
    wb.evaluateFormula("=SUMIF(tax_rate, \">0\")", "Model") match
      case Right(CellValue.Error(_)) => () // Excel-class error value
      case Left(err) =>
        // a clean per-cell Left is also acceptable; a THROW is not (this line proves no throw)
        assert(err.message.contains("tax_rate"), s"error should name the identifier: $err")
      case Right(v) => fail(s"constant name in range slot must not evaluate to $v")
  }

  test("GH-394: unknown name in a range slot is a clean error naming the identifier") {
    val wb = Workbook(model)
    wb.evaluateFormula("=SUM(INDEX(no_such_table, 1))", "Model") match
      case Right(CellValue.Error(_)) => ()
      case Left(err) => assert(err.message.contains("no_such_table"), s"got: ${err.message}")
      case Right(v) => fail(s"unknown name must not evaluate, got $v")
  }

  test("GH-394: name whose target is a single CELL acts as a 1x1 range in range slots") {
    val wb = Workbook(model).withDefinedName("case", "Model!$B$2")
    assertEquals(wb.evaluateFormula("=SUMIF(case, \">0\")", "Model"), Right(num(2)))
  }

  test("GH-394: XLOOKUP with named lookup/return ranges") {
    val sheet = model
      .put(ARef.from0(1, 0), num(101))
      .put(ARef.from0(1, 1), num(102))
      .put(ARef.from0(1, 2), num(103))
    val wb = Workbook(sheet)
      .withDefinedName("keys", "Model!$A$1:$A$3")
      .withDefinedName("vals", "Model!$B$1:$B$3")
    assertEquals(
      wb.evaluateFormula("=XLOOKUP(20, keys, vals)", "Model"),
      Right(num(102))
    )
  }

  test("GH-394: NETWORKDAYS holiday range accepts a sheet-qualified range and a name") {
    val sheet = Sheet(SheetName.unsafe("Cal"))
      .put(ARef.from0(0, 0), CellValue.DateTime(java.time.LocalDateTime.of(2025, 1, 6, 0, 0)))
      .put(ARef.from0(1, 0), CellValue.DateTime(java.time.LocalDateTime.of(2025, 1, 10, 0, 0)))
      .put(ARef.from0(2, 0), CellValue.DateTime(java.time.LocalDateTime.of(2025, 1, 8, 0, 0)))
    val wb = Workbook(sheet).withDefinedName("holidays", "Cal!$C$1:$C$1")
    assertEquals(
      wb.evaluateFormula("=NETWORKDAYS(A1, B1, Cal!C1:C1)", "Cal"),
      Right(num(4))
    )
    assertEquals(
      wb.evaluateFormula("=NETWORKDAYS(A1, B1, holidays)", "Cal"),
      Right(num(4))
    )
  }

  test("GH-394: Excel name characters '.' and '\\' parse and resolve (Sales.Total)") {
    val wb = Workbook(model).withDefinedName("Sales.Total", "Model!$A$1:$A$3")
    assertEquals(wb.evaluateFormula("=SUM(Sales.Total)", "Model"), Right(num(60)))
    assertEquals(wb.evaluateFormula("=Sales.Total*0+1", "Model"), Right(num(1)))
  }

  test("GH-394: dotted LET binding names work and number literals (3.14) are unaffected") {
    val wb = Workbook(model)
    assertEquals(
      wb.evaluateFormula("=LET(sales.total, 5, sales.total*2)", "Model"),
      Right(num(10))
    )
    assertEquals(
      wb.evaluateFormula("=3.14*100", "Model"),
      Right(CellValue.Number(BigDecimal(314)))
    )
  }

  test("GH-394: round-trip — named and sheet-qualified range args print byte-identically") {
    val formulas = List(
      "=VLOOKUP(10, named_table, 2, FALSE)",
      "=SUMIF(rev_range, \">15\")",
      "=XIRR(Model!A1:A3, Model!B1:B3)",
      "=INDEX(rev_range, 2)",
      "=Model!case",
      "=SUM(Sales.Total)"
    )
    formulas.foreach { f =>
      FormulaParser.parse(f) match
        case Right(expr) =>
          assertEquals(FormulaPrinter.print(expr), f, s"print mismatch for $f")
          assertEquals(FormulaParser.parse(FormulaPrinter.print(expr)), Right(expr))
        case Left(err) => fail(s"parse failed for $f: $err")
    }
  }

  test("GH-394: recalculate() orders computed cells before a name-in-range-slot dependent") {
    // Model!A1..A3 are COMPUTED; Calc!A1 aggregates them through a NAMED range in a
    // range-typed slot. Without the workbook-level name edges, Kahn could order Calc!A1
    // first and read stale/uncached precedents.
    val m = Sheet(SheetName.unsafe("Model"))
      .put(ARef.from0(0, 0), formula("=1+9")) // 10
      .put(ARef.from0(0, 1), formula("=A1*2")) // 20
      .put(ARef.from0(0, 2), formula("=A2+10")) // 30
    val calc = Sheet(SheetName.unsafe("Calc"))
      .put(a1, formula("=SUMIF(rev_range, \">15\")"))
    val wb = Workbook(m, calc).withDefinedName("rev_range", "Model!$A$1:$A$3")
    val result = wb.recalculate()
    assert(result.isClean, s"expected clean recalc, got: ${result.errors.map(_.render)}")
    assertEquals(cached(result.workbook, "Calc", a1), Some(num(50))) // 20 + 30
  }

  // ==========================================================================
  // GH-411: name→name chains in RANGE-TYPED argument positions
  // ==========================================================================

  test("GH-411: name→name chain resolves in a RANGE slot (alias -> rev_range -> Model!A1:A3)") {
    val wb = Workbook(model)
      .withDefinedName("rev_range", "Model!$A$1:$A$3")
      .withDefinedName("alias", "rev_range")
    assertEquals(wb.evaluateFormula("=SUMIF(alias, \">15\")", "Model"), Right(num(50)))
    assertEquals(wb.evaluateFormula("=INDEX(alias, 2)", "Model"), Right(num(20)))
  }

  test("GH-411: SUM over a name whose refersTo is another name") {
    val wb = Workbook(model)
      .withDefinedName("rev_range", "Model!$A$1:$A$3")
      .withDefinedName("alias", "rev_range")
    assertEquals(wb.evaluateFormula("=SUM(alias)", "Model"), Right(num(60)))
  }

  test("GH-411: chain through a SHEET-QUALIFIED name link resolves (alias -> Model!rev_range)") {
    val other = Sheet(SheetName.unsafe("Other"))
    val wb = Workbook(model, other)
      .withDefinedName("rev_range", "Model!$A$1:$A$3")
      .withDefinedName("alias", "Model!rev_range")
    assertEquals(wb.evaluateFormula("=SUMIF(alias, \">15\")", "Other"), Right(num(50)))
  }

  test("GH-411: name cycle in a RANGE slot is a clean cycle error, never a throw or hang") {
    val wb = Workbook(model)
      .withDefinedName("aa", "bb")
      .withDefinedName("bb", "aa")
    wb.evaluateFormula("=SUMIF(aa, \">0\")", "Model") match
      case Left(err) => assert(err.message.toLowerCase.contains("cycle"), s"got: ${err.message}")
      case Right(v) => fail(s"range-slot name cycle must not evaluate, got $v")
  }

  test("GH-411: chain ending in a NON-RANGE stays a clean per-cell error naming the identifier") {
    val wb = Workbook(model)
      .withDefinedName("tax_rate", "0.25")
      .withDefinedName("alias", "tax_rate")
    wb.evaluateFormula("=SUMIF(alias, \">0\")", "Model") match
      case Right(CellValue.Error(_)) => () // Excel-class error value
      case Left(err) =>
        assert(err.message.contains("tax_rate"), s"error should name the identifier: $err")
      case Right(v) => fail(s"non-range chain must not evaluate to $v")
  }

  test("GH-411: backslash-led defined name (=\\name) parses and resolves — Sales.Total's twin") {
    val wb = Workbook(model).withDefinedName("\\rng", "Model!$A$1:$A$3")
    assertEquals(wb.evaluateFormula("=SUM(\\rng)", "Model"), Right(num(60)))
    assertEquals(wb.evaluateFormula("=\\rng*0+1", "Model"), Right(num(1)))
  }

  test(
    "GH-411: recalculate() orders computed precedents before a CHAINED-alias range-slot reader"
  ) {
    // Like the GH-394 ordering pin, but the range slot reads a name→name CHAIN — the graph's
    // chain-following name edges and the evaluator's chain-following range resolution must agree.
    val m = Sheet(SheetName.unsafe("Model"))
      .put(ARef.from0(0, 0), formula("=1+9")) // 10
      .put(ARef.from0(0, 1), formula("=A1*2")) // 20
      .put(ARef.from0(0, 2), formula("=A2+10")) // 30
    val calc = Sheet(SheetName.unsafe("Calc"))
      .put(a1, formula("=SUMIF(alias, \">15\")"))
    val wb = Workbook(m, calc)
      .withDefinedName("rev_range", "Model!$A$1:$A$3")
      .withDefinedName("alias", "rev_range")
    val result = wb.recalculate()
    assert(result.isClean, s"expected clean recalc, got: ${result.errors.map(_.render)}")
    assertEquals(cached(result.workbook, "Calc", a1), Some(num(50))) // 20 + 30
  }

  test("GH-411: deprecated ArgSpec.cellRange remains usable (published API, removal deferred)") {
    // Zero in-repo consumers after the GH-394 rangeLocation migration; kept for source
    // compatibility — deprecated with the migration hint, removed next breaking cycle.
    import com.tjclp.xl.formula.ast.TExpr
    import com.tjclp.xl.formula.functions.ArgSpec
    @annotation.nowarn("cat=deprecation")
    def spec: ArgSpec[CellRange] = ArgSpec.cellRange
    CellRange.parse("A1:B2") match
      case Right(r) =>
        assertEquals(spec.parse(List(TExpr.RangeRef(r)), 0, "TEST").map(_._1), Right(r))
      case Left(err) => fail(s"CellRange.parse(A1:B2) must parse, got $err")
  }
