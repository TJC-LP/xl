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
      .collect { case CellValue.Formula(_, Some(v)) => v }

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
