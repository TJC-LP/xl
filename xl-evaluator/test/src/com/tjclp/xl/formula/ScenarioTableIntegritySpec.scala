package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.FormulaKind
import com.tjclp.xl.formula.eval.{DataTableSeedReport, IterativeCalc, SeedTableWarning}

import munit.FunSuite

class ScenarioTableIntegritySpec extends FunSuite:
  private def num(value: Int): CellValue = CellValue.Number(BigDecimal(value))
  private val kind: FormulaKind.DataTable =
    FormulaKind.DataTable(ref"F10:F12", dt2D = false, dtr = false, r1 = Some(ref"A1"), r2 = None)

  private def base: Sheet = Sheet("S")
    .put(ref"A1", num(0))
    .put(ref"B1", CellValue.Formula("A1*10", Some(num(0))))
    .put(ref"F9", CellValue.Formula("B1+1"))
    .put(ref"E10", num(1))
    .put(ref"E11", num(2))
    .put(ref"E12", num(3))
    .put(ref"F10", CellValue.dataTable(kind, None))

  private def seed(wb: Workbook): DataTableSeedReport =
    wb.seedDataTablesReport().fold(error => fail(error.message), identity)

  private def skipped(report: DataTableSeedReport): Int = report.warnings.collect {
    case SeedTableWarning.Skipped(_, _, cells, _) => cells
  }.sum

  private def value(report: DataTableSeedReport, ref: ARef): Option[CellValue] =
    report.workbook.sheets.head(ref).value match
      case CellValue.Formula(_, cached, _) => cached
      case CellValue.Empty => None
      case other => Some(other)

  test("GH-498: a dynamic source is explicitly skipped instead of seeding a flat grid") {
    val wb = Workbook(base.put(ref"F9", CellValue.Formula("INDIRECT(\"B1\")+1")))
    val report = seed(wb)
    assertEquals(skipped(report), 3)
    assertEquals(report.workbook, wb)
    assert(report.warnings.exists {
      case SeedTableWarning.Skipped(_, _, _, reason) =>
        reason.contains("S!F9") && reason.contains("dynamic")
      case _ => false
    })
  }

  test("GH-498: dynamic intermediates and defined-name aliases are included in the check") {
    val indirect = base
      .put(ref"C1", CellValue.Formula("INDIRECT(\"B1\")", Some(num(0))))
      .put(ref"F9", CellValue.Formula("C1+1"))
    assertEquals(skipped(seed(Workbook(indirect))), 3)
    val named = Workbook(base.put(ref"F9", CellValue.Formula("Dynamic+1")))
      .withDefinedName("Dynamic", "INDIRECT(\"B1\")")
    assertEquals(skipped(seed(named)), 3)
  }

  test("GH-498: unrelated dynamic formulas do not prevent a static table from seeding") {
    val report = seed(Workbook(base.put(ref"Z1", CellValue.Formula("INDIRECT(\"B1\")"))))
    assertEquals(report.warnings, Vector.empty)
    List(ref"F10", ref"F11", ref"F12").zip(List(11, 21, 31)).foreach { (ref, expected) =>
      assertEquals(value(report, ref), Some(num(expected)))
    }
  }

  test("GH-498: an input formula replaced by the axis overlay is not a dynamic precedent") {
    val report =
      seed(Workbook(base.put(ref"A1", CellValue.Formula("INDIRECT(\"B1\")", Some(num(0))))))
    assertEquals(report.warnings, Vector.empty)
    assertEquals(value(report, ref"F10"), Some(num(11)))
    assertEquals(value(report, ref"F12"), Some(num(31)))
  }

  test("GH-498: ancestors of an overlaid input are outside the scenario's dependency cone") {
    val sheet = base
      .put(ref"A1", CellValue.Formula("C1", Some(num(0))))
      .put(ref"C1", CellValue.Formula("INDIRECT(\"B1\")", Some(num(0))))
    val report = seed(Workbook(sheet))
    assertEquals(report.warnings, Vector.empty)
    assertEquals(value(report, ref"F10"), Some(num(11)))
    assertEquals(value(report, ref"F12"), Some(num(31)))
  }

  test("GH-506: a failed source reports every unseeded interior") {
    val wb = Workbook(base.put(ref"F9", CellValue.Formula("AND(B1:B1)")))
    val report = seed(wb)
    assertEquals(skipped(report), 3)
    assertEquals(report.workbook, wb)
    assertEquals(value(report, ref"F10"), None)
  }

  test("GH-506: source failure preserves unresolved cone diagnostics across combinations") {
    val sheet = base.put(ref"B1", CellValue.Formula("IF(A1=1,10,AND(A1:A1))"))
    val report = seed(Workbook(sheet))
    assertEquals(value(report, ref"F10"), Some(num(11)))
    assertEquals(value(report, ref"F11"), None)
    assertEquals(value(report, ref"F12"), None)
    assertEquals(skipped(report), 2)
    assert(report.warnings.exists {
      case SeedTableWarning.ConeUnresolved(_, _, cells, refs) => cells == 1 && refs.contains("S!B1")
      case _ => false
    })
  }

  test("GH-506: failed axis values are included in the acyclic skipped count") {
    val report = seed(Workbook(base.put(ref"E10", CellValue.Formula("Missing!A1"))))
    assertEquals(skipped(report), 1)
    assertEquals(value(report, ref"F10"), None)
    assertEquals(value(report, ref"F11"), Some(num(21)))
    assertEquals(value(report, ref"F12"), Some(num(31)))
  }

  test("GH-506: an Excel error value is still a successfully evaluated interior") {
    val report = seed(Workbook(base.put(ref"F9", CellValue.Formula("1/(A1-1)"))))
    assertEquals(report.warnings, Vector.empty)
    assertEquals(value(report, ref"F10"), Some(CellValue.Error(CellError.Div0)))
    assertEquals(value(report, ref"F11"), Some(num(1)))
  }

  test("GH-506: iterative member failures retain their unresolved upstream cone") {
    val sheet = base
      .put(ref"D1", CellValue.Formula("AND(A1:A1)"))
      .put(ref"B1", CellValue.Formula("C1+D1"))
      .put(ref"C1", CellValue.Formula("B1/2"))
    val report = Workbook(sheet)
      .seedDataTablesReport(Clock.system, Some(IterativeCalc(3, BigDecimal("0.001"))))
      .fold(error => fail(error.message), identity)
    assertEquals(skipped(report), 3)
    assert(report.warnings.exists {
      case SeedTableWarning.ConeUnresolved(_, _, _, refs) => refs.contains("S!D1")
      case _ => false
    })
  }
