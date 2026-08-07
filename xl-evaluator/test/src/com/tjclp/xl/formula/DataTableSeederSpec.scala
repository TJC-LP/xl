package com.tjclp.xl.formula

import java.io.ByteArrayInputStream
import java.nio.charset.StandardCharsets
import java.nio.file.Files
import java.time.{LocalDate, LocalDateTime}
import java.util.zip.ZipInputStream

import munit.FunSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.error.XLError
import com.tjclp.xl.ooxml.{TestFixtures, XlsxReader, XlsxWriter}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.{CalcPr, Workbook}

/**
 * GH-419 DataTableSeeder: explicit, tolerant cache seeding for data-table interiors. The
 * Excel-authored fixtures are the oracle — stripping every interior and re-seeding must reproduce
 * the caches Excel itself computed. Seeding is SHAPE-PRESERVING (kinds never change, so foreign
 * every-interior books keep their per-cell faithfulness) and never all-or-nothing (unsupported
 * corners leave their cells untouched; error VALUES are legitimate data and cache as
 * CellValue.Error).
 */
class DataTableSeederSpec extends FunSuite:

  private def range(a1: String): CellRange = CellRange.parse(a1).fold(fail(_), identity)
  private def aref(a1: String): ARef = ARef.parse(a1).fold(fail(_), identity)
  private def num(n: Int): CellValue = CellValue.Number(BigDecimal(n))

  private def sheetNamed(wb: Workbook, n: String): Sheet =
    wb.sheets.find(_.name == SheetName.unsafe(n)).getOrElse(fail(s"missing sheet $n"))

  private def colKind(refA1: String, input: String): FormulaKind.DataTable =
    FormulaKind.DataTable(
      ref = range(refA1),
      dt2D = false,
      dtr = false,
      r1 = Some(aref(input)),
      r2 = None
    )

  /** Strip an interior back to the state a builder ships: records uncached, plain values gone. */
  private def stripInterior(sheet: Sheet, interior: CellRange): Sheet =
    interior.cellsRowMajor.foldLeft(sheet) { (s, r) =>
      s.cells.get(r).map(_.value) match
        case Some(f @ CellValue.Formula(_, _, _: FormulaKind.DataTable)) =>
          s.put(r, f.copy(cachedValue = None))
        case Some(_) => s.remove(r)
        case None => s
    }

  private def readFixture(name: String): Workbook =
    XlsxReader.read(TestFixtures.copyToTemp(name)).fold(err => fail(err.message), identity)

  private def assertInteriorsMatchOracle(
    oracle: Sheet,
    seeded: Sheet,
    interiors: List[CellRange]
  ): Unit =
    interiors.foreach { interior =>
      interior.cellsRowMajor.foreach { r =>
        assertEquals(
          seeded.cells.get(r).map(_.value),
          oracle.cells.get(r).map(_.value),
          s"seeded ${r.toA1} must equal Excel's own computation"
        )
      }
    }

  test("strip-then-seed on datatable-excel.xlsx reproduces Excel's caches (all three shapes)") {
    val wb = readFixture("datatable-excel.xlsx")
    val oracle = wb.sheets.headOption.getOrElse(fail("missing sheet"))
    val interiors = List(range("F2:H4"), range("F10:F12"), range("B21:D21"))
    val stripped = interiors.foldLeft(oracle)(stripInterior)
    // Sanity: the strip really removed Excel's work.
    assertEquals(
      stripped.cells.get(aref("G3")).map(_.value),
      None,
      "strip must remove plain interiors"
    )
    val seeded = wb
      .put(stripped)
      .seedDataTables()
      .fold(err => fail(s"seeding failed: $err"), identity)
    assertInteriorsMatchOracle(oracle, sheetNamed(seeded, "Sheet1"), interiors)
  }

  test("strip-then-seed on datatable-excel-edge.xlsx: empty corner -> 0, 1x1, two-result 1-D") {
    val wb = readFixture("datatable-excel-edge.xlsx")
    val oracle = wb.sheets.headOption.getOrElse(fail("missing sheet"))
    val interiors = List(range("L2:M3"), range("Q2:Q2"), range("T2:U4"))
    val stripped = interiors.foldLeft(oracle)(stripInterior)
    val seeded = wb
      .put(stripped)
      .seedDataTables()
      .fold(err => fail(s"seeding failed: $err"), identity)
    val out = sheetNamed(seeded, "Sheet1")
    assertInteriorsMatchOracle(oracle, out, interiors)
    // Spelled out: the empty-corner grid caches Number(0) everywhere (Excel-verified bytes).
    out.cells.get(aref("L2")).map(_.value) match
      case Some(CellValue.Formula(_, cached, _)) => assertEquals(cached, Some(num(0)))
      case other => fail(s"L2 must stay a record, got $other")
    assertEquals(out.cells.get(aref("M3")).map(_.value), Some(num(0)))
  }

  test("an error-computing corner caches the error VALUE (#NUM! interiors are data)") {
    val sheet = Sheet("S")
      .put(ref"A2", num(4))
      .put(ref"F9", CellValue.Formula("SQRT(A2)"))
      .put(ref"E10", num(-4))
      .put(ref"F10", CellValue.dataTable(colKind("F10:F10", "A2"), None))
    val seeded = Workbook(sheet)
      .seedDataTables()
      .fold(err => fail(s"seeding failed: $err"), identity)
    sheetNamed(seeded, "S").cells.get(ref"F10").map(_.value) match
      case Some(CellValue.Formula(_, cached, _)) =>
        assertEquals(cached, Some(CellValue.Error(CellError.Num)))
      case other => fail(s"F10 must stay a record with an error cache, got $other")
  }

  test("an unsupported-function corner leaves its cells untouched and the op stays Right") {
    val sheet = Sheet("S")
      .put(ref"A2", num(4))
      .put(ref"F9", CellValue.Formula("NOSUCHFN(A2)"))
      .put(ref"E10", num(1))
      .put(ref"E11", num(2))
      .put(ref"F10", CellValue.dataTable(colKind("F10:F11", "A2"), None))
    val seeded = Workbook(sheet)
      .seedDataTables()
      .fold(err => fail(s"tolerant seeding must not fail the op: $err"), identity)
    val out = sheetNamed(seeded, "S")
    assertEquals(out.cells.get(ref"F10").map(_.value), Some(sheet(ref"F10").value))
    assertEquals(out.cells.get(ref"F11"), None)
  }

  test("del1/del2 records and records missing required inputs are skipped untouched") {
    val delKind = colKind("F10:F10", "A2").copy(del1 = true)
    val noInput = colKind("H10:H10", "A2").copy(r1 = None)
    val atOrigin: FormulaKind.DataTable = FormulaKind.DataTable(
      ref = range("A1:A1"),
      dt2D = true,
      dtr = true,
      r1 = Some(aref("X1")),
      r2 = Some(aref("X2"))
    )
    val sheet = Sheet("S")
      .put(ref"A2", num(4))
      .put(ref"F9", CellValue.Formula("A2*2"))
      .put(ref"E10", num(1))
      .put(ref"F10", CellValue.dataTable(delKind, None))
      .put(ref"H10", CellValue.dataTable(noInput, None))
      .put(ref"A1", CellValue.dataTable(atOrigin, None))
    val seeded = Workbook(sheet)
      .seedDataTables()
      .fold(err => fail(s"seeding failed: $err"), identity)
    val out = sheetNamed(seeded, "S")
    assertEquals(out(ref"F10").value, sheet(ref"F10").value)
    assertEquals(out(ref"H10").value, sheet(ref"H10").value)
    assertEquals(out(ref"A1").value, sheet(ref"A1").value)
  }

  test("every-interior foreign books keep their shape: kinds preserved, caches refreshed") {
    // Foreign dtr=0 every-interior dialect over D5:E5 (like the Sentinel books).
    val foreign: FormulaKind.DataTable = FormulaKind.DataTable(
      ref = range("D5:E5"),
      dt2D = true,
      dtr = false,
      r1 = Some(aref("B1")),
      r2 = Some(aref("B2"))
    )
    val sheet = Sheet("S")
      .put(ref"B1", num(1))
      .put(ref"B2", num(2))
      .put(ref"C4", CellValue.Formula("B1*B2"))
      .put(ref"D4", num(10))
      .put(ref"E4", num(20))
      .put(ref"C5", num(100))
      .put(ref"D5", CellValue.dataTable(foreign, Some(num(999))))
      .put(ref"E5", CellValue.dataTable(foreign.copy(ca = true), Some(num(999))))
    val seeded = Workbook(sheet)
      .seedDataTables()
      .fold(err => fail(s"seeding failed: $err"), identity)
    val out = sheetNamed(seeded, "S")
    // dt2D geometry: r1 <- row above (D4/E4), r2 <- column left (C5).
    assertEquals(out(ref"D5").value, CellValue.dataTable(foreign, Some(num(1000))))
    assertEquals(out(ref"E5").value, CellValue.dataTable(foreign.copy(ca = true), Some(num(2000))))
  }

  test("pinned-cache semantics stay landed after seeding: recalc never clobbers the new caches") {
    val sheet = Sheet("S")
      .put(ref"A2", num(4))
      .put(ref"F9", CellValue.Formula("A2*100"))
      .put(ref"E10", num(3))
      .put(ref"F10", CellValue.dataTable(colKind("F10:F10", "A2"), None))
    val seeded = Workbook(sheet)
      .seedDataTables()
      .fold(err => fail(s"seeding failed: $err"), identity)
    val out = sheetNamed(seeded, "S")
    assertEquals(out.evaluateCell(ref"F10"), Right(num(300)))
    val recalced = seeded.recalculate()
    assert(recalced.isClean, s"recalc errors: ${recalced.errors}")
    assertEquals(sheetNamed(recalced.workbook, "S")(ref"F10").value, out(ref"F10").value)
  }

  test("seedDataTables(sheetName) is Left(SheetNotFound) for a bad name") {
    val wb = Workbook(Sheet("S").put(ref"A1", num(1)))
    assertEquals(
      wb.seedDataTables(SheetName.unsafe("NoSuch")),
      Left(XLError.SheetNotFound("NoSuch")): XLResult[Workbook]
    )
  }

  test("the `only` scope seeds just the intersecting table") {
    val sheet = Sheet("S")
      .put(ref"A2", num(4))
      .put(ref"F9", CellValue.Formula("A2*100"))
      .put(ref"E10", num(3))
      .put(ref"F10", CellValue.dataTable(colKind("F10:F10", "A2"), None))
      .put(ref"H9", CellValue.Formula("A2*2"))
      .put(ref"G10", num(5))
      .put(ref"H10", CellValue.dataTable(colKind("H10:H10", "A2"), None))
    val seeded = Workbook(sheet)
      .seedDataTables(SheetName.unsafe("S"), Some(range("F10:F10")), Clock.system)
      .fold(err => fail(s"seeding failed: $err"), identity)
    val out = sheetNamed(seeded, "S")
    assertEquals(out(ref"F10").value, CellValue.dataTable(colKind("F10:F10", "A2"), Some(num(300))))
    assertEquals(out(ref"H10").value, CellValue.dataTable(colKind("H10:H10", "A2"), None))
  }

  test("seedDataTables() covers every sheet and resolves cross-sheet upstream refs") {
    val data = Sheet("Data").put(ref"B5", num(100))
    val s = Sheet("S")
      .put(ref"A2", num(0))
      .put(ref"F9", CellValue.Formula("Data!B5*A2"))
      .put(ref"E10", num(2))
      .put(ref"F10", CellValue.dataTable(colKind("F10:F10", "A2"), None))
    val d2 = data
      .put(ref"C9", CellValue.Formula("B5+A2"))
      .put(ref"B10", num(7))
      .put(ref"C10", CellValue.dataTable(colKind("C10:C10", "A2"), None))
    val seeded = Workbook(Vector(d2, s))
      .seedDataTables()
      .fold(err => fail(s"seeding failed: $err"), identity)
    assertEquals(
      sheetNamed(seeded, "S")(ref"F10").value,
      CellValue.dataTable(colKind("F10:F10", "A2"), Some(num(200)))
    )
    assertEquals(
      sheetNamed(seeded, "Data")(ref"C10").value,
      CellValue.dataTable(colKind("C10:C10", "A2"), Some(num(107)))
    )
  }

  test("a no-op seed never dirties: pristine fixture write stays byte-verbatim") {
    // A pristine Excel book already carries every cache the seeder would compute, so seeding
    // must not mark the sheet modified — otherwise the writer regenerates sheet1.xml instead of
    // riding the verbatim-copy loop (attribute reordering, showGridLines, t="n" drift).
    val in = TestFixtures.copyToTemp("datatable-excel.xlsx")
    val wb = XlsxReader.read(in).fold(err => fail(err.message), identity)
    val seeded = wb.seedDataTables().fold(err => fail(s"seeding failed: $err"), identity)
    val out = Files.createTempFile("xl-dt-seeder-noop-", ".xlsx")
    val outputBytes =
      try
        XlsxWriter.write(seeded, out).fold(err => fail(err.message), identity)
        Files.readAllBytes(out)
      finally Files.deleteIfExists(out)
    assertEquals(
      zipEntry(outputBytes, "xl/worksheets/sheet1.xml"),
      zipEntry(Files.readAllBytes(in), "xl/worksheets/sheet1.xml"),
      "a no-op seed must keep the clean read->write verbatim-copy fidelity law"
    )
  }

  test("a grid-covering corrupt record ref is skipped untouched (interior area cap)") {
    // The record ref comes from the FILE: a crafted B2:XFD1048576 interior is ~1.7e10 cells.
    // The seeder must skip it like any other malformed record (tolerant, Right, fast) instead
    // of materializing the extent — without the area cap this test OOMs/hangs.
    val huge: FormulaKind.DataTable = FormulaKind.DataTable(
      ref = range("B2:XFD1048576"),
      dt2D = true,
      dtr = true,
      r1 = Some(aref("A1")),
      r2 = Some(aref("A2"))
    )
    val sheet = Sheet("S")
      .put(ref"A1", num(1))
      .put(ref"A2", num(2))
      .put(ref"B2", CellValue.dataTable(huge, None))
    val seeded = Workbook(sheet)
      .seedDataTables()
      .fold(err => fail(s"oversized records must skip tolerantly, not fail the op: $err"), identity)
    assertEquals(sheetNamed(seeded, "S")(ref"B2").value, sheet(ref"B2").value)
  }

  // ===== GH-453: circular books — seeding honors the book's CalcPr / an explicit IterativeCalc =====

  /**
   * Canonical circular fixture: C1=100 (plain), B1=C1+B2, B2=$A$1*(C1+B1)/2 (declared-cycle
   * average-balance idiom over rate input A1), corner F9=IFERROR(B1/2,0), 1-var column table
   * F10:F12 over input A1 with left axis E10:E12 = 0.002/0.004/0.006. Analytic per-axis fixpoint:
   * B1 = (100+50r)/(1-r/2), interior = B1/2 — 50.1001/50.2004/50.3009, strictly increasing.
   * Pre-GH-453 every interior silently seeded the base value (FLAT).
   */
  private def circularTableSheet: Sheet =
    Sheet("S")
      .put(ref"A1", num(0))
      .put(ref"C1", num(100))
      .put(ref"B1", CellValue.Formula("C1+B2"))
      .put(ref"B2", CellValue.Formula("$A$1*(C1+B1)/2"))
      .put(ref"F9", CellValue.Formula("IFERROR(B1/2,0)"))
      .put(ref"E10", CellValue.Number(BigDecimal("0.002")))
      .put(ref"E11", CellValue.Number(BigDecimal("0.004")))
      .put(ref"E12", CellValue.Number(BigDecimal("0.006")))
      .put(ref"F10", CellValue.dataTable(colKind("F10:F12", "A1"), None))

  private val tightCalcPr =
    CalcPr(iterativeCalculation = true, Some(200), Some(BigDecimal("0.0000001")))

  private val analyticInteriors = Vector(50.1001001, 50.2004008, 50.3009027)
  private val interiorRefs = Vector(ref"F10", ref"F11", ref"F12")

  private def interiorNum(sheet: Sheet, r: ARef): Option[BigDecimal] =
    sheet.cells.get(r).map(_.value).flatMap {
      case CellValue.Number(n) => Some(n)
      case CellValue.Formula(_, Some(CellValue.Number(n)), _) => Some(n)
      case _ => None
    }

  test("GH-453: declared-cycle table seeds per-axis fixpoints under the book's calcPr") {
    val wb = Workbook(circularTableSheet).withCalcPr(tightCalcPr)
    val seeded = wb.seedDataTables().fold(err => fail(s"seeding failed: $err"), identity)
    val out = sheetNamed(seeded, "S")
    val values = interiorRefs.map(r => interiorNum(out, r).getOrElse(fail(s"${r.toA1} must seed")))
    values.zip(analyticInteriors).foreach { (v, expected) =>
      assert(
        (v.toDouble - expected).abs < 1e-4,
        s"interior $v must be within 1e-4 of the analytic fixpoint $expected"
      )
    }
    assertEquals(
      values.distinct.size,
      3,
      s"interiors must be pairwise distinct — flat interiors are the GH-453 bug: $values"
    )
  }

  test("GH-453: no calcPr -> table skipped untouched, one CircularNotIterated names the cycle") {
    val wb = Workbook(circularTableSheet)
    val report = wb.seedDataTablesReport().fold(err => fail(s"report failed: $err"), identity)
    val out = sheetNamed(report.workbook, "S")
    // Skip keeps interiors uncached, so the data-table-unseeded lint still fires downstream.
    assertEquals(out(ref"F10").value, CellValue.dataTable(colKind("F10:F12", "A1"), None))
    assertEquals(out.cells.get(ref"F11"), None)
    assertEquals(out.cells.get(ref"F12"), None)
    assertEquals(
      report.warnings,
      Vector(
        SeedTableWarning.CircularNotIterated(
          SheetName.unsafe("S"),
          range("F10:F12"),
          Vector("S!B1", "S!B2")
        )
      )
    )
    // The tolerant-Right doctrine holds on the plain overload too: Right, untouched.
    val plain = wb.seedDataTables().fold(err => fail(s"must stay Right: $err"), identity)
    assertEquals(sheetNamed(plain, "S").cells.get(ref"F11"), None)
  }

  test("GH-453: explicit IterativeCalc overload overrides an absent calcPr") {
    val wb = Workbook(circularTableSheet) // no calcPr declared
    val seeded = wb
      .seedDataTables(IterativeCalc(200, BigDecimal("0.0000001")))
      .fold(err => fail(s"seeding failed: $err"), identity)
    val out = sheetNamed(seeded, "S")
    val mid = interiorNum(out, ref"F11").getOrElse(fail("F11 must seed"))
    assert((mid.toDouble - 50.2004008).abs < 1e-4, s"F11=$mid, expected ~50.2004008")
  }

  test("GH-453: acyclic table on a cyclic book still seeds while the cyclic one skips") {
    val sheet = circularTableSheet
      .put(ref"H9", CellValue.Formula("$A$1*10"))
      .put(ref"G10", num(5))
      .put(ref"H10", CellValue.dataTable(colKind("H10:H10", "A1"), None))
    val report =
      Workbook(sheet).seedDataTablesReport().fold(err => fail(s"report failed: $err"), identity)
    val out = sheetNamed(report.workbook, "S")
    assertEquals(
      out(ref"H10").value,
      CellValue.dataTable(colKind("H10:H10", "A1"), Some(CellValue.Number(BigDecimal(50))))
    )
    assertEquals(out.cells.get(ref"F11"), None, "the cyclic table must stay unseeded")
    report.warnings match
      case Vector(SeedTableWarning.CircularNotIterated(_, tableRef, _)) =>
        assertEquals(tableRef, range("F10:F12"))
      case other => fail(s"expected exactly one CircularNotIterated, got $other")
  }

  test("GH-453: non-convergent cycle seeds last-round values and reports NotConverged") {
    // B1/B2 oscillate with period 2 from the (0,0) seed: rounds alternate (1,1)/(0,0);
    // round 5 lands on (1,1) — deterministic last-round values, per Excel.
    val sheet = Sheet("S")
      .put(ref"A1", num(0))
      .put(ref"B1", CellValue.Formula("1-B2"))
      .put(ref"B2", CellValue.Formula("1-B1"))
      .put(ref"F9", CellValue.Formula("B1"))
      .put(ref"E10", num(1))
      .put(ref"E11", num(2))
      .put(ref"E12", num(3))
      .put(ref"F10", CellValue.dataTable(colKind("F10:F12", "A1"), None))
    val wb = Workbook(sheet).withCalcPr(CalcPr(iterativeCalculation = true, Some(5), None))
    val report = wb.seedDataTablesReport().fold(err => fail(s"report failed: $err"), identity)
    val out = sheetNamed(report.workbook, "S")
    interiorRefs.foreach { r =>
      assertEquals(interiorNum(out, r), Some(BigDecimal(1)), s"${r.toA1} must seed the last round")
    }
    assertEquals(
      report.warnings,
      Vector(SeedTableWarning.NotConverged(SheetName.unsafe("S"), range("F10:F12"), 3, 5))
    )
  }

  test("GH-453: a cyclic source seeds its last Jacobi round without one extra evaluation") {
    // F9 is both the table source and a self-cycle member. From the zero seed with A1 pinned to 1,
    // its two Jacobi rounds are 2 then 3; evaluating F9 once more would incorrectly seed 3.5.
    val sheet = Sheet("S")
      .put(ref"A1", num(0))
      .put(ref"F9", CellValue.Formula("F9/2+A1+1"))
      .put(ref"E10", num(1))
      .put(ref"F10", CellValue.dataTable(colKind("F10:F10", "A1"), None))
    val wb = Workbook(sheet).withCalcPr(CalcPr(iterativeCalculation = true, Some(2), None))
    val report = wb.seedDataTablesReport().fold(err => fail(s"report failed: $err"), identity)
    val out = sheetNamed(report.workbook, "S")
    assertEquals(interiorNum(out, ref"F10"), Some(BigDecimal(3)))
    assertEquals(
      report.warnings,
      Vector(SeedTableWarning.NotConverged(SheetName.unsafe("S"), range("F10:F10"), 1, 2))
    )
  }

  test("GH-453: NotConverged reports the normalized one-round minimum") {
    val sheet = Sheet("S")
      .put(ref"A1", num(0))
      .put(ref"F9", CellValue.Formula("F9+1"))
      .put(ref"E10", num(1))
      .put(ref"F10", CellValue.dataTable(colKind("F10:F10", "A1"), None))
    val wb = Workbook(sheet).withCalcPr(CalcPr(iterativeCalculation = true, Some(0), None))
    val report = wb.seedDataTablesReport().fold(err => fail(s"report failed: $err"), identity)
    assertEquals(interiorNum(sheetNamed(report.workbook, "S"), ref"F10"), Some(BigDecimal(1)))
    assertEquals(
      report.warnings,
      Vector(SeedTableWarning.NotConverged(SheetName.unsafe("S"), range("F10:F10"), 1, 1))
    )
  }

  test("GH-453: re-seeding is a no-op and source cycle-member cells are never mutated") {
    val wb = Workbook(circularTableSheet).withCalcPr(tightCalcPr)
    val first = wb.seedDataTables().fold(err => fail(s"seeding failed: $err"), identity)
    val second = first.seedDataTables().fold(err => fail(s"re-seeding failed: $err"), identity)
    assertEquals(second, first, "a re-seed must be a no-op (deterministic fixpoints)")
    val out = sheetNamed(first, "S")
    // The what-if lives entirely on temp sheets: cycle members and the input keep their
    // original uncached formulas / value.
    assertEquals(out(ref"B1").value, CellValue.Formula("C1+B2"))
    assertEquals(out(ref"B2").value, CellValue.Formula("$A$1*(C1+B1)/2"))
    assertEquals(out(ref"A1").value, num(0))
  }

  test("GH-453: input cell inside the cycle stays pinned to the axis value (no silent flat)") {
    // The cycle runs THROUGH the what-if input: A1=B1/2, B1=A1+10. Excel semantics: the axis
    // substitution pins the input cell, breaking the cycle through it — interiors are 11/12/13,
    // NOT the unperturbed fixpoint B1=20 repeated (the silent-FLAT failure class of GH-453).
    val sheet = Sheet("S")
      .put(ref"A1", CellValue.Formula("B1/2"))
      .put(ref"B1", CellValue.Formula("A1+10"))
      .put(ref"F9", CellValue.Formula("B1"))
      .put(ref"E10", num(1))
      .put(ref"E11", num(2))
      .put(ref"E12", num(3))
      .put(ref"F10", CellValue.dataTable(colKind("F10:F12", "A1"), None))
    val wb = Workbook(sheet).withCalcPr(tightCalcPr)
    val report = wb.seedDataTablesReport().fold(err => fail(s"report failed: $err"), identity)
    assertEquals(report.warnings, Vector.empty)
    val out = sheetNamed(report.workbook, "S")
    interiorRefs.zip(Vector(11, 12, 13)).foreach { (r, expected) =>
      assertEquals(
        interiorNum(out, r),
        Some(BigDecimal(expected)),
        s"${r.toA1} must reflect the pinned axis value, not the unperturbed cycle fixpoint"
      )
    }
  }

  // ===== GH-453 review follow-ups: one volatile generation, Skipped visibility, budget guard =====

  /** Every now() read ticks one second — exposes a second volatile generation mid-seed. */
  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  private final class TickingClock extends Clock:
    private var ticks = 0
    def today(): LocalDate = LocalDate.of(2026, 1, 1)
    def now(): LocalDateTime =
      ticks += 1
      LocalDateTime.of(2026, 1, 1, 0, 0, 0).plusSeconds(ticks.toLong)

  test("GH-453: NOW() corner seeds ONE volatile generation across the interior (acyclic path)") {
    val sheet = Sheet("S")
      .put(ref"A1", num(0))
      .put(ref"F9", CellValue.Formula("NOW()"))
      .put(ref"E10", num(1))
      .put(ref"E11", num(2))
      .put(ref"E12", num(3))
      .put(ref"F10", CellValue.dataTable(colKind("F10:F12", "A1"), None))
    val seeded = Workbook(sheet)
      .seedDataTables(SheetName.unsafe("S"), None, new TickingClock)
      .fold(err => fail(s"seeding failed: $err"), identity)
    val out = sheetNamed(seeded, "S")
    val values = interiorRefs.map { r =>
      out.cells.get(r).map(_.value) match
        case Some(CellValue.Formula(_, Some(v), _)) => v
        case Some(v) => v
        case None => fail(s"${r.toA1} must seed")
    }
    assertEquals(
      values.distinct.size,
      1,
      s"one seeding run is one volatile generation — interiors must agree: $values"
    )
  }

  test("GH-453: iterated-path axis failure leaves cells untouched and reports one Skipped") {
    // E10 (the axis value feeding F10) references a missing sheet: its evaluation is a host
    // failure, so F10 stays unseeded — the report must SAY so instead of reading clean.
    val sheet = circularTableSheet.put(ref"E10", CellValue.Formula("Missing!A1"))
    val wb = Workbook(sheet).withCalcPr(tightCalcPr)
    val report = wb.seedDataTablesReport().fold(err => fail(s"report failed: $err"), identity)
    val out = sheetNamed(report.workbook, "S")
    assertEquals(out(ref"F10").value, CellValue.dataTable(colKind("F10:F12", "A1"), None))
    assert(interiorNum(out, ref"F11").isDefined, "F11 must still seed")
    assert(interiorNum(out, ref"F12").isDefined, "F12 must still seed")
    report.warnings match
      case Vector(SeedTableWarning.Skipped(sheet, tableRef, cells, reason)) =>
        assertEquals(sheet, SheetName.unsafe("S"))
        assertEquals(tableRef, range("F10:F12"))
        assertEquals(cells, 1)
        assert(reason.nonEmpty, "the Skipped warning must carry a reason")
      case other => fail(s"expected exactly one Skipped, got $other")
  }

  test("GH-453: iterated table whose worst-case cost blows the budget skips with a Skipped") {
    // 3 interiors x 50,000,000 declared rounds x 2 cycle members = 3e8 member-evaluations —
    // far past the budget; the table must skip untouched instead of honoring a hostile calcPr.
    val wb = Workbook(circularTableSheet)
      .withCalcPr(CalcPr(iterativeCalculation = true, Some(50000000), None))
    val report = wb.seedDataTablesReport().fold(err => fail(s"report failed: $err"), identity)
    val out = sheetNamed(report.workbook, "S")
    assertEquals(out(ref"F10").value, CellValue.dataTable(colKind("F10:F12", "A1"), None))
    assertEquals(out.cells.get(ref"F11"), None)
    assertEquals(out.cells.get(ref"F12"), None)
    report.warnings match
      case Vector(SeedTableWarning.Skipped(sheet, tableRef, cells, reason)) =>
        assertEquals(sheet, SheetName.unsafe("S"))
        assertEquals(tableRef, range("F10:F12"))
        assertEquals(cells, 3)
        assert(reason.contains("budget"), s"the reason must name the budget: $reason")
      case other => fail(s"expected exactly one Skipped, got $other")
  }

  test("GH-453: 20x20 two-var battery over a small cycle stays within budget") {
    // 400 axis combinations x a 2-member fixpoint: guards the narrowed-iteration design —
    // a full-book recalculation per combination would blow this budget.
    val interior = range("D5:W24")
    val rowAxis = (0 until 20).foldLeft(Sheet("S")) { (s, i) =>
      s.put(ARef.from0(3 + i, 3), num(i + 1)) // D4..W4 -> input A1
    }
    val bothAxes = (0 until 20).foldLeft(rowAxis) { (s, j) =>
      s.put(ARef.from0(2, 4 + j), num(j + 1)) // C5..C24 -> input A2
    }
    val twoVar: FormulaKind.DataTable = FormulaKind.DataTable(
      ref = interior,
      dt2D = true,
      dtr = false,
      r1 = Some(aref("A1")),
      r2 = Some(aref("A2"))
    )
    val sheet = bothAxes
      .put(ref"A1", num(0))
      .put(ref"A2", num(0))
      .put(ref"Y1", CellValue.Formula("(Y2+$A$1)/2"))
      .put(ref"Y2", CellValue.Formula("Y1/2"))
      .put(ref"C4", CellValue.Formula("Y1+$A$2")) // corner: cycle fixpoint Y1 = (2/3)*A1
      .put(ref"D5", CellValue.dataTable(twoVar, None))
    val wb = Workbook(sheet).withCalcPr(CalcPr(iterativeCalculation = true, Some(100), None))
    val startNanos = System.nanoTime()
    val report = wb.seedDataTablesReport().fold(err => fail(s"report failed: $err"), identity)
    val elapsedSeconds = (System.nanoTime() - startNanos) / 1e9
    assert(elapsedSeconds < 30.0, s"seeding took ${elapsedSeconds}s, budget is 30s")
    assertEquals(report.warnings, Vector.empty)
    val out = sheetNamed(report.workbook, "S")
    // Spot-check E6 (row axis 2 -> A1, col axis 2 -> A2): (2/3)*2 + 2 = 10/3
    val e6 = interiorNum(out, ref"E6").getOrElse(fail("E6 must seed"))
    assert((e6.toDouble - 10.0 / 3.0).abs < 1e-2, s"E6=$e6, expected ~${10.0 / 3.0}")
  }

  // ===== GH-493 / GH-494: the what-if lane must evaluate the precedent cone =====

  /**
   * GH-494 gate lesson: a seeded-value assertion MUST reject text / error / non-finite reads
   * explicitly. `math.abs(NaN - x) > tol` is false, so a NaN (or a "NM " text arm coerced by a
   * lenient reader) silently satisfies a bare tolerance gate — that is exactly how the field
   * dt-twin gate passed over 100%-text grids.
   */
  private def assertSeededNumber(sheet: Sheet, r: ARef, expected: Double, tol: Double): Unit =
    val seeded = sheet.cells.get(r).map(_.value) match
      case Some(CellValue.Formula(_, Some(inner), _)) => inner
      case Some(CellValue.Formula(_, None, _)) => fail(s"${r.toA1} must seed a cached value")
      case Some(other) => other
      case None => fail(s"${r.toA1} must seed")
    val actual = seeded match
      case CellValue.Number(n) => n.toDouble
      case other => fail(s"${r.toA1} must seed a NUMBER, got $other")
    assert(actual.isFinite, s"${r.toA1} seeded a non-finite value: $actual")
    assert((actual - expected).abs <= tol, s"${r.toA1}=$actual, expected ~$expected (tol $tol)")

  test("GH-493: acyclic table with a CACHED intermediate precedent seeds a live grid, not FLAT") {
    // B1 carries the cache a prior recalculation wrote (the CLI `recalc --tables` lane ALWAYS
    // does this before seeding). Pre-GH-493 the what-if never propagated through B1's cache:
    // every interior seeded 0*10+1 = 1 — a FLAT grid with zero warnings.
    val sheet = Sheet("S")
      .put(ref"A1", num(0)) // what-if input
      .put(ref"B1", CellValue.Formula("A1*10", Some(num(0)))) // cached intermediate
      .put(ref"F9", CellValue.Formula("B1+1")) // corner
      .put(ref"E10", num(1))
      .put(ref"E11", num(2))
      .put(ref"E12", num(3))
      .put(ref"F10", CellValue.dataTable(colKind("F10:F12", "A1"), None))
    val report =
      Workbook(sheet).seedDataTablesReport().fold(err => fail(s"report failed: $err"), identity)
    assertEquals(report.warnings, Vector.empty)
    val out = sheetNamed(report.workbook, "S")
    interiorRefs.zip(Vector(11.0, 21.0, 31.0)).foreach { (r, expected) =>
      assertSeededNumber(out, r, expected, 1e-9)
    }
    // The cone is re-derived on TEMP sheets only: B1 keeps its authored formula AND its cache.
    assertEquals(out(ref"B1").value, CellValue.Formula("A1*10", Some(num(0))))
    assertEquals(out(ref"A1").value, num(0))
  }

  test("GH-493: a cross-sheet cached intermediate is re-derived too") {
    val other = Sheet("Calc").put(ref"A1", CellValue.Formula("S!A1*10", Some(num(0))))
    val table = Sheet("S")
      .put(ref"A1", num(0))
      .put(ref"F9", CellValue.Formula("Calc!A1+1"))
      .put(ref"E10", num(1))
      .put(ref"E11", num(2))
      .put(ref"E12", num(3))
      .put(ref"F10", CellValue.dataTable(colKind("F10:F12", "A1"), None))
    val seeded = Workbook(table, other)
      .seedDataTables()
      .fold(err => fail(s"seeding failed: $err"), identity)
    val out = sheetNamed(seeded, "S")
    interiorRefs.zip(Vector(11.0, 21.0, 31.0)).foreach { (r, expected) =>
      assertSeededNumber(out, r, expected, 1e-9)
    }
    // The other sheet is never mutated by the what-if.
    assertEquals(sheetNamed(seeded, "Calc")(ref"A1").value, other(ref"A1").value)
  }

  /**
   * GH-494: the house-standard guarded headline over a cash-flow strip whose flows are UNCACHED
   * formulas. `numericValues`/`dateValues` drop undecodable cells and SHRINK the array, so XIRR's
   * length guard fires, ISERROR sees TRUE and the seeder banks `"NM "` into every interior.
   */
  private def xirrSeedSheet: Sheet =
    Sheet("S")
      .put(ref"A1", num(1)) // what-if input: exit multiple
      .put(ref"D1", num(1000))
      .put(ref"B1", CellValue.Formula("-D1")) // uncached entry flow
      .put(ref"B2", CellValue.Formula("D1*A1")) // uncached, input-dependent exit flow
      .put(ref"C1", CellValue.DateTime(LocalDate.of(2020, 1, 1).atStartOfDay))
      .put(ref"C2", CellValue.DateTime(LocalDate.of(2022, 1, 1).atStartOfDay))
      .put(ref"F9", CellValue.Formula("IF(ISERROR(XIRR(B1:B2,C1:C2)),\"NM \",XIRR(B1:B2,C1:C2))"))
      .put(ref"E10", CellValue.Number(BigDecimal("1.5")))
      .put(ref"E11", CellValue.Number(BigDecimal("2.0")))
      .put(ref"E12", CellValue.Number(BigDecimal("2.5")))
      .put(ref"F10", CellValue.dataTable(colKind("F10:F12", "A1"), None))

  test("GH-494: guarded XIRR corner seeds real rates, not its ISERROR text arm") {
    val report = Workbook(xirrSeedSheet)
      .seedDataTablesReport()
      .fold(err => fail(s"report failed: $err"), identity)
    val out = sheetNamed(report.workbook, "S")
    // 2020-01-01 -> 2022-01-01 is 731 days (2020 is a leap year): rate = m^(365/731) - 1.
    Vector(1.5, 2.0, 2.5).zip(interiorRefs).foreach { (m, r) =>
      assertSeededNumber(out, r, math.pow(m, 365.0 / 731.0) - 1.0, 1e-6)
    }
    assertEquals(report.warnings, Vector.empty, "a fully-evaluated guard must not warn")
  }

  test("GH-494: a guard that really resolves to its error arm reports ErrorGuardFired") {
    // All-positive flows: XIRR has no sign change, so the guard legitimately fires for every
    // axis combination. The seeded "NM " text is correct — but it must NEVER read as clean.
    val sheet = xirrSeedSheet.put(ref"B1", CellValue.Formula("D1"))
    val report =
      Workbook(sheet).seedDataTablesReport().fold(err => fail(s"report failed: $err"), identity)
    val out = sheetNamed(report.workbook, "S")
    interiorRefs.foreach { r =>
      assertEquals(
        out.cells.get(r).map(_.value).collect {
          case CellValue.Formula(_, Some(v), _) => v
          case v => v
        },
        Some(CellValue.Text("NM ")),
        s"${r.toA1} must bank the guard's text arm"
      )
    }
    report.warnings match
      case Vector(SeedTableWarning.ErrorGuardFired(s, tableRef, cells, guard)) =>
        assertEquals(s, SheetName.unsafe("S"))
        assertEquals(tableRef, range("F10:F12"))
        assertEquals(cells, 3)
        assert(guard.contains("XIRR"), s"the warning must name the guarded expression: $guard")
      case other => fail(s"expected exactly one ErrorGuardFired, got $other")
  }

  test("GH-494: the circular lane evaluates the cone too (guarded XIRR over a cycle)") {
    // The exit flow reads a declared cycle (average-balance idiom); the iterated lane strips the
    // downstream caches, so without cone evaluation the stripped flow VANISHES from the array.
    val sheet = xirrSeedSheet
      .put(ref"B2", CellValue.Formula("G1*A1"))
      .put(ref"G1", CellValue.Formula("1000+G2/1000", Some(num(1000))))
      .put(ref"G2", CellValue.Formula("G1-1000", Some(num(0))))
    val wb = Workbook(sheet).withCalcPr(tightCalcPr)
    val report = wb.seedDataTablesReport().fold(err => fail(s"report failed: $err"), identity)
    val out = sheetNamed(report.workbook, "S")
    Vector(1.5, 2.0, 2.5).zip(interiorRefs).foreach { (m, r) =>
      assertSeededNumber(out, r, math.pow(m * 1000.0 / 1000.0, 365.0 / 731.0) - 1.0, 1e-3)
    }
    assert(
      !report.warnings.exists {
        case _: SeedTableWarning.ErrorGuardFired => true
        case _ => false
      },
      s"the guard must not fire once the cone is evaluated: ${report.warnings}"
    )
  }

  test("GH-494: an IFERROR guard fires per COMBINATION, not per table") {
    // Axis 0 divides by zero (guard fires, "n/a" seeded); axes 1 and 2 are ordinary numbers.
    val sheet = Sheet("S")
      .put(ref"A1", num(1))
      .put(ref"F9", CellValue.Formula("IFERROR(1/A1,\"n/a\")"))
      .put(ref"E10", num(0))
      .put(ref"E11", num(1))
      .put(ref"E12", num(2))
      .put(ref"F10", CellValue.dataTable(colKind("F10:F12", "A1"), None))
    val report =
      Workbook(sheet).seedDataTablesReport().fold(err => fail(s"report failed: $err"), identity)
    val out = sheetNamed(report.workbook, "S")
    assertSeededNumber(out, ref"F11", 1.0, 1e-9)
    assertSeededNumber(out, ref"F12", 0.5, 1e-9)
    report.warnings match
      case Vector(SeedTableWarning.ErrorGuardFired(_, tableRef, cells, guard)) =>
        assertEquals(tableRef, range("F10:F12"))
        assertEquals(cells, 1, "only the divide-by-zero combination took the fallback")
        assert(guard.contains("A1"), s"the warning must name the guarded expression: $guard")
      case other => fail(s"expected exactly one ErrorGuardFired, got $other")
  }

  test("GH-493: a cone the interior cannot afford skips the table with a budget Skipped") {
    // 1,000,000 interior cells x a 21-cell axis-dependent cone = 2.1e7 cell evaluations, past the
    // cone budget: the table must skip untouched rather than grind (or seed a stale FLAT grid).
    val base = Sheet("S").put(ref"A1", num(1)).put(ref"A2", num(1)) // the two what-if inputs
    val chain = (0 until 21).foldLeft(base) { (s, i) =>
      val from = if i == 0 then "A1" else ARef.from0(1, i - 1).toA1
      s.put(ARef.from0(1, i), CellValue.Formula(s"$from+1")) // B1..B21
    }
    val interior = CellRange(ARef.from0(3, 4), ARef.from0(1002, 1003)) // 1000 x 1000
    val huge: FormulaKind.DataTable = FormulaKind.DataTable(
      ref = interior,
      dt2D = true,
      dtr = false,
      r1 = Some(aref("A1")),
      r2 = Some(aref("A2"))
    )
    val sheet = chain
      .put(ref"C4", CellValue.Formula("B21")) // corner: 21 axis-dependent precedents
      .put(ARef.from0(3, 4), CellValue.dataTable(huge, None))
    val report =
      Workbook(sheet).seedDataTablesReport().fold(err => fail(s"report failed: $err"), identity)
    val out = sheetNamed(report.workbook, "S")
    assertEquals(out(ARef.from0(3, 4)).value, CellValue.dataTable(huge, None), "untouched")
    report.warnings match
      case Vector(SeedTableWarning.Skipped(_, tableRef, _, reason)) =>
        assertEquals(tableRef, interior)
        assert(reason.contains("cone budget"), s"the reason must name the budget: $reason")
      case other => fail(s"expected exactly one Skipped, got $other")
  }

  test("GH-494: a chained IFERROR ladder whose inner rung is never taken stays clean") {
    // Review rework: the inner `1/(A1-A1)` ALWAYS resolves to #DIV/0! when probed standalone, but
    // the outer rung holds for every axis value, so no fallback is ever banked. Firing here would
    // brand a perfectly live grid as a fallback grid — the same signal loss #494 filed, inverted.
    val sheet = Sheet("S")
      .put(ref"A1", num(1))
      .put(ref"F9", CellValue.Formula("IFERROR(A1*2,IFERROR(1/(A1-A1),0))"))
      .put(ref"E10", num(1))
      .put(ref"E11", num(2))
      .put(ref"F10", CellValue.dataTable(colKind("F10:F11", "A1"), None))
    val report =
      Workbook(sheet).seedDataTablesReport().fold(err => fail(s"report failed: $err"), identity)
    val out = sheetNamed(report.workbook, "S")
    assertSeededNumber(out, ref"F10", 2.0, 1e-9)
    assertSeededNumber(out, ref"F11", 4.0, 1e-9)
    assertEquals(report.warnings, Vector.empty, "no rung's fallback was ever banked")
  }

  test("GH-494: a guard inside an UNTAKEN IF branch stays clean") {
    // Review rework: the guard lives in the false branch, which no axis value reaches.
    val sheet = Sheet("S")
      .put(ref"A1", num(1))
      .put(ref"F9", CellValue.Formula("IF(A1>0,A1*2,IFERROR(1/(A1-A1),\"x\"))"))
      .put(ref"E10", num(1))
      .put(ref"E11", num(2))
      .put(ref"E12", num(3))
      .put(ref"F10", CellValue.dataTable(colKind("F10:F12", "A1"), None))
    val report =
      Workbook(sheet).seedDataTablesReport().fold(err => fail(s"report failed: $err"), identity)
    val out = sheetNamed(report.workbook, "S")
    interiorRefs.zip(Vector(2.0, 4.0, 6.0)).foreach { (r, expected) =>
      assertSeededNumber(out, r, expected, 1e-9)
    }
    assertEquals(report.warnings, Vector.empty, "the guarded branch was never evaluated")
  }

  /**
   * GH-494 (round-2 review): a guarded column table over `A1`, axis values 0/1/2 — axis 0 is the
   * combination whose protected expression errors, so exactly one interior banks the fallback.
   */
  private def guardedColumnTable(formula: String, extra: Sheet => Sheet = identity): Sheet =
    extra(
      Sheet("S")
        .put(ref"A1", num(1))
        .put(ref"F9", CellValue.Formula(formula))
        .put(ref"E10", num(0))
        .put(ref"E11", num(1))
        .put(ref"E12", num(2))
        .put(ref"F10", CellValue.dataTable(colKind("F10:F12", "A1"), None))
    )

  /** The banked value of an interior cell, reading plain values and record caches alike. */
  private def bankedValue(sheet: Sheet, r: ARef): CellValue =
    sheet.cells.get(r).map(_.value) match
      case Some(CellValue.Formula(_, Some(inner), _)) => inner
      case Some(CellValue.Formula(_, None, _)) => fail(s"${r.toA1} must seed a cached value")
      case Some(other) => other
      case None => fail(s"${r.toA1} must seed")

  private def assertOneGuardFired(warnings: Vector[SeedTableWarning], cells: Int): String =
    warnings match
      case Vector(SeedTableWarning.ErrorGuardFired(_, tableRef, fired, guard)) =>
        assertEquals(tableRef, range("F10:F12"))
        assertEquals(fired, cells)
        guard
      case other => fail(s"expected exactly one ErrorGuardFired, got $other")

  private def seedReport(sheet: Sheet): DataTableSeedReport =
    Workbook(sheet).seedDataTablesReport().fold(err => fail(s"report failed: $err"), identity)

  test("GH-494: IF(ISERROR(x), <cell ref>, x) reports the guard it banked") {
    // Round-2 review: the numeric-fallback house shape. The fallback is a bare cell ref, which
    // parses as a PolyRef inside IF's Any-typed arg slots — a value-equality probe of it goes
    // SILENT even though the guard is exactly what banked -5.
    val report = seedReport(
      guardedColumnTable("IF(ISERROR(1/A1),D1,1/A1)", _.put(ref"D1", num(-5)))
    )
    val out = sheetNamed(report.workbook, "S")
    assertEquals(bankedValue(out, ref"F10"), CellValue.Number(BigDecimal(-5)))
    assertSeededNumber(out, ref"F11", 1.0, 1e-9)
    assertSeededNumber(out, ref"F12", 0.5, 1e-9)
    val guard = assertOneGuardFired(report.warnings, 1)
    assert(guard.contains("A1"), s"the warning must name the guarded expression: $guard")
  }

  test("GH-494: an absolute-ref fallback arm reports the guard too") {
    val report = seedReport(
      guardedColumnTable("IF(ISERROR(1/A1),$D$1,1/A1)", _.put(ref"D1", num(-5)))
    )
    val out = sheetNamed(report.workbook, "S")
    assertEquals(bankedValue(out, ref"F10"), CellValue.Number(BigDecimal(-5)))
    assertOneGuardFired(report.warnings, 1)
  }

  test("GH-494: the guard reports whatever shape its arms take") {
    // Round-2 review: the fallback arm is irrelevant to the verdict — the walk asks only whether
    // ISERROR's own argument errored on the path Excel took.
    Vector(
      ("IF(ISERROR(1/A1),D1,7)", BigDecimal(-5)),
      ("IF(ISERROR(1/A1),D1,D2)", BigDecimal(-5)),
      ("IF(ISERROR(1/A1),7,D1)", BigDecimal(7))
    ).foreach { (formula, banked) =>
      val report = seedReport(
        guardedColumnTable(formula, s => s.put(ref"D1", num(-5)).put(ref"D2", num(3)))
      )
      val out = sheetNamed(report.workbook, "S")
      assertEquals(bankedValue(out, ref"F10"), CellValue.Number(banked), s"$formula at axis 0")
      assertOneGuardFired(report.warnings, 1)
    }
  }

  test("GH-494: a guard that is not the ROOT of the source formula still reports") {
    // `IFERROR(...)+5` banks 5, which no fallback arm equals on its own — the value-equality
    // probe missed every non-root guard, the straight loss of true positives the review found.
    val report = seedReport(guardedColumnTable("IFERROR(1/A1,0)+5"))
    val out = sheetNamed(report.workbook, "S")
    assertSeededNumber(out, ref"F10", 5.0, 1e-9)
    assertSeededNumber(out, ref"F11", 6.0, 1e-9)
    assertOneGuardFired(report.warnings, 1)
  }

  test("GH-494: a guard inside a function argument still reports") {
    val report = seedReport(guardedColumnTable("SUM(IFERROR(1/A1,0),5)"))
    val out = sheetNamed(report.workbook, "S")
    assertSeededNumber(out, ref"F10", 5.0, 1e-9)
    assertSeededNumber(out, ref"F11", 6.0, 1e-9)
    assertOneGuardFired(report.warnings, 1)
  }

  test("GH-494: an untaken inner rung whose fallback COINCIDES with the banked value stays clean") {
    // Round-2 review: at axis 0 the outer rung succeeds with 0, and the untaken inner fallback
    // also evaluates to 0 — value equality false-positives on the coincidence.
    val report = seedReport(guardedColumnTable("IFERROR(A1*2,IFERROR(1/(A1-A1),0))"))
    val out = sheetNamed(report.workbook, "S")
    assertSeededNumber(out, ref"F10", 0.0, 1e-9)
    assertSeededNumber(out, ref"F11", 2.0, 1e-9)
    assertSeededNumber(out, ref"F12", 4.0, 1e-9)
    assertEquals(report.warnings, Vector.empty, "the outer rung held for every combination")
  }

  test("GH-494: CHOOSE does not probe an unselected guarded value") {
    // Review regression: axis 0 makes the third argument's standalone probe divide by zero, but
    // CHOOSE selects 42 without evaluating that argument.
    val report = seedReport(guardedColumnTable("CHOOSE(1,42,IFERROR(1/A1,0))"))
    val out = sheetNamed(report.workbook, "S")
    interiorRefs.foreach(r => assertSeededNumber(out, r, 42.0, 1e-9))
    assertEquals(report.warnings, Vector.empty, "the guarded CHOOSE value was never selected")
  }

  test("GH-494: IFS and SWITCH do not probe unselected guarded values") {
    Vector(
      "IFS(TRUE,42,FALSE,IFERROR(1/A1,0))",
      "SWITCH(1,1,42,2,IFERROR(1/A1,0))"
    ).foreach { formula =>
      val report = seedReport(guardedColumnTable(formula))
      val out = sheetNamed(report.workbook, "S")
      interiorRefs.foreach(r => assertSeededNumber(out, r, 42.0, 1e-9))
      assertEquals(report.warnings, Vector.empty, s"$formula must follow only its selected path")
    }
  }

  test("GH-494: SWITCH scalar-collapses dynamic-array targets and cases") {
    Vector(
      "SWITCH(SEQUENCE(1),1,42,IFERROR(1/A1,0))",
      "SWITCH(1,SEQUENCE(1),42,IFERROR(1/A1,0))"
    ).foreach { formula =>
      val report = seedReport(guardedColumnTable(formula))
      val out = sheetNamed(report.workbook, "S")
      interiorRefs.foreach(r => assertSeededNumber(out, r, 42.0, 1e-9))
      assertEquals(report.warnings, Vector.empty, s"$formula must select 42, not its default")
    }
  }

  test("GH-494: selected CHOOSE, IFS, and SWITCH guards still report") {
    Vector(
      "CHOOSE(1,IFERROR(1/A1,0),42)",
      "IFS(TRUE,IFERROR(1/A1,0),FALSE,42)",
      "SWITCH(1,1,IFERROR(1/A1,0),2,42)"
    ).foreach { formula =>
      val report = seedReport(guardedColumnTable(formula))
      val out = sheetNamed(report.workbook, "S")
      assertSeededNumber(out, ref"F10", 0.0, 1e-9)
      assertSeededNumber(out, ref"F11", 1.0, 1e-9)
      assertOneGuardFired(report.warnings, 1)
    }
  }

  test("GH-494: IFS materializes a range condition before choosing its banked branch") {
    val selected = seedReport(
      guardedColumnTable(
        "IFS(B1:B2,IFERROR(1/A1,0),TRUE,42)",
        s => s.put(ref"B1", num(1)).put(ref"B2", num(0))
      )
    )
    val selectedOut = sheetNamed(selected.workbook, "S")
    assertSeededNumber(selectedOut, ref"F10", 0.0, 1e-9)
    assertOneGuardFired(selected.warnings, 1)

    val selectedAway = seedReport(
      guardedColumnTable(
        "IFS(B1:B2,IFERROR(1/A1,0),TRUE,42)",
        s => s.put(ref"B1", num(0)).put(ref"B2", num(1))
      )
    )
    val selectedAwayOut = sheetNamed(selectedAway.workbook, "S")
    interiorRefs.foreach(r => assertSeededNumber(selectedAwayOut, r, 42.0, 1e-9))
    assertEquals(
      selectedAway.warnings,
      Vector.empty,
      "the guarded top-left branch was not selected"
    )
  }

  test("GH-494: ISERR over #N/A is false and does not report a fired guard") {
    val report = seedReport(
      guardedColumnTable(
        "IF(ISERR(D1),99,42)",
        _.put(ref"D1", CellValue.Error(CellError.NA))
      )
    )
    val out = sheetNamed(report.workbook, "S")
    interiorRefs.foreach(r => assertSeededNumber(out, r, 42.0, 1e-9))
    assertEquals(report.warnings, Vector.empty, "ISERR excludes #N/A, so its TRUE arm was not used")
  }

  test("GH-494: ISERR still reports a non-#N/A error selected into its TRUE arm") {
    val report = seedReport(guardedColumnTable("IF(ISERR(1/A1),99,42)"))
    val out = sheetNamed(report.workbook, "S")
    assertSeededNumber(out, ref"F10", 99.0, 1e-9)
    assertSeededNumber(out, ref"F11", 42.0, 1e-9)
    assertOneGuardFired(report.warnings, 1)
  }

  test("GH-511: an IFNA corner seeds and reports its fired guard") {
    // Was pinned as "seeds NOTHING today" while IFNA had no FunctionSpec: the corner failed to
    // parse, every combination returned Left, and the tolerant acyclic lane left the whole grid
    // uncached with no warning (#506). With IFNA on the roster the walk in `errorGuardFires`
    // becomes reachable and does the discriminating thing — it fires on #N/A only.
    val report = seedReport(
      guardedColumnTable("IFNA(D1,42)", _.put(ref"D1", CellValue.Error(CellError.NA)))
    )
    val out = sheetNamed(report.workbook, "S")
    interiorRefs.foreach(r => assertSeededNumber(out, r, 42.0, 1e-9))
    assertOneGuardFired(report.warnings, interiorRefs.size)
  }

  test("GH-511: an IFNA corner over a NON-#N/A error propagates and reports no guard") {
    val report = seedReport(
      guardedColumnTable("IFNA(D1,42)", _.put(ref"D1", CellValue.Error(CellError.Div0)))
    )
    val out = sheetNamed(report.workbook, "S")
    assertEquals(
      bankedValue(out, ref"F10"),
      CellValue.Error(CellError.Div0),
      "IFNA must propagate #DIV/0! rather than swallow it"
    )
    assertEquals(report.warnings, Vector.empty, s"no fallback was taken: ${report.warnings}")
  }

  test("GH-493: a cone cell that cannot be re-derived is reported, never silently FLAT") {
    // Review rework: B1 is axis-DEPENDENT (it reads the what-if input) so it enters the cone, but
    // its cross-sheet leg is unresolvable — the same Left any unsupported function produces. The
    // cone resolution leaves it on its stale cache, so the grid goes FLAT (8 for every axis). The
    // tolerant seed stands, but it must be VISIBLE: one ConeUnresolved naming the distinct cone
    // refs. Round-2 review: NOT a Skipped — every interior here WAS seeded.
    val sheet = Sheet("S")
      .put(ref"A1", num(0))
      .put(ref"B1", CellValue.Formula("A1*10+Missing!A1", Some(num(7))))
      .put(ref"F9", CellValue.Formula("B1+1"))
      .put(ref"E10", num(1))
      .put(ref"E11", num(2))
      .put(ref"F10", CellValue.dataTable(colKind("F10:F11", "A1"), None))
    val report =
      Workbook(sheet).seedDataTablesReport().fold(err => fail(s"report failed: $err"), identity)
    val out = sheetNamed(report.workbook, "S")
    assertSeededNumber(out, ref"F10", 8.0, 1e-9)
    assertSeededNumber(out, ref"F11", 8.0, 1e-9)
    report.warnings match
      case Vector(SeedTableWarning.ConeUnresolved(s, tableRef, cells, refs)) =>
        assertEquals(s, SheetName.unsafe("S"))
        assertEquals(tableRef, range("F10:F11"))
        assertEquals(cells, 1, "one DISTINCT cone ref, reported once per table not per combination")
        assertEquals(refs, Vector("S!B1"), "the warning must name the stale cone cell")
      case other => fail(s"expected exactly one ConeUnresolved, got $other")
  }

  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def zipEntry(bytes: Array[Byte], name: String): String =
    val zip = new ZipInputStream(new ByteArrayInputStream(bytes))
    var current = zip.getNextEntry
    var result: Option[String] = None
    try
      while current != null && result.isEmpty do
        if current.getName == name then
          result = Some(new String(zip.readAllBytes(), StandardCharsets.UTF_8))
        else
          zip.closeEntry()
          current = zip.getNextEntry
    finally zip.close()
    result.getOrElse(fail(s"missing ZIP entry $name"))
