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
