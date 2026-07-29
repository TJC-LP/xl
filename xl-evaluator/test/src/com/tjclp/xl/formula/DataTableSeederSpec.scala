package com.tjclp.xl.formula

import munit.FunSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.error.XLError
import com.tjclp.xl.ooxml.{TestFixtures, XlsxReader}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook

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
