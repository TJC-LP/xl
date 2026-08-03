package com.tjclp.xl.formula

import munit.FunSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.eval.StructuralEditor
import com.tjclp.xl.formula.graph.DependencyGraph
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook

/**
 * GH-430 evaluator semantics for formula records:
 *   - a DataTable kind pins its Excel-written cache (the GH-353 seam generalized) — evaluation and
 *     recalculation never parse `TABLE(...)` and never clobber the cached `<v>` seeds;
 *   - an ArrayFormula kind evaluates scalar-wise and keeps its kind through cache refresh;
 *   - DependencyGraph excludes DataTable cells (pure value sources, not computation nodes);
 *   - StructuralEditor shifts record payloads (ref/r1/r2), degrades on edits tearing the interior,
 *     and sets del1/del2 when an edit removes an input cell (GH-435).
 */
class FormulaRecordEvalSpec extends FunSuite:

  private val S = SheetName.unsafe("S")

  private def sheetNamed(wb: Workbook, n: String): Sheet =
    wb.sheets.find(_.name == SheetName.unsafe(n)).getOrElse(fail(s"missing sheet $n"))

  private def range(a1: String): CellRange = CellRange.parse(a1).fold(fail(_), identity)
  private def aref(a1: String): ARef = ARef.parse(a1).fold(fail(_), identity)

  private def dtKind(refA1: String, r1: String, r2: String): FormulaKind.DataTable =
    FormulaKind.DataTable(
      range(refA1),
      dt2D = true,
      dtr = false,
      r1 = Some(aref(r1)),
      r2 = Some(aref(r2))
    )

  private def num(n: Int): CellValue = CellValue.Number(BigDecimal(n))

  test("evaluateCell on a dataTable cell returns the pinned cache, never a parse error") {
    val kind = dtKind("F2:G2", "A1", "A2")
    val sheet = Sheet(S)
      .put(ref"A1", num(8))
      .put(ref"F2", CellValue.dataTable(kind, Some(num(42))))
    assertEquals(sheet.evaluateCell(ref"F2"), Right(num(42)))
  }

  test("evaluateCell on an UNCACHED dataTable cell pins to Empty (total, no parse)") {
    val kind = dtKind("F2:G2", "A1", "A2")
    val sheet = Sheet(S).put(ref"F2", CellValue.dataTable(kind, None))
    assertEquals(sheet.evaluateCell(ref"F2"), Right(CellValue.Empty))
  }

  test("DependencyGraph.fromSheet excludes dataTable cells; formulas read them as values") {
    val kind = dtKind("F2:G2", "A1", "A2")
    val sheet = Sheet(S)
      .put(ref"A1", num(8))
      .put(ref"F2", CellValue.dataTable(kind, Some(num(42))))
      .put(ref"B3", CellValue.Formula("F2+1"))
    val graph = DependencyGraph.fromSheet(sheet)
    assert(!graph.dependencies.contains(ref"F2"), "dataTable cell must not be a graph node")
    assert(graph.dependencies.contains(ref"B3"))
    // A formula over the interior reads the pinned cache like any value.
    assertEquals(sheet.evaluateFormula("=F2+1"), Right(num(43)))
  }

  test("recalculate() refreshes Normal caches and leaves dataTable records byte-stable") {
    val kind = dtKind("F2:G2", "A1", "A2")
    val tableCell = CellValue.dataTable(kind, Some(num(42)))
    val uncached = CellValue.dataTable(dtKind("H2:H3", "A1", "A2"), None)
    val sheet = Sheet(S)
      .put(ref"A1", num(8))
      .put(ref"B1", CellValue.Formula("A1*2"))
      .put(ref"F2", tableCell)
      .put(ref"H2", uncached)
      .put(ref"C1", CellValue.Formula("F2+1"))
    val result = Workbook(Vector(sheet)).recalculate()
    assert(result.isClean, s"recalculate reported errors: ${result.errors}")
    val out = sheetNamed(result.workbook, "S")
    assertEquals(out(ref"B1").value, CellValue.Formula("A1*2", Some(num(16))))
    // The dataTable cache is the sole source of truth: value AND kind byte-stable,
    // and an uncached record stays uncached (no synthetic Some(Empty)).
    assertEquals(out(ref"F2").value, tableCell)
    assertEquals(out(ref"H2").value, uncached)
    // Dependents read the pinned cache.
    assertEquals(out(ref"C1").value, CellValue.Formula("F2+1", Some(num(43))))
  }

  test("recalculateDependents keeps the ArrayFormula kind through cache refresh (.copy)") {
    import com.tjclp.xl.formula.eval.DependentRecalculation.*
    val arr = FormulaKind.ArrayFormula(range("C1:C1"))
    val sheet = Sheet(S)
      .put(ref"A1", num(10))
      .put(ref"C1", CellValue.Formula("A1*2", Some(num(20)), arr))
    val updated = sheet.put(ref"A1", num(100)).recalculateDependents(Set(ref"A1"))
    assertEquals(updated(ref"C1").value, CellValue.Formula("A1*2", Some(num(200)), arr))
  }

  test("recalculate strips stale caches from deferred ArrayFormula chains") {
    val a1Kind = FormulaKind.ArrayFormula(range("A1:A1"))
    val b1Kind = FormulaKind.ArrayFormula(range("B1:B1"))
    val sheet = Sheet(S)
      .put(ref"A1", CellValue.Formula("INDIRECT(\"B1\")", Some(num(111)), a1Kind))
      .put(ref"B1", CellValue.Formula("INDIRECT(\"C1\")", Some(num(222)), b1Kind))
      .put(ref"C1", num(42))

    val result = Workbook(Vector(sheet)).recalculate()
    assert(result.isClean, s"recalculate reported errors: ${result.errors}")
    val out = sheetNamed(result.workbook, "S")
    assertEquals(
      out(ref"A1").value,
      CellValue.Formula("INDIRECT(\"B1\")", Some(num(42)), a1Kind)
    )
    assertEquals(
      out(ref"B1").value,
      CellValue.Formula("INDIRECT(\"C1\")", Some(num(42)), b1Kind)
    )
  }

  test("recalculateDependents never touches a dataTable record (pinned, not a dependent)") {
    import com.tjclp.xl.formula.eval.DependentRecalculation.*
    val tableCell = CellValue.dataTable(dtKind("F2:G2", "A1", "A2"), Some(num(42)))
    val sheet = Sheet(S)
      .put(ref"A1", num(8))
      .put(ref"F2", tableCell)
    val updated = sheet.put(ref"A1", num(999)).recalculateDependents(Set(ref"A1"))
    assertEquals(updated(ref"F2").value, tableCell)
  }

  test("evaluateWithDependencyCheck stays clean on a sheet carrying dataTable records") {
    val sheet = Sheet(S)
      .put(ref"A1", num(8))
      .put(ref"F2", CellValue.dataTable(dtKind("F2:G2", "A1", "A2"), Some(num(42))))
      .put(ref"B1", CellValue.Formula("F2*2"))
    sheet.evaluateWithDependencyCheck() match
      case Right(values) => assertEquals(values.get(ref"B1"), Some(num(84)))
      case Left(err) => fail(s"dependency-checked eval failed: $err")
  }

  // ===== StructuralEditor matrix =====

  test("insertRows above a data table shifts ref/r1/r2 and re-synthesizes the expression") {
    val tableCell = CellValue.dataTable(dtKind("F2:G2", "A1", "A2"), Some(num(42)))
    val sheet = Sheet(S).put(ref"A1", num(8)).put(ref"A2", num(3)).put(ref"F2", tableCell)
    val out = StructuralEditor.insertRows(Workbook(Vector(sheet)), S, at = 0, count = 1)
    val s2 = sheetNamed(out, "S")
    assertEquals(
      s2(ref"F3").value,
      CellValue.dataTable(dtKind("F3:G3", "A2", "A3"), Some(num(42)))
    )
  }

  test("insertRows strictly inside the interior degrades the record to its cached constant") {
    val kind: FormulaKind.DataTable = FormulaKind.DataTable(
      range("F2:F3"),
      dt2D = true,
      dtr = false,
      r1 = Some(aref("A1")),
      r2 = Some(aref("A2"))
    )
    val sheet = Sheet(S).put(ref"F2", CellValue.dataTable(kind, Some(num(11))))
    val out = StructuralEditor.insertRows(Workbook(Vector(sheet)), S, at = 2, count = 1)
    // F2 (row index0 1 < at) does not move; the insert tears its interior F2:F3.
    assertEquals(sheetNamed(out, "S")(ref"F2").value, num(11))
  }

  test("deleteRows overlapping the interior degrades the record to its cached constant") {
    val kind: FormulaKind.DataTable = FormulaKind.DataTable(
      range("F2:F3"),
      dt2D = true,
      dtr = false,
      r1 = Some(aref("A1")),
      r2 = Some(aref("A2"))
    )
    val sheet = Sheet(S).put(ref"F2", CellValue.dataTable(kind, Some(num(11))))
    // Delete row 3 (index0 2): band intersects the interior rows [1,2] but not F2 itself.
    val out = StructuralEditor.deleteRows(Workbook(Vector(sheet)), S, at = 2, count = 1)
    assertEquals(sheetNamed(out, "S")(ref"F2").value, num(11))
  }

  test("GH-435: deleteRows removing the row input sets del1, omits r1, keeps the cache") {
    val kind: FormulaKind.DataTable = FormulaKind.DataTable(
      range("F5:F6"),
      dt2D = true,
      dtr = false,
      r1 = Some(aref("A1")),
      r2 = Some(aref("A2"))
    )
    val sheet = Sheet(S).put(ref"F5", CellValue.dataTable(kind, Some(num(7))))
    // Delete row 1 (index0 0): r1=A1 goes with it; the interior itself just shifts up.
    val out = StructuralEditor.deleteRows(Workbook(Vector(sheet)), S, at = 0, count = 1)
    val torn: FormulaKind.DataTable = FormulaKind.DataTable(
      range("F4:F5"),
      dt2D = true,
      dtr = false,
      r1 = None,
      r2 = Some(aref("A1")),
      del1 = true
    )
    assertEquals(sheetNamed(out, "S")(ref"F4").value, CellValue.dataTable(torn, Some(num(7))))
  }

  test("GH-435: deleteColumns removing the column input sets del2, omits r2") {
    val kind: FormulaKind.DataTable = FormulaKind.DataTable(
      range("F5:F6"),
      dt2D = true,
      dtr = false,
      r1 = Some(aref("A1")),
      r2 = Some(aref("B1"))
    )
    val sheet = Sheet(S).put(ref"F5", CellValue.dataTable(kind, Some(num(9))))
    // Delete column B (index0 1): r2=B1 goes with it; r1=A1 sits left of the band and stands.
    val out = StructuralEditor.deleteColumns(Workbook(Vector(sheet)), S, at = 1, count = 1)
    val torn: FormulaKind.DataTable = FormulaKind.DataTable(
      range("E5:E6"),
      dt2D = true,
      dtr = false,
      r1 = Some(aref("A1")),
      r2 = None,
      del2 = true
    )
    assertEquals(sheetNamed(out, "S")(ref"E5").value, CellValue.dataTable(torn, Some(num(9))))
  }

  test("GH-435: deleting both input cells sets del1 and del2, leaving empty display slots") {
    val kind: FormulaKind.DataTable = FormulaKind.DataTable(
      range("F5:F6"),
      dt2D = true,
      dtr = false,
      r1 = Some(aref("A1")),
      r2 = Some(aref("B1"))
    )
    val sheet = Sheet(S).put(ref"F5", CellValue.dataTable(kind, Some(num(5))))
    val out = StructuralEditor.deleteColumns(Workbook(Vector(sheet)), S, at = 0, count = 2)
    val torn: FormulaKind.DataTable = FormulaKind.DataTable(
      range("D5:D6"),
      dt2D = true,
      dtr = false,
      r1 = None,
      r2 = None,
      del1 = true,
      del2 = true
    )
    val moved = sheetNamed(out, "S")(ref"D5").value
    assertEquals(moved, CellValue.dataTable(torn, Some(num(5))))
    moved match
      case CellValue.Formula(expr, _, _) => assertEquals(expr, "TABLE(,)")
      case other => fail(s"record degraded to a constant: $other")
  }

  test("GH-435: a delete missing the interior and both inputs leaves del1/del2 clear") {
    val tableCell = CellValue.dataTable(dtKind("F5:G5", "A1", "A2"), Some(num(42)))
    val sheet = Sheet(S).put(ref"F5", tableCell)
    // Row 20 (index0 19) is below everything the record references.
    val out = StructuralEditor.deleteRows(Workbook(Vector(sheet)), S, at = 19, count = 1)
    assertEquals(sheetNamed(out, "S")(ref"F5").value, tableCell)
  }

  test("GH-435: del flags are sticky — a later unrelated delete keeps them set") {
    val torn: FormulaKind.DataTable = FormulaKind.DataTable(
      range("F5:F6"),
      dt2D = true,
      dtr = false,
      r1 = None,
      r2 = Some(aref("B1")),
      del1 = true
    )
    val tableCell = CellValue.dataTable(torn, Some(num(3)))
    val sheet = Sheet(S).put(ref"F5", tableCell)
    val out = StructuralEditor.deleteRows(Workbook(Vector(sheet)), S, at = 19, count = 1)
    assertEquals(sheetNamed(out, "S")(ref"F5").value, tableCell)
  }

  test("insertRows above an array anchor shifts its record ref and keeps the kind") {
    val arr = FormulaKind.ArrayFormula(range("C8:C9"))
    val sheet = Sheet(S).put(ref"C8", CellValue.Formula("=A8*2", Some(num(2)), arr))
    val out = StructuralEditor.insertRows(Workbook(Vector(sheet)), S, at = 0, count = 1)
    // Shifted formulas re-print equals-free and invalidate their cache because the edit can
    // change both referenced values and position-sensitive results.
    assertEquals(
      sheetNamed(out, "S")(ref"C9").value,
      CellValue.Formula("A9*2", None, FormulaKind.ArrayFormula(range("C9:C10")))
    )
  }

  test("deleteRows tearing an array record degrades the KIND to Normal, keeping the text") {
    val arr = FormulaKind.ArrayFormula(range("C8:C9"))
    val sheet = Sheet(S).put(ref"C8", CellValue.Formula("=A1*2", Some(num(2)), arr))
    // Delete row 9 (index0 8): tears C8:C9; the anchor C8 (row index0 7) survives.
    // The equals-free text survives, but the structural edit invalidates the cache.
    val out = StructuralEditor.deleteRows(Workbook(Vector(sheet)), S, at = 8, count = 1)
    assertEquals(sheetNamed(out, "S")(ref"C8").value, CellValue.Formula("A1*2", None))
  }

  test("a structural edit on ANOTHER sheet leaves dataTable records untouched") {
    val other = SheetName.unsafe("Other")
    val tableCell = CellValue.dataTable(dtKind("F2:G2", "A1", "A2"), Some(num(42)))
    val s = Sheet(S).put(ref"F2", tableCell)
    val o = Sheet(other).put(ref"A1", num(1))
    val out = StructuralEditor.insertRows(Workbook(Vector(s, o)), other, at = 0, count = 2)
    assertEquals(sheetNamed(out, "S")(ref"F2").value, tableCell)
  }
