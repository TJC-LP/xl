package com.tjclp.xl.formula

import com.tjclp.xl.*
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.eval.DependentRecalculation.*
import com.tjclp.xl.formula.graph.DependencyGraph
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * Tests for DependentRecalculation (GH-163).
 *
 * Tests eager recalculation of dependent formulas when cells are modified.
 */
@SuppressWarnings(
  Array("org.wartremover.warts.OptionPartial", "org.wartremover.warts.IterableOps")
)
class DependentRecalculationSpec extends FunSuite:
  val emptySheet = new Sheet(name = SheetName.unsafe("Test"))

  def sheetWith(cells: (ARef, CellValue)*): Sheet =
    cells.foldLeft(emptySheet) { case (s, (ref, value)) =>
      s.put(ref, value)
    }

  // ===== Transitive Dependents Tests (6 tests) =====

  test("transitiveDependents: empty set returns empty") {
    val graph = DependencyGraph(Map.empty, Map.empty)
    assertEquals(DependencyGraph.transitiveDependents(graph, Set.empty), Set.empty[ARef])
  }

  test("transitiveDependents: single node with no dependents") {
    val sheet = sheetWith(ref"A1" -> CellValue.Number(BigDecimal(10)))
    val graph = DependencyGraph.fromSheet(sheet)
    assertEquals(DependencyGraph.transitiveDependents(graph, Set(ref"A1")), Set.empty[ARef])
  }

  test("transitiveDependents: direct dependent") {
    // A1 <- B1 (B1 depends on A1)
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"B1" -> CellValue.Formula("=A1*2")
    )
    val graph = DependencyGraph.fromSheet(sheet)
    val result = DependencyGraph.transitiveDependents(graph, Set(ref"A1"))
    assertEquals(result, Set(ref"B1"))
  }

  test("transitiveDependents: chain of dependents") {
    // A1 <- B1 <- C1 (C1 depends on B1, B1 depends on A1)
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"B1" -> CellValue.Formula("=A1*2"),
      ref"C1" -> CellValue.Formula("=B1+5")
    )
    val graph = DependencyGraph.fromSheet(sheet)
    val result = DependencyGraph.transitiveDependents(graph, Set(ref"A1"))
    // When A1 changes, both B1 and C1 need recalculation
    assertEquals(result, Set(ref"B1", ref"C1"))
  }

  test("transitiveDependents: diamond pattern") {
    // A1 <- B1, A1 <- C1, B1 <- D1, C1 <- D1
    // D1 depends on both B1 and C1, both depend on A1
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"B1" -> CellValue.Formula("=A1+5"),
      ref"C1" -> CellValue.Formula("=A1*2"),
      ref"D1" -> CellValue.Formula("=B1+C1")
    )
    val graph = DependencyGraph.fromSheet(sheet)
    val result = DependencyGraph.transitiveDependents(graph, Set(ref"A1"))
    // When A1 changes, B1, C1, and D1 all need recalculation
    assertEquals(result, Set(ref"B1", ref"C1", ref"D1"))
  }

  test("transitiveDependents: multiple starting refs") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"A2" -> CellValue.Number(BigDecimal(20)),
      ref"B1" -> CellValue.Formula("=A1*2"),
      ref"B2" -> CellValue.Formula("=A2*3")
    )
    val graph = DependencyGraph.fromSheet(sheet)
    val result = DependencyGraph.transitiveDependents(graph, Set(ref"A1", ref"A2"))
    assertEquals(result, Set(ref"B1", ref"B2"))
  }

  // ===== Sheet recalculateDependents Tests (6 tests) =====

  test("recalculateDependents: updates direct dependent cache") {
    // A1=10, B1="=A1*2" (cached: 20)
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"B1" -> CellValue.Formula("A1*2", Some(CellValue.Number(BigDecimal(20))))
    )

    // Change A1 to 50 and recalculate
    val updatedSheet = sheet.put(ref"A1", CellValue.Number(BigDecimal(50)))
    val recalculated = updatedSheet.recalculateDependents(Set(ref"A1"))

    // B1 should now have cached value of 100
    val b1Cell = recalculated.cells.get(ref"B1")
    assert(b1Cell.isDefined)
    b1Cell.get.value match
      case CellValue.Formula(_, Some(CellValue.Number(n))) =>
        assertEquals(n, BigDecimal(100))
      case other =>
        fail(s"Expected Formula with cached Number(100), got $other")
  }

  test("recalculateDependents: updates transitive chain in correct order") {
    // A1=10, B1="=A1*2", C1="=B1+5"
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"B1" -> CellValue.Formula("A1*2", Some(CellValue.Number(BigDecimal(20)))),
      ref"C1" -> CellValue.Formula("B1+5", Some(CellValue.Number(BigDecimal(25))))
    )

    // Change A1 to 50 and recalculate
    val updatedSheet = sheet.put(ref"A1", CellValue.Number(BigDecimal(50)))
    val recalculated = updatedSheet.recalculateDependents(Set(ref"A1"))

    // B1 should be 100 (50 * 2)
    val b1Value = recalculated.cells.get(ref"B1").flatMap(_.value match
      case CellValue.Formula(_, Some(CellValue.Number(n))) => Some(n)
      case _ => None
    )
    assertEquals(b1Value, Some(BigDecimal(100)))

    // C1 should be 105 (100 + 5)
    val c1Value = recalculated.cells.get(ref"C1").flatMap(_.value match
      case CellValue.Formula(_, Some(CellValue.Number(n))) => Some(n)
      case _ => None
    )
    assertEquals(c1Value, Some(BigDecimal(105)))
  }

  test("recalculateDependents: preserves unrelated formula caches") {
    // A1=10, B1="=A1*2", C1="=100" (unrelated)
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"B1" -> CellValue.Formula("A1*2", Some(CellValue.Number(BigDecimal(20)))),
      ref"C1" -> CellValue.Formula("100", Some(CellValue.Number(BigDecimal(100))))
    )

    // Change A1 to 50 and recalculate
    val updatedSheet = sheet.put(ref"A1", CellValue.Number(BigDecimal(50)))
    val recalculated = updatedSheet.recalculateDependents(Set(ref"A1"))

    // C1 should be untouched (still 100)
    val c1Value = recalculated.cells.get(ref"C1").flatMap(_.value match
      case CellValue.Formula(expr, cached) => Some((expr, cached))
      case _ => None
    )
    assertEquals(c1Value.map(_._1), Some("100"))
    assertEquals(
      c1Value.flatMap(_._2),
      Some(CellValue.Number(BigDecimal(100)))
    )
  }

  test("recalculateDependents: handles evaluation errors gracefully") {
    // A1="=1/0" will cause div by zero
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(0)),
      ref"B1" -> CellValue.Formula("1/A1", Some(CellValue.Number(BigDecimal(1))))
    )

    // Change A1 to 0 (will cause div by zero) and recalculate
    val recalculated = sheet.recalculateDependents(Set(ref"A1"))

    // B1 should have its cache cleared (None) due to error
    val b1Value = recalculated.cells.get(ref"B1").map(_.value)
    b1Value match
      case Some(CellValue.Formula(_, None)) => () // Expected: cache cleared
      case Some(CellValue.Formula(_, Some(CellValue.Error(_)))) => () // Also acceptable
      case other => fail(s"Expected Formula with cleared cache or error, got $other")
  }

  test("recalculateDependents: handles empty modifiedRefs") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"B1" -> CellValue.Formula("A1*2", Some(CellValue.Number(BigDecimal(20))))
    )

    // Empty refs should return sheet unchanged
    val result = sheet.recalculateDependents(Set.empty)
    assertEquals(result.cells, sheet.cells)
  }


  // ===== Workbook Cross-Sheet Tests (3 tests) =====

  test("workbook recalculateDependents: updates dependent on same sheet") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"B1" -> CellValue.Formula("A1*2", Some(CellValue.Number(BigDecimal(20))))
    )
    val wb = Workbook(sheet)

    val updatedSheet = sheet.put(ref"A1", CellValue.Number(BigDecimal(50)))
    val updatedWb = wb.put(updatedSheet).recalculateDependents(sheet.name, Set(ref"A1"))

    val resultSheet = updatedWb(sheet.name).toOption.get
    val b1Value = resultSheet.cells.get(ref"B1").flatMap(_.value match
      case CellValue.Formula(_, Some(CellValue.Number(n))) => Some(n)
      case _ => None
    )
    assertEquals(b1Value, Some(BigDecimal(100)))
  }

  test("workbook recalculateDependents: cross-sheet reference") {
    val sheet1 = new Sheet(name = SheetName.unsafe("Sheet1"))
      .put(ref"A1", CellValue.Number(BigDecimal(10)))

    val sheet2 = new Sheet(name = SheetName.unsafe("Sheet2"))
      .put(ref"A1", CellValue.Formula("Sheet1!A1*2", Some(CellValue.Number(BigDecimal(20)))))

    val wb = Workbook(sheet1, sheet2)

    // Change Sheet1!A1 to 50
    val updatedSheet1 = sheet1.put(ref"A1", CellValue.Number(BigDecimal(50)))
    val updatedWb = wb.put(updatedSheet1).recalculateDependents(sheet1.name, Set(ref"A1"))

    // Sheet2!A1 should be recalculated to 100
    val resultSheet2 = updatedWb(SheetName.unsafe("Sheet2")).toOption.get
    val a1Value = resultSheet2.cells.get(ref"A1").flatMap(_.value match
      case CellValue.Formula(_, Some(CellValue.Number(n))) => Some(n)
      case _ => None
    )
    assertEquals(a1Value, Some(BigDecimal(100)))
  }

  test("workbook recalculateDependents: empty refs returns unchanged") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(10))
    )
    val wb = Workbook(sheet)

    val result = wb.recalculateDependents(sheet.name, Set.empty)
    assertEquals(result.sheets.map(_.name), wb.sheets.map(_.name))
  }

  // ===== GH-274: INDIRECT cells are always-dirty in targeted recalculation =====

  test("GH-274: editing the dynamic target refreshes INDIRECT (no static edge exists)") {
    val sheet = sheetWith(
      ref"C1" -> CellValue.Number(BigDecimal(5)),
      ref"B1" -> CellValue.Formula("=INDIRECT(\"C1\")", Some(CellValue.Number(BigDecimal(5))))
    )
    val updated = sheet
      .put(ref"C1", CellValue.Number(BigDecimal(50)))
      .recalculateDependents(Set(ref"C1"))
    assertEquals(
      updated(ref"B1").value,
      CellValue.Formula("=INDIRECT(\"C1\")", Some(CellValue.Number(BigDecimal(50))))
    )
  }

  test("GH-274: editing the text-source cell refreshes INDIRECT via the static edge") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Text("C1"),
      ref"C1" -> CellValue.Number(BigDecimal(5)),
      ref"C2" -> CellValue.Number(BigDecimal(7)),
      ref"B1" -> CellValue.Formula("=INDIRECT(A1)", Some(CellValue.Number(BigDecimal(5))))
    )
    val updated = sheet
      .put(ref"A1", CellValue.Text("C2"))
      .recalculateDependents(Set(ref"A1"))
    assertEquals(
      updated(ref"B1").value,
      CellValue.Formula("=INDIRECT(A1)", Some(CellValue.Number(BigDecimal(7))))
    )
  }

  test("GH-274: dependents of an INDIRECT cell refresh when the dynamic target changes") {
    val sheet = sheetWith(
      ref"C1" -> CellValue.Number(BigDecimal(5)),
      ref"B1" -> CellValue.Formula("=INDIRECT(\"C1\")", Some(CellValue.Number(BigDecimal(5)))),
      ref"D1" -> CellValue.Formula("=B1+1", Some(CellValue.Number(BigDecimal(6))))
    )
    val updated = sheet
      .put(ref"C1", CellValue.Number(BigDecimal(50)))
      .recalculateDependents(Set(ref"C1"))
    assertEquals(
      updated(ref"D1").value,
      CellValue.Formula("=B1+1", Some(CellValue.Number(BigDecimal(51))))
    )
  }

  test("GH-274: workbook recalculateDependents treats INDIRECT cells as always dirty") {
    val s1 = (new Sheet(name = SheetName.unsafe("S1"))).put(ref"A1", CellValue.Number(BigDecimal(5)))
    val s2 = (new Sheet(name = SheetName.unsafe("S2")))
      .put(ref"B1", CellValue.Formula("=INDIRECT(\"S1!A1\")", Some(CellValue.Number(BigDecimal(5)))))
    val wb0 = Workbook(s1, s2)
    val wb1 = wb0
      .put(s1.put(ref"A1", CellValue.Number(BigDecimal(9))))
      .recalculateDependents(SheetName.unsafe("S1"), Set(ref"A1"))
    val out = wb1.sheets.find(_.name.value == "S2").map(_(ref"B1").value)
    assertEquals(
      out,
      Some(CellValue.Formula("=INDIRECT(\"S1!A1\")", Some(CellValue.Number(BigDecimal(9)))))
    )
  }

  // ===== GH-301: OFFSET cells are always-dirty in targeted recalculation (INDIRECT parity) =====

  test("GH-301: editing the dynamic target refreshes OFFSET (no static edge exists)") {
    // B1=OFFSET(C1,1,0) reads C2; the static graph only has the C1 anchor edge, so an
    // edit to C2 is invisible without the always-dirty marking.
    val sheet = sheetWith(
      ref"C1" -> CellValue.Number(BigDecimal(1)),
      ref"C2" -> CellValue.Number(BigDecimal(5)),
      ref"B1" -> CellValue.Formula("=OFFSET(C1,1,0)", Some(CellValue.Number(BigDecimal(5))))
    )
    val updated = sheet
      .put(ref"C2", CellValue.Number(BigDecimal(50)))
      .recalculateDependents(Set(ref"C2"))
    assertEquals(
      updated(ref"B1").value,
      CellValue.Formula("=OFFSET(C1,1,0)", Some(CellValue.Number(BigDecimal(50))))
    )
  }

  test("GH-301: dependents of an OFFSET cell refresh when the dynamic target changes") {
    val sheet = sheetWith(
      ref"C1" -> CellValue.Number(BigDecimal(1)),
      ref"C2" -> CellValue.Number(BigDecimal(5)),
      ref"B1" -> CellValue.Formula("=OFFSET(C1,1,0)", Some(CellValue.Number(BigDecimal(5)))),
      ref"D1" -> CellValue.Formula("=B1+1", Some(CellValue.Number(BigDecimal(6))))
    )
    val updated = sheet
      .put(ref"C2", CellValue.Number(BigDecimal(50)))
      .recalculateDependents(Set(ref"C2"))
    assertEquals(
      updated(ref"D1").value,
      CellValue.Formula("=B1+1", Some(CellValue.Number(BigDecimal(51))))
    )
  }

  test("GH-301: SUM(OFFSET(...)) window refreshes when a window cell changes") {
    // SUM over OFFSET(C1,0,0,3,1) = C1:C3; editing C3 must refresh despite no static edge.
    val sheet = sheetWith(
      ref"C1" -> CellValue.Number(BigDecimal(1)),
      ref"C2" -> CellValue.Number(BigDecimal(2)),
      ref"C3" -> CellValue.Number(BigDecimal(3)),
      ref"B1" -> CellValue.Formula("=SUM(OFFSET(C1,0,0,3,1))", Some(CellValue.Number(BigDecimal(6))))
    )
    val updated = sheet
      .put(ref"C3", CellValue.Number(BigDecimal(30)))
      .recalculateDependents(Set(ref"C3"))
    assertEquals(
      updated(ref"B1").value,
      CellValue.Formula("=SUM(OFFSET(C1,0,0,3,1))", Some(CellValue.Number(BigDecimal(33))))
    )
  }
