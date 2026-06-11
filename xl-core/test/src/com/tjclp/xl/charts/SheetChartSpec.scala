package com.tjclp.xl.charts

import munit.FunSuite

import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.drawings.{Drawing, DrawingAnchor, EditAs, Extent}
import com.tjclp.xl.macros.ref
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.units.Emu
import com.tjclp.xl.workbooks.Workbook

/**
 * GH-222 xl-core chart layer: DataRef parse/print, Chart.validated, Sheet.addChart/charts,
 * structural shifts (anchor remap + same-sheet reference shifts + deletion semantics), and
 * Workbook.rename remapping.
 */
class SheetChartSpec extends FunSuite:

  private val data = SheetName.unsafe("Data")
  private def dref(range: CellRange, sheet: SheetName = data): DataRef = DataRef(sheet, range)

  private val series = Series(
    values = dref(ref"B2:B5"),
    categories = Some(dref(ref"A2:A5")),
    name = Some(SeriesName.FromCell(data, ref"B1"))
  )
  private val chart = Chart(ChartType.Bar(), Vector(series), Some("T"), Some(Legend()))
  private val sheet = Sheet(data)

  // ===== DataRef =====

  test("DataRef.toFormula prints absolute refs with GH-263 quoting") {
    assertEquals(dref(ref"B2:B5").toFormula, "Data!$B$2:$B$5")
    assertEquals(
      DataRef(SheetName.unsafe("Q1 'Report"), ref"A1:A3").toFormula,
      "'Q1 ''Report'!$A$1:$A$3"
    )
    // cell-shaped sheet names must quote (bare Q1!B2 reads as something else entirely)
    assertEquals(DataRef(SheetName.unsafe("Q1"), ref"A1:A3").toFormula, "'Q1'!$A$1:$A$3")
    // single-cell ranges collapse to one absolute ref (Excel's own shape)
    assertEquals(DataRef(data, CellRange(ref"B2", ref"B2")).toFormula, "Data!$B$2")
  }

  test("DataRef.parse strips $ from the ref part only and round-trips toFormula") {
    assertEquals(
      DataRef.parse("'ChartData'!$A$2:$A$5"),
      Some(dref(ref"A2:A5", SheetName.unsafe("ChartData")))
    )
    assertEquals(DataRef.parse("Data!B2:B5"), Some(dref(ref"B2:B5")))
    assertEquals(
      DataRef.parse("'ChartData'!B1"),
      Some(DataRef(SheetName.unsafe("ChartData"), CellRange(ref"B1", ref"B1")))
    )
    // sheet names may legally contain $ — split happens BEFORE the strip
    assertEquals(
      DataRef.parse("'My$Sheet'!$A$1:$A$2"),
      Some(DataRef(SheetName.unsafe("My$Sheet"), ref"A1:A2"))
    )
    val roundTrips = Vector(dref(ref"B2:B5"), DataRef(SheetName.unsafe("Q1 'R"), ref"A1:C1"))
    roundTrips.foreach(r => assertEquals(DataRef.parse(r.toFormula), Some(r)))
  }

  test("DataRef.parse rejects unqualified, union, and external-workbook refs") {
    assertEquals(DataRef.parse("B2:B5"), None)
    assertEquals(DataRef.parse("Data!A1:A3,Data!C1:C3"), None)
    assertEquals(DataRef.parse("[Book1]Data!A1:A3"), None)
    assertEquals(DataRef.parse(""), None)
    assertEquals(DataRef.parse("Data!"), None)
  }

  // ===== Chart.validated =====

  test("Chart.validated enforces non-empty series, pie cardinality, and vector ranges") {
    assert(Chart.validated(ChartType.Line, Vector.empty).isLeft)
    assert(Chart.validated(ChartType.Pie, Vector(series, series)).isLeft)
    assert(Chart.pie(series).isRight)
    val twoD = series.copy(values = dref(ref"A1:B5"))
    assert(Chart.validated(ChartType.Line, Vector(twoD)).isLeft)
    val twoDcats = series.copy(categories = Some(dref(ref"A1:B5")))
    assert(Chart.validated(ChartType.Line, Vector(twoDcats)).isLeft)
    assert(Chart.bar(Vector(series), BarDirection.Bar, BarGrouping.Stacked).isRight)
  }

  // ===== Sheet.addChart / charts =====

  test("addChart appends in z-order; charts collects only ChartFrames") {
    val s = sheet
      .addChart(chart, DrawingAnchor.at(ref"D2", Extent(Emu(5400000L), Emu(2700000L))))
      .addChart(chart, ref"F2:K15")
    assertEquals(s.drawings.size, 2)
    assertEquals(s.charts.size, 2)
    assertEquals(s.pictures.size, 0)
    s.drawings(1) match
      case Drawing.ChartFrame(DrawingAnchor.TwoCell(from, to, EditAs.TwoCell), c, name) =>
        assertEquals(from.cell, ref"F2")
        assertEquals(to.cell, ref"L16") // one past the range end
        assertEquals(c, chart)
        assertEquals(name, "")
      case other => fail(s"expected TwoCell ChartFrame, got $other")
  }

  // ===== structural shifts =====

  private def chartOf(s: Sheet): Chart =
    s.charts.headOption.fold(fail("no chart on sheet"))(_.chart)

  test("insertRows above the data shifts values/categories/name refs down") {
    val s = sheet.addChart(chart, DrawingAnchor.at(ref"D2", Extent(Emu(1), Emu(1))))
    val shifted = s.insertRows(0, 2)
    val c = chartOf(shifted)
    assertEquals(c.series(0).values, dref(ref"B4:B7"))
    assertEquals(c.series(0).categories, Some(dref(ref"A4:A7")))
    assertEquals(c.series(0).name, Some(SeriesName.FromCell(data, ref"B3")))
    // anchor remapped too (row 1 -> row 3, 0-based)
    shifted.drawings(0) match
      case Drawing.ChartFrame(DrawingAnchor.OneCell(from, _), _, _) =>
        assertEquals(from.cell, ref"D4")
      case other => fail(s"unexpected $other")
  }

  test("insertRows within the data grows the range (endpoints shift independently)") {
    val s = sheet.addChart(chart, DrawingAnchor.at(ref"D2", Extent(Emu(1), Emu(1))))
    val c = chartOf(s.insertRows(2, 1)) // 0-based row 2 = inside B2:B5
    assertEquals(c.series(0).values, dref(ref"B2:B6"))
    assertEquals(c.series(0).name, Some(SeriesName.FromCell(data, ref"B1"))) // above, untouched
  }

  test("deleteRows partially overlapping clamps to the surviving span") {
    val s = sheet.addChart(chart, DrawingAnchor.at(ref"D2", Extent(Emu(1), Emu(1))))
    val c = chartOf(s.deleteRows(3, 5)) // deletes rows 4..8 (1-based), B2:B5 -> B2:B3
    assertEquals(c.series(0).values, dref(ref"B2:B3"))
    assertEquals(c.series(0).categories, Some(dref(ref"A2:A3")))
  }

  test("deleteRows covering the values drops the series; covering categories/name clears them") {
    val s = sheet.addChart(chart, DrawingAnchor.at(ref"D2", Extent(Emu(1), Emu(1))))
    // delete rows 2..5 (1-based): values fully deleted -> series dropped entirely
    assertEquals(chartOf(s.deleteRows(1, 4)).series, Vector.empty)
    // delete only row 1 (the name cell): name -> None, ranges shift up
    val c = chartOf(s.deleteRows(0, 1))
    assertEquals(c.series(0).name, None)
    assertEquals(c.series(0).values, dref(ref"B1:B4"))
    // categories-only deletion: shrink data to make cats deletable independently is not possible
    // on the row axis here (cats and values share rows), so exercise the column axis instead:
    val byCol = chartOf(s.deleteColumns(0, 1)) // drop column A -> categories gone, values shift
    assertEquals(byCol.series(0).categories, None)
    assertEquals(byCol.series(0).values, dref(ref"A2:A5"))
    assertEquals(byCol.series(0).name, Some(SeriesName.FromCell(data, ref"A1")))
  }

  test("core-only structural edit leaves cross-sheet chart refs untouched (documented)") {
    val other = SheetName.unsafe("Other")
    val crossSeries = series.copy(
      values = dref(ref"B2:B5", other),
      categories = Some(dref(ref"A2:A5", other)),
      name = Some(SeriesName.FromCell(other, ref"B1"))
    )
    val s = sheet.addChart(
      Chart(ChartType.Line, Vector(crossSeries)),
      DrawingAnchor.at(ref"D2", Extent(Emu(1), Emu(1)))
    )
    val c = chartOf(s.insertRows(0, 3))
    assertEquals(c.series(0), crossSeries) // refs point at Other — this sheet's edit can't see it
  }

  test("shiftChartRefs shifts only refs matching the edited sheet (case-insensitive)") {
    val host = Sheet(SheetName.unsafe("Host")).addChart(
      chart, // refs point at Data
      DrawingAnchor.at(ref"A1", Extent(Emu(1), Emu(1)))
    )
    val shifted = host.shiftChartRefs("DATA", isRow = true, at = 0, delta = 2)
    assertEquals(chartOf(shifted).series(0).values, dref(ref"B4:B7"))
    // anchors on the host sheet are untouched (its geometry did not change)
    shifted.drawings(0) match
      case Drawing.ChartFrame(DrawingAnchor.OneCell(from, _), _, _) =>
        assertEquals(from.cell, ref"A1")
      case other => fail(s"unexpected $other")
    val noMatch = host.shiftChartRefs("Unrelated", isRow = true, at = 0, delta = 2)
    assertEquals(chartOf(noMatch), chart)
    // delete fully covering the values drops the series
    val dropped = host.shiftChartRefs("Data", isRow = true, at = 0, delta = -9)
    assertEquals(chartOf(dropped).series, Vector.empty)
  }

  // ===== rename remap =====

  test("Workbook.rename remaps DataRef/FromCell sheet names across all typed charts") {
    val host = Sheet(SheetName.unsafe("Host"))
      .addChart(chart, DrawingAnchor.at(ref"A1", Extent(Emu(1), Emu(1))))
    val wb = Workbook(Vector(sheet, host))
    val renamed = wb
      .rename(data, SheetName.unsafe("Numbers"))
      .fold(e => fail(s"rename failed: $e"), identity)
    val c = chartOf(renamed.sheets(1))
    val numbers = SheetName.unsafe("Numbers")
    assertEquals(c.series(0).values, DataRef(numbers, ref"B2:B5"))
    assertEquals(c.series(0).categories, Some(DataRef(numbers, ref"A2:A5")))
    assertEquals(c.series(0).name, Some(SeriesName.FromCell(numbers, ref"B1")))
  }
