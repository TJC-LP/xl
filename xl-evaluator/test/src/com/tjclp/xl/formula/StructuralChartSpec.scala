package com.tjclp.xl.formula

import munit.FunSuite

import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.charts.{Chart, ChartType, DataRef, Series, SeriesName}
import com.tjclp.xl.drawings.{Drawing, DrawingAnchor, Extent}
import com.tjclp.xl.formula.eval.StructuralEditor
import com.tjclp.xl.macros.ref
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.units.Emu
import com.tjclp.xl.workbooks.Workbook

/**
 * GH-222: StructuralEditor shifts typed-chart references ACROSS sheets (a chart on Dash tracking
 * data on Data follows Data's structural edits) without double-shifting the edited sheet's own
 * charts.
 */
class StructuralChartSpec extends FunSuite:

  private val data = SheetName.unsafe("Data")
  private val dash = SheetName.unsafe("Dash")

  private val series = Series(
    values = DataRef(data, ref"B2:B5"),
    categories = Some(DataRef(data, ref"A2:A5")),
    name = Some(SeriesName.FromCell(data, ref"B1"))
  )
  private val chart = Chart(ChartType.Bar(), Vector(series))
  private val anchor = DrawingAnchor.at(ref"D2", Extent(Emu(1), Emu(1)))

  private def chartOf(sheet: Sheet): Chart =
    sheet.charts.headOption.fold(fail(s"no chart on ${sheet.name.value}"))(_.chart)

  test("cross-sheet chart refs shift when the data sheet gains rows") {
    val wb = Workbook(Vector(Sheet(data), Sheet(dash).addChart(chart, anchor)))
    val edited = StructuralEditor.insertRows(wb, data, at = 0, count = 2)
    val c = chartOf(edited.sheets(1))
    assertEquals(c.series(0).values, DataRef(data, ref"B4:B7"))
    assertEquals(c.series(0).categories, Some(DataRef(data, ref"A4:A7")))
    assertEquals(c.series(0).name, Some(SeriesName.FromCell(data, ref"B3")))
  }

  test("cross-sheet delete fully covering the values drops the series") {
    val wb = Workbook(Vector(Sheet(data), Sheet(dash).addChart(chart, anchor)))
    val edited = StructuralEditor.deleteRows(wb, data, at = 0, count = 10)
    assertEquals(chartOf(edited.sheets(1)).series, Vector.empty)
  }

  test("edited sheet's own charts shift exactly once (no double-shift)") {
    val selfHosted = Sheet(data).addChart(chart, anchor)
    val wb = Workbook(Vector(selfHosted, Sheet(dash).addChart(chart, anchor)))
    val edited = StructuralEditor.insertRows(wb, data, at = 0, count = 3)
    // both the self-hosted and the cross-sheet chart land on the SAME shifted refs
    assertEquals(chartOf(edited.sheets(0)).series(0).values, DataRef(data, ref"B5:B8"))
    assertEquals(chartOf(edited.sheets(1)).series(0).values, DataRef(data, ref"B5:B8"))
    // the self-hosted anchor moved with its sheet; the cross-sheet anchor did not
    edited.sheets(0).charts(0).anchor match
      case DrawingAnchor.OneCell(from, _) => assertEquals(from.cell, ref"D5")
      case other => fail(s"unexpected $other")
    edited.sheets(1).charts(0).anchor match
      case DrawingAnchor.OneCell(from, _) => assertEquals(from.cell, ref"D2")
      case other => fail(s"unexpected $other")
  }

  test("charts referencing an unrelated sheet are untouched by the edit") {
    val other = SheetName.unsafe("Other")
    val crossChart = Chart(
      ChartType.Line,
      Vector(Series(DataRef(other, ref"B2:B5"), None, None))
    )
    val wb = Workbook(Vector(Sheet(data), Sheet(dash).addChart(crossChart, anchor), Sheet(other)))
    val edited = StructuralEditor.insertColumns(wb, data, at = 0, count = 4)
    assertEquals(chartOf(edited.sheets(1)), crossChart)
  }
