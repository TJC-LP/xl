package com.tjclp.xl.formula

import java.time.LocalDate

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, Column, Row, SheetName}
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.formula.eval.{SheetSummary, WorkbookInspect, WorkbookSummary}
import com.tjclp.xl.ooxml.{TestFixtures, XlsxReader}
import com.tjclp.xl.sheets.{
  AutoFilterState,
  ColumnProperties,
  DataValidation,
  FreezePane,
  RowProperties,
  Sheet
}
import com.tjclp.xl.workbooks.{CalcPr, DefinedName, Workbook}
import munit.FunSuite

/**
 * ADR-017 §2.10: `WorkbookSummary.of(wb)` is the orientation card — one `SheetSummary` per sheet
 * with the counts an agent needs before reading a cell, plus the workbook facts (`definedNames`,
 * `date1904`, `calcPr`). It is total: every workbook the reader accepts summarizes.
 */
class WorkbookSummarySpec extends FunSuite:

  private def a1(s: String): ARef = ARef.parse(s).fold(err => fail(err), identity)
  private def range(s: String): CellRange = CellRange.parse(s).fold(err => fail(err), identity)

  /**
   * The golden corpus's `simple.xlsx` shape: nine cells, one recalculated formula, an empty sheet.
   */
  private def simpleBook(): Workbook =
    val data = Sheet(SheetName.unsafe("Data"))
      .put(a1("A1"), "Hello")
      .put(a1("B1"), 10)
      .put(a1("C1"), LocalDate.of(2024, 1, 15))
      .put(a1("A2"), "World")
      .put(a1("B2"), 20)
      .put(a1("C2"), LocalDate.of(2024, 6, 30))
      .put(a1("A3"), "Total")
      .put(a1("B3"), BigDecimal("12.5"))
      .put(a1("B4"), CellValue.Formula("SUM(B1:B3)", None))
    Workbook(Vector(data, Sheet(SheetName.unsafe("Summary"))))

  test("counts on the simple book: cells, formulas, uncached, dimension, empty sheet") {
    val before = WorkbookSummary.of(simpleBook())
    assertEquals(before.sheets.map(_.name.value), Vector("Data", "Summary"))
    val data = before.sheets.headOption.getOrElse(fail("no Data sheet"))
    assertEquals(data.index, 1)
    assertEquals(data.state, None)
    assertEquals(data.dimension, Some(range("A1:C4")))
    assertEquals(data.cellCount, 9)
    assertEquals(data.formulaCount, 1)
    assertEquals(data.uncachedFormulas, 1)
    val summary = before.sheets.lift(1).getOrElse(fail("no Summary sheet"))
    assertEquals(summary.index, 2)
    assertEquals(summary.dimension, None)
    assertEquals(summary.cellCount, 0)
    assertEquals(summary.formulaCount, 0)
    assertEquals(before.definedNames, Vector.empty)
    assertEquals(before.date1904, false)
    assertEquals(before.calcPr, None)

    val after = WorkbookSummary.of(simpleBook().recalculate().workbook)
    assertEquals(after.sheets.headOption.map(_.uncachedFormulas), Some(0))
    assertEquals(after.sheets.headOption.map(_.formulaCount), Some(1))
  }

  test(
    "every sheet facet is counted: merges, comments, hyperlinks, freeze, tab color, filter, CF, DV, hidden rows/cols, tables"
  ) {
    val sheet = Sheet(SheetName.unsafe("Facets"))
      .put(a1("A1"), "Title")
      .put(Cell(a1("B1"), CellValue.Text("link")).withHyperlink("https://example.com"))
      .put(a1("A2"), 1)
      .put(a1("A3"), 2)
      .merge(range("A1:C1"))
      .merge(range("D1:E2"))
      .comment(a1("A1"), Comment.plainText("heading", Some("qa")))
      .freezeAt(a1("B2"))
      .withTabColor(Color.Rgb(0xff00ff00))
      .conditionalFormat(range("A2:A3"), CfRule.expression("A2>1", Dxf.fill(Color.Rgb(0xffffcc00))))
      .withDataValidation(range("A2:A3"), DataValidation.list("\"1,2,3\""))
      .setRowProperties(Row.from1(5), RowProperties(hidden = true))
      .setRowProperties(Row.from1(6), RowProperties(height = Some(30.0)))
      .setColumnProperties(Column.from0(3), ColumnProperties(hidden = true))
      .setColumnProperties(Column.from0(4), ColumnProperties(hidden = true))
      .copy(autoFilter = Some(AutoFilterState.Ranged(range("A1:C3"))))
    val wb = Workbook(Vector(sheet))
      .withDefinedName("Total", "Facets!$A$3")
      .withCalcPr(CalcPr(iterativeCalculation = true, maxIterations = Some(50)))
    val summary = WorkbookSummary.of(wb)
    val s = summary.sheets.headOption.getOrElse(fail("no sheet"))
    assertEquals(s.cellCount, 4)
    assertEquals(s.mergedRanges, 2)
    assertEquals(s.comments, 1)
    assertEquals(s.hyperlinks, 1)
    assertEquals(s.freeze, Some(FreezePane.At(a1("B2"))))
    assertEquals(s.tabColor, Some(Color.Rgb(0xff00ff00)))
    assertEquals(s.autoFilter, Some(range("A1:C3")))
    assertEquals(s.conditionalFormats, 1)
    assertEquals(s.dataValidations, 1)
    assertEquals(s.hiddenRows, 1)
    assertEquals(s.hiddenCols, 2)
    assertEquals(s.tables, Vector.empty)
    assertEquals(s.charts, 0)
    assertEquals(s.pictures, 0)
    assertEquals(summary.definedNames.map(_.name), Vector("Total"))
    assertEquals(summary.calcPr.map(_.maxIterations), Some(Some(50)))
  }

  test("sheet state, hidden defined names and date1904 pass through") {
    val wb0 = Workbook(
      Vector(
        Sheet(SheetName.unsafe("Visible")).put(a1("A1"), 1),
        Sheet(SheetName.unsafe("Secret")).put(a1("A1"), 2)
      )
    )
    val hidden = wb0
      .setSheetState(SheetName.unsafe("Secret"), Some("hidden"))
      .fold(e => fail(e.message), identity)
    val wb = hidden.copy(metadata =
      hidden.metadata.copy(
        definedNames = Vector(
          DefinedName("Shown", "Visible!$A$1"),
          DefinedName("Hid", "Secret!$A$1", hidden = true)
        ),
        date1904 = true
      )
    )
    val summary = WorkbookSummary.of(wb)
    assertEquals(summary.sheets.map(_.state), Vector(None, Some("hidden")))
    assertEquals(
      summary.definedNames.map(n => (n.name, n.hidden)),
      Vector(("Shown", false), ("Hid", true))
    )
    assertEquals(summary.date1904, true)
  }

  test("wb.describe and wb.audit resolve through WorkbookInspect's extension block") {
    import WorkbookInspect.*
    val wb = simpleBook().recalculate().workbook
    val summary: WorkbookSummary = wb.describe
    assertEquals(summary, WorkbookSummary.of(wb))
    assert(wb.audit.isClean)
  }

  test("the empty workbook summarizes to no sheets") {
    assertEquals(
      WorkbookSummary.of(Workbook(Vector.empty)),
      WorkbookSummary(Vector.empty, Vector.empty, date1904 = false, calcPr = None)
    )
  }

  test(
    "corpus: every readable fixture summarizes, with charts and pictures where the file has them"
  ) {
    val summaries: List[(String, WorkbookSummary)] = TestFixtures.all.flatMap { name =>
      XlsxReader
        .read(TestFixtures.copyToTemp(name))
        .toOption
        .map(wb => name -> WorkbookSummary.of(wb))
    }
    assert(summaries.size >= 3, s"expected at least three readable fixtures, got ${summaries.size}")
    summaries.foreach { (name, summary) =>
      assertEquals(summary.sheets.map(_.index), summary.sheets.indices.map(_ + 1).toVector, name)
      summary.sheets.foreach { s =>
        assert(s.formulaCount <= s.cellCount, s"$name/${s.name.value}: formulas exceed cells")
        assert(
          s.uncachedFormulas <= s.formulaCount,
          s"$name/${s.name.value}: uncached exceed formulas"
        )
      }
    }
    def facet(name: String)(f: SheetSummary => Int): Int =
      summaries
        .find(_._1 == name)
        .map(_._2.sheets.map(f).sum)
        .getOrElse(fail(s"$name did not read"))
    assert(facet("chart-bar.xlsx")(_.charts) >= 1, "chart-bar.xlsx carries a chart")
    assert(facet("image.xlsx")(_.pictures) >= 1, "image.xlsx carries a picture")
    assert(
      facet("comments-hyperlinks.xlsx")(_.comments) >= 1,
      "comments-hyperlinks.xlsx has comments"
    )
    assert(
      facet("comments-hyperlinks.xlsx")(_.hyperlinks) >= 1,
      "comments-hyperlinks.xlsx has hyperlinks"
    )
    assert(facet("formulas.xlsx")(_.formulaCount) >= 1, "formulas.xlsx has formulas")
    assert(
      facet("autofilter.xlsx")(s => s.autoFilter.fold(0)(_ => 1)) >= 1,
      "autofilter.xlsx has an autoFilter"
    )
    assert(facet("condformat.xlsx")(_.conditionalFormats) >= 1, "condformat.xlsx has CF rules")
  }
