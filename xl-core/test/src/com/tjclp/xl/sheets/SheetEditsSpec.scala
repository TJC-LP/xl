package com.tjclp.xl.sheets

import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row, SheetName}
import com.tjclp.xl.cells.{CellValue, Comment}
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.ops.{ClearWhat, ColSpan, Edit, FormulaSupport, OffGridRef, RowSpan}
import com.tjclp.xl.sheets.styleSyntax.{getCellStyle, withCellStyle}
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.numfmt.NumFmt
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * The sheet-local edit semantics that moved from the CLI into `Sheet` (ADR-017 §2.12, W2.1): fill,
 * copy, sort, clear, column auto-fit, outline grouping and the appearance/print merges, pinned at
 * the `Sheet` level so the CLI forwarders and the `Edit` interpreter share one meaning. Formula
 * shifting is exercised through a marking `FormulaSupport` (xl-core has no parser): the marker
 * records the displacement each moved formula was shifted by.
 */
class SheetEditsSpec extends FunSuite:

  private val name = SheetName.unsafe("S")
  private def a1(s: String): ARef = ARef.parse(s).fold(e => fail(e), identity)
  private def rng(s: String): CellRange = CellRange.parse(s).fold(e => fail(e), identity)
  private def num(n: Int): CellValue = CellValue.Number(BigDecimal(n))
  private def formula(text: String): CellValue = CellValue.Formula(text, None)
  private def ok[A](r: XLResult[A]): A = r.fold(e => fail(e.message), identity)

  private given FormulaSupport = FormulaSupport.textOnly

  /** Shifts by appending `|colDelta,rowDelta`, so a test can read the displacement back. */
  private val marking: FormulaSupport = new FormulaSupport:
    def validate(formula: String): XLResult[Unit] = FormulaSupport.textOnly.validate(formula)
    def shift(formula: String, colDelta: Int, rowDelta: Int): XLResult[String] =
      Right(s"$formula|$colDelta,$rowDelta")
    def insertRows(wb: Workbook, sheet: SheetName, at: Row, count: Int): XLResult[Workbook] =
      FormulaSupport.textOnly.insertRows(wb, sheet, at, count)
    def deleteRows(wb: Workbook, sheet: SheetName, at: Row, count: Int): XLResult[Workbook] =
      FormulaSupport.textOnly.deleteRows(wb, sheet, at, count)
    def insertCols(wb: Workbook, sheet: SheetName, at: Column, count: Int): XLResult[Workbook] =
      FormulaSupport.textOnly.insertCols(wb, sheet, at, count)
    def deleteCols(wb: Workbook, sheet: SheetName, at: Column, count: Int): XLResult[Workbook] =
      FormulaSupport.textOnly.deleteCols(wb, sheet, at, count)
    def renameSheet(wb: Workbook, from: SheetName, to: SheetName): XLResult[Workbook] =
      FormulaSupport.textOnly.renameSheet(wb, from, to)

  // ========== fill ==========

  test("fill down cycles the source rows through the target; fill right cycles the columns") {
    val s = Sheet(name).put(a1("A1"), num(1)).put(a1("A2"), num(2))
    val down = ok(s.fill(rng("A1:A2"), rng("A1:A6"), Edit.FillDir.Down))
    assertEquals((1 to 6).map(r => down(a1(s"A$r")).value), Vector(1, 2, 1, 2, 1, 2).map(num))
    val wide = Sheet(name).put(a1("A1"), num(1)).put(a1("B1"), num(2))
    val right = ok(wide.fill(rng("A1:B1"), rng("A1:E1"), Edit.FillDir.Right))
    assertEquals(
      Vector("A1", "B1", "C1", "D1", "E1").map(c => right(a1(c)).value),
      Vector(1, 2, 1, 2, 1).map(num)
    )
    // an empty source cell leaves its targets untouched
    val sparse = Sheet(name).put(a1("A1"), num(1)).put(a1("B3"), num(9))
    val filled = ok(sparse.fill(rng("A1:B1"), rng("A1:B3"), Edit.FillDir.Down))
    assertEquals(filled(a1("B3")).value, num(9))
    assert(!filled.contains(a1("B2")))
  }

  test(
    "fill shifts a formula by its displacement, refuses without support, keeps a record's value"
  ) {
    val s = Sheet(name).put(a1("A1"), formula("B1")).put(a1("A2"), formula("B2"))
    val down = ok(s.fill(rng("A1:A2"), rng("A1:A5"), Edit.FillDir.Down)(using marking))
    assertEquals(down(a1("A3")).value, formula("B1|0,2"))
    assertEquals(down(a1("A4")).value, formula("B2|0,2"))
    assertEquals(down(a1("A5")).value, formula("B1|0,4"))
    val right = ok(s.fill(rng("A1"), rng("A1:C1"), Edit.FillDir.Right)(using marking))
    assertEquals(right(a1("C1")).value, formula("B1|2,0"))
    s.fill(rng("A1"), rng("A1:A3"), Edit.FillDir.Down) match
      case Left(XLError.UnsupportedCapability("shift", _, _)) => ()
      case other => fail(s"expected the text-only refusal, got $other")
    assertEquals(
      s.fill(rng("A1"), rng("B1:B3"), Edit.FillDir.Down),
      Left(
        XLError.InvalidReference(
          "Fill down requires matching columns. Source: A1:A1, Target: B1:B3"
        )
      )
    )
    assertEquals(
      s.fill(rng("A1"), rng("B2:D2"), Edit.FillDir.Right),
      Left(
        XLError.InvalidReference("Fill right requires matching rows. Source: A1:A1, Target: B2:D2")
      )
    )
  }

  /**
   * GH-628: a support that "voids" every reference of a formula spelled `X…`: the shifted text is
   * `#REF!` and the report names the source formula, the way the evaluator's support reports the
   * references a drag carried off the grid.
   */
  private val voiding: FormulaSupport = new FormulaSupport:
    def validate(formula: String): XLResult[Unit] = FormulaSupport.textOnly.validate(formula)
    def shift(formula: String, colDelta: Int, rowDelta: Int): XLResult[String] =
      shiftReporting(formula, colDelta, rowDelta).map(_.formula)
    override def shiftReporting(
      formula: String,
      colDelta: Int,
      rowDelta: Int
    ): XLResult[FormulaSupport.Shifted] =
      if formula.startsWith("X") then Right(FormulaSupport.Shifted("#REF!", Vector(formula)))
      else Right(FormulaSupport.Shifted(s"$formula|$colDelta,$rowDelta", Vector.empty))
    def insertRows(wb: Workbook, sheet: SheetName, at: Row, count: Int): XLResult[Workbook] =
      FormulaSupport.textOnly.insertRows(wb, sheet, at, count)
    def deleteRows(wb: Workbook, sheet: SheetName, at: Row, count: Int): XLResult[Workbook] =
      FormulaSupport.textOnly.deleteRows(wb, sheet, at, count)
    def insertCols(wb: Workbook, sheet: SheetName, at: Column, count: Int): XLResult[Workbook] =
      FormulaSupport.textOnly.insertCols(wb, sheet, at, count)
    def deleteCols(wb: Workbook, sheet: SheetName, at: Column, count: Int): XLResult[Workbook] =
      FormulaSupport.textOnly.deleteCols(wb, sheet, at, count)
    def renameSheet(wb: Workbook, from: SheetName, to: SheetName): XLResult[Workbook] =
      FormulaSupport.textOnly.renameSheet(wb, from, to)

  test("GH-628: fillReporting names each target cell whose formula gained a #REF!, in order") {
    val s = Sheet(name).put(a1("A1"), formula("X1")).put(a1("B1"), formula("B9"))
    val (filled, offGrid) =
      ok(s.fillReporting(rng("A1:B1"), rng("A1:B3"), Edit.FillDir.Down)(using voiding))
    assertEquals(filled(a1("A2")).value, formula("#REF!"))
    assertEquals(filled(a1("A3")).value, formula("#REF!"))
    assertEquals(filled(a1("B3")).value, formula("B9|0,2"))
    assertEquals(
      offGrid,
      Vector(OffGridRef(a1("A2"), Vector("X1")), OffGridRef(a1("A3"), Vector("X1")))
    )
    // the plain fill is the same sheet without the report; a clean fill reports nothing
    assertEquals(ok(s.fill(rng("A1:B1"), rng("A1:B3"), Edit.FillDir.Down)(using voiding)), filled)
    val clean = Sheet(name).put(a1("A1"), formula("B9"))
    assertEquals(
      ok(clean.fillReporting(rng("A1"), rng("A1:A3"), Edit.FillDir.Down)(using voiding))._2,
      Vector.empty
    )
  }

  test("GH-628: copyRangeReporting reports across the target range and across sheets") {
    val s = Sheet(name).put(a1("A1"), formula("X1")).put(a1("A2"), num(4))
    val (copied, offGrid) =
      ok(s.copyRangeReporting(s, rng("A1:A2"), rng("C5:C6"), valuesOnly = false)(using voiding))
    assertEquals(copied(a1("C5")).value, formula("#REF!"))
    assertEquals(copied(a1("C6")).value, num(4))
    assertEquals(offGrid, Vector(OffGridRef(a1("C5"), Vector("X1"))))
    // values-only pastes never shift, so never report
    val (_, none) =
      ok(s.copyRangeReporting(s, rng("A1:A2"), rng("C5:C6"), valuesOnly = true)(using voiding))
    assertEquals(none, Vector.empty)
    val target = Sheet(SheetName.unsafe("T"))
    val (cross, crossReport) =
      ok(
        target.copyRangeReporting(s, rng("A1:A1"), rng("B2:B2"), valuesOnly = false)(using voiding)
      )
    assertEquals(cross.name, target.name)
    assertEquals(crossReport, Vector(OffGridRef(a1("B2"), Vector("X1"))))
  }

  // ========== copy ==========

  test("copyRange snapshots the source, so an overlapping copy reads the pre-copy state") {
    val s = Sheet(name).put(a1("A1"), num(1)).put(a1("A2"), num(2)).put(a1("A3"), num(3))
    val copied = ok(s.copyRange(rng("A1:A3"), rng("A2:A4"), valuesOnly = false))
    assertEquals(Vector("A2", "A3", "A4").map(r => copied(a1(r)).value), Vector(1, 2, 3).map(num))
    assertEquals(copied(a1("A1")).value, num(1))
  }

  test("copyRange moves styles, shifts formulas by the displacement, and pastes values only") {
    val currency = CellStyle.default.withNumFmt(NumFmt.Currency)
    val s = Sheet(name)
      .put(a1("A1"), formula("B1"))
      .put(a1("A2"), CellValue.Formula("B2", Some(num(6))))
      .withCellStyle(a1("A1"), currency)
    val copied = ok(s.copyRange(rng("A1:A2"), rng("D5:D6"), valuesOnly = false)(using marking))
    assertEquals(copied(a1("D5")).value, formula("B1|3,4"))
    assertEquals(copied(a1("D6")).value, formula("B2|3,4"))
    assertEquals(copied.getCellStyle(a1("D5")).map(_.numFmt), Some(NumFmt.Currency))
    val values = ok(s.copyRange(rng("A1:A2"), rng("D5:D6"), valuesOnly = true))
    assertEquals(values(a1("D5")).value, CellValue.Empty) // uncached: nothing to paste
    assertEquals(values(a1("D6")).value, num(6))
    assertEquals(values.getCellStyle(a1("D5")).map(_.numFmt), Some(NumFmt.Currency))
    // the cross-sheet form registers the source style in the target's registry
    val target = Sheet(SheetName.unsafe("T"))
    val cross = ok(target.copyRangeFrom(s, rng("A1:A2"), rng("B1:B2"), valuesOnly = true))
    assertEquals(cross.name, target.name)
    assertEquals(cross(a1("B2")).value, num(6))
    assertEquals(cross.getCellStyle(a1("B1")).map(_.numFmt), Some(NumFmt.Currency))
  }

  // ========== sort ==========

  test("sort is stable, keeps the header, moves styles and comments, and sorts empties last") {
    val bold = CellStyle.default.withFont(com.tjclp.xl.styles.font.Font.default.withBold())
    val s = Sheet(name)
      .put(a1("A1"), CellValue.Text("Name"))
      .put(a1("B1"), CellValue.Text("Qty"))
      .put(a1("A2"), CellValue.Text("pear"))
      .put(a1("B2"), num(5))
      .put(a1("A3"), CellValue.Text("apple"))
      .put(a1("B3"), num(3))
      .put(a1("A4"), CellValue.Text("Apple"))
      .put(a1("B4"), num(7))
      .put(a1("A5"), CellValue.Text("fig"))
      .withCellStyle(a1("A2"), bold)
      .comment(a1("A2"), Comment.plainText("note", None))
    val byName = ok(s.sort(rng("A1:B5"), Vector(Edit.SortKeySpec.ascending(Column.from0(0))), true))
    // case-insensitive and stable: "apple" (row 3) stays ahead of "Apple" (row 4)
    assertEquals(
      (1 to 5).map(r => byName(a1(s"A$r")).value),
      Vector("Name", "apple", "Apple", "fig", "pear").map(CellValue.Text.apply)
    )
    assertEquals(byName(a1("B5")).value, num(5)) // pear's quantity moved with it
    assertEquals(byName.getCellStyle(a1("A5")).map(_.font.bold), Some(true))
    assertEquals(byName.getComment(a1("A5")).map(_.text.toPlainText), Some("note"))
    assertEquals(byName.getComment(a1("A2")), None)
    // an empty key cell (fig has no quantity) sorts last when ascending; the direction flips the
    // whole comparator, so it sorts FIRST when descending — the CLI's rule, moved as it was
    val ascending =
      ok(s.sort(rng("A2:B5"), Vector(Edit.SortKeySpec.ascending(Column.from0(1))), false))
    assertEquals(
      (2 to 5).map(r => ascending(a1(s"B$r")).value),
      Vector(3, 5, 7).map(num) :+ CellValue.Empty
    )
    assertEquals(ascending(a1("A5")).value, CellValue.Text("fig"))
    val descending =
      ok(s.sort(rng("A2:B5"), Vector(Edit.SortKeySpec.descending(Column.from0(1))), false))
    assertEquals(
      (2 to 5).map(r => descending(a1(s"B$r")).value),
      CellValue.Empty +: Vector(7, 5, 3).map(num)
    )
    assertEquals(descending(a1("A2")).value, CellValue.Text("fig"))
  }

  test(
    "sort: numeric mode orders numbers before text, a moved formula shifts and drops its cache"
  ) {
    val s = Sheet(name)
      .put(a1("A1"), CellValue.Text("10"))
      .put(a1("A2"), CellValue.Text("9"))
      .put(a1("A3"), CellValue.Text("x"))
      .put(a1("A4"), num(1))
      .put(a1("B1"), CellValue.Formula("A1*2", Some(num(20))))
    val numeric = ok(
      s.sort(
        rng("A1:B4"),
        Vector(Edit.SortKeySpec(Column.from0(0), Edit.SortDir.Ascending, Edit.SortMode.Numeric)),
        hasHeader = false
      )(using marking)
    )
    assertEquals(
      (1 to 4).map(r => numeric(a1(s"A$r")).value),
      Vector(num(1), CellValue.Text("9"), CellValue.Text("10"), CellValue.Text("x"))
    )
    assertEquals(numeric(a1("B3")).value, formula("A1*2|0,2"))
    assertEquals(
      s.sort(rng("A1:A4"), Vector(Edit.SortKeySpec.ascending(Column.from0(1))), false),
      Left(XLError.InvalidReference("Sort column B is outside range A1:A4"))
    )
    assertEquals(
      s.sort(rng("A1:A4"), Vector.empty, hasHeader = false),
      Left(XLError.InvalidReference("sort requires at least one key column"))
    )
    // only a header row: nothing to sort
    assertEquals(
      s.sort(rng("A1:B1"), Vector(Edit.SortKeySpec.ascending(Column.from0(0))), true),
      Right(s)
    )
  }

  // ========== clear ==========

  test("clearRange: contents unmerge overlapping regions; styles and comments clear alone") {
    val s = Sheet(name)
      .put(a1("A1"), num(1))
      .put(a1("C3"), num(3))
      .withCellStyle(a1("A1"), CellStyle.default.withNumFmt(NumFmt.Percent))
      .comment(a1("A1"), Comment.plainText("note", None))
      .merge(rng("A1:B2"))
      .merge(rng("D1:E1"))
    val contents = s.clearRange(rng("A1:A1"), ClearWhat.contents)
    assert(!contents.contains(a1("A1")))
    assertEquals(contents.mergedRanges.toSet, Set(rng("D1:E1")))
    assertEquals(contents.getComment(a1("A1")).map(_.text.toPlainText), Some("note"))
    val styles = s.clearRange(rng("A1:C3"), ClearWhat.styles)
    assertEquals(styles(a1("A1")).value, num(1))
    assertEquals(styles(a1("A1")).styleId, None)
    assertEquals(styles.mergedRanges.size, 2)
    val comments = s.clearRange(rng("A1:C3"), ClearWhat.comments)
    assertEquals(comments.getComment(a1("A1")), None)
    assertEquals(comments(a1("A1")).value, num(1))
    val all = s.clearRange(rng("A1:C3"), ClearWhat.all)
    assert(all.cells.isEmpty && all.comments.isEmpty)
    assertEquals(all.mergedRanges.toSet, Set(rng("D1:E1")))
  }

  // ========== auto-fit ==========

  test(
    "autoFitWidth: Excel's default for an empty column, the 5.0 floor otherwise; autoFit sets it"
  ) {
    val s = Sheet(name).put(a1("A1"), CellValue.Text("x")).put(a1("B1"), CellValue.Text("a" * 40))
    assertEquals(s.autoFitWidth(Column.from0(5)), 8.43)
    assertEquals(s.autoFitWidth(Column.from0(0)), 5.0)
    assert(s.autoFitWidth(Column.from0(1)) > 20.0)
    val fitted = s.autoFit(Vector(Column.from0(0), Column.from0(1)))
    assertEquals(fitted.getColumnProperties(Column.from0(0)).width, Some(5.0))
    assertEquals(
      fitted.getColumnProperties(Column.from0(1)).width,
      Some(s.autoFitWidth(Column.from0(1)))
    )
    assertEquals(s.autoFitAll, fitted)
    assertEquals(Sheet(name).autoFitAll, Sheet(name))
  }

  // ========== grouping ==========

  test("groupRows/groupCols mark members and the summary after a collapsed group; ungroup clears") {
    val grouped =
      ok(Sheet(name).groupRows(RowSpan(Row.from1(2), Row.from1(4)), 2, collapsed = true))
    (2 to 4).foreach { r =>
      assertEquals(grouped.getRowProperties(Row.from1(r)).outlineLevel, Some(2))
      assert(grouped.getRowProperties(Row.from1(r)).hidden)
    }
    assert(grouped.getRowProperties(Row.from1(5)).collapsed)
    assert(!grouped.getRowProperties(Row.from1(1)).collapsed)
    val open =
      ok(Sheet(name).groupCols(ColSpan(Column.from0(4), Column.from0(7)), 1, collapsed = false))
    assertEquals(open.getColumnProperties(Column.from0(7)).outlineLevel, Some(1))
    assert(!open.getColumnProperties(Column.from0(7)).hidden)
    assert(!open.getColumnProperties(Column.from0(8)).collapsed)
    val ungrouped = grouped.ungroupRows(RowSpan(Row.from1(2), Row.from1(4)))
    assertEquals(ungrouped.getRowProperties(Row.from1(3)).outlineLevel, None)
    assert(!ungrouped.getRowProperties(Row.from1(5)).collapsed)
    assert(ungrouped.getRowProperties(Row.from1(3)).hidden) // like Excel, ungroup does not unhide
    assertEquals(
      Sheet(name).groupRows(RowSpan.single(Row.from1(1)), 8, collapsed = false),
      Left(XLError.Other("Outline level must be 1-7, got: 8"))
    )
    assertEquals(
      Sheet(name).groupCols(ColSpan.single(Column.from0(0)), 0, collapsed = false),
      Left(XLError.Other("Outline level must be 1-7, got: 0"))
    )
  }

  // ========== appearance & print ==========

  test("mergeSheetView / mergePageSetup / mergeHeaderFooter merge into the current settings") {
    val viewed = ok(Sheet(name).mergeSheetView(Some(false), Some(85), None))
    val viewedAgain = ok(viewed.mergeSheetView(None, None, Some(true)))
    assertEquals(
      viewedAgain.viewSettings.map(v => (v.showGridLines, v.zoomScale, v.tabSelected)),
      Some((false, Some(85), Some(true)))
    )
    assertEquals(
      Sheet(name).mergeSheetView(None, Some(401), None),
      Left(XLError.Other("Zoom scale must be 10-400, got: 401"))
    )
    val setUp = ok(Sheet(name).mergePageSetup(Some("landscape"), None, Some(1), Some(0), None))
    val setUpAgain = ok(setUp.mergePageSetup(None, Some(80), None, None, Some(false)))
    assertEquals(
      setUpAgain.pageSetup.map(p =>
        (p.orientation, p.scale, p.fitToWidth, p.fitToHeight, p.fitToPage)
      ),
      Some((Some("landscape"), 80, Some(1), Some(0), Some(false)))
    )
    assertEquals(
      Sheet(name).mergePageSetup(Some("sideways"), None, None, None, None),
      Left(XLError.Other("Orientation must be 'portrait' or 'landscape', got: sideways"))
    )
    assertEquals(
      Sheet(name).mergePageSetup(None, Some(401), None, None, None),
      Left(XLError.Other("Scale must be 10-400, got: 401"))
    )
    assertEquals(
      Sheet(name).mergePageSetup(None, None, None, Some(-1), None),
      Left(XLError.Other("fit-to-height must be >= 0 (0 = automatic), got: -1"))
    )
    val hf = Sheet(name)
      .mergeHeaderFooter(Some("&LLeft"), None, None, Some("even foot"), None, None, false, false)
      .mergeHeaderFooter(None, Some("Page &P"), None, None, None, None, false, true)
    val setup = hf.pageSetup.flatMap(_.headerFooter).getOrElse(fail("no header/footer"))
    assertEquals(setup.oddHeader, Some("&LLeft"))
    assertEquals(setup.oddFooter, Some("Page &P"))
    assertEquals(setup.evenFooter, Some("even foot"))
    assert(setup.differentOddEven) // even text set the flag
    assert(setup.differentFirst) // the explicit flag set it without text
  }

  test("withAutoFilter sets the range; removeAutoFilter actively strips a preserved filter") {
    val s = Sheet(name).withAutoFilter(rng("A1:C10"))
    assertEquals(s.autoFilter, Some(AutoFilterState.Ranged(rng("A1:C10"))))
    assertEquals(s.removeAutoFilter.autoFilter, Some(AutoFilterState.Remove))
  }
