package com.tjclp.xl.ops

import java.time.LocalDate

import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row, SheetName}
import com.tjclp.xl.cells.{CellValue, Comment}
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.sheets.{AutoFilterState, FreezePane, Sheet}
import com.tjclp.xl.sheets.styleSyntax.{getCellStyle, withCellStyle}
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.color.Color
import com.tjclp.xl.styles.font.Font
import com.tjclp.xl.styles.numfmt.NumFmt
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * The interpreter, case by case: THE sheet rule, invariant 4 through `Put`, the text-only refusals,
 * the semantics moved from the CLI (fill, copy, sort, clear, group, autofit, appearance) and the
 * workbook-level edits. Every failure surfaces as `EditFailed(index, name, cause)`.
 */
class EditInterpreterSpec extends FunSuite:

  import EditGenerators.{baseData, baseOther, baseWorkbook, data, missing, other}

  private given FormulaSupport = FormulaSupport.textOnly

  private def a1(s: String): ARef = ARef.parse(s).fold(e => fail(e), identity)
  private def rng(s: String): CellRange = CellRange.parse(s).fold(e => fail(e), identity)
  private def loc(s: String): Loc = Loc.parse(s).fold(e => fail(e.message), identity)
  private def area(s: String): Area = Area.parse(s).fold(e => fail(e.message), identity)
  private def num(n: Int): CellValue = CellValue.Number(BigDecimal(n))

  private def apply(wb: Workbook, edits: Edit*)(using FormulaSupport): XLResult[Workbook] =
    Edit.applyAll(wb, edits.toVector, Scope.of(data)).map(_.workbook)

  private def sheetAfter(wb: Workbook, edits: Edit*)(using FormulaSupport): Sheet =
    apply(wb, edits*).flatMap(_(data)).fold(e => fail(s"apply failed: ${e.message}"), identity)

  private def failure(wb: Workbook, edits: Edit*)(using FormulaSupport): XLError =
    apply(wb, edits*).fold(identity, _ => fail("expected a failure"))

  // ========== THE sheet rule ==========

  test(
    "sheet rule: a qualifier wins over the scope; the scope's default fills an unqualified target"
  ) {
    // Z1/Z2 are empty on both base sheets, so where each put landed is unambiguous
    val edits = Vector(Edit.Put(loc("Other!Z1"), num(1), None), Edit.Put(loc("Z2"), num(2), None))
    val r = Edit.applyAll(baseWorkbook, edits, Scope.of(data)).map(_.workbook)
    val dataSheet = r.flatMap(_(data)).fold(e => fail(e.message), identity)
    val otherSheet = r.flatMap(_(other)).fold(e => fail(e.message), identity)
    assert(!dataSheet.contains(a1("Z1")), "the qualified put must not land on the scope's default")
    assertEquals(dataSheet(a1("Z2")).value, num(2))
    assertEquals(otherSheet(a1("Z1")).value, num(1))
    assert(
      !otherSheet.contains(a1("Z2")),
      "the unqualified put must not land on the qualified sheet"
    )
    assertEquals(r.map(_.sheets.map(_.name)), Right(Vector(data, other)))
  }

  test("sheet rule: no scope on a two-sheet book is SheetRequired naming the candidates") {
    val r = Edit.applyAll(baseWorkbook, Vector(Edit.Put(loc("A1"), num(1), None)), Scope.none)
    assertEquals(
      r.left.map(_.root),
      Left(XLError.SheetRequired("put", Vector("Data", "Other")))
    )
    assertEquals(r.left.map(_.opIndex), Left(Some(1)))
  }

  test("sheet rule: a single-sheet book auto-selects its only sheet") {
    val single = Workbook(baseData)
    val r = Edit.applyAll(single, Vector(Edit.Put(loc("Z9"), num(9), None)), Scope.none)
    assertEquals(r.flatMap(_.workbook(data)).map(_(a1("Z9")).value), Right(num(9)))
    assertEquals(r.map(_.planned.map(_.sheet)), Right(Vector(Some(data))))
  }

  test("sheet rule: an unknown sheet is SheetNotFound at the failing index") {
    assertEquals(
      failure(baseWorkbook, Edit.Merge(area("A1:B2")), Edit.Merge(area("Nope!A1:B2"))),
      XLError.EditFailed(2, "merge", XLError.SheetNotFound("Nope"))
    )
  }

  // ========== put & hints ==========

  test("Edit.put lifts the codec hint: applying equals Sheet.put for typed values") {
    val date = LocalDate.of(2025, 1, 15)
    val viaEdit = sheetAfter(
      Workbook(baseData),
      Edit.put(loc("F1"), date),
      Edit.put(loc("F2"), 42),
      Edit.put(loc("F3"), BigDecimal("10.5")),
      Edit.put(loc("F4"), "text")
    )
    val viaSheet = baseData
      .put(a1("F1"), date)
      .put(a1("F2"), 42)
      .put(a1("F3"), BigDecimal("10.5"))
      .put(a1("F4"), "text")
    assertEquals(viaEdit, viaSheet)
    assertEquals(viaEdit.getCellStyle(a1("F1")).map(_.numFmt), Some(NumFmt.Date))
  }

  test("GH-560: an explicit hint replaces a Currency numFmt, an inferred one does not") {
    val currencyCell = a1("A1") // baseData styles A1 as Currency
    val explicit = sheetAfter(
      Workbook(baseData),
      Edit.Put(loc("A1"), num(5), Some(FormatHint.Explicit(NumFmt.Date)))
    )
    assertEquals(explicit.getCellStyle(currencyCell).map(_.numFmt), Some(NumFmt.Date))
    val inferred = sheetAfter(
      Workbook(baseData),
      Edit.Put(loc("A1"), num(5), Some(FormatHint.Inferred(NumFmt.Date)))
    )
    assertEquals(inferred.getCellStyle(currencyCell).map(_.numFmt), Some(NumFmt.Currency))
    val custom =
      baseData.withCellStyle(a1("B1"), CellStyle.default.withNumFmt(NumFmt.Custom("0.0x")))
    val replaced = sheetAfter(
      Workbook(custom),
      Edit.Put(loc("B1"), num(5), Some(FormatHint.Explicit(NumFmt.Percent)))
    )
    assertEquals(replaced.getCellStyle(a1("B1")).map(_.numFmt), Some(NumFmt.Percent))
  }

  test("put-values fills row-major and refuses a count mismatch") {
    val s = sheetAfter(
      Workbook(baseData),
      Edit.PutValues(area("G1:H2"), Vector(num(1), num(2), num(3), num(4)), None)
    )
    assertEquals(
      Vector("G1", "H1", "G2", "H2").map(r => s(a1(r)).value),
      Vector(1, 2, 3, 4).map(num)
    )
    assertEquals(
      failure(Workbook(baseData), Edit.PutValues(area("G1:H2"), Vector(num(1)), None)),
      XLError.EditFailed(1, "put-values", XLError.ValueCountMismatch(4, 1, "put-values G1:H2"))
    )
  }

  test("put-formula stores the text uncached (leading = dropped); an empty formula is refused") {
    val s = sheetAfter(Workbook(baseData), Edit.PutFormula(loc("G1"), "=SUM(B2:B4)", None))
    assertEquals(s(a1("G1")).value, CellValue.Formula("SUM(B2:B4)", None))
    assert(failure(Workbook(baseData), Edit.PutFormula(loc("G1"), "= ", None)).root match
      case XLError.FormulaError(_, _) => true
      case _ => false)
  }

  test("text-only support refuses a drag, the structural edits and a rename with the capability") {
    def refused(e: Edit, op: String): Unit =
      failure(baseWorkbook, e) match
        case XLError.EditFailed(1, name, XLError.UnsupportedCapability(o, _, _)) =>
          assertEquals(name, EditSchema.nameOf(e))
          assertEquals(o, op)
        case other => fail(s"expected a refusal for $e, got $other")
    refused(Edit.DragFormula(area("G1:G3"), "=B2*2", a1("G1"), None), "shift")
    refused(Edit.InsertRows(None, Row.from0(1), 1), "insert-rows")
    refused(Edit.DeleteRows(None, Row.from0(1), 1), "delete-rows")
    refused(Edit.InsertCols(None, Column.from0(1), 1), "insert-cols")
    refused(Edit.DeleteCols(None, Column.from0(1), 1), "delete-cols")
    refused(Edit.RenameSheet(data, SheetName.unsafe("Renamed")), "rename-sheet")
  }

  /** A formula-blind support for the rule-4 test: renames the tab without rewriting anything. */
  private val blindRename: FormulaSupport = new FormulaSupport:
    def validate(formula: String): XLResult[Unit] = FormulaSupport.textOnly.validate(formula)
    def shift(formula: String, colDelta: Int, rowDelta: Int): XLResult[String] =
      FormulaSupport.textOnly.shift(formula, colDelta, rowDelta)
    def insertRows(wb: Workbook, sheet: SheetName, at: Row, count: Int): XLResult[Workbook] =
      Right(
        wb.put(wb.sheets.find(_.name == sheet).fold(Sheet(sheet))(_.insertRows(at.index0, count)))
      )
    def deleteRows(wb: Workbook, sheet: SheetName, at: Row, count: Int): XLResult[Workbook] =
      FormulaSupport.textOnly.deleteRows(wb, sheet, at, count)
    def insertCols(wb: Workbook, sheet: SheetName, at: Column, count: Int): XLResult[Workbook] =
      FormulaSupport.textOnly.insertCols(wb, sheet, at, count)
    def deleteCols(wb: Workbook, sheet: SheetName, at: Column, count: Int): XLResult[Workbook] =
      FormulaSupport.textOnly.deleteCols(wb, sheet, at, count)
    def renameSheet(wb: Workbook, from: SheetName, to: SheetName): XLResult[Workbook] =
      wb.rename(from, to)

  test("scope rule 4: a rename of the default sheet retargets the edits after it") {
    given FormulaSupport = blindRename
    val renamed = SheetName.unsafe("Renamed")
    val r = Edit.applyAll(
      baseWorkbook,
      Vector(Edit.RenameSheet(data, renamed), Edit.Put(loc("Z1"), num(7), None)),
      Scope.of(data)
    )
    assertEquals(r.map(_.scope), Right(Scope.of(renamed)))
    assertEquals(r.flatMap(_.workbook(renamed)).map(_(a1("Z1")).value), Right(num(7)))
    assertEquals(r.map(_.planned.map(_.sheet)), Right(Vector(None, Some(renamed))))
  }

  test("structural edits go through FormulaSupport and the Planned row is structural") {
    given FormulaSupport = blindRename
    val r =
      Edit.applyAll(baseWorkbook, Vector(Edit.InsertRows(None, Row.from0(1), 2)), Scope.of(data))
    assertEquals(r.flatMap(_.workbook(data)).map(_(a1("A4")).value), Right(CellValue.Text("apple")))
    assert(r.exists(_.structural))
    assertEquals(
      failure(baseWorkbook, Edit.InsertRows(None, Row.from0(1), 0)).root,
      XLError.Other("insert-rows: count must be at least 1, got 0")
    )
  }

  // ========== fill / copy / sort / clear ==========

  test(
    "fill down repeats the source rows; fill right repeats the source columns; formulas need support"
  ) {
    val s = Sheet(data).put(a1("A1"), num(100)).put(a1("B1"), CellValue.Text("x"))
    val down =
      sheetAfter(Workbook(s), Edit.Fill(Area(None, rng("A1:B1")), rng("A1:B4"), Edit.FillDir.Down))
    (1 to 4).foreach(r => assertEquals(down(ARef.from0(0, r - 1)).value, num(100)))
    (1 to 4).foreach(r => assertEquals(down(ARef.from0(1, r - 1)).value, CellValue.Text("x")))
    val right =
      sheetAfter(Workbook(s), Edit.Fill(Area(None, rng("A1")), rng("A1:D1"), Edit.FillDir.Right))
    (0 to 3).foreach(c => assertEquals(right(ARef.from0(c, 0)).value, num(100)))
    // the moved validation texts
    assertEquals(
      failure(Workbook(s), Edit.Fill(Area(None, rng("A1")), rng("B1:B3"), Edit.FillDir.Down)).root,
      XLError.InvalidReference("Fill down requires matching columns. Source: A1:A1, Target: B1:B3")
    )
    assertEquals(
      failure(Workbook(s), Edit.Fill(Area(None, rng("A1")), rng("B2:D2"), Edit.FillDir.Right)).root,
      XLError.InvalidReference("Fill right requires matching rows. Source: A1:A1, Target: B2:D2")
    )
    // a formula source needs shifting, which the text-only support refuses
    val f = s.put(a1("A1"), CellValue.Formula("B1*2", None))
    assert(
      failure(
        Workbook(f),
        Edit.Fill(Area(None, rng("A1")), rng("A1:A3"), Edit.FillDir.Down)
      ).root match
        case XLError.UnsupportedCapability("shift", _, _) => true
        case _ => false
    )
  }

  test("copy moves values and styles, auto-expands the target, and pastes cached values only") {
    val copied =
      sheetAfter(Workbook(baseData), Edit.Copy(area("A1:B2"), loc("F10"), valuesOnly = false))
    assertEquals(copied(a1("F10")).value, CellValue.Text("Item"))
    assertEquals(copied(a1("G11")).value, num(3))
    assertEquals(copied.getCellStyle(a1("F10")).map(_.numFmt), Some(NumFmt.Currency))
    // values-only: the cached formula value lands, an uncached one becomes Empty
    val values =
      sheetAfter(Workbook(baseData), Edit.Copy(area("C2:C3"), loc("H1"), valuesOnly = true))
    assertEquals(values(a1("H1")).value, num(6))
    assert(!values.contains(a1("H2")) || values(a1("H2")).value == CellValue.Empty)
    // cross-sheet: the target sheet is written, the source untouched
    val cross =
      apply(baseWorkbook, Edit.Copy(area("Data!A1:A2"), loc("Other!D1"), valuesOnly = false))
    assertEquals(cross.flatMap(_(other)).map(_(a1("D2")).value), Right(CellValue.Text("apple")))
    assertEquals(cross.flatMap(_(data)), Right(baseData))
    // a formula source under text-only support is refused, a target past the grid is out of bounds
    assert(
      failure(Workbook(baseData), Edit.Copy(area("C2"), loc("H1"), valuesOnly = false)).root match
        case XLError.UnsupportedCapability("shift", _, _) => true
        case _ => false
    )
    assert(
      failure(
        Workbook(baseData),
        Edit.Copy(area("A1:B2"), loc("XFD1"), valuesOnly = false)
      ).root match
        case XLError.OutOfBounds(_, _) => true
        case _ => false
    )
  }

  test("sort orders rows by the key, keeps the header, moves styles and comments, empties last") {
    val sorted = sheetAfter(
      Workbook(baseData),
      Edit.Sort(
        area("A1:B4"),
        Vector(Edit.SortKeySpec.ascending(Column.from0(0))),
        hasHeader = true
      )
    )
    assertEquals(
      Vector("A1", "A2", "A3", "A4").map(r => sorted(a1(r)).value.toString),
      Vector("Text(Item)", "Text(apple)", "Text(fig)", "Text(pear)")
    )
    assertEquals(sorted(a1("B3")).value, num(1)) // fig's quantity moved with it
    assertEquals(sorted.getComment(a1("A2")).map(_.text.toPlainText), Some("first fruit"))
    val desc = sheetAfter(
      Workbook(baseData),
      Edit.Sort(
        area("A2:B4"),
        Vector(Edit.SortKeySpec(Column.from0(1), Edit.SortDir.Descending, Edit.SortMode.Numeric)),
        hasHeader = false
      )
    )
    assertEquals(Vector("B2", "B3", "B4").map(r => desc(a1(r)).value), Vector(5, 3, 1).map(num))
    assertEquals(
      failure(
        Workbook(baseData),
        Edit.Sort(
          area("A1:B4"),
          Vector(Edit.SortKeySpec.ascending(Column.from0(3))),
          hasHeader = false
        )
      ).root,
      XLError.InvalidReference("Sort column D is outside range A1:B4")
    )
  }

  test(
    "sort drops the cache of a moved formula, which is why it is outside the idempotence class"
  ) {
    // A support whose shift succeeds without rewriting: the point here is the cache, not the text
    given FormulaSupport = new FormulaSupport:
      def validate(formula: String): XLResult[Unit] = FormulaSupport.textOnly.validate(formula)
      def shift(formula: String, colDelta: Int, rowDelta: Int): XLResult[String] = Right(formula)
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
    // C1 holds a cached formula (6), C2 a plain 3: ascending puts 3 above 6 and moves the formula
    val s = Sheet(data)
      .put(a1("C1"), CellValue.Formula("B1*2", Some(num(6))))
      .put(a1("C2"), num(3))
    val key = Vector(Edit.SortKeySpec.ascending(Column.from0(2)))
    val once = sheetAfter(Workbook(s), Edit.Sort(area("A1:C2"), key, hasHeader = false))
    assertEquals(once(a1("C1")).value, num(3))
    once(a1("C2")).value match // the moved formula, uncached
      case CellValue.Formula(_, None, _) => ()
      case other => fail(s"expected an uncached formula at C2, got $other")
    // The uncached formula now sorts as empty text, ahead of the number: not a fixed point
    val twice = sheetAfter(Workbook(once), Edit.Sort(area("A1:C2"), key, hasHeader = false))
    assertNotEquals(twice, once)
    assert(!EditSchema.specOf(Edit.Sort(area("A1:C2"), key, hasHeader = false)).idempotent)
  }

  test("clear: contents unmerge overlaps; styles and comments clear independently") {
    val all = sheetAfter(Workbook(baseData), Edit.Clear(area("A1:F6"), ClearWhat.all))
    assert(all.cells.isEmpty)
    assert(all.comments.isEmpty)
    assert(all.mergedRanges.isEmpty)
    val styles = sheetAfter(Workbook(baseData), Edit.Clear(area("A1"), ClearWhat.styles))
    assertEquals(styles(a1("A1")).value, CellValue.Text("Item"))
    assertEquals(styles(a1("A1")).styleId, None)
    val comments = sheetAfter(Workbook(baseData), Edit.Clear(area("A2"), ClearWhat.comments))
    assert(comments.comments.isEmpty)
    assertEquals(comments(a1("A2")).value, CellValue.Text("apple"))
    assertEquals(
      failure(Workbook(baseData), Edit.Clear(area("A1"), ClearWhat(false, false, false))).root,
      XLError.Other("clear: nothing to clear (contents, styles or comments)")
    )
  }

  // ========== style & layout ==========

  test("style merges onto existing styles or replaces them") {
    val merged = sheetAfter(
      Workbook(baseData),
      Edit.Style(area("A1"), StyleOverlay(bold = Some(true)), StyleMode.Merge)
    )
    assertEquals(
      merged.getCellStyle(a1("A1")).map(s => (s.font.bold, s.numFmt)),
      Some((true, NumFmt.Currency))
    )
    val replaced = sheetAfter(
      Workbook(baseData),
      Edit.Style(area("A1"), StyleOverlay(bold = Some(true)), StyleMode.Replace)
    )
    assertEquals(
      replaced.getCellStyle(a1("A1")).map(s => (s.font.bold, s.numFmt)),
      Some((true, NumFmt.General))
    )
    assertEquals(
      failure(
        Workbook(baseData),
        Edit.Style(area("A1"), StyleOverlay(fontSize = Some(-1)), StyleMode.Merge)
      ).root,
      XLError.Other("style: font size must be positive, got: -1.0")
    )
  }

  test(
    "widths, heights and visibility set the properties; group/ungroup follow Excel's outline rules"
  ) {
    val s = sheetAfter(
      Workbook(baseData),
      Edit.ColWidth(None, ColSpan(Column.from0(0), Column.from0(1)), 20.0),
      Edit.RowHeight(None, RowSpan.single(Row.from0(0)), 30.0),
      Edit.HideCols(None, ColSpan.single(Column.from0(2))),
      Edit.HideRows(None, RowSpan.single(Row.from0(5))),
      Edit.GroupRows(None, RowSpan(Row.from0(1), Row.from0(3)), 2, collapsed = true)
    )
    assertEquals(s.getColumnProperties(Column.from0(1)).width, Some(20.0))
    assertEquals(s.getRowProperties(Row.from0(0)).height, Some(30.0))
    assert(s.getColumnProperties(Column.from0(2)).hidden)
    assert(s.getRowProperties(Row.from0(5)).hidden)
    assertEquals(s.getRowProperties(Row.from0(2)).outlineLevel, Some(2))
    assert(s.getRowProperties(Row.from0(2)).hidden)
    assert(s.getRowProperties(Row.from0(4)).collapsed) // the summary row after the group
    val shown = sheetAfter(
      Workbook(s),
      Edit.ShowCols(None, ColSpan.single(Column.from0(2))),
      Edit.UngroupRows(None, RowSpan(Row.from0(1), Row.from0(3)))
    )
    assert(!shown.getColumnProperties(Column.from0(2)).hidden)
    assertEquals(shown.getRowProperties(Row.from0(2)).outlineLevel, None)
    assert(!shown.getRowProperties(Row.from0(4)).collapsed)
    assert(shown.getRowProperties(Row.from0(2)).hidden) // ungroup does not unhide, like Excel
    assertEquals(
      failure(
        Workbook(baseData),
        Edit.GroupCols(None, ColSpan.single(Column.from0(0)), 9, collapsed = false)
      ).root,
      XLError.Other("Outline level must be 1-7, got: 9")
    )
    assertEquals(
      failure(Workbook(baseData), Edit.ColWidth(None, ColSpan.single(Column.from0(0)), 300.0)).root,
      XLError.Other("col-width: width must be 0-255 character units, got 300.0")
    )
  }

  test("autofit sets a width for every used column, or the named span") {
    val all = sheetAfter(Workbook(baseData), Edit.AutoFit(None, None))
    (0 to 3).foreach { c =>
      assert(all.getColumnProperties(Column.from0(c)).width.exists(_ >= 5.0), s"column $c")
    }
    val one =
      sheetAfter(Workbook(baseData), Edit.AutoFit(None, Some(ColSpan.single(Column.from0(0)))))
    assert(one.getColumnProperties(Column.from0(0)).width.isDefined)
    assertEquals(one.getColumnProperties(Column.from0(1)).width, None)
    assertEquals(
      one.getColumnProperties(Column.from0(0)).width,
      Some(baseData.autoFitWidth(Column.from0(0)))
    )
  }

  // ========== annotations & objects ==========

  test("comments and hyperlinks are set and cleared") {
    val s = sheetAfter(
      Workbook(baseData),
      Edit.SetComment(loc("B1"), Comment.plainText("note", None)),
      Edit.RemoveComment(loc("A2")),
      Edit.Hyperlink(loc("A1"), Some("https://example.com")),
      Edit.Hyperlink(loc("A3"), Some("#Other!A1"))
    )
    assertEquals(s.getComment(a1("B1")).map(_.text.toPlainText), Some("note"))
    assertEquals(s.getComment(a1("A2")), None)
    assertEquals(s(a1("A1")).hyperlink, Some("https://example.com"))
    val cleared = sheetAfter(Workbook(s), Edit.Hyperlink(loc("A1"), None))
    assertEquals(cleared(a1("A1")).hyperlink, None)
    assertEquals(
      failure(Workbook(baseData), Edit.Hyperlink(loc("A1"), Some(""))).root,
      XLError.Other("hyperlink: target cannot be empty (omit it to clear the link)")
    )
  }

  test(
    "conditional formats append one block with auto priorities; empty ranges or rules are refused"
  ) {
    import com.tjclp.xl.cf.{CfOperator, CfRule}
    import com.tjclp.xl.styles.Dxf
    val rule = CfRule.cellIs(CfOperator.GreaterThan, "1", Dxf.fill(Color.Rgb(0xffffcc00)))
    val s = sheetAfter(
      Workbook(baseData),
      Edit.AddConditionalFormat(None, Vector(rng("B2:B4")), Vector(rule))
    )
    assertEquals(s.conditionalFormats.size, 1)
    assertEquals(
      failure(Workbook(baseData), Edit.AddConditionalFormat(None, Vector.empty, Vector(rule))).root,
      XLError.Other("add-conditional-format: at least one range and one rule are required")
    )
  }

  // ========== view & print ==========

  test("freeze, view, tab colour, autofilter, page setup and header/footer merge into the sheet") {
    val s = sheetAfter(
      Workbook(baseData),
      Edit.Freeze(loc("B2")),
      Edit.SetSheetView(None, Some(false), Some(85), None),
      Edit.SetTabColor(None, Some(Color.Rgb(0xff1f4e79))),
      Edit.SetAutoFilter(None, Some(rng("A1:C4"))),
      Edit.SetPageSetup(None, Some("landscape"), None, Some(1), Some(0), None),
      Edit.SetHeaderFooter(
        None,
        None,
        Some("Page &P"),
        Some("even"),
        None,
        None,
        None,
        false,
        false
      )
    )
    assertEquals(s.freezePane, Some(FreezePane.At(a1("B2"))))
    assertEquals(s.viewSettings.map(v => (v.showGridLines, v.zoomScale)), Some((false, Some(85))))
    assertEquals(s.tabColor, Some(Color.Rgb(0xff1f4e79)))
    assertEquals(s.autoFilter, Some(AutoFilterState.Ranged(rng("A1:C4"))))
    assertEquals(s.pageSetup.flatMap(_.orientation), Some("landscape"))
    assertEquals(s.pageSetup.flatMap(_.fitToWidth), Some(1))
    assertEquals(s.pageSetup.map(_.scale), Some(100)) // untouched fields keep their defaults
    assertEquals(s.pageSetup.flatMap(_.headerFooter).flatMap(_.oddFooter), Some("Page &P"))
    assertEquals(s.pageSetup.flatMap(_.headerFooter).map(_.differentOddEven), Some(true))
    // the second pass keeps what the first set: merge, not replace
    val again = sheetAfter(
      Workbook(s),
      Edit.SetSheetView(None, None, Some(120), Some(true)),
      Edit.SetPageSetup(None, None, Some(80), None, None, None)
    )
    assertEquals(
      again.viewSettings.map(v => (v.showGridLines, v.zoomScale, v.tabSelected)),
      Some((false, Some(120), Some(true)))
    )
    assertEquals(again.pageSetup.map(p => (p.orientation, p.scale)), Some((Some("landscape"), 80)))
    val cleared = sheetAfter(
      Workbook(s),
      Edit.Unfreeze(None),
      Edit.SetTabColor(None, None),
      Edit.SetAutoFilter(None, None)
    )
    assertEquals(cleared.freezePane, Some(FreezePane.Remove))
    assertEquals(cleared.tabColor, None)
    assertEquals(cleared.autoFilter, Some(AutoFilterState.Remove))
    assertEquals(
      failure(Workbook(baseData), Edit.SetSheetView(None, None, Some(5), None)).root,
      XLError.Other("Zoom scale must be 10-400, got: 5")
    )
    assertEquals(
      failure(
        Workbook(baseData),
        Edit.SetPageSetup(None, Some("sideways"), None, None, None, None)
      ).root,
      XLError.Other("Orientation must be 'portrait' or 'landscape', got: sideways")
    )
    assertEquals(
      failure(Workbook(baseData), Edit.SetPageSetup(None, None, Some(5), None, None, None)).root,
      XLError.Other("Scale must be 10-400, got: 5")
    )
    assertEquals(
      failure(Workbook(baseData), Edit.SetPageSetup(None, None, None, Some(-1), None, None)).root,
      XLError.Other("fit-to-width must be >= 0 (0 = automatic), got: -1")
    )
  }

  // ========== workbook ==========

  test(
    "add-sheet at the end, after or before; duplicates (any case) and unknown anchors are refused"
  ) {
    val newName = SheetName.unsafe("New")
    val atEnd = apply(baseWorkbook, Edit.AddSheet(newName, None, None))
    assertEquals(atEnd.map(_.sheetNames.map(_.value)), Right(Seq("Data", "Other", "New")))
    val after = apply(baseWorkbook, Edit.AddSheet(newName, Some(data), None))
    assertEquals(after.map(_.sheetNames.map(_.value)), Right(Seq("Data", "New", "Other")))
    val before = apply(baseWorkbook, Edit.AddSheet(newName, None, Some(data)))
    assertEquals(before.map(_.sheetNames.map(_.value)), Right(Seq("New", "Data", "Other")))
    assertEquals(
      failure(baseWorkbook, Edit.AddSheet(SheetName.unsafe("data"), None, None)).root,
      XLError.DuplicateSheet("data")
    )
    assertEquals(
      failure(baseWorkbook, Edit.AddSheet(newName, Some(missing), None)).root,
      XLError.SheetNotFound("Nope")
    )
    assertEquals(
      failure(baseWorkbook, Edit.AddSheet(newName, Some(data), Some(other))).root,
      XLError.Other("add-sheet: after and before are mutually exclusive")
    )
  }

  test("remove, rename (with support), move, copy, hide and show sheets") {
    given FormulaSupport = blindRename
    assertEquals(
      apply(baseWorkbook, Edit.RemoveSheet(other)).map(_.sheetNames.map(_.value)),
      Right(Seq("Data"))
    )
    assertEquals(
      failure(Workbook(baseData), Edit.RemoveSheet(data)).root,
      XLError.InvalidWorkbook("Cannot remove last sheet")
    )
    assertEquals(
      apply(baseWorkbook, Edit.MoveSheet(other, Some(0), None, None))
        .map(_.sheetNames.map(_.value)),
      Right(Seq("Other", "Data"))
    )
    assertEquals(
      apply(baseWorkbook, Edit.MoveSheet(data, None, Some(other), None))
        .map(_.sheetNames.map(_.value)),
      Right(Seq("Other", "Data"))
    )
    assertEquals(
      apply(baseWorkbook, Edit.MoveSheet(other, None, None, Some(data)))
        .map(_.sheetNames.map(_.value)),
      Right(Seq("Other", "Data"))
    )
    assertEquals(
      failure(baseWorkbook, Edit.MoveSheet(other, None, None, None)).root,
      XLError.Other("move-sheet: exactly one of to, after or before is required")
    )
    val copied = apply(baseWorkbook, Edit.CopySheet(data, SheetName.unsafe("Data2")))
    assertEquals(copied.flatMap(_(SheetName.unsafe("Data2"))).map(_.cells), Right(baseData.cells))
    assertEquals(
      failure(baseWorkbook, Edit.CopySheet(data, SheetName.unsafe("OTHER"))).root,
      XLError.DuplicateSheet("OTHER")
    )
    val hidden = apply(baseWorkbook, Edit.HideSheet(other, veryHidden = true))
    assertEquals(hidden.map(_.getSheetState(other)), Right(Some("veryHidden")))
    assertEquals(
      hidden.flatMap(wb =>
        Edit
          .applyAll(wb, Vector(Edit.ShowSheet(other)), Scope.none)
          .map(_.workbook.getSheetState(other))
      ),
      Right(None)
    )
    assertEquals(
      failure(Workbook(baseData), Edit.HideSheet(data, veryHidden = false)).root,
      XLError.InvalidWorkbook("Cannot hide the last visible sheet")
    )
  }

  test(
    "defined names: workbook-scoped replace, sheet-scoped add, remove; a missing name is refused"
  ) {
    val wb = apply(
      baseWorkbook,
      Edit.DefineName("Total", "Data!$C$2:$C$4", None),
      Edit.DefineName("Rate", "0.08", Some(other)),
      Edit.DefineName("Rate", "0.09", Some(other))
    ).fold(e => fail(e.message), identity)
    val names = wb.metadata.definedNames
    assertEquals(
      names.filter(_.name == "Total").map(n => (n.formula, n.localSheetId)),
      Vector(("Data!$C$2:$C$4", None))
    )
    assertEquals(
      names.filter(_.name == "Rate").map(n => (n.formula, n.localSheetId)),
      Vector(("0.09", Some(1)))
    )
    val removed = apply(wb, Edit.RemoveName("Total", None), Edit.RemoveName("Rate", Some(other)))
      .fold(e => fail(e.message), identity)
    assert(removed.metadata.definedNames.isEmpty)
    assertEquals(
      failure(baseWorkbook, Edit.RemoveName("Rate", None)).root,
      XLError.Other("Named range 'Rate' not found")
    )
    assertEquals(
      failure(baseWorkbook, Edit.DefineName("", "1", None)).root,
      XLError.Other("define-name: name cannot be empty")
    )
  }

  // ========== the entry points: wb.edit / sheet.edit ==========

  test(
    "wb.edit runs applyAll under Scope.none; editIn names the default sheet; no edits is identity"
  ) {
    val single = Workbook(baseData)
    assertEquals(
      single.edit(Edit.Put(loc("Z1"), num(1), None)).flatMap(_(data)).map(_(a1("Z1")).value),
      Right(num(1))
    )
    assertEquals(
      baseWorkbook.edit(Edit.Put(loc("Z1"), num(1), None)).left.map(_.root),
      Left(XLError.SheetRequired("put", Vector("Data", "Other")))
    )
    assertEquals(
      baseWorkbook
        .editIn(other)(Edit.Put(loc("Z1"), num(1), None))
        .flatMap(_(other))
        .map(_(a1("Z1")).value),
      Right(num(1))
    )
    assertEquals(baseWorkbook.edit(), Right(baseWorkbook))
  }

  test("sheet.edit applies edits to the sheet alone, all-or-nothing, and follows a rename") {
    given FormulaSupport = blindRename
    val edited = baseData.edit(Edit.Put(loc("Z1"), num(1), None), Edit.Merge(area("Z1:Z2")))
    assertEquals(edited.map(_(a1("Z1")).value), Right(num(1)))
    assertEquals(edited.map(_.mergedRanges.contains(rng("Z1:Z2"))), Right(true))
    assertEquals(
      baseData.edit(Edit.Put(loc("Other!A1"), num(1), None)),
      Left(XLError.EditFailed(1, "put", XLError.SheetNotFound("Other")))
    )
    val renamed = SheetName.unsafe("Renamed")
    assertEquals(
      baseData
        .edit(Edit.RenameSheet(data, renamed), Edit.Put(loc("Z1"), num(2), None))
        .map(s => (s.name, s(a1("Z1")).value)),
      Right((renamed, num(2)))
    )
  }

  // ========== plan ==========

  test(
    "plan is applyAll's row list without the workbook, and touched ranges name the content writes"
  ) {
    val edits = Vector(
      Edit.Put(loc("A1"), num(1), None),
      Edit.Style(area("A1:B2"), StyleOverlay(bold = Some(true)), StyleMode.Merge),
      Edit.Copy(area("A1:B2"), loc("Other!C3"), valuesOnly = true)
    )
    val planned = Edit.plan(baseWorkbook, edits, Scope.of(data))
    assertEquals(planned, Edit.applyAll(baseWorkbook, edits, Scope.of(data)).map(_.planned))
    assertEquals(
      planned.map(_.map(p => (p.index, p.sheet, p.touched))),
      Right(
        Vector(
          (1, Some(data), Vector(rng("A1"))),
          (2, Some(data), Vector.empty),
          (3, Some(other), Vector(rng("C3:D4")))
        )
      )
    )
    val applied =
      Edit.applyAll(baseWorkbook, edits, Scope.of(data)).fold(e => fail(e.message), identity)
    assertEquals(
      applied.touchedBySheet,
      Map(data -> Vector(rng("A1")), other -> Vector(rng("C3:D4")))
    )
    assert(!applied.structural)
  }

  test("Edit.validate alone reports the edit-local refusals without a workbook") {
    assert(Edit.validate(Edit.Put(loc("A1"), num(1), None)).isRight)
    assert(Edit.validate(Edit.MoveSheet(data, Some(0), Some(other), None)).isLeft)
    assert(Edit.validate(Edit.RowHeight(None, RowSpan.single(Row.from0(0)), 500.0)).isLeft)
    assert(Edit.validate(Edit.Sort(area("A1:B2"), Vector.empty, hasHeader = false)).isLeft)
    assert(Edit.validate(Edit.DragFormula(area("A1:A3"), "", a1("A1"), None)).isLeft)
  }
