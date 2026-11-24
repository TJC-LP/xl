package com.tjclp.xl.sheets

import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cells.{Cell, CellValue, Comment, Hyperlink}
import com.tjclp.xl.error.XLResult
import com.tjclp.xl.ops.{Area, ClearWhat, Edit, FormulaSupport, Scope}
import com.tjclp.xl.patch.Patch
import com.tjclp.xl.sheets.styleSyntax.{getCellStyle, withCellStyle}
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.numfmt.NumFmt
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * GH-935: merging keeps one value. A non-empty upper-left cell wins, as in Excel's Merge Cells;
 * over an empty one the parser-less Sheet.merge keeps the first non-empty cell (row-major) where it
 * is (Excel moves it and rewrites its readers: the workbook-level merge). Every other cell of the
 * range is cleared the way Clear Contents clears it: the value, formula and hyperlink go; the
 * cell's own formatting and its comment stay. Reading a file whose merge already hides values
 * rewrites nothing: only the merge operation clears.
 */
class MergeClearsHiddenSpec extends FunSuite:

  private val name = SheetName.unsafe("S")
  private def a1(s: String): ARef = ARef.parse(s).fold(e => fail(e), identity)
  private def rng(s: String): CellRange = CellRange.parse(s).fold(e => fail(e), identity)
  private def num(n: Int): CellValue = CellValue.Number(BigDecimal(n))
  private def ok[A](r: XLResult[A]): A = r.fold(e => fail(e.message), identity)

  private given FormulaSupport = FormulaSupport.textOnly

  /** The issue's repro: A1=1, B1=2, A2=3. */
  private val repro = Sheet(name).put(a1("A1"), num(1)).put(a1("B1"), num(2)).put(a1("A2"), num(3))

  private def values(s: Sheet, range: String): Vector[CellValue] =
    rng(range).cellsRowMajor.map(r => s.cells.get(r).fold(CellValue.Empty)(_.value)).toVector

  test("GH-935: Sheet.merge keeps the anchor's value and clears every hidden cell") {
    val merged = repro.merge(rng("A1:B2"))
    assertEquals(
      values(merged, "A1:B2"),
      Vector(num(1), CellValue.Empty, CellValue.Empty, CellValue.Empty)
    )
    assertEquals(merged.mergedRanges, Set(rng("A1:B2")))
    // An unstyled hidden cell leaves no record behind
    assertEquals(merged.cells.get(a1("B1")), None)
    assertEquals(merged.cells.get(a1("A2")), None)
  }

  test("GH-935: hidden formulas are cleared too (a formula is the cell's content)") {
    val s = repro.put(a1("B2"), CellValue.Formula("A1*2", Some(num(2))))
    val merged = s.merge(rng("A1:B2"))
    assertEquals(merged.cells.get(a1("B2")), None)
    assertEquals(merged(a1("A1")).value, num(1))
  }

  // Sheet.merge has no parser, so over an empty upper-left cell it keeps the cell Excel would move
  // in place (Excel's move, checked against Excel for Mac, is CellMover.merge / Edit.Merge, pinned
  // in xl-evaluator's CellMoverSpec): every reader keeps Excel's value.

  test("GH-935: over an empty anchor Sheet.merge keeps the only value where it is") {
    val s = Sheet(name).put(a1("B1"), CellValue.Text("onlyB1"))
    val merged = s.merge(rng("A1:B1"))
    assertEquals(merged.cells.get(a1("A1")), None)
    assertEquals(merged(a1("B1")).value, CellValue.Text("onlyB1"))
  }

  test("GH-935 (Excel): a non-empty anchor wins and the other value is cleared") {
    val s = Sheet(name).put(a1("A3"), CellValue.Text("a3")).put(a1("B3"), CellValue.Text("b3"))
    val merged = s.merge(rng("A3:B3"))
    assertEquals(merged(a1("A3")).value, CellValue.Text("a3"))
    assertEquals(merged.cells.get(a1("B3")), None)
  }

  test("GH-935: over an empty anchor the FIRST non-empty cell (row-major) stays; the rest clear") {
    val s = Sheet(name)
      .put(a1("B5"), CellValue.Text("b5"))
      .put(a1("A6"), CellValue.Text("a6"))
      .put(a1("B6"), num(1))
    val merged = s.merge(rng("A5:B6"))
    assertEquals(merged(a1("B5")).value, CellValue.Text("b5"))
    assert(Vector("A5", "A6", "B6").forall(r => merged.cells.get(a1(r)).isEmpty))
  }

  test("GH-935: readers of the kept cell keep Excel's value; SUM over the area is that value") {
    val s = Sheet(name)
      .put(a1("B1"), num(7))
      .put(a1("B2"), num(3))
      .put(a1("E1"), CellValue.Formula("B1+1", None))
    val merged = s.merge(rng("A1:B2"))
    assertEquals(merged(a1("B1")).value, num(7))
    assertEquals(merged.cells.get(a1("B2")), None)
    assertEquals(merged(a1("E1")).value, CellValue.Formula("B1+1", None))
    // the kept cell keeps its own style and formula untouched
    val styled = Sheet(name)
      .put(a1("B1"), CellValue.Formula("D1*2", Some(num(10))))
      .withCellStyle(a1("B1"), CellStyle.default.withNumFmt(NumFmt.Percent))
      .merge(rng("A1:B1"))
    assertEquals(styled(a1("B1")).value, CellValue.Formula("D1*2", Some(num(10))))
    assertEquals(styled.getCellStyle(a1("B1")).map(_.numFmt), Some(NumFmt.Percent))
  }

  test("GH-935: moveCell (Excel's cut) carries value, format and an anchored array range") {
    import com.tjclp.xl.cells.FormulaKind
    val s = Sheet(name)
      .put(a1("B1"), CellValue.Formula("D1*2", Some(num(10))))
      .withCellStyle(a1("B1"), CellStyle.default.withNumFmt(NumFmt.Percent))
      .withCellStyle(a1("A1"), CellStyle.default.withNumFmt(NumFmt.Date))
    val moved = SheetEdits.moveCell(s, a1("B1"), a1("A1"))
    assertEquals(moved(a1("A1")).value, CellValue.Formula("D1*2", Some(num(10))))
    assertEquals(moved.getCellStyle(a1("A1")).map(_.numFmt), Some(NumFmt.Percent))
    assertEquals(moved.cells.get(a1("B1")), None)
    def dyn(r: String) =
      CellValue.Formula("SEQUENCE(1)", Some(num(1)), FormulaKind.dynamicArray(rng(r)))
    def cse(r: String) = CellValue.Formula("D1*2", Some(num(2)), FormulaKind.ArrayFormula(rng(r)))
    val a = SheetEdits.moveCell(Sheet(name).put(a1("B2"), dyn("B2:B2")), a1("B2"), a1("A1"))
    assertEquals(a(a1("A1")).value, dyn("A1:A1"))
    val b = SheetEdits.moveCell(Sheet(name).put(a1("C3"), cse("C3:C3")), a1("C3"), a1("B2"))
    assertEquals(b(a1("B2")).value, cse("B2:B2"))
  }

  test("GH-935: a moving merge needs the formula support in the Edit algebra (text-only refuses)") {
    val s = Sheet(name).put(a1("B1"), num(7))
    val edit = Edit.Merge(Area(None, rng("A1:B1")))
    assertEquals(Edit.lower(edit, s), None)
    assertEquals(Patch.toEdits(Patch.Merge(rng("A1:B1")), s), None)
    Edit.applyAll(Workbook(s), Vector(edit), Scope.none) match
      case Left(
            com.tjclp.xl.error.XLError
              .EditFailed(1, "merge", _: com.tjclp.xl.error.XLError.UnsupportedCapability)
          ) =>
        ()
      case other => fail(s"expected an UnsupportedCapability refusal, got $other")
  }

  test("GH-935: an empty range merges to nothing; a styled empty anchor is not a value") {
    val s = Sheet(name).withCellStyle(a1("A1"), CellStyle.default.withNumFmt(NumFmt.Percent))
    val merged = s.merge(rng("A1:B2"))
    assertEquals(merged(a1("A1")).value, CellValue.Empty)
    assertEquals(merged.mergedRanges, Set(rng("A1:B2")))
  }

  test("GH-935: a hidden cell keeps its formatting and comment; its hyperlink goes") {
    val s = repro
      .put(
        Cell(a1("B2"), CellValue.Text("site"), hyperlink = Some(Hyperlink("https://example.com")))
      )
      .withCellStyle(a1("B1"), CellStyle.default.withNumFmt(NumFmt.Percent))
      .comment(a1("A2"), Comment.plainText("note", None))
    val percent = s(a1("B1")).styleId
    val merged = s.merge(rng("A1:B2"))
    assertEquals(merged.cells.get(a1("B1")), Some(Cell(a1("B1"), CellValue.Empty, percent)))
    assertEquals(merged.getCellStyle(a1("B1")).map(_.numFmt), Some(NumFmt.Percent))
    assertEquals(merged.cells.get(a1("B2")), None)
    assertEquals(merged.getComment(a1("A2")).map(_.text.toPlainText), Some("note"))
  }

  test("GH-935: the anchor keeps its value, style and hyperlink") {
    val s = Sheet(name)
      .put(Cell(a1("A1"), num(7), hyperlink = Some(Hyperlink("https://example.com"))))
      .withCellStyle(a1("A1"), CellStyle.default.withNumFmt(NumFmt.Percent))
      .put(a1("B1"), num(2))
    val merged = s.merge(rng("A1:B1"))
    assertEquals(merged(a1("A1")).value, num(7))
    assertEquals(merged(a1("A1")).hyperlink, Some(Hyperlink("https://example.com")))
    assertEquals(merged.getCellStyle(a1("A1")).map(_.numFmt), Some(NumFmt.Percent))
  }

  test("GH-935: cells outside the range are untouched") {
    val s = repro.put(a1("C1"), num(9)).put(a1("A3"), num(8))
    val merged = s.merge(rng("A1:B2"))
    assertEquals(merged(a1("C1")).value, num(9))
    assertEquals(merged(a1("A3")).value, num(8))
  }

  test("GH-935: Patch.Merge and the range DSL clear the hidden cells the same way") {
    val viaPatch = Patch.applyPatch(repro, Patch.Merge(rng("A1:B2")))
    assertEquals(viaPatch, repro.merge(rng("A1:B2")))
  }

  test("GH-935: Edit.Merge clears hidden cells and reports the range as content it touched") {
    val wb = Workbook(repro)
    val applied = ok(Edit.applyAll(wb, Vector(Edit.Merge(Area(None, rng("A1:B2")))), Scope.none))
    val out = ok(applied.workbook(name))
    assertEquals(
      values(out, "A1:B2"),
      Vector(num(1), CellValue.Empty, CellValue.Empty, CellValue.Empty)
    )
    assertEquals(applied.touchedBySheet, Map(name -> Vector(rng("A1:B2"))))
  }

  test("GH-935: Edit.lower agrees with applying the edit (coherence law)") {
    val edit = Edit.Merge(Area(None, rng("A1:B2")))
    val lowered = Edit.lower(edit, repro).fold(fail("Merge lowers"))(identity)
    assertEquals(Patch.applyPatch(repro, lowered), repro.merge(rng("A1:B2")))
  }

  test("GH-935: a clear after the merge leaves nothing hidden under the merge") {
    val cleared = repro.merge(rng("A1:B2")).clearRange(rng("A1:A1"), ClearWhat.contents)
    assert(rng("A1:B2").cells.forall(r => cleared.cells.get(r).forall(_.value == CellValue.Empty)))
  }

  test("GH-935: a sheet built with hidden values under a merge keeps them (only merge clears)") {
    val loaded = repro.copy(mergedRanges = Set(rng("A1:B2")))
    assertEquals(values(loaded, "A1:B2"), Vector(num(1), num(2), num(3), CellValue.Empty))
    // unmerge and other edits do not clear them either
    assertEquals(values(loaded.unmerge(rng("A1:B2")), "A1:B2"), values(loaded, "A1:B2"))
  }

  // ===== GH-935 review: range hyperlinks, spill children, overlapping merges =====

  private val link = Hyperlink("https://example.com")

  test("GH-935 review: hidden cells lose range-link coverage as they lose own links") {
    val s = Sheet(name)
      .put(a1("A3"), CellValue.Text("keep"))
      .setHyperlink(rng("B3:C3"), link)
    val merged = s.merge(rng("A3:C3"))
    assertEquals(merged.hyperlinkAt(a1("B3")), None)
    assertEquals(merged.hyperlinkAt(a1("C3")), None)
    assertEquals(merged.rangeHyperlinks.toVector, Vector.empty)
    // a range link reaching outside the merge keeps its outside part
    val wide =
      Sheet(name).put(a1("A1"), num(1)).setHyperlink(rng("B1:D1"), link).merge(rng("A1:C1"))
    assertEquals(wide.hyperlinkAt(a1("B1")), None)
    assertEquals(wide.hyperlinkAt(a1("D1")), Some(link))
  }

  test("GH-935 review: moveCell carries the winner's effective (range) link to the anchor") {
    val s = Sheet(name).put(a1("B1"), CellValue.Text("x")).setHyperlink(rng("B1:C1"), link)
    val moved = SheetEdits.moveCell(s, a1("B1"), a1("A1"))
    assertEquals(moved(a1("A1")).value, CellValue.Text("x"))
    assertEquals(moved.hyperlinkAt(a1("A1")), Some(link))
    assertEquals(moved.hyperlinkAt(a1("B1")), None)
    // the kept-in-place winner of the parser-less merge keeps its link; the other hidden cell
    // loses its coverage
    val kept = s.merge(rng("A1:C1"))
    assertEquals(kept.hyperlinkAt(a1("B1")), Some(link))
    assertEquals(kept.hyperlinkAt(a1("C1")), None)
  }

  test("GH-935 review: a spill child is never the winner and is never made a constant") {
    import com.tjclp.xl.cells.FormulaKind
    val dyn = FormulaKind.dynamicArray(rng("C1:C3"))
    val s = Sheet(name)
      .put(a1("C1"), CellValue.Formula("SEQUENCE(3,1)", Some(num(1)), dyn))
      .put(a1("C2"), num(2))
      .put(a1("C3"), num(3))
    assertEquals(SheetEdits.mergeWinner(s, rng("B2:C3")), None)
    val merged = s.merge(rng("B2:C3"))
    assertEquals(merged.cells.get(a1("B2")), None)
    // a legacy array's members are spill children too
    val cse = Sheet(name)
      .put(a1("B1"), CellValue.Formula("D1:D2*2", None, FormulaKind.ArrayFormula(rng("B1:B2"))))
      .put(a1("B2"), num(4))
    assertEquals(SheetEdits.mergeWinner(cse, rng("A2:B2")), None)
  }

  test("GH-935 review: merging over existing merges absorbs them (never overlapping merges)") {
    val s = Sheet(name).put(a1("A1"), num(1)).put(a1("C2"), num(5)).merge(rng("B1:C2"))
    assertEquals(s.mergedRanges, Set(rng("B1:C2")))
    // A1:B1 touches B1:C2: the merge grows to A1:C2 and absorbs it
    val grown = s.merge(rng("A1:B1"))
    assertEquals(grown.mergedRanges, Set(rng("A1:C2")))
    assertEquals(values(grown, "A1:C2").count(_ != CellValue.Empty), 1)
    assertEquals(grown(a1("A1")).value, num(1))
    // a merge inside an existing one is absorbed by it
    assertEquals(grown.merge(rng("B2:C2")).mergedRanges, Set(rng("A1:C2")))
    // disjoint merges stay
    assertEquals(grown.merge(rng("E1:F1")).mergedRanges, Set(rng("A1:C2"), rng("E1:F1")))
  }

  test("GH-935 review 2: a merge touching a legacy array or a data table is refused") {
    import com.tjclp.xl.cells.FormulaKind
    import com.tjclp.xl.error.XLError
    val cse = Sheet(name)
      .put(
        a1("A2"),
        CellValue.Formula("SEQUENCE(2)", Some(num(1)), FormulaKind.ArrayFormula(rng("A2:A3")))
      )
      .put(a1("A3"), num(2))
      .put(a1("C1"), CellValue.Formula("SUM(A1:A3)", Some(num(3))))
    assert(SheetEdits.mergeRefusal(cse, rng("A1:A3")).isDefined)
    // the total Sheet.merge leaves the sheet as it was rather than merge cells of an array
    assertEquals(cse.merge(rng("A1:A3")), cse)
    // so does a merge that only GROWS into the array through an existing merge
    val grown = cse.copy(mergedRanges = Set(rng("A3:B3")))
    assert(SheetEdits.mergeRefusal(grown, rng("B2:B3")).isDefined)
    val edit = Edit.Merge(Area(None, rng("A1:A3")))
    assertEquals(Edit.lower(edit, cse), None)
    assertEquals(Patch.toEdits(Patch.Merge(rng("A1:A3")), cse), None)
    Edit.applyAll(Workbook(cse), Vector(edit), Scope.none) match
      case Left(XLError.EditFailed(1, "merge", XLError.InvalidRange(_, reason))) =>
        assert(reason.contains("array formula at A2:A3"), reason)
      case other => fail(s"expected an InvalidRange refusal, got $other")
    // a data table too
    val dt = Sheet(name).put(
      a1("B2"),
      CellValue.Formula(
        "TABLE(,A1)",
        None,
        FormulaKind.DataTable(rng("B2:B3"), false, false, None, Some(a1("A1")))
      )
    )
    assert(SheetEdits.mergeRefusal(dt, rng("A2:B2")).exists(_.message.contains("data table")))
    // a dynamic spill is not a refusal (it is blocked: #SPILL!)
    val dyn = Sheet(name)
      .put(
        a1("C1"),
        CellValue.Formula("SEQUENCE(3)", Some(num(1)), FormulaKind.dynamicArray(rng("C1:C3")))
      )
    assertEquals(SheetEdits.mergeRefusal(dyn, rng("B2:C3")), None)
  }

  test("GH-935: merging a whole column walks stored cells, not positions") {
    val s = Sheet(name).put(a1("A1"), num(1)).put(a1("A1000000"), num(2))
    val merged = s.merge(rng("A1:A1048576"))
    assertEquals(merged(a1("A1")).value, num(1))
    assertEquals(merged.cells.get(a1("A1000000")), None)
  }
