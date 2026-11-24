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
 * GH-935: merging keeps one value, as Excel's Merge Cells does: the upper-left cell's when it has
 * one, else the first non-empty cell's (row-major), which is moved into the upper-left cell. Every
 * other cell of the range is cleared the way Clear Contents clears it: the value, formula and
 * hyperlink go; the cell's own formatting and its comment stay. Reading a file whose merge already
 * hides values rewrites nothing: only the merge operation clears.
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

  // The five cases below were checked against Excel for Mac (`merge range`, alerts off).

  test("GH-935 (Excel): an empty anchor receives the only value of the range") {
    val s = Sheet(name).put(a1("B1"), CellValue.Text("onlyB1"))
    val merged = s.merge(rng("A1:B1"))
    assertEquals(merged(a1("A1")).value, CellValue.Text("onlyB1"))
    assertEquals(merged.cells.get(a1("B1")), None)
  }

  test("GH-935 (Excel): a non-empty anchor wins and the other value is cleared") {
    val s = Sheet(name).put(a1("A3"), CellValue.Text("a3")).put(a1("B3"), CellValue.Text("b3"))
    val merged = s.merge(rng("A3:B3"))
    assertEquals(merged(a1("A3")).value, CellValue.Text("a3"))
    assertEquals(merged.cells.get(a1("B3")), None)
  }

  test("GH-935 (Excel): the FIRST non-empty cell in row-major order is moved into the anchor") {
    val s = Sheet(name).put(a1("B5"), CellValue.Text("b5")).put(a1("A6"), CellValue.Text("a6"))
    val merged = s.merge(rng("A5:B6"))
    assertEquals(merged(a1("A5")).value, CellValue.Text("b5"))
    assert(Vector("B5", "A6", "B6").forall(r => merged.cells.get(a1(r)).isEmpty))
  }

  test("GH-935 (Excel): the moved cell keeps its formula verbatim and its format") {
    val s = Sheet(name)
      .put(a1("D1"), num(5))
      .put(a1("B1"), CellValue.Formula("D1*2", Some(num(10))))
      .withCellStyle(a1("B1"), CellStyle.default.withNumFmt(NumFmt.Percent))
      // the empty anchor's own format is replaced by the moved cell's
      .withCellStyle(a1("A1"), CellStyle.default.withNumFmt(NumFmt.Date))
    val merged = s.merge(rng("A1:B1"))
    assertEquals(merged(a1("A1")).value, CellValue.Formula("D1*2", Some(num(10))))
    assertEquals(merged.getCellStyle(a1("A1")).map(_.numFmt), Some(NumFmt.Percent))
    assertEquals(merged.cells.get(a1("B1")), None)
  }

  test("GH-935: a moved array formula carries its result range with it") {
    import com.tjclp.xl.cells.FormulaKind
    def dyn(r: String) =
      CellValue.Formula("SEQUENCE(1)", Some(num(1)), FormulaKind.dynamicArray(rng(r)))
    def cse(r: String) = CellValue.Formula("D1*2", Some(num(2)), FormulaKind.ArrayFormula(rng(r)))
    val a = Sheet(name).put(a1("B2"), dyn("B2:B2")).merge(rng("A1:B2"))
    assertEquals(a(a1("A1")).value, dyn("A1:A1"))
    val b = Sheet(name).put(a1("C3"), cse("C3:C3")).merge(rng("B2:C3"))
    assertEquals(b(a1("B2")).value, cse("B2:B2"))
  }

  test("GH-935: Sheet.merge cannot rewrite the moved cell's readers (no parser in xl-core)") {
    val s = Sheet(name).put(a1("B1"), num(7)).put(a1("E1"), CellValue.Formula("B1+1", None))
    val merged = s.merge(rng("A1:B1"))
    assertEquals(merged(a1("A1")).value, num(7))
    assertEquals(merged(a1("E1")).value, CellValue.Formula("B1+1", None))
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

  test("GH-935: merging a whole column walks stored cells, not positions") {
    val s = Sheet(name).put(a1("A1"), num(1)).put(a1("A1000000"), num(2))
    val merged = s.merge(rng("A1:A1048576"))
    assertEquals(merged(a1("A1")).value, num(1))
    assertEquals(merged.cells.get(a1("A1000000")), None)
  }
