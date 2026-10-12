package com.tjclp.xl.formula

import java.time.{LocalDate, LocalDateTime}

import munit.FunSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cells.{CellValue, FormulaKind}
import com.tjclp.xl.error.XLError
import com.tjclp.xl.formula.eval.DynamicArrayAuthoring
import com.tjclp.xl.formula.eval.DynamicArrayAuthoring.{Refusal, SpillBlocked, SpillBlocker}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.numfmt.NumFmt
import com.tjclp.xl.tables.{TableColumn, TableSpec}
import com.tjclp.xl.workbooks.Workbook

/**
 * GH-714: `DynamicArrayAuthoring.author` stores what Excel 365 stores for a typed formula — a
 * dynamic-array anchor whose `ref` is the spill extent and whose spill cells are constants — and
 * refuses (writing nothing) a spill Excel would show as `#SPILL!`.
 */
class DynamicArrayAuthoringSpec extends FunSuite:

  private val S = SheetName.unsafe("S")
  private val clock = Clock.fixed(LocalDate.of(2026, 1, 2), LocalDateTime.of(2026, 1, 2, 3, 4, 5))
  private def num(n: Int): CellValue = CellValue.Number(BigDecimal(n))
  private def range(a1: String): CellRange = CellRange.parse(a1).fold(fail(_), identity)

  /** B1:B3 = 3, 1, 2 and C1:C3 = 10, 20, 30. */
  private val base: Sheet =
    Sheet(S)
      .put(ref"B1", num(3))
      .put(ref"B2", num(1))
      .put(ref"B3", num(2))
      .put(ref"C1", num(10))
      .put(ref"C2", num(20))
      .put(ref"C3", num(30))

  private def author(sheet: Sheet, at: ARef, formula: String) =
    DynamicArrayAuthoring.author(Workbook(Vector(sheet)), sheet, at, formula, clock)

  private def authored(sheet: Sheet, at: ARef, formula: String): DynamicArrayAuthoring.Authored =
    author(sheet, at, formula).fold(e => fail(e.message), identity)

  test("SORT spills: anchor record, constant children, extent and touched cells") {
    val a = authored(base, ref"A1", "=SORT(B1:B3)")
    assertEquals(a.extent, range("A1:A3"))
    assertEquals(
      a.sheet(ref"A1").value,
      CellValue.Formula("SORT(B1:B3)", Some(num(1)), FormulaKind.dynamicArray(range("A1:A3")))
    )
    assertEquals(a.sheet(ref"A2").value, num(2))
    assertEquals(a.sheet(ref"A3").value, num(3))
    assertEquals(a.touched, Set(ref"A1", ref"A2", ref"A3"))
    assert(a.cached)
  }

  test("a scalar array formula is a 1x1 dynamic record holding the array value") {
    // no implicit intersection: the product sums every row, in any row
    val a = authored(base, ref"D2", "=SUM(B1:B3*C1:C3)")
    assertEquals(a.extent, range("D2:D2"))
    assertEquals(
      a.sheet(ref"D2").value,
      CellValue.Formula("SUM(B1:B3*C1:C3)", Some(num(110)), FormulaKind.dynamicArray(range("D2")))
    )
  }

  test("an empty element spills as 0") {
    val sheet = base.put(ref"E2", CellValue.Empty)
    val a = authored(sheet.put(ref"E1", num(5)).put(ref"E3", num(6)), ref"G1", "=E1:E3")
    assertEquals(a.sheet(ref"G2").value, num(0))
  }

  private def assertBlocked(sheet: Sheet, at: ARef, formula: String, mentions: String*): Unit =
    author(sheet, at, formula) match
      case Left(err) =>
        mentions.foreach(m => assert(err.message.contains(m), err.message))
        assert(err.message.contains("#SPILL!"), err.message)
      case Right(a) => fail(s"expected a refusal, authored ${a.extent}")

  test("a spill into data, an empty string, a merge, another array or a table is refused") {
    assertBlocked(base.put(ref"A3", num(9)), ref"A1", "=SORT(B1:B3)", "A1:A3", "A3")
    assertBlocked(base.put(ref"A2", CellValue.Text("")), ref"A1", "=SORT(B1:B3)", "A2")
    assertBlocked(base.merge(range("A2:A4")), ref"A1", "=SORT(B1:B3)", "merged cells A2:A4")
    // a CSE record anchored outside the extent whose range reaches into it (F3 itself is empty)
    val cse = base.put(
      ref"E3",
      CellValue.Formula("B1:B3*2", None, FormulaKind.ArrayFormula(range("E3:F3")))
    )
    assertBlocked(cse, ref"F1", "=SORT(B1:B3)", "array formula at E3 (E3:F3)")
    val table = TableSpec(
      name = "T1",
      displayName = "T1",
      range = range("H2:I4"),
      columns = Vector(TableColumn(1, "a"), TableColumn(2, "b"))
    )
    assertBlocked(base.copy(tables = Map("T1" -> table)), ref"H1", "=SORT(B1:B3)", "table T1")
  }

  test("a spill past the grid edge is refused") {
    val at = ARef.from0(0, 1048575)
    assertBlocked(base, at, "=SORT(B1:B3)", "XFD1048576")
  }

  test("a refusal lists at most ten blockers, then the rest as a count") {
    val crowded = (1 to 15).foldLeft(base)((s, i) => s.put(ARef.from0(10, i), num(i)))
    author(crowded.put(ref"A1", num(1)), ref"K1", "=SEQUENCE(16)") match
      case Left(err) => assert(err.message.contains("(+5 more)"), err.message)
      case Right(_) => fail("expected a refusal")
  }

  test("a style-only cell does not block, and keeps its style") {
    val styled = base.withCellStyle(ref"A2", CellStyle.default.withNumFmt(NumFmt.Percent))
    val a = authored(styled, ref"A1", "=SORT(B1:B3)")
    assertEquals(a.sheet(ref"A2").value, num(2))
    assertEquals(a.sheet.getCellStyle(ref"A2").map(_.numFmt), Some(NumFmt.Percent))
  }

  test("re-authoring a shrinking spill clears the old children") {
    val first = authored(base, ref"A1", "=SORT(B1:B3)").sheet
    val second = authored(first, ref"A1", "=SORT(B1:B2)")
    assertEquals(second.extent, range("A1:A2"))
    assertEquals(second.sheet.cells.get(ref"A3").map(_.value), None)
    assert(second.touched.contains(ref"A3"), second.touched.toString)
  }

  test("a formula that cannot be evaluated is an uncached 1x1 dynamic record") {
    // an undefined name is a host failure, not an error value
    val a = authored(base, ref"A1", "=NoSuchName*2")
    assertEquals(a.extent, range("A1"))
    assert(!a.cached)
    assertEquals(
      a.sheet(ref"A1").value,
      CellValue.Formula("NoSuchName*2", None, FormulaKind.dynamicArray(range("A1")))
    )
  }

  test("an unparseable formula is refused") {
    author(base, ref"A1", "=SORT(B1:") match
      case Left(Refusal.Invalid(err)) => assertEquals(err.code, "FORMULA_ERROR")
      case other => fail(s"expected an invalid-formula refusal, got $other")
  }

  // ---------------------------------------------------------------------------------------------
  // The typed refusal: one blocker kind each, message and hint derived from it

  private def blockedBy(sheet: Sheet, at: ARef, formula: String): SpillBlocked =
    author(sheet, at, formula) match
      case Left(Refusal.Blocked(spill)) => spill
      case other => fail(s"expected a blocked spill, got $other")

  test("each blocker kind is typed, and its hint names the remedy for that kind") {
    val occupied = blockedBy(base.put(ref"A3", num(9)), ref"A1", "=SORT(B1:B3)")
    assertEquals(occupied.blocker, SpillBlocker.Occupied(Vector(ref"A3")))
    assertEquals(occupied.extent, Some(range("A1:A3")))
    assertEquals(occupied.hint, "clear A3 or choose another anchor")

    val merged = blockedBy(base.merge(range("A2:B4")), ref"A1", "=SORT(B1:B3)")
    assertEquals(merged.blocker, SpillBlocker.Merged(Vector(range("A2:B4"))))
    assertEquals(merged.hint, "unmerge A2:B4 or choose another anchor")
    assert(!merged.hint.contains("clear"), merged.hint)
    assert(merged.message.contains("unmerge them or choose another anchor"), merged.message)

    val cse = base.put(
      ref"E3",
      CellValue.Formula("B1:B3*2", None, FormulaKind.ArrayFormula(range("E3:F3")))
    )
    val array = blockedBy(cse, ref"F1", "=SORT(B1:B3)")
    assertEquals(array.blocker, SpillBlocker.ArrayRecord(Vector(ref"E3" -> range("E3:F3"))))
    assertEquals(array.hint, "clear or move the array formula at E3, or choose another anchor")

    val spec = TableSpec(
      name = "T1",
      displayName = "T1",
      range = range("H2:I4"),
      columns = Vector(TableColumn(1, "a"), TableColumn(2, "b"))
    )
    val table = blockedBy(base.copy(tables = Map("T1" -> spec)), ref"H1", "=SORT(B1:B3)")
    assertEquals(table.blocker, SpillBlocker.Table(Vector("T1" -> range("H2:I4"))))
    assertEquals(
      table.hint,
      "choose an anchor whose spill misses table T1, or convert the table to a range"
    )

    val offGrid = blockedBy(base, ARef.from0(0, 1048575), "=SORT(B1:B3)")
    assertEquals(offGrid.blocker, SpillBlocker.OffGrid(3, 1))
    assertEquals(offGrid.extent, None)
    assertEquals(
      offGrid.hint,
      "choose an anchor with 3 row(s) and 1 column(s) of room before XFD1048576"
    )
    assert(offGrid.message.contains("XFD1048576"), offGrid.message)

    // every message keeps Excel's word for it; the domain error carries the formula and message
    Vector(occupied, merged, array, table, offGrid).foreach { b =>
      assert(b.message.contains("(Excel shows #SPILL!)"), b.message)
      assertEquals(b.toXLError, XLError.FormulaError("=SORT(B1:B3)", b.message))
    }
  }

  // ---------------------------------------------------------------------------------------------
  // releaseForWrite: a plain write over a dynamic array, as Excel applies it

  private def sorted: Sheet = authored(base, ref"A1", "=SORT(B1:B3)").sheet

  private def release(sheet: Sheet, targets: ARef*) =
    DynamicArrayAuthoring.releaseForWrite(sheet, targets.toSet)

  test("a plain write over the anchor releases its spill: the children are cleared") {
    val (released, cleared) = release(sorted, ref"A1").fold(b => fail(b.message), identity)
    assertEquals(cleared, Set(ref"A2", ref"A3"))
    assertEquals(released.cells.get(ref"A2"), None)
    assertEquals(released.cells.get(ref"A3"), None)
    // the anchor itself is the caller's to overwrite
    assertEquals(released(ref"A1").value, sorted(ref"A1").value)
  }

  test("a write covering the anchor and its spill releases it too") {
    assertEquals(release(sorted, ref"A1", ref"A2", ref"A3").map(_._2), Right(Set(ref"A2", ref"A3")))
  }

  test("a write into a spill cell that leaves the anchor is refused as Overwritten") {
    release(sorted, ref"A2", ref"D9") match
      case Left(b) =>
        assertEquals(b.blocker, SpillBlocker.Overwritten(Vector(ref"A2")))
        assertEquals(b.anchor, ref"A1")
        assertEquals(b.extent, Some(range("A1:A3")))
        assertEquals(b.formula, "=SORT(B1:B3)")
        assert(b.message.contains("A2: it is in the spill range A1:A3"), b.message)
        assert(b.message.contains("(Excel shows #SPILL!)"), b.message)
        assertEquals(
          b.hint,
          "overwrite or clear A1 first (its spill goes with it), or write outside A1:A3"
        )
      case Right(_) => fail("a write into a live spill must be refused")
  }

  test("a sheet without dynamic arrays, a 1x1 record and a CSE extent are left alone") {
    assertEquals(release(base, ref"B2"), Right((base, Set.empty[ARef])))
    val single = authored(base, ref"D2", "=SUM(B1:B3*C1:C3)").sheet
    assertEquals(release(single, ref"D2"), Right((single, Set.empty[ARef])))
    val cse = base
      .put(ref"E1", CellValue.Formula("B1:B3*2", None, FormulaKind.ArrayFormula(range("E1:E3"))))
      .put(ref"E2", num(2))
    assertEquals(release(cse, ref"E1", ref"E2"), Right((cse, Set.empty[ARef])))
  }

  test("a spill of the array-cell guard or more is refused before any cell is written") {
    // the guard is ArrayResult.MaxCells (one full column); a small limit shows the rule cheaply
    DynamicArrayAuthoring.author(
      Workbook(Vector(base)),
      base,
      ref"A1",
      "=SORT(B1:B3)",
      clock,
      3
    ) match
      case Left(Refusal.Blocked(b)) =>
        assertEquals(b.blocker, SpillBlocker.TooLarge(3, 1, 3L))
        assertEquals(b.extent, None)
        assert(b.message.contains("a 3x1 spill from A1 is 3 cells"), b.message)
        assert(b.message.contains("fewer than 3 cells"), b.message)
        assert(b.hint.contains("A1:A1000, not A:A"), b.hint)
      case other => fail(s"expected TooLarge, got $other")
    assert(
      DynamicArrayAuthoring
        .author(Workbook(Vector(base)), base, ref"A1", "=SORT(B1:B3)", clock, 4)
        .isRight
    )
    assertEquals(com.tjclp.xl.formula.eval.ArrayResult.MaxCells, 1048576L)
  }

  test("a contents clear into a live spill is refused as Cleared, over the anchor it releases") {
    release(sorted, ref"A2") match
      case Left(_) => ()
      case Right(_) => fail("a write into a live spill must be refused")
    DynamicArrayAuthoring.releaseForWrite(sorted, Set(ref"A3"), clearing = true) match
      case Left(b) =>
        assertEquals(b.blocker, SpillBlocker.Cleared(Vector(ref"A3")))
        assert(b.message.contains("clearing them alone would change nothing"), b.message)
        assertEquals(b.hint, "clear A1 to remove the whole spill, or clear outside A1:A3")
      case Right(_) => fail("a clear of a live spill's cells must be refused")
    assertEquals(
      DynamicArrayAuthoring.releaseForWrite(sorted, Set(ref"A1"), clearing = true).map(_._2),
      Right(Set(ref"A2", ref"A3"))
    )
  }

  test("spillsOf lists multi-cell anchors; releaseRanges agrees with releaseForWrite") {
    import DynamicArrayAuthoring.Spill
    val spills = DynamicArrayAuthoring.spillsOf(sorted)
    assertEquals(spills, Vector(Spill(ref"A1", "SORT(B1:B3)", range("A1:A3"))))
    // a range over the anchor and its spill releases, without expanding the range
    assertEquals(
      DynamicArrayAuthoring
        .releaseRanges(sorted, spills, Vector(range("A1:A1048576")), false)
        .map(_._2),
      release(sorted, ref"A1").map(_._2)
    )
    // a range into the spill alone is refused, naming exactly the overlap
    DynamicArrayAuthoring.releaseRanges(
      sorted,
      spills,
      Vector(range("A2:D2"), range("A3")),
      false
    ) match
      case Left(b) => assertEquals(b.blocker, SpillBlocker.Overwritten(Vector(ref"A2", ref"A3")))
      case Right(_) => fail("a write into a live spill must be refused")
    // no known spill: nothing to scan or release
    assertEquals(
      DynamicArrayAuthoring.releaseRanges(sorted, Vector.empty, Vector(range("A2")), false),
      Right((sorted, Set.empty[ARef]))
    )
  }

  test("author returns the expression it parsed") {
    assertEquals(
      authored(base, ref"A1", "=SORT(B1:B3)").expr,
      com.tjclp.xl.formula.parser.FormulaParser
        .parse("=SORT(B1:B3)")
        .fold(e => fail(e.toString), identity)
    )
  }

  test("withNumFmt formats the whole extent and keeps each cell's other style") {
    val bold = CellStyle.default.withFont(CellStyle.default.font.withBold(true))
    val a = authored(base.withCellStyle(ref"A3", bold), ref"A1", "=SORT(B1:B3)")
      .withNumFmt(NumFmt.Percent)
    Vector(ref"A1", ref"A2", ref"A3").foreach { r =>
      assertEquals(a.sheet.getCellStyle(r).map(_.numFmt), Some(NumFmt.Percent), r.toA1)
    }
    assertEquals(a.sheet.getCellStyle(ref"A3").map(_.font.bold), Some(true))
  }

  test("deterministic under a fixed clock") {
    assertEquals(authored(base, ref"A1", "=SORT(B1:B3)"), authored(base, ref"A1", "=SORT(B1:B3)"))
  }
