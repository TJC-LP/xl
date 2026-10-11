package com.tjclp.xl.formula

import java.time.{LocalDate, LocalDateTime}

import munit.FunSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cells.{CellValue, FormulaKind}
import com.tjclp.xl.formula.eval.DynamicArrayAuthoring
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
    assert(author(base, ref"A1", "=SORT(B1:").isLeft)
  }

  test("deterministic under a fixed clock") {
    assertEquals(authored(base, ref"A1", "=SORT(B1:B3)"), authored(base, ref"A1", "=SORT(B1:B3)"))
  }
