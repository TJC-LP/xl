package com.tjclp.xl.formula

import java.time.LocalDateTime

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * GH-467: MATCH/XLOOKUP lookup-plane fixes.
 *
 * Three sub-bugs, one comparator plane: (a) blank cells "matched" every lookup value (the old
 * comparator returned 0 for any unhandled type pair), so MATCH over a range with blanks stopped at
 * the first blank and every later position looked shifted; (b) date-typed lookup values (DateTime
 * cells, DATE(...) results) hit the same fall-through and always matched position 1; (c) cell-ref
 * lookup values arrive as ExprValue.Cell and were never dereferenced, so MATCH returned 1 and
 * XLOOKUP returned #N/A while the equivalent literal worked.
 */
class LookupBlankDateCoercionSpec extends FunSuite:

  private def num(i: Int): CellValue = CellValue.Number(BigDecimal(i))

  private def sheetWith(cells: (ARef, CellValue)*): Sheet =
    cells.foldLeft(new Sheet(name = SheetName.unsafe("Test"))) { case (s, (ref, value)) =>
      s.put(ref, value)
    }

  // (a) A1:A5 = {10, _, 20, _, 30} — MATCH must index by RANGE position, blanks included
  private val blanky = sheetWith(
    ARef.from0(0, 0) -> num(10), // A1
    ARef.from0(0, 2) -> num(20), // A3 (A2 blank)
    ARef.from0(0, 4) -> num(30) // A5 (A4 blank)
  )

  test("GH-467: MATCH exact over blanks returns the range position (20 -> 3)") {
    assertEquals(
      blanky.evaluateFormula("=MATCH(20, A1:A5, 0)"),
      Right(CellValue.Number(BigDecimal(3)))
    )
  }

  test("GH-467: MATCH exact over blanks returns the range position (30 -> 5)") {
    assertEquals(
      blanky.evaluateFormula("=MATCH(30, A1:A5, 0)"),
      Right(CellValue.Number(BigDecimal(5)))
    )
  }

  test("GH-467: MATCH exact over blanks still misses cleanly (#N/A), blanks match nothing") {
    blanky.evaluateFormula("=MATCH(25, A1:A5, 0)") match
      case Left(error) => assert(error.toString.contains("#N/A"), s"expected #N/A, got $error")
      case other => fail(s"Expected #N/A error, got $other")
  }

  test("GH-467: XLOOKUP exact over blanks returns the positionally aligned value") {
    assertEquals(
      blanky.evaluateFormula("=XLOOKUP(30, A1:A5, A1:A5)"),
      Right(CellValue.Number(BigDecimal(30)))
    )
  }

  // (b) date-typed lookup values: B1:B3 = 2026-02-10 / 2026-03-15 / 2026-04-20, D1 = 2026-03-15
  private val dates = sheetWith(
    ARef.from0(1, 0) -> CellValue.DateTime(LocalDateTime.of(2026, 2, 10, 0, 0)), // B1
    ARef.from0(1, 1) -> CellValue.DateTime(LocalDateTime.of(2026, 3, 15, 0, 0)), // B2
    ARef.from0(1, 2) -> CellValue.DateTime(LocalDateTime.of(2026, 4, 20, 0, 0)), // B3
    ARef.from0(3, 0) -> CellValue.DateTime(LocalDateTime.of(2026, 3, 15, 0, 0)), // D1
    ARef.from0(6, 0) -> num(100), // G1
    ARef.from0(6, 1) -> num(200), // G2
    ARef.from0(6, 2) -> num(300) // G3
  )

  test("GH-467: MATCH with a date CELL-REF lookup value finds the right position") {
    assertEquals(
      dates.evaluateFormula("=MATCH(D1, B1:B3, 0)"),
      Right(CellValue.Number(BigDecimal(2)))
    )
  }

  test("GH-467: MATCH with a DATE(...) lookup value finds the right position") {
    assertEquals(
      dates.evaluateFormula("=MATCH(DATE(2026,4,20), B1:B3, 0)"),
      Right(CellValue.Number(BigDecimal(3)))
    )
  }

  test("GH-467: XLOOKUP with a DATE(...) lookup value matches DateTime cells by serial") {
    assertEquals(
      dates.evaluateFormula("=XLOOKUP(DATE(2026,3,15), B1:B3, G1:G3)"),
      Right(CellValue.Number(BigDecimal(200)))
    )
  }

  // (c) cell-ref lookup values: E1 = "k2", F1:F3 = keys, G1:G3 = values
  private val keyed = sheetWith(
    ARef.from0(4, 0) -> CellValue.Text("k2"), // E1
    ARef.from0(5, 0) -> CellValue.Text("k1"), // F1
    ARef.from0(5, 1) -> CellValue.Text("k2"), // F2
    ARef.from0(5, 2) -> CellValue.Text("k3"), // F3
    ARef.from0(6, 0) -> num(100), // G1
    ARef.from0(6, 1) -> num(200), // G2
    ARef.from0(6, 2) -> num(300) // G3
  )

  test("GH-467: MATCH with a cell-ref lookup value matches like the literal") {
    assertEquals(
      keyed.evaluateFormula("=MATCH(E1, F1:F3, 0)"),
      Right(CellValue.Number(BigDecimal(2)))
    )
  }

  test("GH-467: XLOOKUP with a cell-ref lookup value matches like the literal") {
    assertEquals(
      keyed.evaluateFormula("=XLOOKUP(E1, F1:F3, G1:G3)"),
      Right(CellValue.Number(BigDecimal(200)))
    )
  }

  test("GH-467: XLOOKUP cell-ref lookup value across sheets (exact field repro)") {
    val s1 = new Sheet(name = SheetName.unsafe("S1"))
      .put(ARef.from0(4, 0), CellValue.Text("k2")) // S1!E1
    val s2 = new Sheet(name = SheetName.unsafe("S2"))
      .put(ARef.from0(5, 0), CellValue.Text("k1")) // S2!F1
      .put(ARef.from0(5, 1), CellValue.Text("k2")) // S2!F2
      .put(ARef.from0(5, 2), CellValue.Text("k3")) // S2!F3
      .put(ARef.from0(6, 0), num(100)) // S2!G1
      .put(ARef.from0(6, 1), num(200)) // S2!G2
      .put(ARef.from0(6, 2), num(300)) // S2!G3
    val wb = Workbook(s1).put(s2)
    assertEquals(
      wb.evaluateFormula("=XLOOKUP(E1, S2!F1:F3, S2!G1:G3)", "S1"),
      Right(CellValue.Number(BigDecimal(200)))
    )
  }

  test("GH-467: literal lookup values keep working (sanity)") {
    assertEquals(
      keyed.evaluateFormula("=MATCH(\"k2\", F1:F3, 0)"),
      Right(CellValue.Number(BigDecimal(2)))
    )
    assertEquals(
      keyed.evaluateFormula("=XLOOKUP(\"k2\", F1:F3, G1:G3)"),
      Right(CellValue.Number(BigDecimal(200)))
    )
    assertEquals(
      keyed.evaluateFormula("=XLOOKUP(\"nope\", F1:F3, G1:G3)"),
      Right(CellValue.Error(CellError.NA))
    )
  }
