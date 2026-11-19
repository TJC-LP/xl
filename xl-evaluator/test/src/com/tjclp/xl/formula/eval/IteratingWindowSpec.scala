package com.tjclp.xl.formula.eval

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.Clock
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.{CalcPr, Workbook}
import munit.FunSuite

/**
 * #891: the window evaluation behind `view --eval` in a book whose calcPr iterates.
 *
 * Such a book's cycle members keep their caches (Excel shows the converged values the file holds)
 * and stay out of the window's evaluation order. That rests on every member holding a cache. A
 * member without one (a library-authored or `--no-recalc` book) would be evaluated on demand by
 * whichever reader came first, before its own precedents' slots, from stale caches. The window of
 * such a book is what an iterative recalculation computes.
 */
class IteratingWindowSpec extends FunSuite:
  private val clock = Clock.fixedDate(java.time.LocalDate.of(2026, 1, 2))
  private def at(a1: String): ARef = ARef.parse(a1).fold(e => fail(e), identity)
  private def n(v: Int): CellValue = CellValue.Number(BigDecimal(v))
  private def f(t: String): CellValue = CellValue.Formula(t, None)
  private def stale(t: String, v: Int): CellValue = CellValue.Formula(t, Some(n(v)))
  private def sheet(name: String, cells: (String, CellValue)*): Sheet =
    cells.foldLeft(Sheet(SheetName.unsafe(name)))((s, kv) => s.put(at(kv._1), kv._2))
  private def iterating(wb: Workbook): Workbook =
    wb.copy(metadata = wb.metadata.copy(calcPr = Some(CalcPr(iterativeCalculation = true))))
  private val range = CellRange.parse("A1:F9").fold(e => fail(e), identity)
  private def recalculated(wb: Workbook): Map[ARef, CellValue] =
    val full = wb.recalculate(clock, IterativeCalc.fromCalcPr(CalcPr(iterativeCalculation = true)))
    assertEquals(full.errors, Vector.empty)
    full.evaluated.getOrElse(wb.sheets(0).name, Map.empty).filter((ref, _) => range.contains(ref))

  // S!D1 <-> T!E4 is the cycle. S!D2 is not in it: it feeds it. U!C3 and U!C4 hold stale caches
  // (6 and 9); their constant makes them 2 and 4. S!F1 reads the member and sorts first.
  private def book(e4: CellValue, threshold: Int, reader: CellValue): Workbook =
    iterating(
      Workbook(
        sheet(
          "S",
          "D1" -> stale("COUNTIF(T!E4:E4,\">1\")", 0),
          "D2" -> f("AVERAGE(U!C3:C4)"),
          "F1" -> reader
        )
      ).put(sheet("T", "E4" -> e4))
        .put(sheet("U", "A1" -> n(2), "C3" -> stale("A1*1", 6), "C4" -> stale("A1*2", 9)))
    )
  private def member(threshold: Int): String = s"COUNTIF(S!D1:D2,\">$threshold\")"

  test("an aggregate read early through an uncached cycle member is not kept for its own cell") {
    // the member's on-demand evaluation computed AVERAGE(U!C3:C4) over the stale 6 and 9, and the
    // pass's aggregate memo handed that 7.5 back when S!D2 was evaluated in its turn
    val wb = book(f(member(1)), 1, f("T!E4+0"))
    val shown = wb.sheets(0).evaluateForRangePerCell(range, clock, Some(wb))
    assert(shown.isClean, shown.summary)
    assertEquals(shown.values.get(at("D2")), Some(n(3)))
    assertEquals(shown.values, recalculated(wb))
  }

  test("an uncached cycle member and its reader are computed from fresh precedents") {
    // stale S!D2 is 7.5 (above 5), fresh is 3 (below): the member's count differs between them
    val wb = book(f(member(5)), 5, f("T!E4+0"))
    val shown = wb.sheets(0).evaluateForRangePerCell(range, clock, Some(wb))
    assert(shown.isClean, shown.summary)
    assertEquals(shown.values.get(at("F1")), Some(n(0)))
    assertEquals(shown.values, recalculated(wb))
  }

  test("an uncached cycle member read only through INDIRECT is computed the same way") {
    val wb = book(f(member(5)), 5, f("INDIRECT(\"T!E4\")+0"))
    val shown = wb.sheets(0).evaluateForRangePerCell(range, clock, Some(wb))
    assert(shown.isClean, shown.summary)
    assertEquals(shown.values.get(at("F1")), Some(n(0)))
    assertEquals(shown.values, recalculated(wb))
  }

  test("a window over an uncached cycle converges as the recalculation does") {
    // was: "Circular reference … read again while its own evaluation is in progress"
    val wb = iterating(
      Workbook(sheet("S", "A1" -> f("B1*0.5+1"), "B1" -> f("A1*0.5"), "C1" -> f("A1+1")))
    )
    val shown = wb.sheets(0).evaluateForRangePerCell(range, clock, Some(wb))
    assert(shown.isClean, shown.summary)
    assertEquals(shown.values, recalculated(wb))
    assertEquals(shown.values.keySet, Set(at("A1"), at("B1"), at("C1")))
  }

  test("a cached cycle keeps its values: the file's converged caches are what the window shows") {
    // every member holds a cache, so nothing iterates and the stale-looking 0 stays (Excel parity)
    val wb = book(stale(member(1), 0), 1, f("T!E4+0"))
    val shown = wb.sheets(0).evaluateForRangePerCell(range, clock, Some(wb))
    assert(shown.isClean, shown.summary)
    assertEquals(shown.values.get(at("D2")), Some(n(3)))
    assertEquals(shown.values.get(at("F1")), Some(n(0)))
    assertEquals(shown.values.get(at("D1")), None)
  }

  test("a failure in the iterated book is reported, and what reads it is blocked") {
    val wb = iterating(
      Workbook(
        sheet(
          "S",
          "A1" -> f("B1*0.5+1"),
          "B1" -> f("A1*0.5"),
          "C1" -> f("Missing!A1+1"),
          "D1" -> f("C1+1"),
          "E1" -> f("A1+1")
        )
      )
    )
    val shown = wb.sheets(0).evaluateForRangePerCell(range, clock, Some(wb))
    assertEquals(shown.failures.map(_.ref), Vector(at("C1")))
    assertEquals(shown.blocked, Vector(at("D1")))
    assertEquals(shown.values.keySet, Set(at("A1"), at("B1"), at("E1")))
  }
