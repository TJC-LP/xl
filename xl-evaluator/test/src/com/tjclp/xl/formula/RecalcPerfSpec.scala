package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * GH-346: recalculate() and formula eval on recursive multi-branch cross-sheet chains (the LBO
 * debt-schedule shape) must run in time linear in the cell count, not the path count.
 *
 * Pre-fix, an uncached formula reference was recursively re-evaluated once per PATH (no
 * memoization), and whole-workbook recalculation ordered cells per sheet (not globally), so
 * cross-sheet precedents were always uncached when hit. On this shape — each period's beginning
 * balance read by three formulas, chained backwards period over period — evaluation cost grew ~3^n:
 * at n=24 that is ~10^11 evaluations, hours of CPU. Post-fix both paths complete in milliseconds;
 * the generous wall-clock budgets here only catch a regression to exponential.
 */
class RecalcPerfSpec extends FunSuite:

  private val Periods = 24
  private val BudgetMs = 30000L

  private def num(bd: BigDecimal): CellValue = CellValue.Number(bd)
  private def formula(expr: String): CellValue = CellValue.Formula(expr, None)

  // 0-based (col, row) helpers for the schedule columns; rows are periods (row i = period i+1)
  private def at(col: Int, period: Int): ARef = ARef.from0(col, period - 1)
  private val ColB = 1 // Debt: beginning balance / Forecast: EBITDA
  private val ColD = 3 // Debt: interest
  private val ColF = 5 // Debt: cash available for sweep
  private val ColH = 7 // Debt: sweep
  private val ColJ = 9 // Debt: ending balance

  /**
   * Debt-schedule shape (every formula uncached):
   *   - Forecast!B{i}: EBITDA, 100 growing 10%/period
   *   - Debt!B{i}: beginning balance (period 1: opening 1000; else prior ending =J{i-1})
   *   - Debt!D{i}: interest on beginning balance (=B{i}*0.08)
   *   - Debt!F{i}: cash available (=Forecast!B{i}-D{i}) — cross-sheet
   *   - Debt!H{i}: sweep (=MIN(F{i},B{i}))
   *   - Debt!J{i}: ending balance (=B{i}-H{i})
   *   - Returns!B1: terminal ending + terminal EBITDA — cross-sheet
   */
  private def buildWorkbook(n: Int): Workbook =
    val forecast = (2 to n).foldLeft(
      Sheet(SheetName.unsafe("Forecast")).put(at(ColB, 1), num(BigDecimal(100)))
    ) { (s, i) =>
      s.put(at(ColB, i), formula(s"=B${i - 1}*1.1"))
    }
    val debt = (1 to n).foldLeft(Sheet(SheetName.unsafe("Debt"))) { (s, i) =>
      val withBeginning =
        if i == 1 then s.put(at(ColB, 1), num(BigDecimal(1000)))
        else s.put(at(ColB, i), formula(s"=J${i - 1}"))
      withBeginning
        .put(at(ColD, i), formula(s"=B$i*0.08"))
        .put(at(ColF, i), formula(s"=Forecast!B$i-D$i"))
        .put(at(ColH, i), formula(s"=MIN(F$i,B$i)"))
        .put(at(ColJ, i), formula(s"=B$i-H$i"))
    }
    val returns = Sheet(SheetName.unsafe("Returns"))
      .put(ARef.from0(1, 0), formula(s"=Debt!J$n+Forecast!B$n"))
    Workbook(forecast, debt, returns)

  /** Reference values computed with the same BigDecimal operations the evaluator uses. */
  private case class Terminals(ebitda: BigDecimal, ending: BigDecimal)
  private def expectedTerminals(n: Int): Terminals =
    val init = (BigDecimal(100), BigDecimal(1000))
    val (ebitda, ending) = (1 to n).foldLeft(init) { case ((e, beg), i) =>
      val ebitda = if i == 1 then e else e * BigDecimal("1.1")
      val interest = beg * BigDecimal("0.08")
      val cash = ebitda - interest
      val sweep = cash.min(beg)
      val ending = beg - sweep
      (ebitda, ending)
    }
    // The fold threads (current ebitda, ending balance); ending of period i is beginning of i+1
    Terminals(ebitda, ending)

  private def cached(wb: Workbook, sheetName: String, ref: ARef): Option[CellValue] =
    wb.sheets
      .find(_.name.value == sheetName)
      .flatMap(_.cells.get(ref))
      .map(_.value)
      .collect { case CellValue.Formula(_, Some(v), _) => v }

  test("GH-346: whole-workbook recalculate() on a recursive schedule is linear and correct"):
    val wb = buildWorkbook(Periods)
    val t0 = System.nanoTime()
    val result = wb.recalculate()
    val elapsedMs = (System.nanoTime() - t0) / 1000000L
    assert(result.isClean, s"expected clean, got: ${result.errors.take(5).map(_.render)}")
    val Terminals(ebitda, ending) = expectedTerminals(Periods)
    assertEquals(cached(result.workbook, "Debt", at(ColJ, Periods)), Some(num(ending)))
    assertEquals(cached(result.workbook, "Returns", ARef.from0(1, 0)), Some(num(ending + ebitda)))
    assert(elapsedMs < BudgetMs, s"recalculate took ${elapsedMs}ms (budget ${BudgetMs}ms)")

  test("GH-346: single-shot eval of a deep uncached cross-sheet cell is memoized, not per-path"):
    val wb = buildWorkbook(Periods)
    val t0 = System.nanoTime()
    val evaluated = wb.evaluateFormula(s"=Debt!J$Periods+0", "Returns")
    val elapsedMs = (System.nanoTime() - t0) / 1000000L
    assertEquals(evaluated, Right(num(expectedTerminals(Periods).ending)))
    assert(elapsedMs < BudgetMs, s"eval took ${elapsedMs}ms (budget ${BudgetMs}ms)")

  test("GH-346: cross-sheet aggregate over uncached formula cells computes their values"):
    // SUM over ANOTHER sheet's formula cells: with per-sheet ordering the aggregate read the
    // original uncached snapshot (values silently skipped); the global order must guarantee
    // computed values regardless of sheet position in the workbook.
    val b1 = ARef.from0(1, 0)
    val b2 = ARef.from0(1, 1)
    val b3 = ARef.from0(1, 2)
    val data = Sheet(SheetName.unsafe("Data"))
      .put(b1, num(BigDecimal(10)))
      .put(b2, formula("=B1*2"))
      .put(b3, formula("=B2+5"))
    val summary = Sheet(SheetName.unsafe("Summary")).put(b1, formula("=SUM(Data!B1:B3)"))
    // Summary placed FIRST so a sheet-ordered pass cannot accidentally see Data computed
    val result = Workbook(summary, data).recalculate()
    assert(result.isClean, s"expected clean, got: ${result.errors.map(_.render)}")
    assertEquals(cached(result.workbook, "Summary", b1), Some(num(BigDecimal(55))))
