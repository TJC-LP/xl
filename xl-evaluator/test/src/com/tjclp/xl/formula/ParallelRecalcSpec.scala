package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.eval.RecalcResult
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * GH-520: `recalculateParallel(n)` must be observationally IDENTICAL to `recalculate()` — same
 * workbook (caches included), same evaluated map, same error vector in the same order — for every
 * dependence the static graph can express. The wave partition guarantees it structurally (two
 * same-depth cells have no path between them, so neither can observe the other's write); these
 * gates pin the equality on the shapes most likely to break it: deep chains (waves below the
 * sequential cutoff), wide independent regions (the actually-parallel path), error-bearing books
 * (error ORDER is part of the contract), cycles (circular + blocked reporting precedes the pass),
 * and dynamic INDIRECT cells (excluded from waves, sequential evaluate-last).
 */
class ParallelRecalcSpec extends FunSuite:

  private def num(d: Double): CellValue = CellValue.Number(BigDecimal(d))
  private def formula(expr: String): CellValue = CellValue.Formula(expr, None)

  private def assertSameResult(sequential: RecalcResult, parallel: RecalcResult): Unit =
    assertEquals(parallel.errors, sequential.errors)
    assertEquals(parallel.evaluated, sequential.evaluated)
    assertEquals(parallel.workbook, sequential.workbook)
    assertEquals(parallel.converged, sequential.converged)
    assertEquals(parallel.cycles, sequential.cycles)

  /**
   * 40 independent columns × 30 chained rows + a cross-sheet SUM roll-up: wide waves (40 ≥ cutoff).
   */
  private def wideGrid: Workbook =
    val cols = 40
    val rows = 30
    val data = (0 until cols).foldLeft(Sheet(SheetName.unsafe("Data"))) { (s0, c) =>
      (0 until rows).foldLeft(s0.put(ARef.from0(c, 0), num(100.0 + c))) { (s, r) =>
        if r == 0 then s
        else
          val colName = ARef.from0(c, r).toA1.takeWhile(_.isLetter)
          s.put(ARef.from0(c, r), formula(s"=$colName$r*1.01+SUM(Drivers!$$A$$1:$$A$$5)*0.001"))
      }
    }
    val drivers = (0 until 5).foldLeft(Sheet(SheetName.unsafe("Drivers"))) { (s, i) =>
      s.put(ARef.from0(0, i), num((i + 1) * 1.5))
    }
    val summary = Sheet(SheetName.unsafe("Summary"))
      .put(ARef.from0(0, 0), formula(s"=SUM(Data!A$rows:AN$rows)"))
    Workbook(drivers, data, summary)

  test("GH-520: wide independent grid — parallel(8) equals sequential, element for element"):
    val wb = wideGrid
    assertSameResult(wb.recalculate(), wb.recalculateParallel(8))

  test("GH-520: deep chain (waves below the cutoff) — parallel equals sequential"):
    val n = 200
    val chain =
      (2 to n).foldLeft(Sheet(SheetName.unsafe("Chain")).put(ARef.from0(0, 0), num(1.0))) {
        (s, i) => s.put(ARef.from0(0, i - 1), formula(s"=A${i - 1}*1.1+1"))
      }
    val wb = Workbook(chain)
    assertSameResult(wb.recalculate(), wb.recalculateParallel(4))

  test("GH-520: error-bearing book — errors, Excel errors and their ORDER are identical"):
    val cols = 20
    val s0 = Sheet(SheetName.unsafe("Errs"))
    // Row 0: constants; row 1: a wide wave where every third cell divides by zero and every
    // seventh calls an unknown function — both error planes, interleaved with healthy cells.
    val sheet = (0 until cols).foldLeft(s0) { (s, c) =>
      val colName = ARef.from0(c, 1).toA1.takeWhile(_.isLetter)
      val expr =
        if c % 3 == 0 then s"=$colName" + "1/0"
        else if c % 7 == 0 then s"=NOSUCHFN($colName" + "1)"
        else s"=$colName" + "1*2"
      s.put(ARef.from0(c, 0), num(c.toDouble)).put(ARef.from0(c, 1), formula(expr))
    }
    val wb = Workbook(sheet)
    val seq = wb.recalculate()
    val par = wb.recalculateParallel(8)
    assertSameResult(seq, par)
    assertEquals(par.excelErrors, seq.excelErrors)

  test("GH-520: cyclic core + blocked dependents — parallel equals sequential"):
    val a1 = ARef.from0(0, 0)
    val b1 = ARef.from0(1, 0)
    val c1 = ARef.from0(2, 0)
    // A1 <-> B1 cycle, C1 blocked downstream, plus a wide healthy region.
    val base = Sheet(SheetName.unsafe("Cyc"))
      .put(a1, formula("=B1+1"))
      .put(b1, formula("=A1+1"))
      .put(c1, formula("=A1*2"))
    val sheet = (0 until 20).foldLeft(base) { (s, c) =>
      s.put(ARef.from0(c, 2), num(c.toDouble))
        .put(ARef.from0(c, 3), formula(s"=${ARef.from0(c, 3).toA1.takeWhile(_.isLetter)}3+1"))
    }
    val wb = Workbook(sheet)
    assertSameResult(wb.recalculate(), wb.recalculateParallel(8))

  test("GH-520: INDIRECT bucket stays sequential — parallel equals sequential"):
    val base = Sheet(SheetName.unsafe("Dyn"))
      .put(ARef.from0(0, 0), num(7.0))
      .put(ARef.from0(1, 0), formula("=INDIRECT(\"A1\")*3"))
      .put(ARef.from0(2, 0), formula("=B1+1")) // static dependent of a dynamic cell — bucketed
    val sheet = (0 until 20).foldLeft(base) { (s, c) =>
      s.put(ARef.from0(c, 2), num(c.toDouble))
        .put(ARef.from0(c, 3), formula(s"=${ARef.from0(c, 3).toA1.takeWhile(_.isLetter)}3*2"))
    }
    val wb = Workbook(sheet)
    assertSameResult(wb.recalculate(), wb.recalculateParallel(8))

  test("GH-520: repeated parallel runs are deterministic"):
    val wb = wideGrid
    val first = wb.recalculateParallel(8)
    (1 to 3).foreach { _ =>
      assertSameResult(first, wb.recalculateParallel(8))
    }

  test("GH-520: parallelism 1 and an absurd width both degrade safely"):
    val wb = wideGrid
    assertSameResult(wb.recalculate(), wb.recalculateParallel(1))
    assertSameResult(wb.recalculate(), wb.recalculateParallel(10000))
