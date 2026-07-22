package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.eval.IterativeCalc
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.{CalcPr, Workbook}
import munit.FunSuite

/**
 * GH-373b: opt-in bounded iterative recalculation for circular workbooks.
 *
 * `recalculate(IterativeCalc(maxIter, maxChange))` fixpoints cycle members with Jacobi iteration
 * (every member reads PREVIOUS-iteration values) instead of erroring: members seed to 0 (Excel's
 * uninitialized semantics), iterate until every |Δ| < maxChange or maxIter, and non-convergence
 * KEEPS the last values with NO error (Excel semantics — the deliberate inversion vs the default
 * path). Dependents of a cycle are NOT pre-marked blocked — they evaluate off the converged values.
 * The default `recalculate()` is byte-identical to before (opt-in only; the CalcPr metadata is
 * deliberately NOT auto-honored — bridge explicitly via
 * `wb.metadata.calcPr.filter(_.iterativeCalculation).map(IterativeCalc.fromCalcPr)`).
 */
class IterativeRecalcSpec extends FunSuite:

  private def num(i: Int): CellValue = CellValue.Number(BigDecimal(i))
  private def formula(expr: String): CellValue = CellValue.Formula(expr, None)
  private val a1 = ARef.from0(0, 0)
  private val b1 = ARef.from0(1, 0)
  private val c1 = ARef.from0(2, 0)
  private val d1 = ARef.from0(3, 0)

  private def cached(wb: Workbook, sheetName: String, ref: ARef): Option[CellValue] =
    wb.sheets
      .find(_.name.value == sheetName)
      .flatMap(_.cells.get(ref))
      .map(_.value)
      .collect { case CellValue.Formula(_, Some(v), _) => v }

  private def cachedNum(wb: Workbook, sheetName: String, ref: ARef): Option[BigDecimal] =
    cached(wb, sheetName, ref).collect { case CellValue.Number(n) => n }

  /** The canonical convergent pair: fixpoint B1 = 100/0.95, A1 = B1/10. */
  private def convergentPair: Workbook =
    val sheet = Sheet(SheetName.unsafe("S"))
      .put(a1, formula("=B1*0.1"))
      .put(b1, formula("=100+A1/2"))
    Workbook(sheet)

  private val fixpointB = BigDecimal(100) / BigDecimal("0.95") // 105.26315789...
  private val fixpointA = fixpointB / 10

  test("GH-373: convergent cycle fixpoints to the analytic solution within tolerance") {
    val result = convergentPair.recalculate(IterativeCalc(100, BigDecimal("0.001")))
    assert(result.isClean, s"expected clean, got: ${result.errors.map(_.render)}")
    val a = cachedNum(result.workbook, "S", a1).getOrElse(fail("A1 must cache"))
    val b = cachedNum(result.workbook, "S", b1).getOrElse(fail("B1 must cache"))
    assert((a - fixpointA).abs < BigDecimal("0.005"), s"A1=$a, expected ~$fixpointA")
    assert((b - fixpointB).abs < BigDecimal("0.005"), s"B1=$b, expected ~$fixpointB")
    // evaluated map includes both members
    val evaluated = result.evaluated(SheetName.unsafe("S"))
    assert(evaluated.contains(a1) && evaluated.contains(b1))
  }

  test("GH-373: non-convergent oscillator hits maxIter, keeps last values, NO circular error") {
    // A1 = 1-B1, B1 = A1 oscillates with period 4 from the (0,0) seed; s10 = (1,1).
    val sheet = Sheet(SheetName.unsafe("S"))
      .put(a1, formula("=1-B1"))
      .put(b1, formula("=A1"))
    val result = Workbook(sheet).recalculate(IterativeCalc(10, BigDecimal("0.001")))
    assert(
      result.errors.isEmpty,
      s"non-convergence must NOT error (Excel semantics), got: ${result.errors.map(_.render)}"
    )
    assertEquals(cachedNum(result.workbook, "S", a1), Some(BigDecimal(1)))
    assertEquals(cachedNum(result.workbook, "S", b1), Some(BigDecimal(1)))
  }

  test("GH-373: loose vs tight maxChange trade precision (loose stops far from the fixpoint)") {
    val loose = convergentPair.recalculate(IterativeCalc(100, BigDecimal(10)))
    val tight = convergentPair.recalculate(IterativeCalc(100, BigDecimal("0.0001")))
    val looseB = cachedNum(loose.workbook, "S", b1).getOrElse(fail("loose B1 must cache"))
    val tightB = cachedNum(tight.workbook, "S", b1).getOrElse(fail("tight B1 must cache"))
    assert((looseB - fixpointB).abs > BigDecimal("0.1"), s"loose stopped too precisely: $looseB")
    assert((tightB - fixpointB).abs < BigDecimal("0.001"), s"tight not precise enough: $tightB")
  }

  test("GH-373: default recalculate() is unchanged — cycles error exactly as before") {
    val result = convergentPair.recalculate()
    assertEquals(cached(result.workbook, "S", a1), None)
    assertEquals(cached(result.workbook, "S", b1), None)
    assertEquals(result.errors.map(_.ref).toSet, Set(a1, b1))
    assert(result.errors.forall(_.error.message.contains("Circular")))
  }

  test("GH-373: iterative recalc is total under numeric blowup (repeated squaring, no throw)") {
    val sheet = Sheet(SheetName.unsafe("S"))
      .put(a1, formula("=POWER(B1,2)+2"))
      .put(b1, formula("=A1"))
    // Doubling digit-count per round; must return a RecalcResult, never unwind (GH-388 posture).
    val result = Workbook(sheet).recalculate(IterativeCalc(15, BigDecimal("0.001")))
    assert(result.workbook.sheets.nonEmpty) // reached: no exception escaped
  }

  test("GH-373: acyclic remainder still computes in one pass alongside an iterated cycle") {
    val sheet = Sheet(SheetName.unsafe("S"))
      .put(a1, formula("=B1*0.1"))
      .put(b1, formula("=100+A1/2"))
      .put(c1, num(7))
      .put(d1, formula("=C1*3")) // independent of the cycle
    val result = Workbook(sheet).recalculate(IterativeCalc(100, BigDecimal("0.001")))
    assert(result.isClean, s"expected clean, got: ${result.errors.map(_.render)}")
    assertEquals(cachedNum(result.workbook, "S", d1), Some(BigDecimal(21)))
  }

  test("GH-373: dependents of the cycle get converged values (blocked-inversion pin)") {
    val sheet = Sheet(SheetName.unsafe("S"))
      .put(a1, formula("=B1*0.1"))
      .put(b1, formula("=100+A1/2"))
      .put(c1, formula("=B1*2")) // downstream of the cycle — blocked in the default path
    val result = Workbook(sheet).recalculate(IterativeCalc(100, BigDecimal("0.001")))
    assert(result.isClean, s"expected clean, got: ${result.errors.map(_.render)}")
    val c = cachedNum(result.workbook, "S", c1).getOrElse(fail("C1 must cache"))
    assert((c - fixpointB * 2).abs < BigDecimal("0.01"), s"C1=$c, expected ~${fixpointB * 2}")
    // Contrast: the default path reports this exact cell as blocked
    val blocked = Workbook(sheet).recalculate()
    assert(blocked.errors.exists(e => e.ref == c1 && e.error.message.contains("Blocked")))
  }

  test("GH-373: cross-sheet cycle iterates too (workbook-level SCC)") {
    val s1 = Sheet(SheetName.unsafe("S1")).put(a1, formula("=S2!A1*0.1"))
    val s2 = Sheet(SheetName.unsafe("S2")).put(a1, formula("=100+S1!A1/2"))
    val result = Workbook(s1, s2).recalculate(IterativeCalc(100, BigDecimal("0.001")))
    assert(result.isClean, s"expected clean, got: ${result.errors.map(_.render)}")
    val b = cachedNum(result.workbook, "S2", a1).getOrElse(fail("S2!A1 must cache"))
    assert((b - fixpointB).abs < BigDecimal("0.005"), s"S2!A1=$b, expected ~$fixpointB")
  }

  test("GH-373: clock is pinned across iterations — NOW() must not re-tick per iteration") {
    // A ticking clock advances one day per call. Unpinned, each Jacobi iteration would read a
    // different NOW() and B1 (previous-iteration copy of A1) could never equal A1.
    final class TickingClock extends Clock:
      private var calls = 0
      def today(): java.time.LocalDate =
        calls += 1
        java.time.LocalDate.of(2026, 1, 1).plusDays(calls.toLong)
      def now(): java.time.LocalDateTime =
        calls += 1
        java.time.LocalDateTime.of(2026, 1, 1, 0, 0).plusDays(calls.toLong)
    val sheet = Sheet(SheetName.unsafe("S"))
      .put(a1, formula("=B1*0+NOW()"))
      .put(b1, formula("=A1"))
    val result =
      Workbook(sheet).recalculate(new TickingClock, IterativeCalc(20, BigDecimal("0.001")))
    assert(
      result.isClean,
      s"expected clean (pinned clock converges): ${result.errors.map(_.render)}"
    )
    assertEquals(
      cachedNum(result.workbook, "S", a1),
      cachedNum(result.workbook, "S", b1),
      "A1 and B1 must agree — a re-ticking clock would leave them one iteration apart"
    )
  }

  test("GH-373: IterativeCalc.fromCalcPr uses Excel defaults when attributes are absent") {
    assertEquals(
      IterativeCalc.fromCalcPr(CalcPr(iterativeCalculation = true)),
      IterativeCalc(100, BigDecimal("0.001"))
    )
    assertEquals(
      IterativeCalc.fromCalcPr(
        CalcPr(iterativeCalculation = true, Some(50), Some(BigDecimal("0.01")))
      ),
      IterativeCalc(50, BigDecimal("0.01"))
    )
  }

  test("GH-373: acyclic workbook under iterative settings behaves exactly like the default") {
    val sheet = Sheet(SheetName.unsafe("S"))
      .put(a1, num(10))
      .put(b1, formula("=A1*2"))
    val iterative = Workbook(sheet).recalculate(IterativeCalc(100, BigDecimal("0.001")))
    val default = Workbook(sheet).recalculate()
    assertEquals(iterative.workbook, default.workbook)
    assertEquals(iterative.errors, default.errors)
  }
