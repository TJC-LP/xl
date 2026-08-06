package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue}
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

  test("GH-454: convergent cycle reports converged=true with iterationsUsed in (0, maxIter]") {
    val result = convergentPair.recalculate(IterativeCalc(100, BigDecimal("0.001")))
    assert(result.converged, "a fixpointed cycle must report converged")
    assert(
      result.iterationsUsed > 0 && result.iterationsUsed <= 100,
      s"iterationsUsed must count the rounds actually run, got ${result.iterationsUsed}"
    )
  }

  test("GH-454: exhausted maxIter reports converged=false and iterationsUsed == maxIter") {
    // Same oscillator as the non-convergence test above: values are kept (Excel semantics)
    // but the result must now SAY it under-converged — exhaustion was previously
    // indistinguishable from convergence (the field bug: a stale debt schedule with r.errors
    // empty).
    val sheet = Sheet(SheetName.unsafe("S"))
      .put(a1, formula("=1-B1"))
      .put(b1, formula("=A1"))
    val result = Workbook(sheet).recalculate(IterativeCalc(10, BigDecimal("0.001")))
    assert(result.errors.isEmpty, "non-convergence still must NOT error")
    assert(!result.converged, "maxIter exhaustion must report converged=false")
    assertEquals(result.iterationsUsed, 10)
  }

  test("GH-454: non-iterative runs report converged=true, iterationsUsed=0") {
    // Default path (cycle errors instead of iterating): no iteration happened.
    val defaultRun = convergentPair.recalculate()
    assert(defaultRun.converged)
    assertEquals(defaultRun.iterationsUsed, 0)
    // Acyclic workbook under iterative settings: nothing to iterate.
    val acyclic = Workbook(Sheet(SheetName.unsafe("S")).put(a1, num(1)).put(b1, formula("=A1*2")))
    val acyclicRun = acyclic.recalculate(IterativeCalc(100, BigDecimal("0.001")))
    assert(acyclicRun.converged)
    assertEquals(acyclicRun.iterationsUsed, 0)
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

  // ===== GH-469: warm start — seed cycle members from their loaded caches (Excel semantics) =====

  private val a3 = ARef.from0(0, 2)
  private val b3 = ARef.from0(1, 2)

  /**
   * The reported fixture: a mutually `IF(ISERROR(...))`-guarded pair sitting at a VALID numeric
   * fixpoint (A3 = B3*0.5+10 = 20, B3 = A3*0.5+10 = 20). Cold-seeded at 0 the guards see 0/0 =
   * #DIV/0! in round 1, both members flip to the text branch, and `"NA "` is itself a fixpoint — so
   * the run "converges" having destroyed two valid caches with `errors` and `excelErrors` both
   * empty.
   */
  private val guardedPairAtFixpoint: Workbook =
    Workbook(
      Sheet(SheetName.unsafe("S"))
        .put(a3, CellValue.Formula("=IF(ISERROR(B3/A3-1),\"NA \",B3*0.5+10)", Some(num(20))))
        .put(b3, CellValue.Formula("=IF(ISERROR(A3/B3-1),\"NA \",A3*0.5+10)", Some(num(20))))
        .put(c1, CellValue.Formula("=A3*2+1", Some(num(41))))
    )

  test("GH-469: a guarded pair at a valid numeric fixpoint keeps its caches (warm start)") {
    val result = guardedPairAtFixpoint.recalculate(IterativeCalc(200, BigDecimal("1E-10")))
    assertEquals(cachedNum(result.workbook, "S", a3), Some(BigDecimal(20)), "A3 must stay 20")
    assertEquals(cachedNum(result.workbook, "S", b3), Some(BigDecimal(20)), "B3 must stay 20")
    assert(result.converged, "a book already at its fixpoint must converge")
    assertEquals(result.errors, Vector.empty[com.tjclp.xl.formula.eval.CellEvalError])
    assertEquals(result.excelErrors, Vector.empty)
    assertEquals(result.iterationsUsed, 1, "seeding at the fixpoint converges in one round")
    assertEquals(result.cycles.map(_.rounds), Vector(1))
    // the naked dependent recomputes off the (unchanged) member values
    assertEquals(cachedNum(result.workbook, "S", c1), Some(BigDecimal(41)))
  }

  test("GH-469: seedFromCaches = false reproduces the cold-start text flip") {
    val cold = guardedPairAtFixpoint
      .recalculate(IterativeCalc(200, BigDecimal("1E-10"), seedFromCaches = false))
    assertEquals(cached(cold.workbook, "S", a3), Some(CellValue.Text("NA ")))
    assertEquals(cached(cold.workbook, "S", b3), Some(CellValue.Text("NA ")))
    assert(cold.converged, "the text branch is itself a fixpoint — that is why it was silent")
  }

  test("GH-469: warm start still lands on the fixpoint when the caches are WRONG (contraction)") {
    // Same convergent pair, but poisoned with caches far from the fixpoint.
    val poisoned = Workbook(
      Sheet(SheetName.unsafe("S"))
        .put(a1, CellValue.Formula("=B1*0.1", Some(num(-5000))))
        .put(b1, CellValue.Formula("=100+A1/2", Some(num(9999))))
    )
    val warm = poisoned.recalculate(IterativeCalc(200, BigDecimal("1E-12")))
    val cold = convergentPair.recalculate(IterativeCalc(200, BigDecimal("1E-12")))
    val warmB = cachedNum(warm.workbook, "S", b1).getOrElse(fail("warm B1"))
    val coldB = cachedNum(cold.workbook, "S", b1).getOrElse(fail("cold B1"))
    assert(warm.converged && cold.converged)
    assert(
      (warmB - coldB).abs < BigDecimal("1E-9"),
      s"warm start reached a different fixpoint: $warmB vs $coldB"
    )
    assert((warmB - fixpointB).abs < BigDecimal("1E-9"), s"B1=$warmB, expected ~$fixpointB")
  }

  test("GH-469: a member inside the stripped dynamic bucket seeds 0 despite carrying a cache") {
    // P1 is dynamic (INDIRECT), so the whole component defers with the bucket and its caches are
    // declared stale and stripped — warm start must NOT resurrect them.
    val sheet = Sheet(SheetName.unsafe("S"))
      .put(a1, num(0))
      .put(
        a3,
        CellValue.Formula(
          "=IF(ISERROR(B3/A3-1),\"NA \",B3*0.5+10+INDIRECT(\"A1\"))",
          Some(num(20))
        )
      )
      .put(b3, CellValue.Formula("=IF(ISERROR(A3/B3-1),\"NA \",A3*0.5+10)", Some(num(20))))
    val result = Workbook(sheet).recalculate(IterativeCalc(200, BigDecimal("1E-10")))
    assertEquals(cached(result.workbook, "S", a3), Some(CellValue.Text("NA ")))
    assertEquals(cached(result.workbook, "S", b3), Some(CellValue.Text("NA ")))
  }

  /**
   * A perfectly healthy linear cycle (fixpoint A1 = 80/3, B1 = 100/3) whose LOADED caches are junk.
   * Warm seeding must not let the junk decide the answer: a non-numeric seed makes the arithmetic
   * propagate the junk forever (`#DIV/0! * 0.5 + 10` is `#DIV/0!` — a fixpoint), so seeding it
   * would wedge the cycle at its poison AND report success.
   */
  private def poisonedHealthyCycle(cache: CellValue): Workbook =
    Workbook(
      Sheet(SheetName.unsafe("S"))
        .put(a1, CellValue.Formula("=B1*0.5+10", Some(cache)))
        .put(b1, CellValue.Formula("=A1*0.5+20", Some(cache)))
    )

  private val healedA = BigDecimal(80) / BigDecimal(3)
  private val healedB = BigDecimal(100) / BigDecimal(3)

  test(
    "GH-469: a healthy cycle carrying stale #DIV/0! caches HEALS (non-numeric seeds fall to 0)"
  ) {
    val result = poisonedHealthyCycle(CellValue.Error(CellError.Div0))
      .recalculate(IterativeCalc(200, BigDecimal("1E-12")))
    val a = cachedNum(result.workbook, "S", a1).getOrElse(
      fail(s"A1 must heal to a number, got ${cached(result.workbook, "S", a1)}")
    )
    val b = cachedNum(result.workbook, "S", b1).getOrElse(
      fail(s"B1 must heal to a number, got ${cached(result.workbook, "S", b1)}")
    )
    assert((a - healedA).abs < BigDecimal("1E-9"), s"A1=$a, expected ~$healedA")
    assert((b - healedB).abs < BigDecimal("1E-9"), s"B1=$b, expected ~$healedB")
    assert(result.converged, "the cycle is a contraction — it must converge")
    assertEquals(result.excelErrors, Vector.empty, "no error value may survive the recalculation")
    assert(result.certified, "a healed, error-free contraction is at its global fixpoint")
  }

  test("GH-469: a healthy cycle carrying stale text caches HEALS (non-numeric seeds fall to 0)") {
    val result = poisonedHealthyCycle(CellValue.Text("junk"))
      .recalculate(IterativeCalc(200, BigDecimal("1E-12")))
    val a = cachedNum(result.workbook, "S", a1).getOrElse(
      fail(s"A1 must heal to a number, got ${cached(result.workbook, "S", a1)}")
    )
    assert((a - healedA).abs < BigDecimal("1E-9"), s"A1=$a, expected ~$healedA")
    assert(result.converged && result.isClean, s"errors: ${result.errors.map(_.render)}")
  }

  test("GH-469: fromCalcPr inherits the Excel-parity warm-start default") {
    assert(IterativeCalc.fromCalcPr(CalcPr(iterativeCalculation = true)).seedFromCaches)
    assert(IterativeCalc(100, BigDecimal("0.001")).seedFromCaches)
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
