package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.eval.{IterationScheme, IterativeCalc}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.{CalcPr, DefinedName, Workbook}
import munit.FunSuite

/**
 * GH-373b: opt-in bounded iterative recalculation for circular workbooks.
 *
 * `recalculate(IterativeCalc(maxIter, maxChange))` fixpoints cycle members — a Gauss–Seidel sweep
 * in the graph's within-component order by default (GH-482), `IterationScheme.Jacobi` for
 * previous-round reads — instead of erroring: members seed to 0 (Excel's uninitialized semantics;
 * GH-469 warm-starts from cached numbers), iterate until every |Δ| < maxChange or maxIter, and
 * non-convergence KEEPS the last values with NO error (Excel semantics — the deliberate inversion
 * vs the default path). Dependents of a cycle are NOT pre-marked blocked — they evaluate off the
 * converged values. The default `recalculate()` is byte-identical to before (opt-in only; the
 * CalcPr metadata is deliberately NOT auto-honored — bridge explicitly via
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
    // A1 = 1-B1, B1 = A1. GH-482: the default Gauss–Seidel sweep evaluates A1 then B1 within a
    // round, so from the (0,0) seed the rounds alternate (1,1)/(0,0) and round 10 lands on (0,0);
    // Jacobi (every member reads the previous round) has period 4 and lands on (1,1).
    val sheet = Sheet(SheetName.unsafe("S"))
      .put(a1, formula("=1-B1"))
      .put(b1, formula("=A1"))
    val result = Workbook(sheet).recalculate(IterativeCalc(10, BigDecimal("0.001")))
    assert(
      result.errors.isEmpty,
      s"non-convergence must NOT error (Excel semantics), got: ${result.errors.map(_.render)}"
    )
    assertEquals(cachedNum(result.workbook, "S", a1), Some(BigDecimal(0)))
    assertEquals(cachedNum(result.workbook, "S", b1), Some(BigDecimal(0)))
    val jacobi = Workbook(sheet)
      .recalculate(IterativeCalc(10, BigDecimal("0.001"), scheme = IterationScheme.Jacobi))
    assertEquals(cachedNum(jacobi.workbook, "S", a1), Some(BigDecimal(1)))
    assertEquals(cachedNum(jacobi.workbook, "S", b1), Some(BigDecimal(1)))
  }

  test("GH-373: loose vs tight maxChange trade precision (loose stops far from the fixpoint)") {
    val loose = convergentPair.recalculate(IterativeCalc(100, BigDecimal(10)))
    val tight = convergentPair.recalculate(IterativeCalc(100, BigDecimal("0.0001")))
    val looseB = cachedNum(loose.workbook, "S", b1).getOrElse(fail("loose B1 must cache"))
    val tightB = cachedNum(tight.workbook, "S", b1).getOrElse(fail("tight B1 must cache"))
    // GH-482: a Gauss–Seidel sweep lands closer than Jacobi did for the same loose tolerance
    // (105.25 vs the 105.263… fixpoint); the trade-off itself is what this pins.
    val looseResidual = (looseB - fixpointB).abs
    val tightResidual = (tightB - fixpointB).abs
    assert(looseResidual > BigDecimal("0.005"), s"loose stopped too precisely: $looseB")
    assert(tightResidual < BigDecimal("0.001"), s"tight not precise enough: $tightB")
    assert(
      looseResidual > tightResidual,
      s"loose ($looseB) must stop farther out than tight ($tightB)"
    )
    assert(loose.iterationsUsed < tight.iterationsUsed, "the loose run must also stop sooner")
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

  // ===== GH-537: a permanently failing member stalls its component at stationarity =====

  /**
   * A1/B1 cycle whose B1 fails EVERY round: a reference to a sheet that does not exist is a host
   * failure (never an Excel error value), and the graph keeps B1 cyclic through its `A1` edge.
   * Round 1 moves A1 from the 0 seed to 10; round 2 replays round 1 exactly — and so would every
   * further round, the clock being pinned and the formulas deterministic.
   */
  private def failingMemberCycle: Workbook =
    Workbook(
      Sheet(SheetName.unsafe("S"))
        .put(a1, formula("=B1*0.5+10"))
        .put(b1, formula("=A1*0.5+Nowhere!A1"))
    )

  test("GH-537: a member that fails every round stalls the component at stationarity") {
    val result = failingMemberCycle.recalculate(IterativeCalc(400, BigDecimal("0.001")))
    val scc = result.cycles.headOption.getOrElse(fail("one cyclic component expected"))
    assert(!result.converged, "a failing member can never converge")
    assert(scc.stalled, s"the component must report stalled: ${scc.render}")
    assertEquals(scc.rounds, 2, "round 2 is the first exact replay of round 1")
    assertEquals(result.iterationsUsed, 2)
    // The failing member is the host failure; GH-563 then invalidates its dependents (here A1,
    // the other member of the cycle) exactly as it does after a failed acyclic cell.
    assertEquals(result.errors.map(_.ref), Vector(b1, a1), result.errors.map(_.render).toString)
    assert(
      result.errors.headOption.exists(_.error.message.contains("Nowhere")),
      result.errors.map(_.render).toString
    )
    assert(
      result.errors.lift(1).exists(_.error.message.contains("Blocked by an upstream")),
      result.errors.map(_.render).toString
    )
    assert(scc.render.contains("stalled after 2 round(s)"), scc.render)
    assert(!scc.render.contains("exhausted"), scc.render)
  }

  test("GH-537: a stall on the last budgeted round is still a stall — rounds can equal maxIter") {
    // Round 2 is the first exact replay AND the budget's last round: the stall verdict wins and
    // `rounds == maxIter` is a legitimate shape (`rounds <= maxIter`, not strictly below). With a
    // budget of ONE round no replay can be observed yet — that is exhaustion, not a stall.
    val result = failingMemberCycle.recalculate(IterativeCalc(2, BigDecimal("0.001")))
    val scc = result.cycles.headOption.getOrElse(fail("one cyclic component expected"))
    assert(!result.converged)
    assert(scc.stalled, scc.render)
    assertEquals(scc.rounds, 2)
    assertEquals(result.iterationsUsed, 2)
    assert(scc.render.contains("stalled after 2 round(s)"), scc.render)
    val single = failingMemberCycle.recalculate(IterativeCalc(1, BigDecimal("0.001")))
    val onlyRound = single.cycles.headOption.getOrElse(fail("one cyclic component expected"))
    assert(!onlyRound.stalled, onlyRound.render)
    assertEquals(onlyRound.rounds, 1)
    assert(onlyRound.render.contains("exhausted 1 round(s)"), onlyRound.render)
  }

  test("GH-537: a member drawing randomness keeps a failing cycle running to its budget") {
    // Every round consumes randomness: even if two values match, that does not certify that the
    // next round will replay. The loop must run to maxIter and report exhaustion, not a stall.
    val wb = Workbook(
      Sheet(SheetName.unsafe("S"))
        .put(a1, formula("=B1*0+RAND()"))
        .put(b1, formula("=A1*0.5+Nowhere!A1"))
    )
    val result = wb.recalculate(IterativeCalc(20, BigDecimal("0.001")))
    val scc = result.cycles.headOption.getOrElse(fail("one cyclic component expected"))
    assert(!result.converged)
    assert(!scc.stalled, scc.render)
    assertEquals(scc.rounds, 20)
    assert(scc.render.contains("exhausted 20 round(s)"), scc.render)
    assert(result.summary.contains("exhausted 20 round(s)"), result.summary)
  }

  Vector(IterationScheme.Jacobi, IterationScheme.GaussSeidel).foreach { scheme =>
    Vector("RANDBETWEEN(0,1)", "ROUND(RAND(),0)", "randomChoice").foreach { random =>
      test(s"GH-537: repeated $random values can recover a failed cycle under $scheme") {
        var draws = 0
        val rng = new Rng:
          def nextDouble(): Double =
            draws += 1
            if draws == 1 then 0.1 else 0.9
        val base = Workbook(
          Sheet(SheetName.unsafe("S"))
            .put(a1, formula(s"B1*0+$random"))
            .put(b1, formula("IF(A1=0,Nowhere!A1,A1)"))
        )
        // A workbook name reaches a sheet-scoped name: tracking must follow actual evaluation,
        // including name resolution, rather than scanning only the cycle's formula text.
        val wb = base.copy(metadata =
          base.metadata.copy(definedNames =
            Vector(
              DefinedName("randomChoice", "randomSource"),
              DefinedName("randomSource", "RANDBETWEEN(0,1)", Some(0))
            )
          )
        )
        val result = wb.recalculate(
          Clock.system,
          rng,
          IterativeCalc(10, BigDecimal("0.001"), scheme = scheme)
        )
        assert(result.isClean, s"draws=$draws ${result.summary}")
        assert(result.converged, result.summary)
        assert(draws > 1, "the first draw repeats the zero seed but later draws heal B1")
        assertEquals(cachedNum(result.workbook, "S", a1), Some(BigDecimal(1)))
        assertEquals(cachedNum(result.workbook, "S", b1), Some(BigDecimal(1)))
        assert(result.cycles.forall(!_.stalled))
      }
    }

    test(s"GH-537: a deterministic replay can stall after an earlier random round under $scheme") {
      var draws = 0
      val rng = new Rng:
        def nextDouble(): Double =
          draws += 1
          0.5
      val wb = Workbook(
        Sheet(SheetName.unsafe("S"))
          .put(a1, formula("IF(A1=0,RAND()+1,A1)+B1*0"))
          .put(b1, formula("A1*0+Nowhere!A1"))
      )
      val result = wb.recalculate(
        Clock.system,
        rng,
        IterativeCalc(10, BigDecimal("0.001"), scheme = scheme)
      )
      assertEquals(draws, 1, "only round 1 reaches RAND")
      assertEquals(result.cycles.map(_.rounds), Vector(2))
      assert(result.cycles.forall(_.stalled), result.summary)
    }
  }

  test("GH-537: the summary names a stall as a stall, never as exhaustion") {
    val stalled = failingMemberCycle.recalculate(IterativeCalc(400, BigDecimal("0.001")))
    assert(
      stalled.summary.contains(
        "WARNING: iterative calculation stalled after 2 round(s): a cyclic member fails every round"
      ),
      stalled.summary
    )
    assert(!stalled.summary.contains("exhausted"), stalled.summary)
    // One stalled and one genuinely oscillating component: exhaustion stays the headline.
    val mixed = Workbook(
      Sheet(SheetName.unsafe("S"))
        .put(a1, formula("=B1*0.5+10"))
        .put(b1, formula("=A1*0.5+Nowhere!A1"))
        .put(c1, formula("=1-D1"))
        .put(d1, formula("=C1"))
    ).recalculate(IterativeCalc(10, BigDecimal("0.001")))
    assertEquals(mixed.cycles.map(_.stalled), Vector(true, false))
    assertEquals(mixed.iterationsUsed, 10)
    assert(mixed.summary.contains("exhausted 10 round(s) without converging"), mixed.summary)
    assert(!mixed.summary.contains("stalled"), mixed.summary)
  }

  test("GH-537: the stall verdict names the STALLED component's rounds, not the longest run") {
    // A1/B1 stalls at round 2; C1/D1 is a contraction (error shrinks 4x per sweep) that converges
    // only after many rounds at 1E-9. `iterationsUsed` is the worst component's rounds, but the
    // stall the summary reports happened at round 2 — and the strict reason must agree with it.
    val mixed = Workbook(
      Sheet(SheetName.unsafe("S"))
        .put(a1, formula("=B1*0.5+10"))
        .put(b1, formula("=A1*0.5+Nowhere!A1"))
        .put(c1, formula("=D1*0.5+10"))
        .put(d1, formula("=C1*0.5+20"))
    ).recalculate(IterativeCalc(400, BigDecimal("1E-9")))
    assertEquals(
      mixed.cycles.map(c => (c.stalled, c.converged)),
      Vector((true, false), (false, true))
    )
    val healthy = mixed.cycles.lift(1).getOrElse(fail("two components expected"))
    assert(healthy.rounds > 2, s"the fixture must converge late, took ${healthy.rounds}")
    assertEquals(mixed.iterationsUsed, healthy.rounds)
    val verdict = "stalled after 2 round(s): a cyclic member fails every round"
    assertEquals(mixed.unconvergedVerdict, Some(verdict))
    assert(mixed.summary.contains(s"WARNING: iterative calculation $verdict"), mixed.summary)
    assert(!mixed.summary.contains(s"${healthy.rounds} round"), mixed.summary)
    // Per component, each report renders its own rounds.
    assertEquals(
      mixed.cycles.map(_.verdict),
      Vector(
        "stalled after 2 round(s): a member fails every round",
        s"converged in ${healthy.rounds} round(s)"
      )
    )
    // Exhaustion keeps its own number: the budget, not the stall's round.
    val exhausted = Workbook(
      Sheet(SheetName.unsafe("S")).put(c1, formula("=1-D1")).put(d1, formula("=C1"))
    ).recalculate(IterativeCalc(10, BigDecimal("0.001")))
    assertEquals(
      exhausted.unconvergedVerdict,
      Some("exhausted 10 round(s) without converging (last values kept)")
    )
    // A converged run has no verdict to gate on.
    assertEquals(
      convergentPair.recalculate(IterativeCalc(100, BigDecimal("0.001"))).unconvergedVerdict,
      None
    )
  }

  test("GH-537: a cached external-ref member is a constant inside its cycle (pinned cache)") {
    // GH-353: the closed-workbook cache IS the value. Parsing each member once per fixpoint must
    // keep the pinned short-circuit: A1 never evaluates, B1 settles on half of it.
    val wb = Workbook(
      Sheet(SheetName.unsafe("S"))
        .put(a1, CellValue.Formula("=[2]Book!A1*B1", Some(num(7))))
        .put(b1, formula("=A1*0.5"))
    )
    val result = wb.recalculate(IterativeCalc(100, BigDecimal("0.001")))
    assert(result.isClean, result.errors.map(_.render).toString)
    assert(result.converged)
    assertEquals(cachedNum(result.workbook, "S", a1), Some(BigDecimal(7)))
    assertEquals(cachedNum(result.workbook, "S", b1), Some(BigDecimal("3.5")))
    assertEquals(result.cycles.map(_.rounds), Vector(2))
  }

  // ===== GH-482: Gauss–Seidel within a component (Excel's sequential sweep) is the default =====

  private def gaussSeidel(maxIter: Int, maxChange: BigDecimal): IterativeCalc =
    IterativeCalc(maxIter, maxChange, scheme = IterationScheme.GaussSeidel)
  private def jacobi(maxIter: Int, maxChange: BigDecimal): IterativeCalc =
    IterativeCalc(maxIter, maxChange, scheme = IterationScheme.Jacobi)

  test("GH-482: the chain A1 = B1+1, B1 = A1 counts to 100 in 100 rounds under the sweep") {
    // A sequential sweep with the latest values (Excel's iteration model, in Excel's documented
    // left-to-right, top-to-bottom order — xl has not been checked against Excel itself on this
    // shape): A1 steps once per iteration and B1 copies it, so 100 iterations from empty cells
    // give 100/100. Jacobi hands B1 the PREVIOUS A1, so the pair advances one unit every two
    // rounds — ~50/50.
    val sheet = Sheet(SheetName.unsafe("S"))
      .put(a1, formula("=B1+1"))
      .put(b1, formula("=A1"))
    val excel = Workbook(sheet).recalculate(IterativeCalc(100, BigDecimal("0.001")))
    assertEquals(cachedNum(excel.workbook, "S", a1), Some(BigDecimal(100)))
    assertEquals(cachedNum(excel.workbook, "S", b1), Some(BigDecimal(100)))
    assert(!excel.converged && excel.iterationsUsed == 100, excel.cycles.map(_.render).toString)
    val old = Workbook(sheet).recalculate(jacobi(100, BigDecimal("0.001")))
    assertEquals(cachedNum(old.workbook, "S", a1), Some(BigDecimal(50)))
    assertEquals(cachedNum(old.workbook, "S", b1), Some(BigDecimal(50)))
  }

  /**
   * The two-cell counter `x = y+1`, `y = x` from empty cells: sweeping x first leaves (n, n) after
   * n rounds, sweeping y first leaves (n, n-1) — the trailing cell is one round old. So the kept
   * values of this non-convergent cycle expose the within-component order directly.
   */
  private def counter(sheetOf: (SheetName, SheetName), x: ARef, y: ARef): Workbook =
    val (sx, sy) = sheetOf
    val yText = s"=${SheetName.quoteForFormula(sy.value)}!${y.toA1}+1"
    val xText = s"=${SheetName.quoteForFormula(sx.value)}!${x.toA1}"
    if sx == sy then Workbook(Sheet(sx).put(x, formula(yText)).put(y, formula(xText)))
    else Workbook(Sheet(sx).put(x, formula(yText)), Sheet(sy).put(y, formula(xText)))

  private def counted(wb: Workbook, sheet: SheetName, ref: ARef): Option[BigDecimal] =
    cachedNum(wb.recalculate(IterativeCalc(100, BigDecimal("0.001"))).workbook, sheet.value, ref)

  test("GH-482: the sweep order is the grid's, so A2/A10 and Z1/AA1 count exactly like A1/B1") {
    // An isomorphic model must not change its kept values because a row has two digits or a
    // column two letters: the A1 STRING order ("A10" < "A2", "AA1" < "Z1") swept the trailing
    // cell first and left 100/99; the grid order sweeps the top-left cell first — 100/100.
    val s = SheetName.unsafe("S")
    val a2 = ARef.from0(0, 1)
    val a10 = ARef.from0(0, 9)
    val rows = counter((s, s), a2, a10)
    assertEquals(counted(rows, s, a2), Some(BigDecimal(100)))
    assertEquals(counted(rows, s, a10), Some(BigDecimal(100)))
    val z1 = ARef.from0(25, 0)
    val aa1 = ARef.from0(26, 0)
    val cols = counter((s, s), z1, aa1)
    assertEquals(counted(cols, s, z1), Some(BigDecimal(100)))
    assertEquals(counted(cols, s, aa1), Some(BigDecimal(100)))
    // Row before column: B2 (row 2) is swept before A10 (row 10) although "A10" < "B2".
    val b2 = ARef.from0(1, 1)
    val rowMajor = counter((s, s), b2, a10)
    assertEquals(counted(rowMajor, s, b2), Some(BigDecimal(100)))
    assertEquals(counted(rowMajor, s, a10), Some(BigDecimal(100)))
    // The control the CHANGELOG example rests on.
    val classic = counter((s, s), a1, b1)
    assertEquals(counted(classic, s, a1), Some(BigDecimal(100)))
    assertEquals(counted(classic, s, b1), Some(BigDecimal(100)))
  }

  test("GH-482: a cross-sheet cycle sweeps sheet order first, then row-major within the sheet") {
    // T!A1 = S!A10+1 and S!A10 = T!A1: sheet S sorts before T, so S!A10 is swept first even
    // though T!A1 sits on row 1 — it copies the PREVIOUS T!A1 and trails by a round (99/100).
    // Swap the roles (S!A10 = T!A1+1) and the counter cell is swept first: 100/100.
    val sS = SheetName.unsafe("S")
    val sT = SheetName.unsafe("T")
    val a10 = ARef.from0(0, 9)
    val trailing = counter((sT, sS), a1, a10)
    assertEquals(counted(trailing, sT, a1), Some(BigDecimal(100)))
    assertEquals(counted(trailing, sS, a10), Some(BigDecimal(99)))
    val leading = counter((sS, sT), a10, a1)
    assertEquals(counted(leading, sS, a10), Some(BigDecimal(100)))
    assertEquals(counted(leading, sT, a1), Some(BigDecimal(100)))
  }

  test("GH-482: fromCalcPr selects Gauss–Seidel, and so does the constructor default") {
    assertEquals(
      IterativeCalc.fromCalcPr(CalcPr(iterativeCalculation = true)).scheme,
      IterationScheme.GaussSeidel
    )
    assertEquals(IterativeCalc(100, BigDecimal("0.001")).scheme, IterationScheme.GaussSeidel)
    assertEquals(jacobi(100, BigDecimal("0.001")).scheme, IterationScheme.Jacobi)
  }

  /** The banker's average-balance idiom: interest on the average of opening and closing balance. */
  private def averageBalance: Workbook =
    Workbook(
      Sheet(SheetName.unsafe("S"))
        .put(ARef.from0(0, 0), num(1000)) // A1 opening
        .put(ARef.from0(0, 1), CellValue.Number(BigDecimal("0.08"))) // A2 rate
        .put(ARef.from0(0, 2), num(50)) // A3 cash
        .put(ARef.from0(1, 0), formula("=A1")) // B1 beg
        .put(ARef.from0(1, 1), formula("=(B1+B4)/2*A2")) // B2 interest
        .put(ARef.from0(1, 2), formula("=A3-B2")) // B3 pay
        .put(ARef.from0(1, 3), formula("=B1-B3")) // B4 end
    )

  test("GH-482: the average-balance idiom converges in strictly fewer rounds under Gauss–Seidel") {
    val tolerance = BigDecimal("1E-9")
    val gs = averageBalance.recalculate(gaussSeidel(400, tolerance))
    val jac = averageBalance.recalculate(jacobi(400, tolerance))
    assert(
      gs.converged && jac.converged,
      s"${gs.cycles.map(_.render)} / ${jac.cycles.map(_.render)}"
    )
    val gsRounds = gs.cycles.map(_.rounds).sum
    val jacRounds = jac.cycles.map(_.rounds).sum
    assert(gsRounds < jacRounds, s"Gauss–Seidel took $gsRounds rounds, Jacobi $jacRounds")
    // Same fixpoint: B4 (1 - r/2) = B1 (1 + r/2) - cash
    val analytic = (BigDecimal(1000) * BigDecimal("1.04") - BigDecimal(50)) / BigDecimal("0.96")
    val b4 = ARef.from0(1, 3)
    val gsEnd = cachedNum(gs.workbook, "S", b4).getOrElse(fail("B4 (G-S) must cache"))
    val jacEnd = cachedNum(jac.workbook, "S", b4).getOrElse(fail("B4 (Jacobi) must cache"))
    assert((gsEnd - analytic).abs < BigDecimal("1E-6"), s"G-S B4=$gsEnd, expected ~$analytic")
    assert((gsEnd - jacEnd).abs < BigDecimal("1E-6"), s"schemes disagree: $gsEnd vs $jacEnd")
  }

  test("GH-482: on every pinned convergent fixture Gauss–Seidel needs no more rounds than Jacobi") {
    val tolerance = BigDecimal("1E-9")
    val sumOverCoMembers = Workbook(
      Sheet(SheetName.unsafe("S"))
        .put(a1, formula("=SUM(B1:C1)*0.5+10"))
        .put(b1, formula("=A1*0.5"))
        .put(c1, formula("=A1*0.25"))
    )
    Vector(
      "convergentPair" -> convergentPair,
      "poisoned caches" -> Workbook(
        Sheet(SheetName.unsafe("S"))
          .put(a1, CellValue.Formula("=B1*0.1", Some(num(-5000))))
          .put(b1, CellValue.Formula("=100+A1/2", Some(num(9999))))
      ),
      "sum over co-members" -> sumOverCoMembers,
      "average balance" -> averageBalance
    ).foreach { (label, wb) =>
      val gs = wb.recalculate(gaussSeidel(400, tolerance))
      val jac = wb.recalculate(jacobi(400, tolerance))
      assert(gs.converged && jac.converged, s"$label: ${gs.cycles} / ${jac.cycles}")
      assert(
        gs.iterationsUsed <= jac.iterationsUsed,
        s"$label: Gauss–Seidel ${gs.iterationsUsed} rounds > Jacobi ${jac.iterationsUsed}"
      )
      gs.evaluated(SheetName.unsafe("S")).foreach { (ref, value) =>
        (value, jac.evaluated(SheetName.unsafe("S")).get(ref)) match
          case (CellValue.Number(g), Some(CellValue.Number(j))) =>
            assert((g - j).abs < BigDecimal("1E-6"), s"$label ${ref.toA1}: $g vs $j")
          case (g, j) => fail(s"$label ${ref.toA1}: non-numeric fixpoint $g / $j")
      }
    }
  }
