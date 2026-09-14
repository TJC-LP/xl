package com.tjclp.xl.formula.eval

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * GH-537: the fixpoint engine's per-round contract, pinned at the `private[eval]` seam the two
 * callers (`recalculate(IterativeCalc)` and the data-table seeder) share. Members are parsed once
 * per fixpoint and evaluated per round; a member that fails every round stalls the loop at the
 * first exact replay that consumes no randomness; a pinned (cached closed-workbook) member is a
 * constant.
 */
class FixpointEngineSpec extends FunSuite:

  private val s = SheetName.unsafe("S")
  private val a1 = ARef.from0(0, 0)
  private val b1 = ARef.from0(1, 0)
  private def q(ref: ARef): QualifiedRef = QualifiedRef(s, ref)
  private def num(i: Int): CellValue = CellValue.Number(BigDecimal(i))
  private val tolerance = IterativeCalc(50, BigDecimal("0.001"))

  private def run(
    sheet: Sheet,
    members: List[(QualifiedRef, Int, String)],
    seed: Map[QualifiedRef, CellValue] = Map.empty,
    iterative: IterativeCalc = tolerance,
    rng: Rng = Rng.system
  ): WorkbookEvaluator.FixpointOutcome =
    val wb = Workbook(sheet)
    WorkbookEvaluator.jacobiFixpoint(
      wb,
      wb.sheets,
      members,
      iterative,
      Clock.system,
      rng,
      roundRng => Evaluator.instance(roundRng),
      seed
    )

  test("GH-537: an unparseable member reports the acyclic path's exact `Parse error:` XLError") {
    // The static graph never makes an unparseable cell cyclic (it has no edges), so this seam is
    // the only way to hand the engine one. Parsing once per fixpoint must surface the identical
    // message the one-shot path produces — and, failing every round, stall the loop at once: A1
    // seeds 0 and evaluates to 0, so round 1 already replays the seed.
    val sheet = Sheet(s)
      .put(a1, CellValue.Formula("=B1*0.5", None))
      .put(b1, CellValue.Formula("=A1+", None))
    val outcome = run(sheet, List((q(a1), 0, "=B1*0.5"), (q(b1), 0, "=A1+")))
    val expected =
      sheet.evaluateFormula("=A1+").swap.getOrElse(fail("the acyclic path must reject `=A1+`"))
    assert(expected.message.contains("Parse error"), expected.message)
    assertEquals(outcome.results(q(b1)), Left(expected))
    assertEquals(outcome.results(q(a1)), Right(num(0)))
    assert(!outcome.converged)
    assert(outcome.stalled, "an unparseable member fails every round")
    assertEquals(outcome.rounds, 1)
  }

  test("GH-537: a pinned external-ref member is the constant its seed says, decided once") {
    val text = "=[2]Book!A1*B1"
    val sheet = Sheet(s)
      .put(a1, CellValue.Formula(text, Some(num(7))))
      .put(b1, CellValue.Formula("=A1*0.5", None))
    val members = List((q(a1), 0, text), (q(b1), 0, "=A1*0.5"))
    // Warm-seeded from its cache (what `recalculate` does): the constant is 7.
    val warm = run(sheet, members, seed = Map(q(a1) -> num(7)))
    assertEquals(warm.results(q(a1)), Right(num(7)))
    assertEquals(warm.results(q(b1)), Right(CellValue.Number(BigDecimal("3.5"))))
    assert(warm.converged)
    assertEquals(warm.rounds, 2)
    // Cold-seeded (the seeder's posture) the same member is the constant 0: a pinned cell returns
    // its overlay's `Some(previous)`, round after round, exactly as before parse-once.
    val cold = run(sheet, members)
    assertEquals(cold.results(q(a1)), Right(num(0)))
    assertEquals(cold.results(q(b1)), Right(num(0)))
    assert(cold.converged)
    assertEquals(cold.rounds, 1)
  }

  test("GH-537/GH-482: the factory runs once per Jacobi round, once per Gauss–Seidel evaluation") {
    // A memo is sound only while its overlay is fixed: a whole round under Jacobi, a single
    // member evaluation under Gauss–Seidel (members change mid-round).
    val sheet = Sheet(s)
      .put(a1, CellValue.Formula("=B1*0.5+10", None))
      .put(b1, CellValue.Formula("=A1*0.5+20", None))
    val wb = Workbook(sheet)
    val members = List((q(a1), 0, "=B1*0.5+10"), (q(b1), 0, "=A1*0.5+20"))
    def run(scheme: IterationScheme): (Int, WorkbookEvaluator.FixpointOutcome) =
      var calls = 0
      val outcome = WorkbookEvaluator.jacobiFixpoint(
        wb,
        wb.sheets,
        members,
        IterativeCalc(100, BigDecimal("1E-9"), scheme = scheme),
        Clock.system,
        Rng.system,
        rng =>
          calls += 1
          Evaluator.instance(rng)
        ,
        Map.empty
      )
      (calls, outcome)
    val (jacobiCalls, jacobi) = run(IterationScheme.Jacobi)
    assert(jacobi.converged, "a contraction converges")
    assert(jacobi.rounds > 1, s"the fixture must take several rounds, took ${jacobi.rounds}")
    assertEquals(jacobiCalls, jacobi.rounds, "one fresh evaluator (and memo) per Jacobi round")
    val (gsCalls, gs) = run(IterationScheme.GaussSeidel)
    assert(gs.converged)
    assertEquals(gsCalls, gs.rounds * members.size, "one fresh evaluator per member evaluation")
    assert(gs.rounds <= jacobi.rounds, s"G-S ${gs.rounds} rounds vs Jacobi ${jacobi.rounds}")
  }

  test("GH-482: a Gauss–Seidel sweep publishes each value to the members after it") {
    // A1 = B1+1, B1 = A1 swept (A1, B1) from 0: round r leaves (r, r); Jacobi's B1 reads the
    // previous A1 and round 3 leaves (2, 1).
    val sheet = Sheet(s)
      .put(a1, CellValue.Formula("=B1+1", None))
      .put(b1, CellValue.Formula("=A1", None))
    val members = List((q(a1), 0, "=B1+1"), (q(b1), 0, "=A1"))
    val gs = run(sheet, members, iterative = IterativeCalc(3, BigDecimal("0.001")))
    assertEquals(gs.results(q(a1)), Right(num(3)))
    assertEquals(gs.results(q(b1)), Right(num(3)))
    assert(!gs.converged && gs.rounds == 3)
    val jacobi = run(
      sheet,
      members,
      iterative = IterativeCalc(3, BigDecimal("0.001"), scheme = IterationScheme.Jacobi)
    )
    assertEquals(jacobi.results(q(a1)), Right(num(2)))
    assertEquals(jacobi.results(q(b1)), Right(num(1)))
  }

  Vector(IterationScheme.Jacobi, IterationScheme.GaussSeidel).foreach { scheme =>
    test(s"GH-537: a draw in a dynamically resolved formula prevents a false stall under $scheme") {
      var draws = 0
      val rng = new Rng:
        def nextDouble(): Double =
          draws += 1
          if draws == 1 then 0.1 else 0.9
      val aText = "B1*0+INDIRECT(\"C1\")"
      val bText = "IF(A1=0,Nowhere!A1,A1)"
      val sheet = Sheet(s)
        .put(a1, CellValue.Formula(aText))
        .put(b1, CellValue.Formula(bText))
        .put(ARef.from0(2, 0), CellValue.Formula("RANDBETWEEN(0,1)"))
      // C1 is outside the component and uncached: INDIRECT recursively evaluates it on every
      // visit. A text/AST scan of the component cannot see its use of randomness.
      val outcome = run(
        sheet,
        List((q(a1), 0, aText), (q(b1), 0, bText)),
        iterative = tolerance.copy(scheme = scheme),
        rng = rng
      )
      assert(outcome.converged)
      assert(!outcome.stalled)
      assertEquals(outcome.results(q(a1)), Right(num(1)))
      assertEquals(outcome.results(q(b1)), Right(num(1)))
      assert(draws > 1)
    }

    test(s"GH-537: a random draw that throws cannot certify a permanent failure under $scheme") {
      var attempts = 0
      val rng = new Rng:
        def nextDouble(): Double =
          attempts += 1
          if attempts == 1 then throw new IllegalStateException("transient random source failure")
          else 0.0
      val text = "A1*0+RAND()"
      val sheet = Sheet(s).put(a1, CellValue.Formula(text))
      val outcome = run(
        sheet,
        List((q(a1), 0, text)),
        iterative = tolerance.copy(scheme = scheme),
        rng = rng
      )
      assert(outcome.converged)
      assert(!outcome.stalled)
      assertEquals(outcome.rounds, 2)
      assertEquals(attempts, 2)
      assertEquals(outcome.results(q(a1)), Right(num(0)))
    }
  }
