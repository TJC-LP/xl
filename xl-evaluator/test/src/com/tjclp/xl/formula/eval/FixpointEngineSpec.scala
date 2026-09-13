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
 * first exact replay; a pinned (cached closed-workbook) member is a constant.
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
    iterative: IterativeCalc = tolerance
  ): WorkbookEvaluator.FixpointOutcome =
    val wb = Workbook(sheet)
    WorkbookEvaluator.jacobiFixpoint(
      wb,
      wb.sheets,
      members,
      iterative,
      Clock.system,
      () => Evaluator.instance,
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

  test("GH-537: the round evaluator factory is invoked exactly once per round") {
    var calls = 0
    val sheet = Sheet(s)
      .put(a1, CellValue.Formula("=B1*0.5+10", None))
      .put(b1, CellValue.Formula("=A1*0.5+20", None))
    val wb = Workbook(sheet)
    val outcome = WorkbookEvaluator.jacobiFixpoint(
      wb,
      wb.sheets,
      List((q(a1), 0, "=B1*0.5+10"), (q(b1), 0, "=A1*0.5+20")),
      IterativeCalc(100, BigDecimal("1E-9")),
      Clock.system,
      () =>
        calls += 1
        Evaluator.instance
      ,
      Map.empty
    )
    assert(outcome.converged, "a contraction converges")
    assert(outcome.rounds > 1, s"the fixture must take several rounds, took ${outcome.rounds}")
    assertEquals(calls, outcome.rounds, "one fresh evaluator (and memo) per round, never fewer")
  }
