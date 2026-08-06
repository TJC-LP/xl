package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellValue, FormulaKind}
import com.tjclp.xl.formula.eval.IterativeCalc
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * GH-492 / GH-491: one iterative recalculation walks the SCC condensation in dependency-first
 * order, so every precedent is a FRESHLY COMPUTED value before anything reads it.
 *
 * Pre-fix the pass was split three ways (`preOrder` / one flat Jacobi over the whole cyclic core /
 * `postOrder`). An acyclic cell BETWEEN two SCCs is a transitive dependent of the first, so it
 * landed in `postOrder` and was not evaluated before the fixpoint — yet the second SCC read it
 * during iteration, off whatever cache the loaded workbook happened to carry. On a fresh book that
 * read recursively evaluated live and pass 1 was exact; on a CACHED book it returned the previous
 * generation's value, so the workbook walked one SCC of wavefront per pass (and, on interleaved
 * tax-bank/year-cycle topologies, parked in a stationary WRONG state that reported
 * `converged = false` forever with no error).
 *
 * The gates here are: (i) one pass reaches the global fixpoint from BOTH a fresh and an
 * already-cached input, (ii) re-recalculation is idempotent, (iii) per-component convergence
 * budgets and verdicts, (iv) determinism, (v) one volatile generation per iterative run.
 */
class GlobalFixpointSpec extends FunSuite:

  private def num(n: BigDecimal): CellValue = CellValue.Number(n)
  private def formula(expr: String): CellValue = CellValue.Formula(expr, None)
  private def aref(a1: String): ARef = ARef.parse(a1).fold(fail(_), identity)

  private val Tight = IterativeCalc(400, BigDecimal("1E-12"))
  private val Eps = BigDecimal("1E-6")

  private def cached(wb: Workbook, sheetName: String, a1: String): Option[CellValue] =
    wb.sheets
      .find(_.name.value == sheetName)
      .flatMap(_.cells.get(aref(a1)))
      .map(_.value)
      .collect { case CellValue.Formula(_, Some(v), _) => v }

  private def n(wb: Workbook, sheetName: String, a1: String): BigDecimal =
    cached(wb, sheetName, a1) match
      case Some(CellValue.Number(v)) => v
      case other => fail(s"$sheetName!$a1 must cache a number, got $other")

  private def assertClose(actual: BigDecimal, expected: BigDecimal, what: String): Unit =
    assert((actual - expected).abs < Eps, s"$what: got $actual, expected ~$expected")

  // ------------------------------------------------------------------
  // The interleaved fixture: year cycles with tax-bank SCCs BETWEEN them,
  // separated by acyclic bridge cells (the shape #491 reports; a single
  // chain of SCCs does NOT reproduce it).
  //
  //   A1 opening | A2 rate | A3 cash                    (constants)
  //   B1 = A1                                           acyclic
  //   SCC1 (year 1): B2 interest, B3 pay, B4 end
  //   C1 = B4*0.5                                       acyclic BRIDGE
  //   SCC2 (tax bank): D1 limit, D2 deduction, D3 carry  reads C1
  //   E1 = B4                                           acyclic BRIDGE
  //   SCC3 (year 2): E2 interest, E3 pay, E4 end         reads D2, E1
  //   F1 = E4+D3                                        acyclic headline
  // ------------------------------------------------------------------
  private def interleaved(opening: BigDecimal): Workbook =
    val s = Sheet(SheetName.unsafe("M"))
      .put(aref("A1"), num(opening))
      .put(aref("A2"), num(BigDecimal("0.08")))
      .put(aref("A3"), num(BigDecimal(50)))
      .put(aref("B1"), formula("=A1"))
      .put(aref("B2"), formula("=(B1+B4)/2*A2"))
      .put(aref("B3"), formula("=A3-B2"))
      .put(aref("B4"), formula("=B1-B3"))
      .put(aref("C1"), formula("=B4*0.5"))
      .put(aref("D1"), formula("=C1*0.3+D3*0.1"))
      .put(aref("D2"), formula("=MIN(D1,500)"))
      .put(aref("D3"), formula("=D2*0.4"))
      .put(aref("E1"), formula("=B4"))
      .put(aref("E2"), formula("=(E1+E4)/2*A2"))
      .put(aref("E3"), formula("=A3-E2-D2*0.01"))
      .put(aref("E4"), formula("=E1-E3"))
      .put(aref("F1"), formula("=E4+D3"))
    Workbook(s)

  /** The workbook's own internal-consistency law: end == beg - pay, for both year cycles. */
  private def assertSchedule(wb: Workbook, label: String): Unit =
    assertClose(n(wb, "M", "B4"), n(wb, "M", "B1") - n(wb, "M", "B3"), s"$label year-1 end=beg-pay")
    assertClose(n(wb, "M", "E4"), n(wb, "M", "E1") - n(wb, "M", "E3"), s"$label year-2 end=beg-pay")

  private def snapshot(wb: Workbook): Vector[(String, BigDecimal)] =
    Vector("B1", "B2", "B3", "B4", "C1", "D1", "D2", "D3", "E1", "E2", "E3", "E4", "F1")
      .map(a1 => a1 -> n(wb, "M", a1))

  // ================= GH-491: idempotence on an interleaved book =================

  test("GH-491: re-recalculating a solved interleaved multi-SCC book is stable and converged") {
    val fresh = interleaved(BigDecimal(1000)).recalculate(Tight)
    assert(fresh.converged, "the fresh solve must converge")
    assert(fresh.isClean, s"fresh solve errors: ${fresh.errors.map(_.render)}")
    assertSchedule(fresh.workbook, "fresh")

    // PERTURB THE DRIVER ON THE SOLVED BOOK. Re-solving a book that is ALREADY at its fixpoint
    // does not discriminate — the buggy pass operator is stationary there too, so the same five
    // passes below pass against the pre-GH-492 engine. Moving the opening balance on a fully
    // CACHED book is what exposes the split pass: the bridges C1/E1 (acyclic cells BETWEEN two
    // SCCs) landed in post-order, so the downstream SCCs iterated against the 1000-generation
    // caches and pass 1 parked one SCC of wavefront behind the answer.
    val perturbed = fresh.workbook.sheets
      .find(_.name.value == "M")
      .map(s => fresh.workbook.put(s.put(aref("A1"), num(BigDecimal(1200)))))
      .getOrElse(fail("sheet M"))
    val baseline = snapshot(interleaved(BigDecimal(1200)).recalculate(Tight).workbook)

    // Five successive re-solves starting from the perturbed cached book. EVERY one of them — the
    // first included — must already be at the 1200 global fixpoint, and stay there.
    val passes =
      (1 to 5).scanLeft(perturbed.recalculate(Tight))((r, _) => r.workbook.recalculate(Tight))
    passes.zipWithIndex.foreach { (r, i) =>
      assert(r.converged, s"pass ${i + 1} must report converged=true")
      assert(r.isClean, s"pass ${i + 1} errors: ${r.errors.map(_.render)}")
      assertSchedule(r.workbook, s"pass ${i + 1}")
      snapshot(r.workbook).zip(baseline).foreach { case ((a1, got), (_, want)) =>
        assert(
          (got - want).abs < Eps,
          s"pass ${i + 1} is not at the global fixpoint at $a1: $got vs $want"
        )
      }
    }
  }

  test("GH-491: one pass reaches the global fixpoint from BOTH a fresh and a cached input") {
    // Solve at opening = 1000, then perturb the driver on the SOLVED (fully cached) book.
    val solved = interleaved(BigDecimal(1000)).recalculate(Tight).workbook
    val perturbedCached = solved.sheets
      .find(_.name.value == "M")
      .map(s => solved.put(s.put(aref("A1"), num(BigDecimal(1200)))))
      .getOrElse(fail("sheet M"))

    val oneFromCached = perturbedCached.recalculate(Tight)
    val fromFresh = interleaved(BigDecimal(1200)).recalculate(Tight)

    assert(oneFromCached.converged, "the cached-input pass must converge")
    assertSchedule(oneFromCached.workbook, "cached-input")
    // Pre-fix: the bridges C1/E1 stayed frozen at the 1000-opening generation for this whole
    // pass, so D*/E*/F1 landed on the OLD answer and only pass 2 (or 3) caught up.
    snapshot(oneFromCached.workbook).zip(snapshot(fromFresh.workbook)).foreach {
      case ((a1, got), (_, want)) =>
        assert(
          (got - want).abs < Eps,
          s"one cached pass did not reach the global fixpoint at $a1: $got vs $want"
        )
    }
  }

  test("GH-492 twin exactness: a COLD re-solve of the fresh result reproduces it bit-for-bit") {
    // With cold (0) seeding both runs take the identical trajectory, so identical output is the
    // sharpest possible statement that the pass no longer depends on the input's loaded caches.
    // Pre-fix this failed: the second run's SCCs iterated against the first run's cached bridges.
    val cold = Tight.copy(seedFromCaches = false)
    val fresh = interleaved(BigDecimal(1000)).recalculate(cold)
    val again = fresh.workbook.recalculate(cold)
    assertEquals(again.workbook, fresh.workbook)
    assertEquals(again.evaluated, fresh.evaluated)
    assertEquals(again.errors.map(_.render), fresh.errors.map(_.render))
    assertEquals(again.cycles, fresh.cycles)
  }

  test("GH-492/GH-469: a WARM re-solve converges at round 1 and stays inside maxChange") {
    // Warm-started from its own output every component recognises the input as a fixpoint in one
    // round. It still writes that round's values, so the book creeps by strictly less than
    // maxChange — toward the true fixpoint, not away from it. That is the honest idempotence
    // claim: bit-exact for the cold path, |Δ| < maxChange for the warm one.
    val fresh = interleaved(BigDecimal(1000)).recalculate(Tight)
    val again = fresh.workbook.recalculate(Tight)
    assert(again.converged && again.certified)
    assertEquals(again.cycles.map(_.rounds), Vector(1, 1, 1), again.cycles.map(_.render).toString)
    assert(
      again.cycles.forall(_.maxDelta.exists(_ < Tight.maxChange)),
      again.cycles.map(_.render).toString
    )
    snapshot(again.workbook).zip(snapshot(fresh.workbook)).foreach { case ((a1, got), (_, want)) =>
      assert((got - want).abs < Tight.maxChange, s"$a1 moved by more than maxChange: $got vs $want")
    }
  }

  test("GH-492 bridge freshness: the bridge caches equal f(the SCC fixpoint) exactly") {
    val r = interleaved(BigDecimal(1000)).recalculate(Tight)
    val b4 = n(r.workbook, "M", "B4")
    assertEquals(n(r.workbook, "M", "C1"), b4 * BigDecimal("0.5"))
    assertEquals(n(r.workbook, "M", "E1"), b4)
    // Analytic year-1 fixpoint: B4*(1 - rate/2) = opening*(1 + rate/2) - cash
    val analytic = (BigDecimal(1000) * BigDecimal("1.04") - BigDecimal(50)) / BigDecimal("0.96")
    assertClose(b4, analytic, "year-1 ending balance")
  }

  // ================= GH-492: chain of SCCs, one pass =================

  /** Five year-cycles chained through acyclic bridges; every formula carries a STALE cache. */
  private def chainOfSccs(opening: BigDecimal, staleCache: BigDecimal): Workbook =
    val cols = Vector("B", "C", "D", "E", "F")
    val withCells = cols.zipWithIndex.foldLeft(
      Sheet(SheetName.unsafe("Y"))
        .put(aref("A1"), num(opening))
        .put(aref("A2"), num(BigDecimal("0.1")))
    ) { case (s, (c, i)) =>
      val prevEnd = if i == 0 then "A1" else s"${cols(i - 1)}4"
      s.put(aref(s"${c}1"), formula(s"=$prevEnd")) // beg (acyclic bridge)
        .put(aref(s"${c}2"), formula(s"=(${c}1+${c}4)/2*A2")) // interest (cyclic)
        .put(aref(s"${c}3"), formula(s"=20-${c}2")) // pay (cyclic)
        .put(aref(s"${c}4"), formula(s"=${c}1-${c}3")) // end (cyclic)
    }
    // Poison every formula cache: a wavefront engine needs one pass per SCC to burn this off.
    val poisoned = withCells.cells.foldLeft(withCells) { case (s, (r, cell)) =>
      cell.value match
        case f @ CellValue.Formula(_, _, _) => s.put(r, f.copy(cachedValue = Some(num(staleCache))))
        case _ => s
    }
    Workbook(poisoned)

  test("GH-492: a five-SCC chain with poisoned caches reaches the global fixpoint in ONE pass") {
    val poisoned = chainOfSccs(BigDecimal(1000), BigDecimal(-777))
    val one = poisoned.recalculate(Tight)
    val five = (1 to 5).foldLeft(one)((r, _) => r.workbook.recalculate(Tight))
    assert(one.converged, "one pass must certify convergence")
    Vector("B", "C", "D", "E", "F").foreach { c =>
      assertClose(n(one.workbook, "Y", s"${c}4"), n(five.workbook, "Y", s"${c}4"), s"${c}4")
      // and the schedule law inside every year
      assertClose(
        n(one.workbook, "Y", s"${c}4"),
        n(one.workbook, "Y", s"${c}1") - n(one.workbook, "Y", s"${c}3"),
        s"${c}: end=beg-pay"
      )
    }
  }

  // ================= GH-492: per-component budgets and verdicts =================

  /** Two independent components: A1/A2 contract (fixpoint 12.1212…), C1/C2 oscillate forever. */
  private def convergentPlusOscillator: Workbook =
    Workbook(
      Sheet(SheetName.unsafe("S"))
        .put(aref("A1"), formula("=A2*0.1+10"))
        .put(aref("A2"), formula("=A1*0.1+20"))
        .put(aref("C1"), formula("=1-C2"))
        .put(aref("C2"), formula("=C1"))
    )
  private val Budget = IterativeCalc(20, BigDecimal("1E-9"))

  test("GH-492: an exhausted component does not burn the budget of a convergent one") {
    val r = convergentPlusOscillator.recalculate(Budget)
    assertEquals(r.cycles.size, 2, s"expected two cyclic components, got ${r.cycles.map(_.render)}")
    assertEquals(r.errors, Vector.empty, "exhaustion is a data condition, never a CellEvalError")
    assert(!r.converged, "a globally unconverged run must say so")
    assertEquals(r.iterationsUsed, 20)
    assertEquals(r.unconverged.size, 1, s"exactly one component must be unconverged: ${r.cycles}")
    val bad = r.unconverged.headOption.getOrElse(fail("unconverged component"))
    assertEquals(bad.members.map((_, ref) => ref.toA1), Vector("C1", "C2"))
    assertEquals(bad.rounds, 20)
    assert(bad.maxDelta.exists(_ > BigDecimal(0)), s"an oscillator has a residual: ${bad.maxDelta}")
    assert(bad.render.contains("exhausted 20 round(s)"), bad.render)
    val good = r.cycles.filter(_.converged).headOption.getOrElse(fail("converged component"))
    assertEquals(good.members.map((_, ref) => ref.toA1), Vector("A1", "A2"))
    // The pin against the pre-GH-492 shared `members.forall` gate: one permanently-oscillating
    // member used to force the WHOLE core through every one of the maxIter rounds.
    assert(good.rounds < 20, s"the convergent component must stop early, took ${good.rounds}")
    // Exhausted members still cache their last values (Excel semantics)
    assert(cached(r.workbook, "S", "C1").isDefined, "exhausted members keep their last values")
    assertClose(n(r.workbook, "S", "A1"), BigDecimal("12.121212121"), "A1 fixpoint")
  }

  test("GH-492: summary fields are exactly the aggregate of the per-component verdicts") {
    val r = convergentPlusOscillator.recalculate(Budget)
    assertEquals(r.converged, r.cycles.forall(_.converged))
    assertEquals(r.iterationsUsed, r.cycles.map(_.rounds).maxOption.getOrElse(0))
    assertEquals(r.certified, r.errors.isEmpty && r.converged)
    val clean = interleaved(BigDecimal(1000)).recalculate(Tight)
    assertEquals(clean.converged, clean.cycles.forall(_.converged))
    assertEquals(clean.iterationsUsed, clean.cycles.map(_.rounds).maxOption.getOrElse(0))
    assert(clean.certified)
  }

  test("GH-492: non-iterative and acyclic runs report no cycles at all") {
    assertEquals(interleaved(BigDecimal(1000)).recalculate().cycles, Vector.empty)
    val acyclic = Workbook(
      Sheet(SheetName.unsafe("S"))
        .put(aref("A1"), num(BigDecimal(1)))
        .put(aref("B1"), formula("=A1*2"))
    )
    val r = acyclic.recalculate(IterativeCalc(100, BigDecimal("0.001")))
    assertEquals(r.cycles, Vector.empty)
    assert(r.converged)
    assertEquals(r.iterationsUsed, 0)
  }

  // ================= GH-492: determinism =================

  test("GH-492: the result is independent of sheet/cell insertion order") {
    def build(order: Vector[String]): Workbook =
      val cells: Map[String, CellValue] = Map(
        "A1" -> num(BigDecimal(1000)),
        "A2" -> num(BigDecimal("0.08")),
        "A3" -> num(BigDecimal(50)),
        "B1" -> formula("=A1"),
        "B2" -> formula("=(B1+B4)/2*A2"),
        "B3" -> formula("=A3-B2"),
        "B4" -> formula("=B1-B3"),
        "C1" -> formula("=B4*0.5"),
        "D1" -> formula("=C1*0.3+D3*0.1"),
        "D2" -> formula("=MIN(D1,500)"),
        "D3" -> formula("=D2*0.4")
      )
      Workbook(
        order.foldLeft(Sheet(SheetName.unsafe("M")))((s, a1) =>
          s.put(aref(a1), cells.getOrElse(a1, fail(s"missing $a1")))
        )
      )
    val keys = Vector("A1", "A2", "A3", "B1", "B2", "B3", "B4", "C1", "D1", "D2", "D3")
    val a = build(keys).recalculate(Tight)
    val b = build(keys.reverse).recalculate(Tight)
    val c = build(keys.sortBy(_.reverse)).recalculate(Tight)
    assertEquals(b.evaluated, a.evaluated)
    assertEquals(c.evaluated, a.evaluated)
    assertEquals(b.errors.map(_.render), a.errors.map(_.render))
    assertEquals(c.errors.map(_.render), a.errors.map(_.render))
    assertEquals(b.cycles, a.cycles)
    assertEquals(c.cycles, a.cycles)
    assertEquals(b.iterationsUsed, a.iterationsUsed)
  }

  test("GH-492: a seeded Rng makes a multi-SCC iterative run repeatable") {
    val wb = Workbook(
      Sheet(SheetName.unsafe("S"))
        .put(aref("A1"), formula("=A2*0.5+10+RAND()*0"))
        .put(aref("A2"), formula("=A1*0.5+20"))
        .put(aref("B1"), formula("=A1*2"))
        .put(aref("C1"), formula("=C2*0.5+B1+RAND()*0"))
        .put(aref("C2"), formula("=C1*0.25"))
    )
    val one = wb.recalculate(Clock.system, Rng.seeded(42), Tight)
    val two = wb.recalculate(Clock.system, Rng.seeded(42), Tight)
    assertEquals(two.evaluated, one.evaluated)
    assertEquals(two.errors.map(_.render), one.errors.map(_.render))
    assertEquals(two.cycles, one.cycles)
  }

  // ================= GH-492: one volatile generation =================

  test("GH-492: one iterative recalculation is ONE volatile generation, bridges included") {
    final class TickingClock extends Clock:
      private var calls = 0
      def today(): java.time.LocalDate =
        calls += 1
        java.time.LocalDate.of(2026, 1, 1).plusDays(calls.toLong)
      def now(): java.time.LocalDateTime =
        calls += 1
        java.time.LocalDateTime.of(2026, 1, 1, 0, 0).plusDays(calls.toLong)

    val wb = Workbook(
      Sheet(SheetName.unsafe("S"))
        .put(aref("A1"), formula("=B1*0+NOW()")) // SCC 1
        .put(aref("B1"), formula("=A1"))
        .put(aref("C1"), formula("=A1*0+NOW()")) // acyclic bridge between the SCCs
        .put(aref("D1"), formula("=E1*0+C1*0+NOW()")) // SCC 2
        .put(aref("E1"), formula("=D1"))
    )
    val r = wb.recalculate(new TickingClock, IterativeCalc(50, BigDecimal("0.001")))
    val a = cached(r.workbook, "S", "A1")
    val c = cached(r.workbook, "S", "C1")
    val d = cached(r.workbook, "S", "D1")
    assert(a.isDefined && c.isDefined && d.isDefined, s"all three must cache: $a / $c / $d")
    assertEquals(c, a, "the acyclic bridge must see the same volatile generation as SCC 1")
    assertEquals(d, a, "SCC 2 must see the same volatile generation as SCC 1")
  }

  // ================= GH-492: dynamic bucket upstream of an SCC (§3.3) =================

  test("GH-492: an INDIRECT cell upstream of a cyclic component defers WITH its component") {
    val s = Sheet(SheetName.unsafe("S"))
      .put(aref("A1"), num(BigDecimal(5)))
      .put(aref("Z1"), CellValue.Formula("=INDIRECT(\"A1\")*2", Some(num(BigDecimal(999)))))
      .put(aref("B1"), formula("=C1*0.5+Z1"))
      .put(aref("C1"), formula("=B1*0.5"))
    val r = Workbook(s).recalculate(Tight)
    assert(r.isClean, s"expected clean, got ${r.errors.map(_.render)}")
    assertClose(n(r.workbook, "S", "Z1"), BigDecimal(10), "Z1 must recompute from INDIRECT")
    // B1 = 0.5*C1 + 10, C1 = 0.5*B1  =>  B1 = 40/3
    assertClose(n(r.workbook, "S", "B1"), BigDecimal("13.3333333"), "B1")
    assertClose(n(r.workbook, "S", "C1"), BigDecimal("6.66666666"), "C1")
  }

  // ================= must-not-move regressions =================

  test("GH-492: the non-iterative path on a cyclic workbook is unchanged") {
    val wb = interleaved(BigDecimal(1000))
    val r = wb.recalculate()
    // Every cyclic member reports circular; its dependents report blocked; nothing caches.
    assertEquals(
      r.errors.filter(_.error.message.contains("Circular")).map(_.ref.toA1).sorted,
      Vector("B2", "B3", "B4", "D1", "D2", "D3", "E2", "E3", "E4")
    )
    assertEquals(
      r.errors.filter(_.error.message.contains("Blocked")).map(_.ref.toA1).sorted,
      Vector("C1", "E1", "F1")
    )
    assertEquals(cached(r.workbook, "M", "B1"), Some(num(BigDecimal(1000))))
    assertEquals(cached(r.workbook, "M", "B4"), None)
    assertEquals(r.cycles, Vector.empty)
    assert(r.converged)
    assertEquals(r.iterationsUsed, 0)
    // Deterministic and repeatable
    val again = wb.recalculate()
    assertEquals(again.workbook, r.workbook)
    assertEquals(again.errors.map(_.render), r.errors.map(_.render))
  }

  test("GH-430: an iterative recalculation leaves data-table caches byte-identical") {
    val tableKind = FormulaKind.DataTable(
      ref = CellRange.parse("B10:B11").fold(fail(_), identity),
      dt2D = false,
      dtr = false,
      r1 = Some(aref("A1")),
      r2 = None
    )
    val s = Sheet(SheetName.unsafe("S"))
      .put(aref("A1"), num(BigDecimal(3)))
      .put(aref("A2"), formula("=A3*0.5+10"))
      .put(aref("A3"), formula("=A2*0.5+20"))
      .put(aref("B10"), CellValue.Formula("=TABLE(A1,)", Some(num(BigDecimal(42))), tableKind))
      .put(aref("B11"), CellValue.Formula("=TABLE(A1,)", None, tableKind))
    val r = Workbook(s).recalculate(Tight)
    assertEquals(cached(r.workbook, "S", "B10"), Some(num(BigDecimal(42))))
    assertEquals(cached(r.workbook, "S", "B11"), None)
    assert(
      r.cycles.forall(_.members.forall((_, ref) => ref != aref("B10") && ref != aref("B11"))),
      s"data-table records must never be fixpoint members: ${r.cycles.map(_.render)}"
    )
  }
