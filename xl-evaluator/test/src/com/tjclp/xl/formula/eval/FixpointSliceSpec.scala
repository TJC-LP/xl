package com.tjclp.xl.formula.eval

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.{Clock, Rng}
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * #891: one iterative fixpoint shares a column slice and equality index per criteria range that
 * holds no member ([[Evaluator.AggregateMemo.fixpointCache]]). Law: every outcome equals the
 * fixpoint's outcome with the cache switched off, under both schemes, whether a criteria range
 * holds no member (columns A and B: sliced), a member (column C: scanned, it changes during the
 * fixpoint) or an uncached formula outside the fixpoint (column D, the criteria range of C5:
 * scanned, the read evaluates it). Only criteria ranges are sliced: C3 reads column D as a sum
 * range, which no slice serves. `FixpointCacheLawSpec` states the same law over generated
 * fixpoints.
 */
class FixpointSliceSpec extends FunSuite:

  private val s = SheetName.unsafe("S")
  private def at(c: Int, r: Int): ARef = ARef.from0(c, r)
  private def num(n: Int): CellValue = CellValue.Number(BigDecimal(n))
  private def text(t: String): CellValue = CellValue.Text(t)
  private def formula(f: String): CellValue = CellValue.Formula(f, None)

  // Columns A (keys) and B (amounts) are constants over 40 rows; column C holds members; column D
  // holds an uncached formula outside the cycle.
  private val rows = 40
  private val base: Sheet = (0 until rows).foldLeft(Sheet(s)) { (sheet, r) =>
    sheet
      .put(at(0, r), text(if r % 3 == 0 then "x" else if r % 3 == 1 then "y" else "z"))
      .put(at(1, r), num(r + 1))
      .put(at(3, r), formula(s"=B${r + 1}*2"))
  }

  // The cycle: C1..C5 read each other through criteria calls over each kind of column.
  private val memberFormulas = Vector(
    at(2, 0) -> "=SUMIFS(B1:B40,A1:A40,\"x\")/100+C2*0.5",
    at(2, 1) -> "=COUNTIFS(A1:A40,\"y\",B1:B40,\">10\")+C3*0.25",
    at(2, 2) -> "=SUMIF(C1:C40,\">0\")*0.1+SUMIFS(D1:D40,A1:A40,\"z\")/1000",
    at(2, 3) -> "=AVERAGEIFS(B1:B40,A1:A40,\"x\",B1:B40,\"<30\")+C1*0.1",
    // criteria over the uncached formulas of column D (B*2: 2..80)
    at(2, 4) -> "=COUNTIF(D1:D40,\">40\")/10+SUMIFS(B1:B40,D1:D40,\"<=20\")/100+C4*0.1"
  )

  private val sheet: Sheet =
    memberFormulas.foldLeft(base)((acc, m) => acc.put(m._1, formula(m._2)))

  private val members: List[(QualifiedRef, Int, String)] =
    memberFormulas.toList.map((ref, f) => (QualifiedRef(s, ref), 0, f))

  private def run(
    scheme: IterationScheme,
    columnCache: Boolean
  ): WorkbookEvaluator.FixpointOutcome =
    val wb = Workbook(sheet)
    WorkbookEvaluator.jacobiFixpoint(
      wb,
      wb.sheets,
      members,
      IterativeCalc(100, BigDecimal("1E-12"), scheme = scheme),
      Clock.system,
      Rng.system,
      Evaluator.recalculationInstance(_, _),
      Map.empty,
      columnCache = columnCache
    )

  List(IterationScheme.Jacobi, IterationScheme.GaussSeidel).foreach { scheme =>
    test(s"#891: $scheme fixpoint with the shared criteria cache equals the uncached fixpoint") {
      val cached = run(scheme, columnCache = true)
      val uncached = run(scheme, columnCache = false)
      assertEquals(cached.results, uncached.results)
      assertEquals(cached.converged, uncached.converged)
      assertEquals(cached.rounds, uncached.rounds)
      assert(cached.converged, cached.toString)
      assert(cached.results.values.forall(_.isRight), cached.results.toString)
    }
  }

  test("#891: a fixpoint readiness refuses boxes holding members and nothing else") {
    val ready = Readiness.fixpoint(members.map(_._1))
    val anyPosition = EvalPosition.None
    assert(ready.allows(s, 0, 0, rows - 1, anyPosition))
    assert(ready.allows(s, 2, 10, rows - 1, anyPosition))
    assert(!ready.allows(s, 2, 0, rows - 1, anyPosition))
    assert(!ready.allows(s, 2, 3, 3, EvalPosition.at(5)))
    assert(!Readiness.build(Map.empty, Set.empty).allows(s, 0, 0, 1, anyPosition))
  }

  test("#891: the fixpoint cache slices constant columns and refuses member or uncached columns") {
    val cache = Evaluator.AggregateMemo.fixpointCache(members.map(_._1))
    def answer(col: Int) = cache.slice(sheet, col, 0, rows - 1, EvalPosition.None)
    assert(answer(0).isInstanceOf[PassCache.SliceAnswer.Use], answer(0).toString)
    assert(answer(1).isInstanceOf[PassCache.SliceAnswer.Use], answer(1).toString)
    assertEquals(answer(2), PassCache.SliceAnswer.Refused)
    assertEquals(answer(3), PassCache.SliceAnswer.Refused)
  }

  test("#891: the law's fixpoint slices A and B and nothing of C or D, whose calls then scan") {
    // the law above covers the refusals only if the fixpoint really asks for them
    List(IterationScheme.Jacobi, IterationScheme.GaussSeidel).foreach { scheme =>
      val counters = RecalcCounters.create()
      val cache = new PassCache(Evaluator.AggregateMemo.FixpointLimits, counters)
      cache.begin(Readiness.fixpoint(members.map(_._1)))
      val wb = Workbook(sheet)
      val outcome = WorkbookEvaluator.jacobiFixpoint(
        wb,
        wb.sheets,
        members,
        IterativeCalc(100, BigDecimal("1E-12"), scheme = scheme),
        Clock.system,
        Rng.system,
        (rng, _) =>
          Evaluator.recalculationInstance(
            rng,
            new Evaluator.AggregateMemo(counters = counters, shared = Some(cache))
          ),
        Map.empty
      )
      assertEquals(outcome, run(scheme, columnCache = false))
      assertEquals(cache.builtSlices.map(_.col).sorted, Vector(0, 1))
      // C3's SUMIF over column C and C5's two calls over column D, every round
      val snap = counters.snapshot
      assert(snap(RecalcCounters.Counter.ScanFallbacks) >= 3L * outcome.rounds, snap.render)
    }
  }
