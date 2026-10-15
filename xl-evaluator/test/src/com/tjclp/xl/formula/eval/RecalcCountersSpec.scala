package com.tjclp.xl.formula.eval

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.Clock
import com.tjclp.xl.formula.eval.RecalcCounters.{Counter, Phase}
import com.tjclp.xl.formula.eval.WorkbookEvaluator.*
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/** #767: the recalculation work counters count exactly and never change a result. */
class RecalcCountersSpec extends FunSuite:

  private val data = SheetName.unsafe("Data")
  private val out = SheetName.unsafe("Out")
  private val fixedClock = Clock.fixedDate(java.time.LocalDate.of(2026, 1, 2))
  private val options = RecalcOptions(clock = fixedClock, iterative = IterativeMode.Off)

  private def num(n: Int): CellValue = CellValue.Number(BigDecimal(n))
  private def ref(col: Int, row: Int): ARef = ARef.from0(col, row)

  /**
   * Data!A1:A5 = 1..5, Data!B1:B5 = x y x x y, `rng` = Data!$A$1:$A$5. Out!A1:A5 hold four criteria
   * calls and one plain formula that reads two of them.
   */
  private val book: Workbook =
    val dataSheet = (0 until 5).foldLeft(Sheet(data)) { (s, i) =>
      s.put(ref(0, i), num(i + 1))
        .put(ref(1, i), CellValue.Text(Vector("x", "y", "x", "x", "y")(i)))
    }
    val formulas = Vector(
      """COUNTIF(Data!A1:A5,">2")""",
      """COUNTIFS(Data!A1:A5,">2",Data!B1:B5,"x")""",
      """SUMIF(rng,">2")""",
      """COUNTIF(Data!A:A,">2")""",
      "A1+A2"
    )
    val outSheet = formulas.zipWithIndex.foldLeft(Sheet(out)) { case (s, (f, i)) =>
      s.put(ref(0, i), CellValue.Formula(f))
    }
    Workbook(Vector(dataSheet, outSheet)).withDefinedName("rng", "Data!$A$1:$A$5")

  private def assertSame(actual: RecalcResult, expected: RecalcResult): Unit =
    assertEquals(actual.workbook, expected.workbook)
    assertEquals(actual.evaluated, expected.evaluated)
    assertEquals(actual.errors, expected.errors)
    assertEquals(actual.converged, expected.converged)

  test("a measured recalculation returns exactly what the public entry point returns") {
    assertSame(
      recalculateCounted(book, options, RecalcCounters.create()),
      book.recalculate(options)
    )
    val parallel = options.copy(parallelism = 2)
    assertSame(
      recalculateCounted(book, parallel, RecalcCounters.create()),
      book.recalculate(parallel)
    )
    assertSame(
      recalculateUncachedCounted(book, options, RecalcCounters.create()),
      book.recalculateUncached(options)
    )
  }

  test("the results the counters observe are the expected values") {
    val result = recalculateCounted(book, options, RecalcCounters.create())
    val values = result.evaluated.getOrElse(out, Map.empty)
    assertEquals(
      (0 until 5).map(i => values.get(ref(0, i))).toVector,
      Vector(3, 2, 12, 3, 5).map(n => Some(num(n)))
    )
  }

  test("counts graph size, criteria calls, matcher calls, used-range scans and name parses") {
    val counters = RecalcCounters.create()
    recalculateCounted(book, options, counters)
    val snap = counters.snapshot
    assertEquals(snap(Counter.GraphNodes), 5L)
    // compact graph: point edges Out!A1 -> A5 and Out!A2 -> A5, three range nodes (Data!A1:A5,
    // which `rng` resolves to as well, Data!B1:B5, Data!A:A) and five reader edges; Data holds no
    // formula, so no segment node
    assertEquals(snap(Counter.GraphEdges), 7L)
    assertEquals(snap(Counter.RangeNodes), 3L)
    // Tarjan walks each edge once, Kahn twice (in-degrees, then the order)
    assertEquals(snap(Counter.GraphWork), 3L * 7L)
    assertEquals(snap(Counter.CriteriaScan), 4L)
    // COUNTIF 5 rows; COUNTIFS 5 rows on the first pair, then the 3 rows that passed it;
    // SUMIF 5; COUNTIF(Data!A:A) 5 rows after the used-range bound
    assertEquals(snap(Counter.MatcherCalls), 5L + 8L + 5L + 5L)
    assertEquals(snap(Counter.UsedRangeScans), 1L)
    // SUMIF resolves `rng` once as its range and once as its (absent, defaulted) sum range
    assertEquals(snap(Counter.NameParses), 2L)
    Vector(
      Counter.VirtualNodes,
      Counter.CriteriaIndexed,
      Counter.CriteriaWindowOnly,
      Counter.CriteriaNotIndexable,
      Counter.SlicesBuilt,
      Counter.ScanFallbacks
    ).foreach(c => assertEquals(snap(c), 0L, c.toString))
  }

  test("the legacy engine counts formula-to-formula edges and no virtual work") {
    val counters = RecalcCounters.create()
    recalculateCounted(book, options, counters, RecalcEngine.Legacy)
    val snap = counters.snapshot
    assertEquals(snap(Counter.GraphNodes), 5L)
    // Out!A5 reads Out!A1 and Out!A2; the criteria calls read constants only
    assertEquals(snap(Counter.GraphEdges), 2L)
    Vector(Counter.VirtualNodes, Counter.RangeNodes, Counter.GraphWork)
      .foreach(c => assertEquals(snap(c), 0L, c.toString))
  }

  test("parallel and uncached passes count the same work") {
    def counted(run: RecalcCounters => RecalcResult): Map[Counter, Long] =
      val counters = RecalcCounters.create()
      run(counters)
      counters.snapshot.counters
    val sequential = counted(recalculateCounted(book, options, _))
    assertEquals(counted(recalculateCounted(book, options.copy(parallelism = 2), _)), sequential)
    assertEquals(counted(recalculateUncachedCounted(book, options, _)), sequential)
  }

  test("phases are timed when measured, and the disabled instance never records anything") {
    val counters = RecalcCounters.create()
    recalculateCounted(book, options, counters)
    val snap = counters.snapshot
    Vector(Phase.Parse, Phase.GraphBuild, Phase.Kahn, Phase.Eval, Phase.Criteria).foreach(p =>
      assert(snap.nanos(p) > 0L, s"$p was not timed")
    )
    assertEquals(snap.nanos(Phase.Write), 0L)

    book.recalculate(options)
    book.recalculateUncached(options)
    val disabled = RecalcCounters.disabled.snapshot
    assert(disabled.counters.values.forall(_ == 0L), disabled.render)
    assert(disabled.phaseNanos.values.forall(_ == 0L), disabled.render)
  }

  test("a tally adds to its counter once, on flush") {
    val counters = RecalcCounters.create()
    val tally = counters.tally(Counter.MatcherCalls)
    (1 to 7).foreach(_ => tally.increment())
    assertEquals(counters.snapshot(Counter.MatcherCalls), 0L)
    tally.flush()
    assertEquals(counters.snapshot(Counter.MatcherCalls), 7L)
    tally.flush()
    assertEquals(counters.snapshot(Counter.MatcherCalls), 7L)
  }

  test("render lists every counter and phase") {
    val lines = RecalcCounters.create().snapshot.render.linesIterator.toVector
    assertEquals(lines.size, Counter.values.length + Phase.values.length)
    assert(lines.contains("MatcherCalls=0"))
    assert(lines.contains("EvalSeconds=0.000"))
  }
