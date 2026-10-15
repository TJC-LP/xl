package com.tjclp.xl.formula.perf

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.eval.RecalcCounters.{Counter, Phase}
import com.tjclp.xl.formula.perf.LoadsBookHarness.Mode
import munit.FunSuite

/**
 * #767 benchmark gates, characterisation mode: on `Tiny` only, every counter is pinned to the value
 * today's engine produces, and the 40 oracle cells to the values it computes. A PR that changes the
 * algorithm updates the pinned counts in its own diff; a PR that changes a value fails the oracle.
 * The fork's `-Xmx2g` ceiling is the heap gate. Wall-clock gates and the larger shapes switch on
 * with the criteria index (design §12.3).
 */
class RecalcPerfGatesSpec extends FunSuite:

  private val tiny = LoadsBook.build(LoadsBook.Tiny)

  private val oracle: Vector[(String, String)] =
    val stream = getClass.getResourceAsStream("/loads-oracle-tiny.tsv")
    val text =
      try new String(stream.readAllBytes(), java.nio.charset.StandardCharsets.UTF_8)
      finally stream.close()
    LoadsBookHarness.parseOracle(text)

  private def runTiny(mode: Mode): LoadsBookHarness.Run =
    val run = LoadsBookHarness.run(tiny.workbook, mode)
    println(LoadsBookHarness.report(s"Tiny mode=$mode", run))
    run

  private lazy val recalc = runTiny(Mode.Recalc)
  private lazy val checked = runTiny(Mode.Checked)
  private lazy val parallel = runTiny(Mode.Parallel)

  /** Today's engine on Tiny: the legacy per-cell graph, every criteria call a full scan. */
  private val TinyCounts: Map[Counter, Long] = Map(
    Counter.GraphNodes -> 1600L,
    Counter.GraphEdges -> 186400L,
    Counter.VirtualNodes -> 0L,
    Counter.RangeNodes -> 0L,
    Counter.CriteriaIndexed -> 0L,
    Counter.CriteriaWindowOnly -> 0L,
    Counter.CriteriaNotIndexable -> 0L,
    Counter.CriteriaScan -> 400L,
    Counter.MatcherCalls -> 180118L,
    Counter.SlicesBuilt -> 0L,
    Counter.ScanFallbacks -> 0L,
    Counter.UsedRangeScans -> 0L,
    Counter.NameParses -> 3456L
  )

  private def assertOracle(run: LoadsBookHarness.Run): Unit =
    assertEquals(oracle.size, 40)
    assertEquals(oracle.map(_._1), tiny.sampleRefs)
    oracle.foreach { case (ref, expected) =>
      assertEquals(LoadsBookHarness.sampleValue(run.result.workbook, ref), expected, ref)
    }

  test("Tiny recalculate: 40/40 oracle cells, no errors") {
    assertEquals(recalc.result.errors, Vector.empty)
    assertOracle(recalc)
  }

  test("Tiny writeChecked and parallel recalculate: the same 40 oracle cells") {
    assertEquals(checked.result.errors, Vector.empty)
    assertOracle(checked)
    assertEquals(parallel.result.errors, Vector.empty)
    assertOracle(parallel)
  }

  test("Tiny recalculate: counters characterise today's engine") {
    assertEquals(recalc.counters.counters, TinyCounts)
  }

  test("Tiny writeChecked and parallel recalculate count the same work as recalculate") {
    assertEquals(checked.counters.counters, TinyCounts)
    assertEquals(parallel.counters.counters, TinyCounts)
  }

  test("scanFallbacks is 0 on recalculate and on writeChecked") {
    assertEquals(recalc.counters(Counter.ScanFallbacks), 0L)
    assertEquals(checked.counters(Counter.ScanFallbacks), 0L)
  }

  test("every phase of a measured pass is timed; Write only when the file is written") {
    Vector(Phase.Parse, Phase.GraphBuild, Phase.Tarjan, Phase.Kahn, Phase.Eval, Phase.Criteria)
      .foreach(p => assert(recalc.counters.nanos(p) > 0L, s"$p not timed"))
    assertEquals(recalc.counters.nanos(Phase.Write), 0L)
    assert(checked.counters.nanos(Phase.Write) > 0L, "Write not timed")
  }

  test("measuring changes nothing: counters off and on give the same workbook") {
    val unmeasured = LoadsBookHarness.run(tiny.workbook, Mode.Recalc, measured = false)
    assertEquals(unmeasured.result.workbook, recalc.result.workbook)
    assertEquals(unmeasured.result.errors, recalc.result.errors)
    assert(unmeasured.counters.counters.values.forall(_ == 0L))
  }

  test("whole-column names: the same oracle cells, every criteria range bounded by a scan") {
    val whole = LoadsBook.build(LoadsBook.Tiny.copy(wholeColumns = true))
    val run = LoadsBookHarness.run(whole.workbook, Mode.Recalc)
    println(LoadsBookHarness.report("Tiny wholeColumns", run))
    assertOracle(run)
    // every criteria call scans the used range once per range it reads (value range plus criteria
    // ranges): 34 per Summary row x 8 rows + 64 per Rate History row x 24 rows
    assertEquals(run.counters(Counter.UsedRangeScans), 34L * 8 + 64L * 24)
    assertEquals(run.counters(Counter.CriteriaScan), 400L)
  }

  test("RunTotal: running totals equal the closed form; the legacy graph is quadratic") {
    val rows = 300
    val run = LoadsBookHarness.run(LoadsBook.runTotal(rows), Mode.Recalc)
    assertEquals(run.result.errors, Vector.empty)
    val sheet = run.result.workbook.sheets.find(_.name == SheetName.unsafe("Run"))
    Vector(2, 3, 150, rows + 1).foreach { r =>
      val value = sheet.flatMap(_.cells.get(ARef.from0(3, r - 1))).map(_.value)
      val expected: Option[CellValue] =
        Some(CellValue.Number(LoadsBook.runTotalAt(r)))
      assertEquals(value.map(LoadsBookHarness.show), expected.map(LoadsBookHarness.show), s"D$r")
    }
    assertEquals(run.counters(Counter.GraphNodes), 2L * rows)
    // D_n reads C_2..C_n: 1 + 2 + ... + rows edges
    assertEquals(run.counters(Counter.GraphEdges), rows.toLong * (rows + 1) / 2)
  }
