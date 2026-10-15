package com.tjclp.xl.formula.perf

import com.tjclp.xl.formula.eval.RecalcCounters.Counter
import com.tjclp.xl.formula.perf.LoadsBookHarness.Mode
import munit.FunSuite

/**
 * #767 criteria-index gates (design §12.3) on `R20`: 20k Loads rows, raw criteria, 8,000 *IFS
 * formulas over whole-column names. Lanes and carriers scale with the rows, so an equality bucket
 * stays about the same size as the book grows while a scan grows with it; the counter gates tell
 * the two apart deterministically. Wall clock is a coarse backstop.
 *
 * The `H10` gates of the design wait for the compact graph: on the legacy per-cell graph `H10`
 * needs about 25 GB of heap, past this module's `-Xmx2g`.
 */
class CriteriaIndexGatesSpec extends FunSuite:

  override val munitTimeout = scala.concurrent.duration.Duration(600, "s")

  private val r20 = LoadsBook.build(LoadsBook.R20)
  private val r5 = LoadsBook.build(LoadsBook.R20.copy(name = "R5", rows = 5000))

  private def resource(name: String): String =
    val stream = getClass.getResourceAsStream(name)
    try new String(stream.readAllBytes(), java.nio.charset.StandardCharsets.UTF_8)
    finally stream.close()

  private val oracle: Vector[(String, String)] =
    LoadsBookHarness.parseOracle(resource("/loads-oracle-r20.tsv"))

  private def runR20(mode: Mode, criteriaIndex: Boolean = true): LoadsBookHarness.Run =
    val run = LoadsBookHarness.run(r20.workbook, mode, criteriaIndex = criteriaIndex)
    println(LoadsBookHarness.report(s"R20 mode=$mode index=$criteriaIndex", run))
    run

  private lazy val recalc = runR20(Mode.Recalc)
  private lazy val checked = runR20(Mode.Checked)

  private def assertOracle(run: LoadsBookHarness.Run): Unit =
    assertEquals(oracle.size, 40)
    assertEquals(oracle.map(_._1), r20.sampleRefs)
    oracle.foreach { case (ref, expected) =>
      assertEquals(LoadsBookHarness.sampleValue(run.result.workbook, ref), expected, ref)
    }

  test("R20 recalculate and writeChecked: 40/40 oracle cells, no errors") {
    assertEquals(recalc.result.errors, Vector.empty)
    assertOracle(recalc)
    assertEquals(checked.result.errors, Vector.empty)
    assertOracle(checked)
  }

  test("R20: every criteria call is indexed, none falls back to a scan") {
    Vector(recalc, checked).foreach { run =>
      assertEquals(
        run.counters(Counter.CriteriaIndexed),
        r20.summaryFormulas + r20.rhFormulas.toLong
      )
      assertEquals(run.counters(Counter.CriteriaScan), 0L)
      assertEquals(run.counters(Counter.ScanFallbacks), 0L)
    }
  }

  test("R20: at most one slice per criteria column") {
    // the criteria ranges are nine whole-column names (every ld_* name but ld_miles)
    assert(recalc.counters(Counter.SlicesBuilt) <= 9L, recalc.counters.render)
  }

  test("matcher calls per indexed call stay flat from 5k to 20k rows") {
    val small = LoadsBookHarness.run(r5.workbook, Mode.Recalc)
    println(LoadsBookHarness.report("R5 mode=Recalc", small))
    def perCall(run: LoadsBookHarness.Run): Double =
      run.counters(Counter.MatcherCalls).toDouble / run.counters(Counter.CriteriaIndexed).toDouble
    assertEquals(small.counters(Counter.CriteriaIndexed), r5.summaryFormulas + r5.rhFormulas.toLong)
    val ratio = perCall(recalc) / perCall(small)
    // a scan grows 4x with the rows; the index's buckets stay about the same size
    assert(ratio <= 1.5, f"per-call matcher calls grew $ratio%.2fx from 5k to 20k rows")
  }

  test("R20: the index tests at most a twentieth of the rows the scan tests") {
    val off = runR20(Mode.Recalc, criteriaIndex = false)
    assertEquals(off.result.workbook, recalc.result.workbook)
    assertEquals(off.counters(Counter.CriteriaIndexed), 0L)
    val (on, scanned) = (recalc.counters(Counter.MatcherCalls), off.counters(Counter.MatcherCalls))
    assert(on * 20L <= scanned, s"index on: $on matcher calls; off: $scanned")
  }

  test("R20 recalculate: under 30 s, the fastest of three after a warm-up") {
    val times = (1 to 3).map(_ => LoadsBookHarness.run(r20.workbook, Mode.Recalc).wallSeconds)
    println(f"R20 recalc wall: ${times.map(t => f"$t%.2f").mkString(", ")} s")
    assert(times.min <= 30.0, s"R20 recalc took ${times.min} s at best")
  }
