package com.tjclp.xl.formula.perf

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.eval.{IterativeMode, RecalcCounters, RecalcEngine, RecalcOptions}
import com.tjclp.xl.formula.eval.RecalcCounters.{Counter, Phase}
import com.tjclp.xl.formula.eval.WorkbookEvaluator
import com.tjclp.xl.formula.perf.LoadsBookHarness.Mode
import munit.FunSuite

/**
 * #767 benchmark gates on `Tiny`, characterisation mode: every counter is pinned to the value
 * today's engine produces, and the 40 oracle cells to the values it computes. A PR that changes the
 * algorithm updates the pinned counts in its own diff; a PR that changes a value fails the oracle.
 * The fork's `-Xmx2g` ceiling is the heap gate. The `R20` gates are in [[CriteriaIndexGatesSpec]].
 */
class RecalcPerfGatesSpec extends FunSuite:

  private val tiny = LoadsBook.build(LoadsBook.Tiny)

  private def resource(name: String): String =
    val stream = getClass.getResourceAsStream(name)
    try new String(stream.readAllBytes(), java.nio.charset.StandardCharsets.UTF_8)
    finally stream.close()

  private val oracle: Vector[(String, String)] =
    LoadsBookHarness.parseOracle(resource("/loads-oracle-tiny.tsv"))

  /** Every formula cell of Tiny as 0.24.1 computes it (1,600 lines, 765 of them non-zero). */
  private val fullOracle: Vector[String] =
    resource("/loads-values-tiny.tsv").linesIterator.filter(_.nonEmpty).toVector

  private def runTiny(mode: Mode): LoadsBookHarness.Run =
    val run = LoadsBookHarness.run(tiny.workbook, mode)
    println(LoadsBookHarness.report(s"Tiny mode=$mode", run))
    run

  private lazy val recalc = runTiny(Mode.Recalc)
  private lazy val checked = runTiny(Mode.Checked)
  private lazy val parallel = runTiny(Mode.Parallel)

  /**
   * Today's engine on Tiny: the compact graph (11 range nodes, 1,197 segment nodes over the three
   * helper columns, 5,001 edges where the legacy per-cell graph had 186,400; Tarjan and Kahn walk
   * them three times); the per-pass column slices (#767 PR 1), one per criteria column the book
   * reads (8 of the Loads columns), one used-range scan (the Loads sheet, sizing those slices), and
   * one parse per distinct defined-name body; every criteria call answered from the equality index
   * (#767 PR 2), testing only its candidate rows (180,118 matcher calls when every row was tested).
   */
  private val TinyCounts: Map[Counter, Long] = Map(
    Counter.GraphNodes -> 1600L,
    Counter.GraphEdges -> 5001L,
    Counter.VirtualNodes -> 1197L,
    Counter.RangeNodes -> 11L,
    Counter.GraphWork -> 15003L,
    Counter.CriteriaIndexed -> 400L,
    Counter.CriteriaWindowOnly -> 0L,
    Counter.CriteriaNotIndexable -> 0L,
    Counter.CriteriaScan -> 0L,
    Counter.MatcherCalls -> 1478L,
    Counter.SlicesBuilt -> 8L,
    Counter.ScanFallbacks -> 0L,
    Counter.UsedRangeScans -> 1L,
    Counter.NameParses -> 15L,
    // #790: one parse per formula, made by the graph build and reused by evaluation
    Counter.FormulaParses -> 1600L
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

  test("Tiny recalculate, writeChecked and parallel: every formula cell equals the full oracle") {
    assertEquals(fullOracle.size, tiny.totalFormulas)
    Vector(recalc, checked, parallel).foreach { run =>
      val actual = LoadsBookHarness.valueLines(run.result.workbook)
      val diff = actual.zip(fullOracle).filter(_ != _).take(5)
      assertEquals(diff, Vector.empty)
      assertEquals(actual.size, fullOracle.size)
    }
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

  test("whole-column names: the same values, one used-range scan for the whole pass") {
    val whole = LoadsBook.build(LoadsBook.Tiny.copy(wholeColumns = true))
    val run = LoadsBookHarness.run(whole.workbook, Mode.Recalc)
    println(LoadsBookHarness.report("Tiny wholeColumns", run))
    assertOracle(run)
    // the header row the whole columns add is text, which no criterion and no MIN/MAX picks up
    assertEquals(LoadsBookHarness.valueLines(run.result.workbook), fullOracle)
    // the whole-column ranges are all on Loads, whose populated cells never change during the
    // pass: one scan (before the per-pass cache, one per range per call: 34 per Summary row x 8
    // rows + 64 per Rate History row x 24 rows = 1,808)
    assertEquals(run.counters(Counter.UsedRangeScans), 1L)
    assertEquals(run.counters(Counter.CriteriaIndexed), 400L)
    assertEquals(run.counters(Counter.ScanFallbacks), 0L)
  }

  test("criteria index on and off: the same workbook and errors on Tiny, in every mode") {
    Vector(Mode.Recalc, Mode.Parallel, Mode.Uncached).foreach { mode =>
      val on = LoadsBookHarness.run(tiny.workbook, mode)
      val off = LoadsBookHarness.run(tiny.workbook, mode, criteriaIndex = false)
      assertEquals(on.result.workbook, off.result.workbook, mode.toString)
      assertEquals(on.result.errors, off.result.errors, mode.toString)
      assertEquals(off.counters(Counter.CriteriaIndexed), 0L)
      assertEquals(off.counters(Counter.MatcherCalls), 180118L)
    }
  }

  test("RunTotal: running totals equal the closed form; the compact graph is O(n log n)") {
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
    assertEquals(run.counters(Counter.RangeNodes), rows.toLong)
    // C's segment tree (rows - 1 nodes, 2 * rows - 2 edges), at most 2 log2(rows) canonical
    // pieces per D_n, one reader edge each
    val log2 = 32 - Integer.numberOfLeadingZeros(rows)
    assert(run.counters(Counter.GraphEdges) <= 2L * rows * (log2 + 2), run.counters.render)
    // the legacy graph: D_n reads C_2..C_n, 1 + 2 + ... + rows edges
    val legacy = RecalcCounters.create()
    WorkbookEvaluator.recalculateCounted(
      LoadsBook.runTotal(rows),
      RecalcOptions(iterative = IterativeMode.Off),
      legacy,
      RecalcEngine.Legacy
    )
    assertEquals(legacy.snapshot(Counter.GraphEdges), rows.toLong * (rows + 1) / 2)
  }
