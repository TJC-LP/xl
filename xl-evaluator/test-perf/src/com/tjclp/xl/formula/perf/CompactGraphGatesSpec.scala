package com.tjclp.xl.formula.perf

import scala.collection.immutable.BitSet
import scala.concurrent.duration.Duration

import com.tjclp.xl.formula.eval.RecalcCounters
import com.tjclp.xl.formula.eval.RecalcCounters.Counter
import com.tjclp.xl.formula.graph.CompactGraph
import munit.FunSuite

/**
 * #767 graph gates (design §12.3, from the compact-graph PR): the recalculation graph of the H10
 * book and of a 200k-row running total is built, cycle-checked and ordered inside the fork's
 * `-Xmx2g` ceiling, with sizes and walk lengths bounded by (formulas + ranges), times log2(rows)
 * for the running total. Only the graph runs here: evaluating H10 still scans every criteria row
 * until the criteria index lands, and the running total's SUMs are quadratic to evaluate.
 */
class CompactGraphGatesSpec extends FunSuite:

  override val munitTimeout: Duration = Duration(600, "s")

  private final case class Measured(
    graph: CompactGraph,
    counters: RecalcCounters.Snapshot,
    work: Long
  )

  /** Build, Tarjan and Kahn, as a recalculation runs them; `work` is the edges they walk. */
  private def measure(label: String, wb: com.tjclp.xl.workbooks.Workbook): Measured =
    val counters = RecalcCounters.create()
    val start = System.nanoTime()
    val graph = CompactGraph.build(wb, counters)
    val components = graph.components
    val kahn = graph.kahn(components.cyclic)
    val seconds = (System.nanoTime() - start).toDouble / 1e9
    val snap = counters.snapshot
    val work = components.work + kahn.work
    println(
      f"== $label graph: formulas=${graph.formulaCount}%,d ranges=${graph.rangeCount}%,d " +
        f"segments=${graph.segmentCount}%,d edges=${graph.edgeCount}%,d work=$work%,d " +
        f"in $seconds%.2f s"
    )
    assertEquals(kahn.remaining, BitSet.empty)
    assertEquals(kahn.formulaOrder.length, graph.formulaCount - components.cyclic.size)
    Measured(graph, snap, work)

  test("H10: graphEdges + virtualNodes <= 40 x (formulas + ranges)") {
    val built = LoadsBook.build(LoadsBook.H10)
    val m = measure("H10", built.workbook)
    val f = m.counters(Counter.GraphNodes)
    val r = m.counters(Counter.RangeNodes)
    assertEquals(f, built.totalFormulas.toLong)
    val size = m.counters(Counter.GraphEdges) + m.counters(Counter.VirtualNodes)
    assert(size <= 40L * (f + r), s"$size > 40 x ($f + $r)")
    // the legacy graph of this book held ~92.5M edges
    assert(m.counters(Counter.GraphEdges) < 1000000L, m.counters.render)
  }

  test("RunTotal(200k): graph work <= 3 x (formulas + ranges) x log2(rows), inside -Xmx2g") {
    val rows = 200000
    val m = measure("RunTotal(200k)", LoadsBook.runTotal(rows))
    val f = m.graph.formulaCount.toLong
    val r = m.graph.rangeCount.toLong
    assertEquals(f, 2L * rows)
    assertEquals(r, rows.toLong)
    val log2 = 32 - Integer.numberOfLeadingZeros(rows)
    assert(m.work <= 3L * (f + r) * log2, s"work ${m.work} > 3 x ${f + r} x $log2")
    assert(m.graph.edgeCount.toLong <= (f + r) * log2, s"${m.graph.edgeCount} edges")
  }
