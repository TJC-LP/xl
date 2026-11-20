package com.tjclp.xl.formula.perf

import java.lang.management.ManagementFactory

import scala.collection.immutable.BitSet
import scala.concurrent.duration.Duration

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.formula.graph.CompactGraph
import com.tjclp.xl.sheets.{CellStore, Sheet}
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * #891 sweep-order gate: ordering the cyclic components of a book costs what the cycles cost,
 * whatever the size of the acyclic part. The first id-based `CompactGraph.withinComponentOrder`
 * allocated arrays over every node of the graph for each component, so 10,000 two-cell cycles in a
 * million-formula book took 2.8 s to order.
 *
 * Two measures, at 100,000 and 400,000 acyclic rows around the same 10,000 cycles:
 *
 *   - the walk's scratch as the graph reports it (`SweepScratch`), which must be equal cell for
 *     cell (`SweepOrderCostSpec` pins the same at small sizes in the blocking test module);
 *   - the bytes the ordering thread allocated (the JVM's per-thread allocation counter), which also
 *     sees an array the report does not instrument. Bytes vary a little with JIT state, so the gate
 *     is a ratio: at most 1.5x for a 4x larger graph, where scratch sized to the graph gives about
 *     4x.
 */
class SweepOrderGatesSpec extends FunSuite:

  override val munitTimeout: Duration = Duration(600, "s")

  private val cycles = 10000

  private def formula(c: Int, r: Int, text: String): Cell =
    Cell(ARef.from0(c, r), CellValue.Formula(text))

  /**
   * `rows` acyclic rows (C = A + 1, D = a SUM over a four-row window of C, G = a SUM over the whole
   * cycle block, so the block's range node has a reader on every row) beside `cycles` independent
   * two-cell cycles in E and F.
   */
  private def mixedBook(rows: Int): Workbook =
    val block = s"$$E$$1:$$F$$$cycles"
    val acyclic = (0 until rows).flatMap { r =>
      Vector(
        formula(2, r, s"=A${r + 1}+1"),
        formula(3, r, s"=SUM(C${r + 1}:C${math.min(rows, r + 4)})"),
        formula(6, r, s"=SUM($block)+C${r + 1}")
      )
    }
    val cyclic = (0 until cycles).flatMap { r =>
      Vector(formula(4, r, s"=F${r + 1}*0.5+1"), formula(5, r, s"=E${r + 1}*0.5"))
    }
    Workbook(Vector(Sheet(SheetName.unsafe("S"), cells = CellStore.from(acyclic ++ cyclic))))

  private final case class Measured(
    nodes: Int,
    components: Int,
    scratchCells: Long,
    allocatedBytes: Long,
    millis: Long
  )

  private def allocatedBytes(): Long =
    ManagementFactory.getThreadMXBean match
      case bean: com.sun.management.ThreadMXBean =>
        bean.getThreadAllocatedBytes(Thread.currentThread().threadId())
      case _ => -1L

  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  private def measure(rows: Int): Measured =
    val graph = CompactGraph.build(mixedBook(rows))
    val components = graph.condensation.collect { case (ids, true) =>
      BitSet.fromSpecific(ids.iterator)
    }
    def orderAll(): Long =
      var cells = 0L
      components.foreach(c => cells += graph.withinComponentOrderMeasured(c)._2.cells)
      cells
    // warm: Tarjan's components are built, the code is loaded and compiled
    orderAll()
    orderAll()
    val before = allocatedBytes()
    val start = System.nanoTime()
    val cells = orderAll()
    val millis = (System.nanoTime() - start) / 1000000L
    val bytes = allocatedBytes() - before
    val measured = Measured(graph.nodeCount, components.size, cells, bytes, millis)
    println(f"== sweep order, $rows%,d acyclic rows: $measured")
    measured

  test("10,000 two-cell cycles: scratch and allocation do not grow with the acyclic part") {
    val small = measure(100000)
    val large = measure(400000)
    assertEquals(small.components, cycles)
    assertEquals(large.components, cycles)
    assert(large.nodes > 3L * small.nodes, s"$small, $large")
    assertEquals(large.scratchCells, small.scratchCells, s"$small, $large")
    assertEquals(small.scratchCells, 80L * cycles)
    assert(small.allocatedBytes > 0L, "the JVM does not report per-thread allocation")
    assert(
      large.allocatedBytes * 2L <= small.allocatedBytes * 3L,
      s"allocation grew with the graph: $small, $large"
    )
  }
