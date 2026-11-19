package com.tjclp.xl.formula.graph

import scala.collection.immutable.BitSet

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.formula.graph.CompactGraph.SweepScratch
import com.tjclp.xl.sheets.{CellStore, Sheet}
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * #891: what [[CompactGraph.withinComponentOrder]] allocates is a function of the cycle it orders,
 * not of the workbook around it. A recalculation orders every cyclic component, so scratch sized to
 * the graph (a stamp array over every node, as the first id-based version had) makes a book of many
 * small cycles quadratic: 10,000 two-cell cycles in a million-formula book took 2.8 s to order. The
 * pin is a count, not a clock: the same cycles inside acyclic parts of two sizes report the same
 * scratch, cell for cell.
 *
 * The scratch counted is the walk's (the visited set with every table it outgrew, the queue, the
 * edge lists), read from the buffers themselves. An array allocated beside them would not be
 * counted here; `SweepOrderGatesSpec` in the perf module measures allocated bytes for that.
 */
class SweepOrderCostSpec extends FunSuite:

  /** Every cyclic component's scratch, in condensation order. */
  private def scratches(wb: Workbook): Vector[SweepScratch] =
    val graph = CompactGraph.build(wb)
    graph.condensation.collect { case (ids, true) =>
      val members = BitSet.fromSpecific(ids.iterator)
      // the measured call is the production walk itself (`withinComponentOrder` is its first half)
      graph.withinComponentOrderMeasured(members)._2
    }

  test("two-cell cycles: the same scratch in a 2,000-row and a 20,000-row acyclic book") {
    val cycles = 200
    val small = scratches(SweepOrderCostSpec.mixedBook(2000, cycles))
    val large = scratches(SweepOrderCostSpec.mixedBook(20000, cycles))
    assertEquals(small.size, cycles)
    // a visited set of 16 slots (keys and stamps), a queue and two edge lists of 16 cells each
    assertEquals(small.distinct, Vector(SweepScratch(members = 2, edges = 2, cells = 80L)))
    assertEquals(large, small)
  }

  test("one ring of 300 cells: the same scratch in a 2,000-row and a 20,000-row acyclic book") {
    val small = scratches(SweepOrderCostSpec.ringBook(2000, 300))
    val large = scratches(SweepOrderCostSpec.ringBook(20000, 300))
    // every member reads every other through the ring's range node: 300 x 299 reads
    assertEquals(small, Vector(SweepScratch(members = 300, edges = 89700, cells = 313856L)))
    assertEquals(large, small)
  }

object SweepOrderCostSpec:

  private def formula(c: Int, r: Int, text: String): Cell =
    Cell(ARef.from0(c, r), CellValue.Formula(text))

  private def workbook(cells: IndexedSeq[Cell]): Workbook =
    Workbook(Vector(Sheet(SheetName.unsafe("S"), cells = CellStore.from(cells))))

  /**
   * The acyclic part, `rows` rows: C = A + 1, D = a SUM over a four-row window of C (range nodes
   * and their segments), and G = a SUM over the whole cycle block `block`, so the block's range
   * node has one reader per row outside every cycle.
   */
  private def acyclic(rows: Int, block: String): IndexedSeq[Cell] =
    (0 until rows).flatMap { r =>
      Vector(
        formula(2, r, s"=A${r + 1}+1"),
        formula(3, r, s"=SUM(C${r + 1}:C${math.min(rows, r + 4)})"),
        formula(6, r, s"=SUM($block)+C${r + 1}")
      )
    }

  /** `cycles` independent two-cell cycles (E reads F, F reads E) beside `rows` acyclic rows. */
  def mixedBook(rows: Int, cycles: Int): Workbook =
    val cyclic = (0 until cycles).flatMap { r =>
      Vector(formula(4, r, s"=F${r + 1}*0.5+1"), formula(5, r, s"=E${r + 1}*0.5"))
    }
    workbook(acyclic(rows, s"$$E$$1:$$F$$$cycles") ++ cyclic)

  /**
   * One cycle of `size` cells in column E (each reads the next, the last the first, and every one a
   * SUM over the whole ring, so the ring's range node sits inside the component) beside `rows`
   * acyclic rows.
   */
  def ringBook(rows: Int, size: Int): Workbook =
    val block = s"$$E$$1:$$E$$$size"
    val ring = (0 until size).map { r =>
      formula(4, r, s"=E${(r + 1) % size + 1}*0.5+SUM($block)/1000000")
    }
    workbook(acyclic(rows, block) ++ ring)
