package com.tjclp.xl.sheets

import java.lang.ref.WeakReference
import java.util.concurrent.atomic.{AtomicInteger, AtomicReferenceArray}

import com.tjclp.xl.addressing.ARef

/**
 * The range hyperlink covering a cell, found without scanning every link (#931). A sheet can hold
 * one range link per row (80k links on an 80k-row export), and `Sheet.apply` asks for the covering
 * link on every read of a cell without its own, so a linear scan made whole-sheet passes quadratic.
 *
 * The links are ordered by first row, over an implicit segment tree that keeps each subtree's
 * largest last row. A lookup for row `r` descends only into subtrees whose links can start at or
 * before `r` and end at or after it: O(log n + k) nodes for the k links whose rows span `r`. Among
 * those, the latest link (in model order) that covers the cell wins, as in
 * [[Sheet.rangeHyperlinkAt]].
 */
final class RangeHyperlinkIndex private (val links: Vector[RangeHyperlink]):

  private val n = links.size

  /** Model indices, sorted by first row. */
  private val order: Array[Int] = links.indices.sortBy(i => links(i).range.rowStart.index0).toArray
  private val firstRows: Array[Int] = order.map(i => links(i).range.rowStart.index0)
  private val width: Int = Iterator.iterate(1)(_ * 2).dropWhile(_ < math.max(n, 1)).next()

  /** Largest last row under each node of the implicit tree (leaves at `width + k`). */
  private val maxLastRow: Array[Int] =
    val tree = Array.fill(2 * width)(-1)
    order.indices.foreach(k => tree(width + k) = links(order(k)).range.rowEnd.index0)
    (width - 1 to 1 by -1).foreach(node =>
      tree(node) = math.max(tree(2 * node), tree(2 * node + 1))
    )
    tree

  /** The latest link covering `ref`, and the number of tree nodes the lookup visited. */
  private[sheets] def lookup(ref: ARef): (Option[RangeHyperlink], Int) =
    val row = ref.row.index0
    // links at positions [0, startsBefore) begin at or above `row`
    val startsBefore = upperBound(row)

    def visit(node: Int, lo: Int, hi: Int, best: Int, visited: Int): (Int, Int) =
      if lo >= startsBefore || maxLastRow(node) < row then (best, visited + 1)
      else if node >= width then
        val i = order(lo)
        (if i > best && links(i).covers(ref) then i else best, visited + 1)
      else
        val mid = (lo + hi) / 2
        val (leftBest, leftVisited) = visit(2 * node, lo, mid, best, visited + 1)
        visit(2 * node + 1, mid, hi, leftBest, leftVisited)

    val (best, visited) = visit(1, 0, width, -1, 0)
    (Option.when(best >= 0)(links(best)), visited)

  /** The latest link covering `ref`. */
  def at(ref: ARef): Option[RangeHyperlink] = lookup(ref)._1

  /** How many links start at or above `row` (binary search over [[firstRows]]). */
  private def upperBound(row: Int): Int =
    @annotation.tailrec
    def go(lo: Int, hi: Int): Int =
      if lo >= hi then lo
      else
        val mid = (lo + hi) >>> 1
        if firstRows(mid) <= row then go(mid + 1, hi) else go(lo, mid)
    go(0, n)

object RangeHyperlinkIndex:

  /** Below this many links a scan is cheaper than building an index. */
  private[sheets] val ScanLimit: Int = 16

  /** An index over `links` (built now, whatever their number). */
  def apply(links: Vector[RangeHyperlink]): RangeHyperlinkIndex = new RangeHyperlinkIndex(links)

  /**
   * The latest link in `links` covering `ref`. A few links are scanned; more go through an index
   * built once per distinct `links` vector and shared by every sheet that holds the same vector (a
   * sheet's copies share it, so a run of edits and reads does not rebuild it).
   */
  def covering(links: Vector[RangeHyperlink], ref: ARef): Option[RangeHyperlink] =
    if links.sizeIs <= ScanLimit then links.findLast(_.covers(ref))
    else Memo.of(links).at(ref)

  /**
   * A few recently used indices, matched by the identity of their links vector. Entries are weak,
   * so the memo never keeps a sheet's links alive; a miss only costs a rebuild, so races between
   * threads are harmless.
   */
  private object Memo:
    private val Slots = 8
    private val slots = new AtomicReferenceArray[WeakReference[RangeHyperlinkIndex]](Slots)
    private val next = new AtomicInteger(0)

    def of(links: Vector[RangeHyperlink]): RangeHyperlinkIndex =
      (0 until Slots).iterator
        .flatMap(i => Option(slots.get(i)).flatMap(ref => Option(ref.get)))
        .find(_.links eq links)
        .getOrElse {
          val built = RangeHyperlinkIndex(links)
          slots.set(Math.floorMod(next.getAndIncrement(), Slots), new WeakReference(built))
          built
        }
