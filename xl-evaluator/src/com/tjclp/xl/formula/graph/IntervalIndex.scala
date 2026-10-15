package com.tjclp.xl.formula.graph

/**
 * A static set of row intervals `[start, end]`, each carrying an `Int` id, answering "which
 * intervals overlap rows `a..b`?" in O(log n + k) for the common shapes (#767).
 *
 * The intervals are sorted by (start, end, id); an implicit segment tree over that order holds the
 * largest end below each node. A query first finds `p`, the number of intervals starting at or
 * before `b`, then walks the tree from the root, skipping a node whose leaves all start after `b`
 * (index at or past `p`) or whose largest end is before `a`.
 *
 * Ported from GridCraft's `Intervals<T>` (`crates/calc/src/recalc.rs`, MIT/Apache-2.0) without its
 * pending insertions and dead marks: an index is built once per graph and never changes.
 */
private[formula] final class IntervalIndex private (
  starts: Array[Int],
  ends: Array[Int],
  ids: Array[Int],
  maxEnd: Array[Int],
  leafBase: Int,
  stackDepth: Int
):
  /** Number of intervals. */
  def size: Int = starts.length

  /** Calls `f` with the id of every interval containing `row`. */
  def foreachStab(row: Int)(f: Int => Unit): Unit =
    walk(row, row, f)
    ()

  /** Calls `f` with the id of every interval overlapping rows `a..b` (`start <= b && end >= a`). */
  def foreachOverlap(a: Int, b: Int)(f: Int => Unit): Unit =
    walk(a, b, f)
    ()

  /** Tree nodes a stab at `row` visits: the cost the interval law bounds. */
  private[graph] def visitCount(row: Int): Int = walk(row, row, _ => ())

  /** First index whose start is after `b`. */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def partitionPoint(b: Int): Int =
    var lo = 0
    var hi = starts.length
    while lo < hi do
      val mid = (lo + hi) >>> 1
      if starts(mid) <= b then lo = mid + 1 else hi = mid
    lo

  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def walk(a: Int, b: Int, f: Int => Unit): Int =
    val p = partitionPoint(b)
    if p == 0 then 0
    else
      // explicit stack of (node, lo, hi); a depth-first walk keeps at most one pending sibling
      // per level
      val nodes = new Array[Int](stackDepth)
      val los = new Array[Int](stackDepth)
      val his = new Array[Int](stackDepth)
      var sp = 0
      nodes(0) = 1
      los(0) = 0
      his(0) = leafBase
      sp = 1
      var visited = 0
      while sp > 0 do
        sp -= 1
        val node = nodes(sp)
        val lo = los(sp)
        val hi = his(sp)
        visited += 1
        if lo < p && maxEnd(node) >= a then
          if hi - lo <= 1 then
            if lo < starts.length && ends(lo) >= a then f(ids(lo))
          else
            val mid = (lo + hi) >>> 1
            // right child first, so the left subtree is walked first
            nodes(sp) = 2 * node + 1
            los(sp) = mid
            his(sp) = hi
            nodes(sp + 1) = 2 * node
            los(sp + 1) = lo
            his(sp + 1) = mid
            sp += 2
      visited

private[formula] object IntervalIndex:

  val empty: IntervalIndex = build(Array.emptyIntArray, Array.emptyIntArray, Array.emptyIntArray)

  /**
   * The index of intervals `[starts(i), ends(i)]` with ids `ids(i)`. The three arrays have equal
   * length and are copied; `ends(i) >= starts(i)` is assumed, as for every normalised range.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  def build(starts: Array[Int], ends: Array[Int], ids: Array[Int]): IntervalIndex =
    val n = math.min(starts.length, math.min(ends.length, ids.length))
    val order = (0 until n).toArray.sortWith { (x, y) =>
      if starts(x) != starts(y) then starts(x) < starts(y)
      else if ends(x) != ends(y) then ends(x) < ends(y)
      else ids(x) < ids(y)
    }
    val s = order.map(starts)
    val e = order.map(ends)
    val d = order.map(ids)
    val leafBase = if n <= 1 then 1 else Integer.highestOneBit(n - 1) << 1
    val maxEnd = Array.fill(2 * leafBase)(Int.MinValue)
    var i = 0
    while i < n do
      maxEnd(leafBase + i) = e(i)
      i += 1
    var node = leafBase - 1
    while node >= 1 do
      maxEnd(node) = math.max(maxEnd(2 * node), maxEnd(2 * node + 1))
      node -= 1
    // a depth-first walk holds at most two entries per level of a tree of height log2(leafBase)
    val depth = 2 * (Integer.numberOfTrailingZeros(leafBase) + 2)
    new IntervalIndex(s, e, d, maxEnd, leafBase, depth)
