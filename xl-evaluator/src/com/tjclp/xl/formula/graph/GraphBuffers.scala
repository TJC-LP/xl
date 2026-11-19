package com.tjclp.xl.formula.graph

/**
 * Growable primitive buffers and a binary min-heap for the compact graph's builders and walks
 * (#767). Each instance is local to one build or one walk and never escapes it: the results are
 * copied into immutable arrays.
 */
@SuppressWarnings(Array("org.wartremover.warts.Var"))
private[graph] final class IntBuf(initial: Int = 16):
  private var data = new Array[Int](math.max(1, initial))
  private var n = 0

  def length: Int = n
  def apply(i: Int): Int = data(i)
  def clear(): Unit = n = 0

  /** Cells of the backing array, whatever `length` is. */
  def capacity: Int = data.length

  def +=(x: Int): Unit =
    if n == data.length then data = java.util.Arrays.copyOf(data, data.length * 2)
    data(n) = x
    n += 1

  def toArray: Array[Int] = java.util.Arrays.copyOf(data, n)

  /** Sort the first `length` elements and drop repeats. */
  @SuppressWarnings(Array("org.wartremover.warts.While"))
  def sortDistinct(): Unit =
    java.util.Arrays.sort(data, 0, n)
    var out = 0
    var i = 0
    while i < n do
      if out == 0 || data(out - 1) != data(i) then
        data(out) = data(i)
        out += 1
      i += 1
    n = out

@SuppressWarnings(Array("org.wartremover.warts.Var"))
private[graph] final class LongBuf(initial: Int = 16):
  private var data = new Array[Long](math.max(1, initial))
  private var n = 0

  def length: Int = n
  def apply(i: Int): Long = data(i)

  def +=(x: Long): Unit =
    if n == data.length then data = java.util.Arrays.copyOf(data, data.length * 2)
    data(n) = x
    n += 1

  def toArray: Array[Long] = java.util.Arrays.copyOf(data, n)

/** A binary min-heap of `Int`s with room for `capacity` elements. */
@SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
private[graph] final class IntMinHeap(capacity: Int):
  private val heap = new Array[Int](math.max(1, capacity))
  private var n = 0

  def isEmpty: Boolean = n == 0
  def nonEmpty: Boolean = n > 0

  def push(x: Int): Unit =
    var i = n
    n += 1
    heap(i) = x
    while i > 0 && heap((i - 1) >>> 1) > heap(i) do
      val parent = (i - 1) >>> 1
      val t = heap(parent)
      heap(parent) = heap(i)
      heap(i) = t
      i = parent

  /** Remove and return the smallest element; the heap must be non-empty. */
  def pop(): Int =
    val top = heap(0)
    n -= 1
    heap(0) = heap(n)
    var i = 0
    var done = false
    while !done do
      val l = 2 * i + 1
      val r = l + 1
      var smallest = i
      if l < n && heap(l) < heap(smallest) then smallest = l
      if r < n && heap(r) < heap(smallest) then smallest = r
      if smallest == i then done = true
      else
        val t = heap(smallest)
        heap(smallest) = heap(i)
        heap(i) = t
        i = smallest
    top

/**
 * #891: the nodes a walk has reached, each with the stamp of the walk that reached it last — a
 * visited set for many short walks over a few nodes of a large graph. Open addressing on the node
 * id, so its size follows the nodes touched, not the graph: a stamp array over every node costs a
 * graph-sized allocation per use.
 */
@SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
private[graph] final class StampSet(expected: Int):
  private var mask = Integer.highestOneBit(math.max(8, expected) * 2 - 1) * 2 - 1
  private var keys = StampSet.empty(mask + 1)
  private var stamps = new Array[Int](mask + 1)
  private var used = 0
  private var cells = 2L * (mask + 1)

  /** Array cells this set has allocated since it was made, the tables it outgrew included. */
  def allocated: Long = cells

  /** Stamp `node` with `stamp`; false when it already carried that stamp. `node` is >= 0. */
  def mark(node: Int, stamp: Int): Boolean =
    var h = StampSet.mix(node) & mask
    while keys(h) >= 0 && keys(h) != node do h = (h + 1) & mask
    if keys(h) == node then
      if stamps(h) == stamp then false
      else
        stamps(h) = stamp
        true
    else
      keys(h) = node
      stamps(h) = stamp
      used += 1
      if used * 2 > mask then grow()
      true

  private def grow(): Unit =
    val oldKeys = keys
    val oldStamps = stamps
    mask = mask * 2 + 1
    keys = StampSet.empty(mask + 1)
    stamps = new Array[Int](mask + 1)
    cells += 2L * (mask + 1)
    var i = 0
    while i < oldKeys.length do
      val k = oldKeys(i)
      if k >= 0 then
        var h = StampSet.mix(k) & mask
        while keys(h) >= 0 do h = (h + 1) & mask
        keys(h) = k
        stamps(h) = oldStamps(i)
      i += 1

private[graph] object StampSet:
  private def empty(size: Int): Array[Int] =
    val a = new Array[Int](size)
    java.util.Arrays.fill(a, -1)
    a

  private def mix(x: Int): Int =
    val h = x * 0x9e3779b9
    h ^ (h >>> 16)
