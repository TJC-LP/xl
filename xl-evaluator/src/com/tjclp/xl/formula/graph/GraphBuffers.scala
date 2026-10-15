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
