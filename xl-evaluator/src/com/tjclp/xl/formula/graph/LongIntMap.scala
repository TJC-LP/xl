package com.tjclp.xl.formula.graph

/**
 * An immutable map from `Long` keys to non-negative `Int` values, stored as one open-addressing
 * table (#767).
 *
 * Keys hash through the SplitMix64 finaliser, so packed cell keys (sheet, row, column) spread over
 * the table however their bits are laid out; `ARef`'s own `row ^ col` hash puts one cell per column
 * in each bucket. [[indexOf]] neither boxes nor allocates. The arrays are filled once by
 * [[LongIntMap.from]] and never written afterwards.
 */
private[formula] final class LongIntMap private (
  keys: Array[Long],
  values: Array[Int],
  val size: Int
):
  private val mask = keys.length - 1

  /** The value stored for `key`, or -1 when the map has none. */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  def indexOf(key: Long): Int =
    var slot = (LongIntMap.mix(key) & mask).toInt
    var found = -2
    while found == -2 do
      val v = values(slot)
      if v < 0 then found = -1
      else if keys(slot) == key then found = v
      else slot = (slot + 1) & mask
    found

  /** The value stored for `key`. */
  def get(key: Long): Option[Int] =
    val v = indexOf(key)
    if v < 0 then None else Some(v)

  def contains(key: Long): Boolean = indexOf(key) >= 0

private[formula] object LongIntMap:

  val empty: LongIntMap = from(Array.emptyLongArray, Array.emptyIntArray, 0)

  /**
   * SplitMix64's finaliser (Steele, Lea and Flood, "Fast splittable pseudorandom number
   * generators").
   */
  def mix(key: Long): Long =
    val a = key + 0x9e3779b97f4a7c15L
    val b = (a ^ (a >>> 30)) * 0xbf58476d1ce4e5b9L
    val c = (b ^ (b >>> 27)) * 0x94d049bb133111ebL
    c ^ (c >>> 31)

  /**
   * The map of the first `n` pairs `keys(i) -> values(i)`; a repeated key keeps its last value.
   * Every value must be non-negative: -1 marks an empty slot.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  def from(keys: Array[Long], values: Array[Int], n: Int): LongIntMap =
    // load factor at most 1/2: a probe sequence stays short and always reaches an empty slot
    val capacity = Integer.highestOneBit(math.max(2, n) * 2 - 1) << 1
    val mask = capacity - 1
    val tableKeys = new Array[Long](capacity)
    val tableValues = Array.fill(capacity)(-1)
    var size = 0
    var i = 0
    while i < n do
      val key = keys(i)
      var slot = (mix(key) & mask).toInt
      while tableValues(slot) >= 0 && tableKeys(slot) != key do slot = (slot + 1) & mask
      if tableValues(slot) < 0 then size += 1
      tableKeys(slot) = key
      tableValues(slot) = values(i)
      i += 1
    new LongIntMap(tableKeys, tableValues, size)
