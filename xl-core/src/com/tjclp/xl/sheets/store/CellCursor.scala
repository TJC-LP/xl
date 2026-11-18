package com.tjclp.xl.sheets.store

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.sheets.CellStore
import com.tjclp.xl.styles.units.StyleId

/**
 * A zero-allocation walk over the stored cells of a [[CellStore]] inside a window of rows and
 * columns (#891), in the store's row-major order: rows ascending as unsigned 32-bit numbers, then
 * columns ascending as signed `Int`s.
 *
 * {{{
 * val c = CellCursor.column(store, col = 2, rowFrom = 0, rowTo = 49999)
 * while c.next() do use(c.row, c.value)
 * }}}
 *
 * The accessors describe the cell the last successful `next()` stopped on. The cursor is mutable
 * and belongs to one thread; the store it walks is immutable and is not affected. Finding the next
 * row costs O(trie height); every cell after that within a row costs a few bit operations.
 */
// A cursor is mutable by design (one per walk, never shared); its state is vars and its walk loops.
@SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
private[xl] final class CellCursor private (
  root: Node,
  height: Int,
  lo1: Long,
  hi1: Long,
  lo2: Long,
  hi2: Long,
  colLo: Int,
  colHi: Int
):
  private val chunkLo = colLo >> 6
  private val chunkHi = colHi >> 6
  private val emptyWindow = colLo > colHi

  /** 0: walking `[lo1, hi1]`; 1: walking `[lo2, hi2]`; 2: done. */
  private var span = if emptyWindow || height == 0 then 2 else 0
  private var nextKey = lo1
  private var leaf: RowLeaf = RowLeaf.Empty
  private var rowKey = 0L
  private var chunkIdx = 0
  private var chunkCol = 0
  private var chunkMask = 0L
  private var chunkBase = 0
  private var pending = 0L
  private var pos = 0
  private var curCol = 0
  private var leafActive = false

  /** Moves to the next cell; false once the window is exhausted. */
  def next(): Boolean =
    var result = false
    var going = true
    while going do
      if pending != 0L then
        val b = java.lang.Long.numberOfTrailingZeros(pending)
        pending &= pending - 1
        pos = chunkBase + java.lang.Long.bitCount(chunkMask & ((1L << b) - 1))
        curCol = chunkCol + b
        result = true
        going = false
      else if leafActive && chunkIdx + 1 < leaf.chunkCount && leaf.chunkAt(chunkIdx + 1) <= chunkHi
      then loadChunk(chunkIdx + 1)
      else if !advanceRow() then going = false
    result

  /** The row index (signed, as `ARef` holds it). */
  def row: Int = rowKey.toInt
  def col: Int = curCol
  def ref: ARef = ARef.from0(curCol, rowKey.toInt)
  def value: CellValue = leaf.values(pos)
  def meta: Long = leaf.meta(pos)
  def tag: Int = Meta.tag(leaf.meta(pos))
  def hasStyle: Boolean = Meta.hasStyle(leaf.meta(pos))

  /** The raw style index; meaningful only when [[hasStyle]]. */
  def rawStyle: Int = Meta.rawStyle(leaf.meta(pos))
  def styleId: Option[StyleId] = Meta.style(leaf.meta(pos))
  def hasLink: Boolean = Meta.hasLink(leaf.meta(pos))

  /** The hyperlink; "" when [[hasLink]] is false. */
  def link: String = leaf.linkAt(pos)
  def hyperlink: Option[String] = if hasLink then Some(link) else None

  /** The current cell as a [[Cell]] (allocates). */
  def cell: Cell = Cell(ref, value, styleId, hyperlink)

  private def loadChunk(i: Int): Unit =
    chunkIdx = i
    val c = leaf.chunkAt(i)
    chunkCol = c << 6
    chunkMask = leaf.maskAt(i)
    chunkBase = leaf.baseAt(i)
    val lowBit = if c == chunkLo then colLo & 63 else 0
    val highBit = if c == chunkHi then colHi & 63 else 63
    pending = chunkMask & (-1L << lowBit) & (-1L >>> (63 - highBit))

  private def advanceRow(): Boolean =
    leafActive = false
    pending = 0L
    var found = false
    while !found && span < 2 do
      val hi = if span == 0 then hi1 else hi2
      if nextKey > hi || !ceil(nextKey) || rowKey > hi then
        span += 1
        nextKey = lo2
      else
        nextKey = rowKey + 1
        if leaf.maxCol >= colLo && leaf.minCol <= colHi then
          val i = leaf.chunkFrom(chunkLo)
          if i < leaf.chunkCount && leaf.chunkAt(i) <= chunkHi then
            leafActive = true
            loadChunk(i)
            found = true
    found

  /** Positions `leaf`/`rowKey` on the first row at or after `k`; false when there is none. */
  private def ceil(k: Long): Boolean =
    Trie.covers(height, k) && ceilIn(root, 5 * (height - 1), 0L, k)

  private def ceilIn(node: Node, shift: Int, prefix: Long, k: Long): Boolean =
    val start = if k <= prefix then 0 else ((k >>> shift) & 31L).toInt
    var bm = node.bitmap & (-1 << start)
    var found = false
    while !found && bm != 0 do
      val slot = Integer.numberOfTrailingZeros(bm)
      bm &= bm - 1
      val childLo = prefix + (slot.toLong << shift)
      node.children(Trie.slotIndex(node.bitmap, slot)) match
        case l: RowLeaf =>
          leaf = l
          rowKey = childLo
          found = true
        case inner: Node => found = ceilIn(inner, shift - 5, childLo, k)
    found

private[xl] object CellCursor:
  private val U = Trie.UnsignedMask

  /** Every stored cell. */
  def all(store: CellStore): CellCursor =
    new CellCursor(store.root, store.height, 0L, U, 1L, 0L, Int.MinValue, Int.MaxValue)

  /** The cells of rows `rowFrom..rowTo` (signed, inclusive) in columns `colFrom..colTo`. */
  def window(store: CellStore, rowFrom: Int, rowTo: Int, colFrom: Int, colTo: Int): CellCursor =
    if rowFrom > rowTo then new CellCursor(store.root, store.height, 1L, 0L, 1L, 0L, colFrom, colTo)
    else if rowFrom >= 0 || rowTo < 0 then
      new CellCursor(
        store.root,
        store.height,
        Trie.key(rowFrom),
        Trie.key(rowTo),
        1L,
        0L,
        colFrom,
        colTo
      )
    else
      // the signed span crosses zero: rows 0..rowTo first, then the negative rows (high keys)
      new CellCursor(
        store.root,
        store.height,
        0L,
        Trie.key(rowTo),
        Trie.key(rowFrom),
        U,
        colFrom,
        colTo
      )

  /** The cells of rows `rowFrom..rowTo` (signed, inclusive), every column. */
  def rows(store: CellStore, rowFrom: Int, rowTo: Int): CellCursor =
    window(store, rowFrom, rowTo, Int.MinValue, Int.MaxValue)

  /** The cells of column `col` in rows `rowFrom..rowTo` (signed, inclusive). */
  def column(store: CellStore, col: Int, rowFrom: Int, rowTo: Int): CellCursor =
    window(store, rowFrom, rowTo, col, col)

  /** The cells inside `range` (an unnormalized range with start after end holds nothing). */
  def range(store: CellStore, range: CellRange): CellCursor =
    window(
      store,
      range.start.row.index0,
      range.end.row.index0,
      range.start.col.index0,
      range.end.col.index0
    )
