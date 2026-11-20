package com.tjclp.xl.sheets.store

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.{Cell, CellValue, Hyperlink}
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
 * row costs O(trie height); every cell after that within a row costs a few bit operations (plus a
 * step to the next band's leaf every [[Band.Width]] columns of a row that spans several bands).
 */
// A cursor is mutable by design (one per walk, never shared); its state is vars and its walk loops.
@SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
private[xl] final class CellCursor private (
  root: Entry,
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
  private val bandLo = Band.of(colLo)
  private val bandHi = Band.of(colHi)
  private val emptyWindow = colLo > colHi

  /** 0: walking `[lo1, hi1]`; 1: walking `[lo2, hi2]`; 2: done. */
  private var span = if emptyWindow || height == 0 then 2 else 0
  private var nextKey = lo1

  // The current row: a leaf, band `bandIdx` of a banded row (then `leaf` is that band's leaf), or
  // row `blockRow` (key bits `blockLow`) of a block. `cells` holds the arrays in every case.
  private var cells: Cells = RowLeaf.Empty
  private var leaf: RowLeaf = RowLeaf.Empty
  private var inBands = false
  private var bandIds: Array[Int] = Array.emptyIntArray
  private var bandLeaves: Array[RowLeaf] = CellCursor.NoBands
  private var bandIdx = 0
  private var block: Block = Group.Placeholder
  private var inBlock = false
  private var blockRow = 0
  private var blockLow = 0

  private var rowKey = 0L
  private var chunkIdx = 0
  private var chunkCol = 0
  private var chunkMask = 0L
  private var chunkBase = 0
  private var pending = 0L
  private var pos = 0
  private var curCol = 0
  private var rowActive = false

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
      else if rowActive && !inBlock && chunkIdx + 1 < leaf.chunkCount &&
        leaf.chunkAt(chunkIdx + 1) <= chunkHi
      then loadChunk(chunkIdx + 1)
      else if rowActive && inBands && nextBand() then () // a chunk of a later band is loaded
      else if !advanceRow() then going = false
    result

  /** The row index (signed, as `ARef` holds it). */
  def row: Int = rowKey.toInt
  def col: Int = curCol
  def ref: ARef = ARef.from0(curCol, rowKey.toInt)
  def value: CellValue = cells.values(pos)
  def meta: Long = cells.meta(pos)
  def tag: Int = Meta.tag(cells.meta(pos))
  def hasStyle: Boolean = Meta.hasStyle(cells.meta(pos))

  /** The raw style index; meaningful only when [[hasStyle]]. */
  def rawStyle: Int = Meta.rawStyle(cells.meta(pos))
  def styleId: Option[StyleId] = Meta.style(cells.meta(pos))
  def hasLink: Boolean = Meta.hasLink(cells.meta(pos))

  /** The hyperlink; "" when [[hasLink]] is false. */
  def link: String = cells.linkAt(pos)
  def hyperlink: Option[Hyperlink] = if hasLink then Some(Hyperlink(link)) else None

  /** The current cell as a [[Cell]] (allocates). */
  def cell: Cell = Cell(ref, value, styleId, hyperlink)

  private def loadChunk(i: Int): Unit =
    chunkIdx = i
    val c = if inBlock then block.chunk else leaf.chunkAt(i)
    chunkCol = c << 6
    chunkMask = if inBlock then block.rowMaskAt(blockRow) else leaf.maskAt(i)
    chunkBase = if inBlock then block.rowBaseAt(blockRow) else leaf.baseAt(i)
    val lowBit = if c == chunkLo then colLo & 63 else 0
    val highBit = if c == chunkHi then colHi & 63 else 63
    pending = chunkMask & (-1L << lowBit) & (-1L >>> (63 - highBit))

  /** Loads the current row's first chunk inside the column window; false when it has none. */
  private def enterRow(): Boolean =
    if inBlock then
      if block.chunk < chunkLo || block.chunk > chunkHi then false
      else
        loadChunk(0)
        true
    else if inBands then
      bandIdx = BandedRow.indexFrom(bandIds, bandLo) - 1
      nextBand()
    else if leaf.maxCol < colLo || leaf.minCol > colHi then false
    else enterLeaf()

  /** Loads the current leaf's first chunk inside the column window; false when it has none. */
  private def enterLeaf(): Boolean =
    val i = leaf.chunkFrom(chunkLo)
    if i < leaf.chunkCount && leaf.chunkAt(i) <= chunkHi then
      loadChunk(i)
      true
    else false

  /**
   * Moves to the first band after `bandIdx` that has a chunk inside the column window and loads
   * that chunk; false when the row has no such band.
   */
  private def nextBand(): Boolean =
    var found = false
    var j = bandIdx + 1
    while !found && j < bandIds.length && bandIds(j) <= bandHi do
      bandIdx = j
      leaf = bandLeaves(j)
      cells = leaf
      found = enterLeaf()
      j += 1
    found

  private def advanceRow(): Boolean =
    pending = 0L
    var found = false
    while !found && span < 2 do
      val hi = if span == 0 then hi1 else hi2
      // the next row of the current block, when there is one, needs no descent from the root
      val inBlockNext =
        if rowActive && inBlock && blockLow < 31 then block.rowMask & (-1 << (blockLow + 1)) else 0
      rowActive = false
      if inBlockNext != 0 && rowKey - blockLow + Integer.numberOfTrailingZeros(inBlockNext) <= hi
      then
        val low = Integer.numberOfTrailingZeros(inBlockNext)
        rowKey = rowKey - blockLow + low
        blockLow = low
        blockRow += 1
        nextKey = rowKey + 1
        rowActive = true
        found = enterRow()
      else if nextKey > hi || !ceil(nextKey) || rowKey > hi then
        span += 1
        nextKey = lo2
      else
        nextKey = rowKey + 1
        rowActive = true
        found = enterRow()
    found

  /** Positions the cursor on the first row at or after `k`; false when there is none. */
  private def ceil(k: Long): Boolean =
    Trie.covers(height, k) && ceilIn(root, 5 * (height - 1), 0L, k)

  private def ceilIn(e: Entry, shift: Int, prefix: Long, k: Long): Boolean = e match
    case b: Block =>
      val start = if k <= prefix then 0 else (k & 31L).toInt
      val bm = b.rowMask & (-1 << start)
      if bm == 0 then false
      else
        val low = Integer.numberOfTrailingZeros(bm)
        block = b
        cells = b
        inBlock = true
        inBands = false
        blockLow = low
        blockRow = b.rowIndex(low)
        rowKey = prefix + low
        true
    case l: RowLeaf =>
      leaf = l
      cells = l
      inBlock = false
      inBands = false
      rowKey = prefix
      true
    case r: BandedRow =>
      bandIds = r.ids
      bandLeaves = r.bands
      inBlock = false
      inBands = true
      rowKey = prefix
      true
    case node: Node =>
      val start = if k <= prefix then 0 else ((k >>> shift) & 31L).toInt
      var bm = node.bitmap & (-1 << start)
      var found = false
      while !found && bm != 0 do
        val slot = Integer.numberOfTrailingZeros(bm)
        bm &= bm - 1
        val childLo = prefix + (slot.toLong << shift)
        found = ceilIn(node.children(Trie.slotIndex(node.bitmap, slot)), shift - 5, childLo, k)
      found

private[xl] object CellCursor:
  private val U = Trie.UnsignedMask
  private val NoBands: Array[RowLeaf] = Array.empty[RowLeaf]

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
