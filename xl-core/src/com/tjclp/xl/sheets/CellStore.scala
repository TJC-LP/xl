package com.tjclp.xl.sheets

import scala.util.hashing.MurmurHash3

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.{Cell, CellValue, Hyperlink}
import com.tjclp.xl.sheets.store.{
  Band,
  BandedRow,
  CellCursor,
  Entry,
  Meta,
  Block,
  Cells,
  Group,
  Node,
  RowLeaf,
  Shift,
  StableSort,
  StoredRow,
  Trie,
  WideRow
}
import com.tjclp.xl.styles.units.StyleId

/**
 * A sheet's cells: an immutable, persistent map from cell reference to cell, laid out for the
 * access patterns of a spreadsheet engine (#891).
 *
 * '''Layout.''' A bitmap trie keyed by the row index read as an unsigned 32-bit number, five bits
 * per level, so a lookup is a few shifts and popcounts with no hashing. Cells live in parallel
 * arrays (a metadata word with the value tag, style and link flags; the value; the hyperlink)
 * indexed by 64-column bitmaps: one set per row (one per band of [[Band.Width]] columns when a row
 * of more than [[Band.FlatMax]] cells spans several), or, for a narrow group of 32 rows that shares
 * one 64-column chunk (tall narrow data), one set for the whole group. Every trie node carries
 * summaries (cell count, valued count, column bounds, a value-tag mask), so `size`, `usedRange`,
 * `extent` and `hasTag` read them instead of scanning. Updates copy one row band (or one narrow
 * group) and the path above it; the rest is shared with the previous store.
 *
 * '''Order.''' [[iterator]], [[refs]], [[inRange]] and `toString` go row-major: rows ascending as
 * unsigned 32-bit numbers, then columns ascending as signed `Int`s. On Excel's grid that is the
 * familiar A1, B1, ..., A2 order; a cell with a negative row (only [[ARef.shift]] can mint one)
 * comes after every grid row, and a negative column comes first in its row.
 *
 * '''Bounds.''' [[usedRange]] and [[extent]] take signed minima and maxima, as [[Sheet.usedRange]]
 * always has: a negative row is the minimum row even though it iterates last.
 *
 * '''Equality.''' Two stores are equal when they hold equal cells at the same references, however
 * they were built: the layout is canonical (no empty rows or nodes, minimal height, exactly sized
 * arrays). `hashCode` agrees with `equals` and costs a pass over the cells.
 *
 * Every reference is accepted, including ones off Excel's grid (negative, or beyond row 1,048,576
 * or column XFD); nothing here fails or throws.
 */
// Local loops over the trie's arrays (the house style for flat-array code, see SccKernel).
@SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
final class CellStore private (
  private[xl] val root: Entry,
  private[xl] val height: Int
):
  import CellStore.*

  // ----- reads -----

  /** The number of stored cells (style-only Empty cells included). */
  def size: Int = root.count

  /** The number of stored cells whose value is not Empty. */
  def valuedCount: Int = root.valued

  def isEmpty: Boolean = root.count == 0
  def nonEmpty: Boolean = root.count != 0

  private def cellsAt(ref: ARef): Cells =
    Trie.lookup(root, height, Trie.key(ref.row.index0), ref.col.index0)

  /** The array position of `ref`'s cell within `c`, or -1. */
  private def posIn(c: Cells, ref: ARef): Int =
    c.find(Trie.key(ref.row.index0), ref.col.index0)

  def contains(ref: ARef): Boolean = posIn(cellsAt(ref), ref) >= 0

  /** The cell at `ref`, built from the stored fields. */
  def get(ref: ARef): Option[Cell] =
    val c = cellsAt(ref)
    val p = posIn(c, ref)
    if p < 0 then None
    else
      val m = c.meta(p)
      Some(
        Cell(
          ref,
          c.values(p),
          Meta.style(m),
          if Meta.hasLink(m) then Some(Hyperlink(c.links(p))) else None
        )
      )

  /** The value at `ref`, Empty when no cell is stored there. Allocates nothing. */
  def valueAt(ref: ARef): CellValue =
    val c = cellsAt(ref)
    val p = posIn(c, ref)
    if p < 0 then CellValue.Empty else c.values(p)

  def styleAt(ref: ARef): Option[StyleId] =
    val c = cellsAt(ref)
    val p = posIn(c, ref)
    if p < 0 then None else Meta.style(c.meta(p))

  def hyperlinkAt(ref: ARef): Option[Hyperlink] =
    val c = cellsAt(ref)
    val p = posIn(c, ref)
    if p < 0 || !Meta.hasLink(c.meta(p)) then None else Some(Hyperlink(c.links(p)))

  /** Every stored cell, row-major (see the class notes for the order off the grid). */
  def iterator: Iterator[Cell] = cellsOf(CellCursor.all(this))

  /** Every stored reference, row-major. */
  def refs: Iterator[ARef] = refsOf(CellCursor.all(this))

  /**
   * The stored cells inside `range`, row-major. Only stored cells are visited, never the empty
   * positions between them. A range whose start lies after its end holds nothing.
   */
  def inRange(range: CellRange): Iterator[Cell] = cellsOf(CellCursor.range(this, range))

  /** The bounding box of the cells holding a non-Empty value, as [[Sheet.usedRange]] defines it. */
  def usedRange: Option[CellRange] =
    if root.valued == 0 then None else Some(box(root.vMinCol, root.vMaxCol, valuedOnly = true))

  /** The bounding box of every stored cell, style-only Empty cells included. */
  def extent: Option[CellRange] =
    if root.count == 0 then None else Some(box(root.minCol, root.maxCol, valuedOnly = false))

  /**
   * Whether some stored cell's value tag is in `tagMask` (a union of [[CellStore.tagBit]]s, e.g.
   * [[CellStore.FormulaTags]]). Reads the root summary: O(1).
   */
  def hasTag(tagMask: Int): Boolean = (root.tagMask & tagMask) != 0

  private def box(minCol: Int, maxCol: Int, valuedOnly: Boolean): CellRange =
    val shift = 5 * (height - 1)
    val neg = 0x80000000L
    val top = Trie.UnsignedMask
    // signed minimum row: the most negative row if any (keys 2^31..2^32-1), else the least key
    val firstNeg = Trie.firstKey(root, shift, 0L, neg, top, valuedOnly)
    val minRow =
      if firstNeg >= 0 then firstNeg else Trie.firstKey(root, shift, 0L, 0L, neg - 1, valuedOnly)
    val lastPos = Trie.lastKey(root, shift, 0L, 0L, neg - 1, valuedOnly)
    val maxRow =
      if lastPos >= 0 then lastPos else Trie.lastKey(root, shift, 0L, neg, top, valuedOnly)
    CellRange(ARef.from0(minCol, minRow.toInt), ARef.from0(maxCol, maxRow.toInt))

  // ----- writes -----

  private def modifyRow(row: Int, f: StoredRow => StoredRow): CellStore =
    val key = Trie.key(row)
    val h = math.max(height, Trie.heightFor(key))
    val grown = Trie.grow(root, height, h)
    val next = Trie.modify(grown, 5 * (h - 1), key, f)
    if next eq grown then this else CellStore.normalized(next, h)

  /** This store with `cell` stored at `cell.ref`, replacing whatever was there. */
  def updated(cell: Cell): CellStore =
    val meta = Meta.of(cell.value, cell.styleId, cell.hyperlink.isDefined)
    val link = cell.hyperlink.fold("")(_.target)
    modifyRow(
      cell.ref.row.index0,
      row => StoredRow.set(row, cell.ref.col.index0, meta, cell.value, link)
    )

  /**
   * This store with `value` at `ref`. An existing cell keeps its style and hyperlink; an absent one
   * becomes a plain cell (even when `value` is Empty), as [[Sheet.put]] does.
   */
  def updatedValue(ref: ARef, value: CellValue): CellStore =
    modifyRow(ref.row.index0, row => StoredRow.setValue(row, ref.col.index0, value))

  /** This store without a cell at `ref`; `this` when there is none. */
  def removed(ref: ARef): CellStore =
    val col = ref.col.index0
    modifyRow(ref.row.index0, row => StoredRow.remove(row, col))

  /** This store without the cells inside `range`; `this` when there are none. */
  def removedIn(range: CellRange): CellStore =
    val c0 = range.start.col.index0
    val c1 = range.end.col.index0
    val skip = (e: Entry) => e.count == 0 || e.maxCol < c0 || e.minCol > c1
    val drop = (row: StoredRow) => StoredRow.retain(row, c => c < c0 || c > c1)
    if c0 > c1 then this
    else mapRowSpans(range.start.row.index0, range.end.row.index0, skip, drop)

  /**
   * This store with every cell of `cells` stored, later cells replacing earlier ones at the same
   * reference: the result equals a fold of [[updated]]. Sorts the incoming references once and
   * merges them in, sharing every row they do not touch.
   */
  def putAll(cells: IterableOnce[Cell]): CellStore =
    val incoming = CellStore.from(cells)
    unionWith(incoming, valuesOnly = false)

  /**
   * This store with `values(i)` at `refs(i)`, in any order (a later pair wins over an earlier one
   * at the same reference). Existing cells keep their style and hyperlink; absent references become
   * plain cells. When the arrays differ in length the extra elements of the longer one are ignored.
   */
  def withValues(refs: IArray[ARef], values: IArray[CellValue]): CellStore =
    val n = math.min(refs.length, values.length)
    val b = CellStore.newBuilder
    var i = 0
    while i < n do
      b.addOne(Cell(refs(i), values(i)))
      i += 1
    unionWith(b.result(), valuesOnly = true)

  private def unionWith(incoming: CellStore, valuesOnly: Boolean): CellStore =
    if incoming.isEmpty then this
    else if isEmpty then incoming
    else
      val h = math.max(height, incoming.height)
      val a = Trie.grow(root, height, h)
      val b = Trie.grow(incoming.root, incoming.height, h)
      val merged = Trie.union(a, b, 5 * (h - 1), valuesOnly)
      if merged eq a then this else new CellStore(merged, h)

  /**
   * This store with `f` applied to the value of every cell whose value tag is in `tagMask`; styles
   * and hyperlinks are kept. Rows (and whole subtrees) where `f` changes nothing, by reference, are
   * shared; `this` when nothing changes.
   */
  def transformValues(tagMask: Int)(f: CellValue => CellValue): CellStore =
    if (root.tagMask & tagMask) == 0 then this
    else
      val skip = (e: Entry) => (e.tagMask & tagMask) == 0
      mapRowSpans(Int.MinValue, Int.MaxValue, skip, row => StoredRow.mapValues(row, tagMask, f))

  /**
   * This store with its rows moved as a row insert or delete moves them: a row at or after `at`
   * moves by `by`. With `by < 0` the rows `[at + by, at)` are deleted first, so a delete of `n`
   * rows starting at `r` is `shiftRows(r + n, -n)` and an insert of `n` rows at `r` is
   * `shiftRows(r, n)`. A row whose destination leaves the `Int` range is dropped (on Excel's grid
   * this never happens), so no two rows can land on one index. Rows are re-keyed, not rebuilt.
   */
  def shiftRows(at: Int, by: Int): CellStore =
    if by == 0 || isEmpty then this
    else
      val n = rowCount
      val keys = new Array[Long](n)
      val leaves = new Array[StoredRow](n)
      var k = 0
      Trie.foreachRow(
        root,
        5 * (height - 1),
        0L,
        (key, leaf) =>
          val dest = Shift.dest(key.toInt, at, by)
          if dest != Shift.Dropped then
            keys(k) = Trie.key(dest.toInt)
            leaves(k) = leaf
            k += 1
      )
      val order = StableSort.order(keys, k)
      val sortedKeys = new Array[Long](k)
      val sortedLeaves = new Array[StoredRow](k)
      var i = 0
      while i < k do
        sortedKeys(i) = keys(order(i))
        sortedLeaves(i) = leaves(order(i))
        i += 1
      val (r, h) = Trie.build(sortedKeys, sortedLeaves, k)
      new CellStore(r, h)

  /**
   * This store with its columns moved as a column insert or delete moves them, by the rule of
   * [[shiftRows]] applied to every row's columns. Rows with no cell at or after the first affected
   * column are shared.
   */
  def shiftCols(at: Int, by: Int): CellStore =
    if by == 0 || isEmpty then this
    else
      val first = if by > 0 then at.toLong else at.toLong + by
      val skip = (e: Entry) => e.count == 0 || e.maxCol.toLong < first
      mapRowSpans(Int.MinValue, Int.MaxValue, skip, row => StoredRow.shift(row, at, by))

  // ----- internals -----

  /** The number of rows holding a cell. */
  private def rowCount: Int =
    var n = 0
    Trie.foreachRow(root, 5 * (height - 1), 0L, (_, _) => n += 1)
    n

  /** `f` over the rows `rowFrom..rowTo` (signed, inclusive). */
  private def mapRowSpans(
    rowFrom: Int,
    rowTo: Int,
    skip: Entry => Boolean,
    f: StoredRow => StoredRow
  ): CellStore =
    if isEmpty || rowFrom > rowTo then this
    else
      val shift = 5 * (height - 1)
      def span(r: Entry, lo: Long, hi: Long): Entry = Trie.mapRows(r, shift, 0L, lo, hi, skip, f)
      val next =
        if rowFrom >= 0 || rowTo < 0 then span(root, Trie.key(rowFrom), Trie.key(rowTo))
        else span(span(root, 0L, Trie.key(rowTo)), Trie.key(rowFrom), Trie.UnsignedMask)
      if next eq root then this else CellStore.normalized(next, height)

  /** The trie height (0 for the empty store); exposed for the canonical-shape laws. */
  private[xl] def debugHeight: Int = height

  /** Every structural invariant, or a description of the first one broken (for the law specs). */
  private[xl] def checkInvariants: Either[String, Unit] = CellStore.check(this)

  override def equals(that: Any): Boolean = that match
    case other: CellStore =>
      (this eq other) || (height == other.height && Trie.same(root, other.root))
    case _ => false

  override def hashCode: Int =
    var h = MurmurHash3.mix(0x5ce11, size)
    val c = CellCursor.all(this)
    while c.next() do
      h = MurmurHash3.mix(h, c.row)
      h = MurmurHash3.mix(h, c.col)
      h = MurmurHash3.mix(h, (c.meta ^ (c.meta >>> 32)).toInt)
      h = MurmurHash3.mix(h, c.value.hashCode)
      h = MurmurHash3.mix(h, c.link.hashCode)
    MurmurHash3.finalizeHash(h, size)

  override def toString: String =
    val sb = new StringBuilder("CellStore(")
    val c = CellCursor.all(this)
    var first = true
    while c.next() do
      if !first then sb.append(", ")
      first = false
      sb.append(label(c.row, c.col)).append(" -> ").append(c.value)
      if c.hasStyle then sb.append(" style=").append(c.rawStyle)
      if c.hasLink then sb.append(" link=").append(c.link)
    sb.append(")").toString

end CellStore

// The builder is a mutable accumulator (like any collection builder) and the invariant checker
// loops over arrays; both keep pure signatures.
@SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
object CellStore:

  // ----- value tags (for hasTag / transformValues) -----

  final val TagEmpty = Meta.Empty
  final val TagText = Meta.Text
  final val TagRichText = Meta.RichText
  final val TagNumber = Meta.Number
  final val TagBool = Meta.Bool
  final val TagDateTime = Meta.DateTime
  final val TagFormulaUncached = Meta.FormulaUncached
  final val TagFormulaCached = Meta.FormulaCached
  final val TagFormulaDataTable = Meta.FormulaDataTable
  final val TagError = Meta.Error

  /** The mask bit of one tag. */
  def tagBit(tag: Int): Int = 1 << tag

  /** Every formula tag: uncached, cached and data-table. */
  final val FormulaTags: Int =
    (1 << Meta.FormulaUncached) | (1 << Meta.FormulaCached) | (1 << Meta.FormulaDataTable)

  /** Every tag. */
  final val AllTags: Int = (1 << 10) - 1

  /**
   * The tag of a value: Empty 0, Text 1, RichText 2, Number 3, Bool 4, DateTime 5, Formula 6..8,
   * Error 9.
   */
  def tagOf(value: CellValue): Int = Meta.tagOf(value)

  // ----- construction -----

  val empty: CellStore = new CellStore(Node.Empty, 0)

  def of(cells: Cell*): CellStore = from(cells)

  /** The store of `cells`; a later cell replaces an earlier one at the same reference. */
  def from(cells: IterableOnce[Cell]): CellStore = newBuilder.addAll(cells).result()

  def newBuilder: Builder = new Builder

  private def normalized(root: Entry, height: Int): CellStore =
    val (r, h) = Trie.shrink(root, height)
    if r.count == 0 then empty else new CellStore(r, h)

  private def label(row: Int, col: Int): String =
    if row >= 0 && row <= 1048575 && col >= 0 && col <= 16383 then ARef.from0(col, row).toA1
    else s"R${row}C$col"

  /**
   * The sort key of a cell in the store's order: the unsigned row in the high word, the signed
   * column in the low word, each flipped at its sign bit so a signed `Long` comparison orders them.
   * (Raw `ARef` longs do not order this way off the grid.)
   */
  private[sheets] def orderKey(row: Int, col: Int): Long =
    ((row ^ Int.MinValue).toLong << 32) | ((col ^ Int.MinValue).toLong & 0xffffffffL)

  private[sheets] def rowOfKey(k: Long): Int = (k >>> 32).toInt ^ Int.MinValue
  private[sheets] def colOfKey(k: Long): Int = k.toInt ^ Int.MinValue

  /**
   * Accumulates cells into a store. A later cell replaces an earlier one at the same reference.
   * [[result]] leaves the builder as it was, so it can keep accumulating.
   *
   * A cell that comes after every cell before it, in the store's row-major order, is streamed (a
   * worksheet reader produces nothing else): the builder holds the cells of one slot group (32
   * rows) in scratch arrays, turns the group into its final [[Block]] or rows when the next group
   * begins, and keeps only the finished groups. So beyond the store being built it holds one group
   * of cells, not a slot per cell, and a block's arrays are copied once, straight from the scratch.
   *
   * A cell that comes before the last streamed one is buffered instead (the tail). [[result]] sorts
   * the tail once, stably, streams it into a store of its own and merges that over the streamed
   * cells, the tail winning. That is last-write-wins: a tail cell sorts before the cell streamed
   * last when it arrived, and every cell streamed since sorts after that one, so a streamed cell at
   * a tail cell's reference can only have arrived earlier.
   */
  final class Builder private[CellStore] ():
    // ----- the streamed cells: finished groups, then the open group's scratch -----

    /** The finished groups: `groupKeys(i)` is the group's row keys above their five low bits. */
    private var groupKeys = new Array[Long](8)
    private var groups = new Array[Entry](8)
    private var groupCount = 0

    /**
     * The open group's cells in order. `links` is empty until a cell brings a hyperlink and grows
     * to the scratch's size only when one does, so it may be shorter than the other three; it is
     * read only where the meta word has its link bit.
     */
    private var cols = new Array[Int](16)
    private var metas = new Array[Long](16)
    private var values = new Array[CellValue](16)
    private var links = RowLeaf.NoLinks
    private var n = 0

    /** The open group's rows: row `i` is `rowLows(i)` and starts at cell `rowStarts(i)`. */
    private val rowLows = new Array[Int](32)
    private val rowStarts = new Array[Int](32)
    private var rowCount = 0
    private var openGroup = 0L

    /** Whether every cell of the open group lies in the 64-column chunk `openChunk`. */
    private var oneChunk = true
    private var openChunk = 0

    /** The order key of the last streamed cell, meaningful once `any`. */
    private var lastKey = 0L
    private var any = false

    // ----- the tail: the cells that arrived before the last streamed one, in arrival order -----

    private var tailKeys = Array.emptyLongArray
    private var tailMetas = Array.emptyLongArray
    private var tailValues = Array.empty[CellValue]
    private var tailLinks = RowLeaf.NoLinks
    private var tailCount = 0

    def addOne(cell: Cell): this.type =
      add(
        orderKey(cell.ref.row.index0, cell.ref.col.index0),
        Meta.of(cell.value, cell.styleId, cell.hyperlink.isDefined),
        cell.value,
        cell.hyperlink.fold("")(_.target)
      )
      this

    def addAll(cells: IterableOnce[Cell]): this.type =
      val it = cells.iterator
      while it.hasNext do addOne(it.next())
      this

    def +=(cell: Cell): this.type = addOne(cell)

    /** One cell by its stored fields; `link` is read only when `meta` has the link bit. */
    private def add(key: Long, meta: Long, value: CellValue, link: String): Unit =
      if !any || key > lastKey then
        val row = rowOfKey(key)
        val col = colOfKey(key)
        val group = Trie.key(row) >>> 5
        if n > 0 && group != openGroup then closeGroup()
        if n == 0 then
          openGroup = group
          oneChunk = true
          openChunk = col >> 6
          rowCount = 0
        if n == 0 || rowOfKey(lastKey) != row then
          rowLows(rowCount) = (Trie.key(row) & 31L).toInt
          rowStarts(rowCount) = n
          rowCount += 1
        if (col >> 6) != openChunk then oneChunk = false
        if n == cols.length then growGroup()
        cols(n) = col
        n += 1
        setStreamed(n - 1, meta, value, link)
        lastKey = key
        any = true
      else if key == lastKey then setStreamed(n - 1, meta, value, link) // last write wins in place
      else addTail(key, meta, value, link)

    private def setStreamed(at: Int, meta: Long, value: CellValue, link: String): Unit =
      metas(at) = meta
      values(at) = value
      if Meta.hasLink(meta) then
        if links.length < metas.length then links = java.util.Arrays.copyOf(links, metas.length)
        links(at) = link

    private def growGroup(): Unit =
      val cap = cols.length * 2
      cols = java.util.Arrays.copyOf(cols, cap)
      metas = java.util.Arrays.copyOf(metas, cap)
      values = java.util.Arrays.copyOf(values, cap)

    private def addTail(key: Long, meta: Long, value: CellValue, link: String): Unit =
      val at =
        if tailCount > 0 && tailKeys(tailCount - 1) == key then tailCount - 1
        else
          if tailCount == tailKeys.length then
            val cap = math.max(16, tailCount * 2)
            tailKeys = java.util.Arrays.copyOf(tailKeys, cap)
            tailMetas = java.util.Arrays.copyOf(tailMetas, cap)
            tailValues = java.util.Arrays.copyOf(tailValues, cap)
          tailCount += 1
          tailCount - 1
      tailKeys(at) = key
      tailMetas(at) = meta
      tailValues(at) = value
      if Meta.hasLink(meta) then
        if tailLinks.length < tailKeys.length then
          tailLinks = java.util.Arrays.copyOf(tailLinks, tailKeys.length)
        tailLinks(at) = link

    /** The open group becomes a finished one. */
    private def closeGroup(): Unit =
      if groupCount == groups.length then
        groupKeys = java.util.Arrays.copyOf(groupKeys, groupCount * 2)
        groups = java.util.Arrays.copyOf(groups, groupCount * 2)
      groupKeys(groupCount) = openGroup
      groups(groupCount) = openEntry()
      groupCount += 1
      n = 0

    /**
     * The open group (which holds a cell) in its canonical form, read from the scratch without
     * changing it: a [[Block]] when its cells share one chunk and fit the cap (the rule of
     * [[Group.of]]), its arrays copied from the scratch once; otherwise its rows.
     */
    private def openEntry(): Entry =
      if oneChunk && n <= Group.MaxBlockCells then
        var rowMask = 0
        val rowMasks = new Array[Long](rowCount)
        var anyLink = false
        var r = 0
        while r < rowCount do
          rowMask |= 1 << rowLows(r)
          val until = if r + 1 < rowCount then rowStarts(r + 1) else n
          var mask = 0L
          var p = rowStarts(r)
          while p < until do
            mask |= 1L << (cols(p) & 63)
            if Meta.hasLink(metas(p)) then anyLink = true
            p += 1
          rowMasks(r) = mask
          r += 1
        val blockLinks =
          if !anyLink then RowLeaf.NoLinks
          else
            val out = new Array[String](n)
            var p = 0
            while p < n do
              out(p) = if Meta.hasLink(metas(p)) then links(p) else ""
              p += 1
            out
        Group.make(
          rowMask,
          openChunk,
          rowMasks,
          java.util.Arrays.copyOf(metas, n),
          java.util.Arrays.copyOf(values, n),
          blockLinks
        )
      else
        val leaves = new Array[StoredRow](rowCount)
        var r = 0
        while r < rowCount do
          val until = if r + 1 < rowCount then rowStarts(r + 1) else n
          leaves(r) = StoredRow.build(cols, metas, values, links, rowStarts(r), until)
          r += 1
        Group.of(rowLows, leaves, rowCount)

    /** The store of the streamed cells. */
    private def streamed(): CellStore =
      if n == 0 then
        if groupCount == 0 then empty
        else
          val (root, h) = Trie.buildGroups(groupKeys, groups, groupCount)
          new CellStore(root, h)
      else
        // the open group goes last, in copies: the builder's own arrays stay as they are
        val keys = java.util.Arrays.copyOf(groupKeys, groupCount + 1)
        val entries = java.util.Arrays.copyOf(groups, groupCount + 1)
        keys(groupCount) = openGroup
        entries(groupCount) = openEntry()
        val (root, h) = Trie.buildGroups(keys, entries, groupCount + 1)
        new CellStore(root, h)

    /** The store of the tail: sorted stably, the last of each run of equal references kept. */
    private def tail(): CellStore =
      val order = StableSort.order(tailKeys, tailCount)
      val b = new Builder
      var i = 0
      while i < tailCount do
        val p = order(i)
        if i + 1 >= tailCount || tailKeys(order(i + 1)) != tailKeys(p) then
          val meta = tailMetas(p)
          b.add(tailKeys(p), meta, tailValues(p), if Meta.hasLink(meta) then tailLinks(p) else "")
        i += 1
      b.streamed()

    def result(): CellStore =
      if tailCount == 0 then streamed() else streamed().unionWith(tail(), valuesOnly = false)
  end Builder

  // ----- invariants (checked by the law spec) -----

  private def check(s: CellStore): Either[String, Unit] =
    if s.height == 0 then
      if s.root.count == 0 && (s.root eq Node.Empty) then Right(())
      else Left("height 0 but non-empty")
    else if s.root.count == 0 then Left("empty root with height > 0")
    else if Trie.heightFor(
        Trie.lastKey(s.root, 5 * (s.height - 1), 0L, 0L, Trie.UnsignedMask, false)
      ) != s.height
    then Left("height is not the minimal height for the largest row key")
    else checkEntry(s.root, 5 * (s.height - 1), 0L)

  private def checkEntry(e: Entry, shift: Int, prefix: Long): Either[String, Unit] = e match
    case n: Node => checkNode(n, shift, prefix)
    case b: Block => if shift != 0 then Left("block above shift 0") else checkBlock(b, prefix)
    case l: RowLeaf => Left(s"leaf outside a shift-0 node at $prefix")
    case r: BandedRow => Left(s"banded row outside a shift-0 node at $prefix")

  private def checkNode(node: Node, shift: Int, prefix: Long): Either[String, Unit] =
    if node.children.length != Integer.bitCount(node.bitmap) then Left("children size != popcount")
    else if node.bitmap == 0 then Left("empty inner node")
    else if shift == 0 && blockShaped(node) then Left(s"shift-0 node at $prefix should be a block")
    else
      var count = 0
      var valued = 0
      var minCol = Int.MaxValue
      var maxCol = Int.MinValue
      var vMin = Int.MaxValue
      var vMax = Int.MinValue
      var tags = 0
      var err: Either[String, Unit] = Right(())
      var bm = node.bitmap
      var i = 0
      while err.isRight && bm != 0 do
        val slot = Integer.numberOfTrailingZeros(bm)
        bm &= bm - 1
        val child = node.children(i)
        val childLo = prefix + (slot.toLong << shift)
        err = child match
          case l: RowLeaf =>
            if shift != 0 then Left("leaf above shift 0") else checkLeaf(l, childLo)
          case r: BandedRow =>
            if shift != 0 then Left("banded row above shift 0") else checkBanded(r, childLo)
          case other =>
            if shift == 0 then Left("group below shift 0")
            else checkEntry(other, shift - 5, childLo)
        if child.count == 0 then err = Left("empty child")
        count += child.count
        valued += child.valued
        minCol = math.min(minCol, child.minCol)
        maxCol = math.max(maxCol, child.maxCol)
        vMin = math.min(vMin, child.vMinCol)
        vMax = math.max(vMax, child.vMaxCol)
        tags |= child.tagMask
        i += 1
      if err.isLeft then err
      else if count != node.count || valued != node.valued || minCol != node.minCol ||
        maxCol != node.maxCol || vMin != node.vMinCol || vMax != node.vMaxCol || tags != node.tagMask
      then Left(s"node summary mismatch at prefix $prefix")
      else Right(())

  /** The block rule, restated independently of [[Group]]: one shared chunk, at most the cap. */
  private def blockShaped(node: Node): Boolean =
    val chunks = node.children.toList.map {
      case l: RowLeaf => if l.chunkCount == 1 then Some(l.chunkAt(0)) else None
      case _ => None
    }
    chunks.forall(_.isDefined) && chunks.distinct.size == 1 &&
    node.count <= Group.MaxBlockCells

  /** Per-cell checks shared by leaves and blocks: meta bits, tags, canonical links. */
  private def checkCells(c: Cells): Either[String, Unit] =
    val n = c.count
    if c.values.length != n then Left("values length")
    else if c.links.length != 0 && c.links.length != n then Left("links length")
    else
      var anyLink = false
      var err: Either[String, Unit] = Right(())
      var p = 0
      while p < n do
        val m = c.meta(p)
        if Meta.hasLink(m) then anyLink = true
        else if c.links.length > 0 && c.links(p) != "" then err = Left("filler link not empty")
        if !Meta.hasStyle(m) && Meta.rawStyle(m) != 0 then err = Left("style bits without flag")
        if (m & ~(Meta.TagBits | Meta.StyleBit | Meta.LinkBit | 0xffffffffL)) != 0L then
          err = Left("stray meta bits")
        if Meta.tag(m) != Meta.tagOf(c.values(p)) then err = Left("tag mismatch")
        p += 1
      if anyLink != (c.links.length > 0) then err = Left("links array not canonical")
      err

  private def checkLeaf(l: RowLeaf, key: Long): Either[String, Unit] =
    val n = l.count
    if n == 0 then Left(s"empty leaf at $key")
    else
      var err = checkCells(l)
      // chunk index: ascending chunks, non-zero masks, cumulative bases, narrow iff one chunk
      var base = 0
      var i = 0
      while i < l.chunkCount do
        if l.maskAt(i) == 0L then err = Left("zero chunk mask")
        if i > 0 && l.chunkAt(i) <= l.chunkAt(i - 1) then err = Left("chunks not ascending")
        if l.baseAt(i) != base then err = Left("bad chunk base")
        base += java.lang.Long.bitCount(l.maskAt(i))
        i += 1
      if base != n then err = Left("masks do not cover the arrays")
      l match
        case _: WideRow if l.chunkCount < 2 => err = Left("one-chunk WideRow")
        case _ => ()
      // the band rule, restated independently of StoredRow: a leaf's columns share one band
      // unless it holds no more than Band.FlatMax cells
      if n > Band.FlatMax && (l.cols(0) >> Band.Shift) != (l.cols(n - 1) >> Band.Shift) then
        err = Left(s"leaf of $n cells at $key spans several bands")
      // summaries recomputed from scratch
      val again = RowLeaf.build(l.cols, l.meta, l.values, l.links, 0, n)
      if again.valued != l.valued || again.minCol != l.minCol || again.maxCol != l.maxCol ||
        again.vMinCol != l.vMinCol || again.vMaxCol != l.vMaxCol || again.tagMask != l.tagMask
      then err = Left("leaf summary mismatch")
      err

  /**
   * A banded row: two or more bands, more than `Band.FlatMax` cells, ids strictly ascending, each
   * band a valid non-empty leaf whose columns all lie in its band, and summaries equal to the ones
   * recomputed from the bands.
   */
  private def checkBanded(r: BandedRow, key: Long): Either[String, Unit] =
    val n = r.bands.length
    if n < 2 then Left(s"banded row at $key with fewer than two bands")
    else if r.count <= Band.FlatMax then Left(s"banded row at $key holds only ${r.count} cells")
    else if r.ids.length != n then Left("band ids and bands differ in length")
    else
      var err: Either[String, Unit] = Right(())
      var count = 0
      var valued = 0
      var minCol = Int.MaxValue
      var maxCol = Int.MinValue
      var vMin = Int.MaxValue
      var vMax = Int.MinValue
      var tags = 0
      var i = 0
      while err.isRight && i < n do
        val band = r.bands(i)
        if i > 0 && r.ids(i) <= r.ids(i - 1) then err = Left("band ids not ascending")
        else if band.count == 0 then err = Left(s"empty band ${r.ids(i)} at $key")
        else
          err = checkLeaf(band, key)
          val cols = band.cols
          if (cols(0) >> Band.Shift) != r.ids(i) || (cols(cols.length - 1) >> Band.Shift) != r.ids(
              i
            )
          then err = Left(s"band ${r.ids(i)} at $key holds a column of another band")
          count += band.count
          valued += band.valued
          minCol = math.min(minCol, band.minCol)
          maxCol = math.max(maxCol, band.maxCol)
          vMin = math.min(vMin, band.vMinCol)
          vMax = math.max(vMax, band.vMaxCol)
          tags |= band.tagMask
        i += 1
      if err.isLeft then err
      else if count != r.count || valued != r.valued || minCol != r.minCol || maxCol != r.maxCol ||
        vMin != r.vMinCol || vMax != r.vMaxCol || tags != r.tagMask
      then Left(s"banded row summary mismatch at $key")
      else Right(())

  private def checkBlock(b: Block, prefix: Long): Either[String, Unit] =
    val rows = b.rows
    if b.rowMask == 0 || b.count == 0 then Left(s"empty block at $prefix")
    else if b.count > Group.MaxBlockCells then Left(s"block over the cell cap at $prefix")
    else if b.uniform && b.masks.length != 1 then Left("uniform block with several masks")
    else if !b.uniform && (b.masks.length != rows || b.bases.length != rows) then
      Left("per-row masks/bases sized wrong")
    else
      var err = checkCells(b)
      val rowMasks = new Array[Long](rows)
      var base = 0
      var i = 0
      while i < rows do
        rowMasks(i) = b.rowMaskAt(i)
        if rowMasks(i) == 0L then err = Left("empty row in block")
        if b.rowBaseAt(i) != base then err = Left("bad row base")
        base += java.lang.Long.bitCount(rowMasks(i))
        i += 1
      if base != b.count then err = Left("row masks do not cover the arrays")
      // canonical masks and summaries: rebuild from the per-row masks and compare
      if err.isRight then
        val again = Group.make(b.rowMask, b.chunk, rowMasks, b.meta, b.values, b.links)
        if !java.util.Arrays.equals(again.masks, b.masks) ||
          !java.util.Arrays.equals(again.bases, b.bases)
        then err = Left("block masks not canonical")
        else if again.valued != b.valued || again.minCol != b.minCol || again.maxCol != b.maxCol ||
          again.vMinCol != b.vMinCol || again.vMaxCol != b.vMaxCol || again.tagMask != b.tagMask
        then err = Left("block summary mismatch")
      err

  private def cellsOf(c: CellCursor): Iterator[Cell] = new Iterator[Cell]:
    private var ready = false
    private var done = false
    def hasNext: Boolean =
      if !ready && !done then if c.next() then ready = true else done = true
      ready
    def next(): Cell =
      if !hasNext then Iterator.empty.next()
      else
        ready = false
        c.cell

  private def refsOf(c: CellCursor): Iterator[ARef] = new Iterator[ARef]:
    private var ready = false
    private var done = false
    def hasNext: Boolean =
      if !ready && !done then if c.next() then ready = true else done = true
      ready
    def next(): ARef =
      if !hasNext then Iterator.empty.next()
      else
        ready = false
        c.ref
