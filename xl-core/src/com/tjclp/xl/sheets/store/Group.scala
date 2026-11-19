package com.tjclp.xl.sheets.store

import com.tjclp.xl.cells.CellValue

/**
 * Shift-0 slot groups (#891): the up-to-32 rows under one trie slot at shift 5, stored either as a
 * [[Block]] (when every cell of the group sits in one 64-column chunk) or as a shift-0 [[Node]] of
 * rows ([[StoredRow]]s) otherwise. The form is a function of the content, so equal groups have
 * equal shapes. Rows are addressed by `low`, the row key's five low bits.
 */
// Flat-array loops over the trie (house style, as in SccKernel/LongIntMap): local vars and while
// loops behind pure signatures; nothing mutable escapes.
@SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
private[xl] object Group:
  val NoBases: Array[Int] = Array.emptyIntArray

  /**
   * The most cells a block holds. A write copies its whole block, so blocks are kept to narrow
   * groups (eight full columns of 32 rows); wider groups stay per-row leaves, whose rows already
   * amortise their overhead. Measured on 100k-cell sheets: at 640 cells per block (20 columns) a
   * row-major append chain ran 2.6x slower than per-row leaves.
   */
  final val MaxBlockCells = 256

  /** A one-cell block a cursor holds before it reaches its first block (never part of a trie). */
  val Placeholder: Block =
    make(1, 0, Array(1L), Array(0L), Array[CellValue](CellValue.Empty), RowLeaf.NoLinks)

  /**
   * The canonical group of rows `lows(i) -> leaves(i)` (`i < n`, lows ascending, leaves non-empty).
   */
  def of(lows: Array[Int], leaves: Array[StoredRow], n: Int): Entry =
    if n == 0 then Node.Empty
    else if blockable(leaves, n) then blockOf(lows, leaves, n)
    else
      var bitmap = 0
      val children = new Array[Entry](n)
      var i = 0
      while i < n do
        bitmap |= 1 << lows(i)
        children(i) = leaves(i)
        i += 1
      Node.of(bitmap, children)

  /** A shift-0 node in canonical form: a [[Block]] when its rows allow one. */
  def compact(node: Node): Entry =
    if node.count == 0 then Node.Empty
    else if node.count > MaxBlockCells || !sharesOneChunk(node.children) then node
    else
      val (lows, leaves, n) = explode(node)
      blockOf(lows, leaves, n)

  /** Whether these leaves are all narrow rows over one chunk (checked without copying). */
  private def sharesOneChunk(children: Array[Entry]): Boolean =
    var ok = true
    var chunk = 0
    var i = 0
    while ok && i < children.length do
      children(i) match
        case r: NarrowRow => if i == 0 then chunk = r.chunk else ok = r.chunk == chunk
        case _ => ok = false
      i += 1
    ok

  /** Whether these rows all lie in one common 64-column chunk and fit in a block. */
  private def blockable(leaves: Array[StoredRow], n: Int): Boolean =
    var ok = true
    var chunk = 0
    var total = 0
    var i = 0
    while ok && i < n do
      total += leaves(i).count
      if total > MaxBlockCells then ok = false
      leaves(i) match
        case r: NarrowRow =>
          if i == 0 then chunk = r.chunk else ok = ok && r.chunk == chunk
        case _ => ok = false
      i += 1
    ok

  /** The block of these rows, which [[blockable]] has found to be narrow rows over one chunk. */
  private def blockOf(lows: Array[Int], leaves: Array[StoredRow], n: Int): Block =
    var rowMask = 0
    var chunk = 0
    var total = 0
    var anyLink = false
    var i = 0
    val rowMasks = new Array[Long](n)
    while i < n do
      rowMask |= 1 << lows(i)
      total += leaves(i).count
      leaves(i) match
        case l: NarrowRow =>
          chunk = l.chunk
          rowMasks(i) = l.mask
          if l.links.length > 0 then anyLink = true
        case _ => () // unreachable: a block's rows are narrow
      i += 1
    val meta = new Array[Long](total)
    val values = new Array[CellValue](total)
    val links = if anyLink then new Array[String](total) else RowLeaf.NoLinks
    var p = 0
    i = 0
    while i < n do
      leaves(i) match
        case l: NarrowRow =>
          System.arraycopy(l.meta, 0, meta, p, l.count)
          System.arraycopy(l.values, 0, values, p, l.count)
          if anyLink then
            var q = 0
            while q < l.count do
              links(p + q) = l.linkAt(q)
              q += 1
          p += l.count
        case _ => () // unreachable: a block's rows are narrow
      i += 1
    make(rowMask, chunk, rowMasks, meta, values, links)

  /**
   * The block with these rows (`rowMasks` one per present row of `rowMask`, all non-zero), its
   * arrays (owned from now on; `links` already canonical), canonical masks and summaries.
   */
  def make(
    rowMask: Int,
    chunk: Int,
    rowMasks: Array[Long],
    meta: Array[Long],
    values: Array[CellValue],
    links: Array[String]
  ): Block =
    val (vMin, vMax) = valuedBounds(chunk, rowMasks, meta)
    assemble(
      rowMask,
      chunk,
      rowMasks,
      meta,
      values,
      links,
      countValued(meta),
      vMin,
      vMax,
      tagsOf(meta)
    )

  /**
   * The block with these arrays and value summaries; canonicalizes the masks, derives col bounds.
   */
  private def assemble(
    rowMask: Int,
    chunk: Int,
    rowMasks: Array[Long],
    meta: Array[Long],
    values: Array[CellValue],
    links: Array[String],
    valued: Int,
    vMin: Int,
    vMax: Int,
    tagMask: Int
  ): Block =
    val rows = rowMasks.length
    var allEqual = true
    var orMask = 0L
    var i = 0
    while i < rows do
      if rowMasks(i) != rowMasks(0) then allEqual = false
      orMask |= rowMasks(i)
      i += 1
    val (masks, bases) =
      if allEqual then (Array(rowMasks(0)), NoBases)
      else
        val b = new Array[Int](rows)
        var acc = 0
        var j = 0
        while j < rows do
          b(j) = acc
          acc += java.lang.Long.bitCount(rowMasks(j))
          j += 1
        (rowMasks, b)
    val first = chunk << 6
    val minCol = first + java.lang.Long.numberOfTrailingZeros(orMask)
    val maxCol = first + 63 - java.lang.Long.numberOfLeadingZeros(orMask)
    new Block(
      rowMask,
      chunk,
      masks,
      bases,
      meta,
      values,
      links,
      valued,
      minCol,
      maxCol,
      vMin,
      vMax,
      tagMask
    )

  private def tagsOf(meta: Array[Long]): Int =
    var t = 0
    var p = 0
    while p < meta.length do
      t |= 1 << Meta.tag(meta(p))
      p += 1
    t

  private def countValued(meta: Array[Long]): Int =
    var n = 0
    var p = 0
    while p < meta.length do
      if Meta.tag(meta(p)) != Meta.Empty then n += 1
      p += 1
    n

  /** The (min, max) column of the valued cells; (Int.MaxValue, Int.MinValue) when none. */
  private def valuedBounds(chunk: Int, rowMasks: Array[Long], meta: Array[Long]): (Int, Int) =
    var vMin = Int.MaxValue
    var vMax = Int.MinValue
    val first = chunk << 6
    var p = 0
    var i = 0
    while i < rowMasks.length do
      var m = rowMasks(i)
      while m != 0L do
        if Meta.tag(meta(p)) != Meta.Empty then
          val col = first + java.lang.Long.numberOfTrailingZeros(m)
          if col < vMin then vMin = col
          if col > vMax then vMax = col
        m &= m - 1
        p += 1
      i += 1
    (vMin, vMax)

  /** Present row `i` of a block as a leaf (a fresh copy of its cells). */
  def rowLeaf(b: Block, i: Int): RowLeaf =
    val m = b.rowMaskAt(i)
    val base = b.rowBaseAt(i)
    val n = java.lang.Long.bitCount(m)
    val meta = java.util.Arrays.copyOfRange(b.meta, base, base + n)
    val values = java.util.Arrays.copyOfRange(b.values, base, base + n)
    val links =
      if b.links.length == 0 || !anyLink(meta) then RowLeaf.NoLinks
      else java.util.Arrays.copyOfRange(b.links, base, base + n)
    RowLeaf.fromIndex(Array(m, b.chunk.toLong << 32), meta, values, links)

  private def fillEmpty(a: Array[String], from: Int, until: Int): Unit =
    var q = from
    while q < until do
      a(q) = ""
      q += 1

  private def anyLink(meta: Array[Long]): Boolean =
    var found = false
    var i = 0
    while !found && i < meta.length do
      found = Meta.hasLink(meta(i))
      i += 1
    found

  /** A shift-0 group's rows as (lows, leaves, n); a block's rows are copied out. */
  def explode(e: Entry): (Array[Int], Array[StoredRow], Int) = e match
    case b: Block =>
      val n = b.rows
      val lows = new Array[Int](n)
      val leaves = new Array[StoredRow](n)
      var bm = b.rowMask
      var i = 0
      while i < n do
        lows(i) = Integer.numberOfTrailingZeros(bm)
        bm &= bm - 1
        leaves(i) = rowLeaf(b, i)
        i += 1
      (lows, leaves, n)
    case node: Node =>
      val n = node.children.length
      val lows = new Array[Int](n)
      val leaves = new Array[StoredRow](n)
      var bm = node.bitmap
      var i = 0
      while i < n do
        lows(i) = Integer.numberOfTrailingZeros(bm)
        bm &= bm - 1
        node.children(i) match
          case l: RowLeaf => leaves(i) = l
          case r: BandedRow => leaves(i) = r
          case _ => () // unreachable: a shift-0 node holds rows
        i += 1
      (lows, leaves, n)
    case l: RowLeaf => (Array(0), Array[StoredRow](l), if l.count == 0 then 0 else 1)
    case r: BandedRow => (Array(0), Array[StoredRow](r), 1)

  /**
   * `b` with row `low` (currently `old`, empty when absent) replaced by `row` (an empty leaf
   * removes the row). A row that has come to span several bands cannot stay in a block, so the
   * group is rebuilt as a node.
   */
  def withRow(b: Block, low: Int, old: RowLeaf, row: StoredRow): Entry = row match
    case leaf: RowLeaf => withLeaf(b, low, old, leaf)
    case banded: BandedRow =>
      val (lows, leaves, n) = explode(b)
      val (l2, v2, n2) = replaceRow(lows, leaves, n, low, banded)
      of(l2, v2, n2)

  /**
   * `b` with row `low` (currently `old`, empty when absent) replaced by `leaf` (an empty leaf
   * removes the row). The value summaries are derived from the old ones and the two rows, with a
   * rescan only when the old row may have held the block's only instance of an extreme or tag.
   * Splices the block's arrays directly when the result is still a block; otherwise rebuilds the
   * group as a node.
   */
  private def withLeaf(b: Block, low: Int, old: RowLeaf, leaf: RowLeaf): Entry =
    val bit = 1 << low
    val present = (b.rowMask & bit) != 0
    val others = b.rowMask & ~bit
    val chunk = if others == 0 then leaf.chunkAt(0) else b.chunk
    val fits = leaf.count == 0 || (b.count - old.count + leaf.count <= MaxBlockCells && (leaf match
      case r: NarrowRow => r.chunk == chunk
      case _: WideRow => false))
    if !fits then
      val (lows, leaves, n) = explode(b)
      val (l2, v2, n2) = replaceRow(lows, leaves, n, low, leaf)
      of(l2, v2, n2)
    else
      val newRowMask = if leaf.count == 0 then others else b.rowMask | bit
      if newRowMask == 0 then Node.Empty
      else
        val i = b.rowIndex(low)
        val oldCount = if present then java.lang.Long.bitCount(b.rowMaskAt(i)) else 0
        val at = if i < b.rows then b.rowBaseAt(i) else b.count
        val newCount = leaf.count
        val total = b.count - oldCount + newCount
        val tail = b.count - at - oldCount
        val meta = new Array[Long](total)
        val values = new Array[CellValue](total)
        System.arraycopy(b.meta, 0, meta, 0, at)
        System.arraycopy(b.values, 0, values, 0, at)
        System.arraycopy(leaf.meta, 0, meta, at, newCount)
        System.arraycopy(leaf.values, 0, values, at, newCount)
        System.arraycopy(b.meta, at + oldCount, meta, at + newCount, tail)
        System.arraycopy(b.values, at + oldCount, values, at + newCount, tail)
        // links stay canonical: dropped when the old row held the block's last ones
        val noLinks =
          (b.links.length == 0 && leaf.links.length == 0) ||
            (leaf.links.length == 0 && old.links.length > 0 && !anyLink(meta))
        val links =
          if noLinks then RowLeaf.NoLinks
          else
            val out = new Array[String](total)
            if b.links.length > 0 then
              System.arraycopy(b.links, 0, out, 0, at)
              System.arraycopy(b.links, at + oldCount, out, at + newCount, tail)
            else
              fillEmpty(out, 0, at)
              fillEmpty(out, at + newCount, total)
            if leaf.links.length > 0 then System.arraycopy(leaf.links, 0, out, at, newCount)
            else fillEmpty(out, at, at + newCount)
            out
        val rows = Integer.bitCount(newRowMask)
        val rowMasks = new Array[Long](rows)
        var src = 0
        var dst = 0
        var bm = b.rowMask | newRowMask
        while bm != 0 do
          val r = Integer.numberOfTrailingZeros(bm)
          bm &= bm - 1
          if r == low then
            if present then src += 1
            if leaf.count > 0 then
              rowMasks(dst) = leaf.maskAt(0)
              dst += 1
          else
            rowMasks(dst) = b.rowMaskAt(src)
            src += 1
            dst += 1
        val valued = b.valued - old.valued + leaf.valued
        val tagMask =
          if (old.tagMask & ~leaf.tagMask) == 0 then b.tagMask | leaf.tagMask else tagsOf(meta)
        val minOk = old.valued == 0 || old.vMinCol > b.vMinCol ||
          (leaf.valued > 0 && leaf.vMinCol <= old.vMinCol)
        val maxOk = old.valued == 0 || old.vMaxCol < b.vMaxCol ||
          (leaf.valued > 0 && leaf.vMaxCol >= old.vMaxCol)
        val (vMin, vMax) =
          if valued == 0 then (Int.MaxValue, Int.MinValue)
          else if minOk && maxOk then
            (math.min(b.vMinCol, leaf.vMinCol), math.max(b.vMaxCol, leaf.vMaxCol))
          else valuedBounds(chunk, rowMasks, meta)
        assemble(newRowMask, chunk, rowMasks, meta, values, links, valued, vMin, vMax, tagMask)

  /** (lows, leaves) with row `low` set to `leaf` (removed when empty). */
  def replaceRow(
    lows: Array[Int],
    leaves: Array[StoredRow],
    n: Int,
    low: Int,
    leaf: StoredRow
  ): (Array[Int], Array[StoredRow], Int) =
    val outLows = new Array[Int](n + 1)
    val outLeaves = new Array[StoredRow](n + 1)
    var k = 0
    var inserted = false
    def put(l: Int, leaf: StoredRow): Unit =
      outLows(k) = l
      outLeaves(k) = leaf
      k += 1
    var i = 0
    while i < n do
      if lows(i) != low then
        if !inserted && lows(i) > low then
          inserted = true
          if leaf.count > 0 then put(low, leaf)
        put(lows(i), leaves(i))
      i += 1
    if !inserted && leaf.count > 0 then put(low, leaf)
    (outLows, outLeaves, k)

  /** Whether present row `i` of `b` holds a non-Empty value. */
  def rowValued(b: Block, i: Int): Boolean =
    val base = b.rowBaseAt(i)
    val end = base + java.lang.Long.bitCount(b.rowMaskAt(i))
    var found = false
    var p = base
    while !found && p < end do
      found = Meta.tag(b.meta(p)) != Meta.Empty
      p += 1
    found

  /** Structural equality of two blocks. */
  def same(a: Block, b: Block): Boolean =
    if a eq b then true
    else
      var ok = a.rowMask == b.rowMask && a.chunk == b.chunk && a.count == b.count &&
        java.util.Arrays.equals(a.masks, b.masks) && java.util.Arrays.equals(a.bases, b.bases) &&
        java.util.Arrays.equals(a.meta, b.meta) && a.links.length == b.links.length
      var p = 0
      while ok && p < a.count do
        ok = a.values(p) == b.values(p) && a.linkAt(p) == b.linkAt(p)
        p += 1
      ok
