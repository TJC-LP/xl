package com.tjclp.xl.sheets.store

import com.tjclp.xl.cells.{CellValue, FormulaKind}
import com.tjclp.xl.styles.units.StyleId

/**
 * The per-cell metadata word of a [[RowLeaf]] (#891).
 *
 * Bits 63..60 hold the value tag ([[Meta.tagOf]]), bit 59 says the cell has a style, bit 58 that it
 * has a hyperlink, and the low 32 bits hold the raw style index. The style travels behind its own
 * flag because every `Int`, -1 included, is a legal `StyleId`. When the flag is off the low word is
 * zero, so two cells that are equal have equal words.
 */
private[xl] object Meta:
  final val TagShift = 60
  final val TagBits: Long = 0xfL << TagShift
  final val StyleBit: Long = 1L << 59
  final val LinkBit: Long = 1L << 58

  final val Empty = 0
  final val Text = 1
  final val RichText = 2
  final val Number = 3
  final val Bool = 4
  final val DateTime = 5
  final val FormulaUncached = 6
  final val FormulaCached = 7
  final val FormulaDataTable = 8
  final val Error = 9

  /** The tag of a value. A data-table formula is tagged as one whether or not it holds a cache. */
  def tagOf(value: CellValue): Int = value match
    case CellValue.Empty => Empty
    case _: CellValue.Text => Text
    case _: CellValue.RichText => RichText
    case _: CellValue.Number => Number
    case _: CellValue.Bool => Bool
    case _: CellValue.DateTime => DateTime
    case CellValue.Formula(_, _, _: FormulaKind.DataTable) => FormulaDataTable
    case CellValue.Formula(_, None, _) => FormulaUncached
    case _: CellValue.Formula => FormulaCached
    case _: CellValue.Error => Error

  def of(value: CellValue, style: Option[StyleId], hasLink: Boolean): Long =
    val tagged = tagOf(value).toLong << TagShift
    val styled = style match
      case Some(s) => tagged | StyleBit | (s.value.toLong & 0xffffffffL)
      case None => tagged
    if hasLink then styled | LinkBit else styled

  def tag(meta: Long): Int = (meta >>> TagShift).toInt
  def hasStyle(meta: Long): Boolean = (meta & StyleBit) != 0L
  def hasLink(meta: Long): Boolean = (meta & LinkBit) != 0L
  def rawStyle(meta: Long): Int = meta.toInt
  def style(meta: Long): Option[StyleId] =
    if hasStyle(meta) then Some(StyleId(rawStyle(meta))) else None

  /** `meta` with its value tag replaced by the tag of `value`; style and link bits are kept. */
  def retag(meta: Long, value: CellValue): Long =
    (meta & ~TagBits) | (tagOf(value).toLong << TagShift)

/**
 * A trie entry: an inner [[Node]] or a [[RowLeaf]]. Every entry carries the summaries of the cells
 * below it. They are invariants, set when the entry is built and never changed:
 *
 *   - `count` cells stored, `valued` of them holding a non-Empty value
 *   - `minCol`/`maxCol` over all stored cells (meaningful only when `count > 0`)
 *   - `vMinCol`/`vMaxCol` over the valued cells (meaningful only when `valued > 0`)
 *   - `tagMask`, bit `t` set when some cell has value tag `t`
 *
 * The min/max sentinels for "none" are `Int.MaxValue`/`Int.MinValue`, which are also legal columns,
 * so readers check `count`/`valued` first.
 */
private[xl] sealed abstract class Entry(
  val count: Int,
  val valued: Int,
  val minCol: Int,
  val maxCol: Int,
  val vMinCol: Int,
  val vMaxCol: Int,
  val tagMask: Int
)

/**
 * An inner trie node. `children` holds one entry per set bit of `bitmap`, in slot order; slot `s`
 * of a node at `shift` covers row keys whose bits `shift+4..shift` equal `s`. A node at shift 0
 * holds [[RowLeaf]]s, every other node holds nodes. No child is ever empty.
 */
private[xl] final class Node private (
  val bitmap: Int,
  val children: Array[Entry],
  count: Int,
  valued: Int,
  minCol: Int,
  maxCol: Int,
  vMinCol: Int,
  vMaxCol: Int,
  tagMask: Int
) extends Entry(count, valued, minCol, maxCol, vMinCol, vMaxCol, tagMask)

// Flat-array loops over the trie (house style, as in SccKernel/LongIntMap): local vars and while
// loops behind pure signatures; nothing mutable escapes.
@SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
private[xl] object Node:
  val Empty: Node =
    new Node(0, Array.empty[Entry], 0, 0, Int.MaxValue, Int.MinValue, Int.MaxValue, Int.MinValue, 0)

  /** The node with these children (one per set bit of `bitmap`), its summaries computed. */
  def of(bitmap: Int, children: Array[Entry]): Node =
    if bitmap == 0 then Empty
    else
      var count = 0
      var valued = 0
      var minCol = Int.MaxValue
      var maxCol = Int.MinValue
      var vMinCol = Int.MaxValue
      var vMaxCol = Int.MinValue
      var tagMask = 0
      var i = 0
      while i < children.length do
        val c = children(i)
        count += c.count
        valued += c.valued
        if c.minCol < minCol then minCol = c.minCol
        if c.maxCol > maxCol then maxCol = c.maxCol
        if c.vMinCol < vMinCol then vMinCol = c.vMinCol
        if c.vMaxCol > vMaxCol then vMaxCol = c.vMaxCol
        tagMask |= c.tagMask
        i += 1
      new Node(bitmap, children, count, valued, minCol, maxCol, vMinCol, vMaxCol, tagMask)

/**
 * One row's cells as a structure of arrays (#891). The row number is not stored: the trie path is
 * the key, so a row shift can re-key a leaf without rebuilding it.
 *
 * Columns are split into 64-wide chunks: chunk `col >> 6` (an arithmetic shift, so negative columns
 * keep their order) and bit `col & 63`. A cell's position in the parallel arrays is its rank: the
 * chunk's base plus the set bits below its own. `meta(p)` is the [[Meta]] word, `values(p)` the
 * value, and `links(p)` the hyperlink when `links` is non-empty; `links` has length 0 unless some
 * cell of the row carries a link, and holds "" for the cells that do not.
 */
// Flat-array loops over the trie (house style, as in SccKernel/LongIntMap): local vars and while
// loops behind pure signatures; nothing mutable escapes.
@SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
private[xl] sealed abstract class RowLeaf(
  val meta: Array[Long],
  val values: Array[CellValue],
  val links: Array[String],
  valued: Int,
  minCol: Int,
  maxCol: Int,
  vMinCol: Int,
  vMaxCol: Int,
  tagMask: Int
) extends Entry(meta.length, valued, minCol, maxCol, vMinCol, vMaxCol, tagMask):
  def chunkCount: Int
  def chunkAt(i: Int): Int
  def maskAt(i: Int): Long
  def baseAt(i: Int): Int

  /** The index of the first chunk numbered `chunk` or later (`chunkCount` when there is none). */
  def chunkFrom(chunk: Int): Int

  /** The array position of column `col`, or -1 when the row has no cell there. */
  def rank(col: Int): Int

  def linkAt(p: Int): String = if links.length == 0 then "" else links(p)

  /** The row's columns, ascending. */
  def cols: Array[Int] =
    val out = new Array[Int](count)
    var p = 0
    var i = 0
    while i < chunkCount do
      val first = chunkAt(i) << 6
      var m = maskAt(i)
      while m != 0L do
        out(p) = first + java.lang.Long.numberOfTrailingZeros(m)
        m &= m - 1
        p += 1
      i += 1
    out

/** A row whose cells all sit in one 64-column chunk. */
private[xl] final class NarrowRow(
  val chunk: Int,
  val mask: Long,
  meta: Array[Long],
  values: Array[CellValue],
  links: Array[String],
  valued: Int,
  minCol: Int,
  maxCol: Int,
  vMinCol: Int,
  vMaxCol: Int,
  tagMask: Int
) extends RowLeaf(meta, values, links, valued, minCol, maxCol, vMinCol, vMaxCol, tagMask):
  def chunkCount: Int = 1
  def chunkAt(i: Int): Int = chunk
  def maskAt(i: Int): Long = mask
  def baseAt(i: Int): Int = 0
  def chunkFrom(c: Int): Int = if chunk >= c then 0 else 1
  def rank(col: Int): Int =
    if (col >> 6) != chunk then -1
    else
      val bit = 1L << (col & 63)
      if (mask & bit) == 0L then -1 else java.lang.Long.bitCount(mask & (bit - 1))

/**
 * A row spanning several chunks. `index` interleaves one pair per chunk, sorted by chunk:
 * `index(2i)` is the chunk's mask and `index(2i+1)` is `(chunk << 32) | base`.
 */
// Flat-array loops over the trie (house style, as in SccKernel/LongIntMap): local vars and while
// loops behind pure signatures; nothing mutable escapes.
@SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
private[xl] final class WideRow(
  val index: Array[Long],
  meta: Array[Long],
  values: Array[CellValue],
  links: Array[String],
  valued: Int,
  minCol: Int,
  maxCol: Int,
  vMinCol: Int,
  vMaxCol: Int,
  tagMask: Int
) extends RowLeaf(meta, values, links, valued, minCol, maxCol, vMinCol, vMaxCol, tagMask):
  def chunkCount: Int = index.length >> 1
  def chunkAt(i: Int): Int = (index(2 * i + 1) >> 32).toInt
  def maskAt(i: Int): Long = index(2 * i)
  def baseAt(i: Int): Int = index(2 * i + 1).toInt

  def chunkFrom(c: Int): Int =
    var lo = 0
    var hi = chunkCount
    while lo < hi do
      val mid = (lo + hi) >>> 1
      if chunkAt(mid) < c then lo = mid + 1 else hi = mid
    lo

  def rank(col: Int): Int =
    val c = col >> 6
    val i = chunkFrom(c)
    if i >= chunkCount || chunkAt(i) != c then -1
    else
      val m = maskAt(i)
      val bit = 1L << (col & 63)
      if (m & bit) == 0L then -1 else baseAt(i) + java.lang.Long.bitCount(m & (bit - 1))

// Flat-array loops over the trie (house style, as in SccKernel/LongIntMap): local vars and while
// loops behind pure signatures; nothing mutable escapes.
@SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
private[xl] object RowLeaf:
  val NoLinks: Array[String] = Array.empty[String]

  val Empty: RowLeaf = new NarrowRow(
    0,
    0L,
    Array.emptyLongArray,
    Array.empty[CellValue],
    NoLinks,
    0,
    Int.MaxValue,
    Int.MinValue,
    Int.MaxValue,
    Int.MinValue,
    0
  )

  /**
   * The leaf holding cells `from until until` of the parallel arrays, whose columns must be
   * strictly ascending. `links` is read only where the meta word has its link bit, so it may be
   * empty when no cell has one. The arrays are copied; the caller keeps ownership.
   */
  def build(
    cols: Array[Int],
    meta: Array[Long],
    values: Array[CellValue],
    links: Array[String],
    from: Int,
    until: Int
  ): RowLeaf =
    val n = until - from
    if n <= 0 then Empty
    else
      val m = java.util.Arrays.copyOfRange(meta, from, until)
      val v = java.util.Arrays.copyOfRange(values, from, until)
      var anyLink = false
      var chunks = 1
      var i = from
      while i < until do
        if Meta.hasLink(meta(i)) then anyLink = true
        if i > from && (cols(i) >> 6) != (cols(i - 1) >> 6) then chunks += 1
        i += 1
      val l =
        if !anyLink then NoLinks
        else
          val out = new Array[String](n)
          var j = 0
          while j < n do
            out(j) = if Meta.hasLink(m(j)) then links(from + j) else ""
            j += 1
          out
      val index = new Array[Long](2 * chunks)
      var c = 0
      var k = from
      while k < until do
        val chunk = cols(k) >> 6
        var mask = 0L
        val base = k - from
        while k < until && (cols(k) >> 6) == chunk do
          mask |= 1L << (cols(k) & 63)
          k += 1
        index(2 * c) = mask
        index(2 * c + 1) = (chunk.toLong << 32) | (base.toLong & 0xffffffffL)
        c += 1
      fromIndex(index, m, v, l)

  /**
   * The leaf with this chunk index and these arrays (owned by the leaf from now on), demoting a
   * one-chunk index to a [[NarrowRow]] and computing the summaries. Empty arrays give [[Empty]].
   */
  def fromIndex(
    index: Array[Long],
    meta: Array[Long],
    values: Array[CellValue],
    links: Array[String]
  ): RowLeaf =
    val n = meta.length
    if n == 0 then Empty
    else
      var valued = 0
      var tagMask = 0
      var firstValued = -1
      var lastValued = -1
      var p = 0
      while p < n do
        val t = Meta.tag(meta(p))
        tagMask |= 1 << t
        if t != Meta.Empty then
          valued += 1
          if firstValued < 0 then firstValued = p
          lastValued = p
        p += 1
      val chunks = index.length >> 1
      val firstChunk = (index(1) >> 32).toInt
      val lastChunk = (index(2 * chunks - 1) >> 32).toInt
      val minCol = (firstChunk << 6) + java.lang.Long.numberOfTrailingZeros(index(0))
      val maxCol =
        (lastChunk << 6) + 63 - java.lang.Long.numberOfLeadingZeros(index(2 * chunks - 2))
      val vMin = if valued == 0 then Int.MaxValue else colAt(index, firstValued)
      val vMax = if valued == 0 then Int.MinValue else colAt(index, lastValued)
      if chunks == 1 then
        new NarrowRow(
          firstChunk,
          index(0),
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
      else new WideRow(index, meta, values, links, valued, minCol, maxCol, vMin, vMax, tagMask)

  /** The interleaved chunk index of a leaf (a fresh array). */
  def indexOf(leaf: RowLeaf): Array[Long] = leaf match
    case w: WideRow => w.index
    case n: NarrowRow => Array(n.mask, n.chunk.toLong << 32)

  /** The column at array position `p` of a leaf with this chunk index. */
  def colAt(index: Array[Long], p: Int): Int =
    var i = 0
    while i + 1 < (index.length >> 1) && index(2 * (i + 1) + 1).toInt <= p do i += 1
    var m = index(2 * i)
    var k = p - index(2 * i + 1).toInt
    while k > 0 do
      m &= m - 1
      k -= 1
    ((index(2 * i + 1) >> 32).toInt << 6) + java.lang.Long.numberOfTrailingZeros(m)

  /** The leaf with the same columns as `leaf` and these arrays in place of its own. */
  def withArrays(
    leaf: RowLeaf,
    meta: Array[Long],
    values: Array[CellValue],
    links: Array[String]
  ): RowLeaf =
    val l = if hasAnyLink(meta) then links else NoLinks
    fromIndex(indexOf(leaf), meta, values, l)

  private def hasAnyLink(meta: Array[Long]): Boolean =
    var i = 0
    var found = false
    while !found && i < meta.length do
      found = Meta.hasLink(meta(i))
      i += 1
    found

  /** The leaf's links as one entry per cell (a fresh array of "" when it has none). */
  def linksOf(leaf: RowLeaf): Array[String] =
    if leaf.links.length == 0 then Array.fill(leaf.count)("") else leaf.links.clone()

  /** `leaf` with the cell at `col` set to (`meta`, `value`, `link`), added if absent. */
  def set(leaf: RowLeaf, col: Int, meta: Long, value: CellValue, link: String): RowLeaf =
    val p = leaf.rank(col)
    if p >= 0 then
      val m = leaf.meta.clone()
      val v = leaf.values.clone()
      m(p) = meta
      v(p) = value
      val l =
        if Meta.hasLink(meta) then
          val out = linksOf(leaf)
          out(p) = link
          out
        else if leaf.links.length == 0 then NoLinks
        else
          val out = leaf.links.clone()
          out(p) = ""
          out
      withArrays(leaf, m, v, l)
    else insert(leaf, col, meta, value, link)

  /** `leaf` with `value` at `col`: an existing cell keeps its style and link. */
  def setValue(leaf: RowLeaf, col: Int, value: CellValue): RowLeaf =
    val p = leaf.rank(col)
    if p >= 0 then
      val m = leaf.meta.clone()
      val v = leaf.values.clone()
      m(p) = Meta.retag(m(p), value)
      v(p) = value
      fromIndex(indexOf(leaf), m, v, leaf.links)
    else insert(leaf, col, Meta.of(value, None, false), value, "")

  /**
   * `leaf` with a new cell at `col` (absent from it). Rank arithmetic finds the position, each
   * array is allocated once, the chunk index is patched (the chunk's mask gains the bit, later
   * bases move up by one) and the summaries are extended incrementally: an insert can only grow
   * them.
   */
  private def insert(leaf: RowLeaf, col: Int, meta: Long, value: CellValue, link: String): RowLeaf =
    if leaf.count == 0 then build(Array(col), Array(meta), Array(value), Array(link), 0, 1)
    else
      val n = leaf.count
      val c = col >> 6
      val bit = 1L << (col & 63)
      val chunks = leaf.chunkCount
      val i = leaf.chunkFrom(c)
      val existing = i < chunks && leaf.chunkAt(i) == c
      val p =
        if existing then leaf.baseAt(i) + java.lang.Long.bitCount(leaf.maskAt(i) & (bit - 1))
        else if i < chunks then leaf.baseAt(i)
        else n
      val newChunks = if existing then chunks else chunks + 1
      val index = new Array[Long](2 * newChunks)
      var j = 0
      var k = 0
      while j < newChunks do
        if !existing && j == i then
          index(2 * j) = bit
          index(2 * j + 1) = (c.toLong << 32) | p.toLong
        else
          val mask = if existing && k == i then leaf.maskAt(k) | bit else leaf.maskAt(k)
          val base = leaf.baseAt(k) + (if k > i || (!existing && k >= i) then 1 else 0)
          index(2 * j) = mask
          index(2 * j + 1) = (leaf.chunkAt(k).toLong << 32) | base.toLong
          k += 1
        j += 1
      val m = new Array[Long](n + 1)
      val v = new Array[CellValue](n + 1)
      System.arraycopy(leaf.meta, 0, m, 0, p)
      System.arraycopy(leaf.values, 0, v, 0, p)
      m(p) = meta
      v(p) = value
      System.arraycopy(leaf.meta, p, m, p + 1, n - p)
      System.arraycopy(leaf.values, p, v, p + 1, n - p)
      val l =
        if leaf.links.length == 0 && !Meta.hasLink(meta) then NoLinks
        else
          val out = new Array[String](n + 1)
          if leaf.links.length == 0 then
            var q = 0
            while q <= n do
              out(q) = ""
              q += 1
          else
            System.arraycopy(leaf.links, 0, out, 0, p)
            System.arraycopy(leaf.links, p, out, p + 1, n - p)
          out(p) = if Meta.hasLink(meta) then link else ""
          out
      val t = Meta.tag(meta)
      val isValued = t != Meta.Empty
      make(
        index,
        m,
        v,
        l,
        leaf.valued + (if isValued then 1 else 0),
        math.min(leaf.minCol, col),
        math.max(leaf.maxCol, col),
        if isValued then math.min(leaf.vMinCol, col) else leaf.vMinCol,
        if isValued then math.max(leaf.vMaxCol, col) else leaf.vMaxCol,
        leaf.tagMask | (1 << t)
      )

  /** The leaf with this index, arrays and (already computed) summaries. */
  private def make(
    index: Array[Long],
    meta: Array[Long],
    values: Array[CellValue],
    links: Array[String],
    valued: Int,
    minCol: Int,
    maxCol: Int,
    vMinCol: Int,
    vMaxCol: Int,
    tagMask: Int
  ): RowLeaf =
    if index.length == 2 then
      new NarrowRow(
        (index(1) >> 32).toInt,
        index(0),
        meta,
        values,
        links,
        valued,
        minCol,
        maxCol,
        vMinCol,
        vMaxCol,
        tagMask
      )
    else new WideRow(index, meta, values, links, valued, minCol, maxCol, vMinCol, vMaxCol, tagMask)

  /** `leaf` without its cell at `col` ([[Empty]] when that was the last one). */
  def remove(leaf: RowLeaf, col: Int): RowLeaf =
    if leaf.rank(col) < 0 then leaf else retain(leaf, c => c != col)

  /** `leaf` keeping only the cells whose column passes `keep`; `leaf` itself when all pass. */
  def retain(leaf: RowLeaf, keep: Int => Boolean): RowLeaf =
    val cols = leaf.cols
    val n = cols.length
    val outCols = new Array[Int](n)
    val m = new Array[Long](n)
    val v = new Array[CellValue](n)
    val l = new Array[String](n)
    var k = 0
    var i = 0
    while i < n do
      if keep(cols(i)) then
        outCols(k) = cols(i)
        m(k) = leaf.meta(i)
        v(k) = leaf.values(i)
        l(k) = leaf.linkAt(i)
        k += 1
      i += 1
    if k == n then leaf else build(outCols, m, v, l, 0, k)

  /**
   * `leaf` with its columns moved by the shift rule of [[com.tjclp.xl.sheets.CellStore.shiftCols]]:
   * a column at or after `at` moves by `by`, a column in the vacated window `[at + by, at)` (when
   * `by < 0`) is dropped, and a column whose destination leaves the `Int` range is dropped.
   */
  def shift(leaf: RowLeaf, at: Int, by: Int): RowLeaf =
    val cols = leaf.cols
    val n = cols.length
    val outCols = new Array[Int](n)
    val m = new Array[Long](n)
    val v = new Array[CellValue](n)
    val l = new Array[String](n)
    var k = 0
    var i = 0
    while i < n do
      val dest = Shift.dest(cols(i), at, by)
      if dest != Shift.Dropped then
        outCols(k) = dest.toInt
        m(k) = leaf.meta(i)
        v(k) = leaf.values(i)
        l(k) = leaf.linkAt(i)
        k += 1
      i += 1
    build(outCols, m, v, l, 0, k)

  /**
   * `leaf` with `f` applied to every value whose tag is in `tagMask`; `leaf` itself when `f`
   * returns every such value unchanged (by reference).
   */
  def mapValues(leaf: RowLeaf, tagMask: Int, f: CellValue => CellValue): RowLeaf =
    if (leaf.tagMask & tagMask) == 0 then leaf
    else
      var m: Array[Long] = leaf.meta
      var v: Array[CellValue] = leaf.values
      var changed = false
      var p = 0
      while p < leaf.count do
        if (tagMask & (1 << Meta.tag(leaf.meta(p)))) != 0 then
          val old = leaf.values(p)
          val nv = f(old)
          if !(nv eq old) then
            if !changed then
              changed = true
              m = leaf.meta.clone()
              v = leaf.values.clone()
            m(p) = Meta.retag(m(p), nv)
            v(p) = nv
        p += 1
      if changed then fromIndex(indexOf(leaf), m, v, leaf.links) else leaf

  /**
   * The union of two rows, `b` winning where both hold a column. With `valuesOnly`, a column held
   * by both takes `b`'s value but keeps `a`'s style and link.
   */
  def merge(a: RowLeaf, b: RowLeaf, valuesOnly: Boolean): RowLeaf =
    val ca = a.cols
    val cb = b.cols
    val n = ca.length + cb.length
    val cols = new Array[Int](n)
    val m = new Array[Long](n)
    val v = new Array[CellValue](n)
    val l = new Array[String](n)
    var i = 0
    var j = 0
    var k = 0
    while i < ca.length || j < cb.length do
      if j >= cb.length || (i < ca.length && ca(i) < cb(j)) then
        cols(k) = ca(i); m(k) = a.meta(i); v(k) = a.values(i); l(k) = a.linkAt(i)
        i += 1
      else if i >= ca.length || cb(j) < ca(i) then
        cols(k) = cb(j); m(k) = b.meta(j); v(k) = b.values(j); l(k) = b.linkAt(j)
        j += 1
      else
        cols(k) = cb(j)
        if valuesOnly then
          m(k) = Meta.retag(a.meta(i), b.values(j)); v(k) = b.values(j); l(k) = a.linkAt(i)
        else
          m(k) = b.meta(j); v(k) = b.values(j); l(k) = b.linkAt(j)
        i += 1
        j += 1
      k += 1
    build(cols, m, v, l, 0, k)

  /** Structural equality of two leaves. */
  def same(a: RowLeaf, b: RowLeaf): Boolean =
    if a eq b then true
    else if a.count != b.count || a.chunkCount != b.chunkCount then false
    else
      var ok = true
      var i = 0
      while ok && i < a.chunkCount do
        ok = a.chunkAt(i) == b.chunkAt(i) && a.maskAt(i) == b.maskAt(i)
        i += 1
      ok = ok && java.util.Arrays.equals(a.meta, b.meta) && a.links.length == b.links.length
      var p = 0
      while ok && p < a.count do
        ok = a.values(p) == b.values(p) && a.linkAt(p) == b.linkAt(p)
        p += 1
      ok

/**
 * The index map of a row or column shift (#891): an index at or after `at` moves by `by`; when
 * `by < 0` the indices in `[at + by, at)` are vacated and dropped; an index whose destination falls
 * outside the `Int` range is dropped too, so the map is injective and no two cells collide.
 */
private[xl] object Shift:
  /** Marks a dropped index (outside the `Int` range, so no destination equals it). */
  final val Dropped: Long = Long.MinValue

  def dest(i: Int, at: Int, by: Int): Long =
    if i < at then if by < 0 && i.toLong >= at.toLong + by then Dropped else i.toLong
    else
      val d = i.toLong + by
      if d < Int.MinValue || d > Int.MaxValue then Dropped else d
