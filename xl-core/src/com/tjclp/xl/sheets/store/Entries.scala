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
 * A trie entry: an inner [[Node]], a [[Block]], or a row (a [[RowLeaf]] or a [[BandedRow]] of
 * them). Every entry carries the summaries of the cells below it. They are invariants, set when the
 * entry is built and never changed:
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
 * holds rows ([[StoredRow]]s); a node at shift 5 holds shift-0 groups ([[Node]]s or [[Block]]s);
 * every other node holds nodes. No child is ever empty. A shift-0 node exists only when its rows
 * cannot form a [[Block]] (they span more than one 64-column chunk between them).
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
 * An entry that holds cells: a [[RowLeaf]] (one row) or a [[Block]] (the rows of one 32-row slot
 * group). `meta(p)` is the [[Meta]] word, `values(p)` the value and `links(p)` the hyperlink when
 * `links` is non-empty; `links` has length 0 unless some cell carries a link, and holds "" for the
 * cells that do not.
 */
private[xl] sealed abstract class Cells(
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
  /** The array position of the cell at (`key`, `col`), or -1. A leaf ignores `key`. */
  def find(key: Long, col: Int): Int

  def linkAt(p: Int): String = if links.length == 0 then "" else links(p)

/**
 * The rows of one 32-row slot group whose cells all sit in one 64-column chunk, at most
 * [[Group.MaxBlockCells]] of them, stored as one set of parallel arrays (#891). It takes the place
 * of a shift-0 [[Node]] and its [[RowLeaf]]s, so a tall narrow sheet pays one object and three
 * arrays per 32 rows instead of per row.
 *
 * `rowMask` bit `r` says row `prefix + r` is present; present rows are numbered `i` in ascending
 * order. Row `i`'s columns are the set bits of its chunk mask over `chunk`; its cells sit at
 * `base(i) ..` in the arrays, ordered by column. When every present row has the same mask (a dense
 * table or column) `masks` holds that one mask and `bases` is empty, and `base(i)` is `i` times the
 * row width; otherwise `masks(i)` and `bases(i)` are stored per row. The choice is canonical: a
 * block is uniform exactly when its rows' masks are all equal.
 */
private[xl] final class Block(
  val rowMask: Int,
  val chunk: Int,
  val masks: Array[Long],
  val bases: Array[Int],
  meta: Array[Long],
  values: Array[CellValue],
  links: Array[String],
  valued: Int,
  minCol: Int,
  maxCol: Int,
  vMinCol: Int,
  vMaxCol: Int,
  tagMask: Int
) extends Cells(meta, values, links, valued, minCol, maxCol, vMinCol, vMaxCol, tagMask):
  def rows: Int = Integer.bitCount(rowMask)
  def uniform: Boolean = bases.length == 0

  /** The index of present row `low` (0..31); meaningful only when the row is present. */
  def rowIndex(low: Int): Int = Integer.bitCount(rowMask & ((1 << low) - 1))

  def rowMaskAt(i: Int): Long = if bases.length == 0 then masks(0) else masks(i)

  def rowBaseAt(i: Int): Int =
    if bases.length == 0 then i * java.lang.Long.bitCount(masks(0)) else bases(i)

  def find(key: Long, col: Int): Int =
    val bit = 1 << (key & 31L).toInt
    if (rowMask & bit) == 0 || (col >> 6) != chunk then -1
    else
      val i = Integer.bitCount(rowMask & (bit - 1))
      val m = rowMaskAt(i)
      val cb = 1L << (col & 63)
      if (m & cb) == 0L then -1 else rowBaseAt(i) + java.lang.Long.bitCount(m & (cb - 1))

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
  meta: Array[Long],
  values: Array[CellValue],
  links: Array[String],
  valued: Int,
  minCol: Int,
  maxCol: Int,
  vMinCol: Int,
  vMaxCol: Int,
  tagMask: Int
) extends Cells(meta, values, links, valued, minCol, maxCol, vMinCol, vMaxCol, tagMask):
  def find(key: Long, col: Int): Int = rank(col)

  def chunkCount: Int
  def chunkAt(i: Int): Int
  def maskAt(i: Int): Long
  def baseAt(i: Int): Int

  /** The index of the first chunk numbered `chunk` or later (`chunkCount` when there is none). */
  def chunkFrom(chunk: Int): Int

  /** The array position of column `col`, or -1 when the row has no cell there. */
  def rank(col: Int): Int

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
 * `index(2i)` is the chunk's mask and `index(2i+1)` is `(chunk << 32) | base`. A dense row holds
 * consecutive chunks, so [[rank]] first looks for a chunk that many places after the first one and
 * bisects the index only when it is not there.
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
    val guess = c - chunkAt(0)
    val i = if guess >= 0 && guess < chunkCount && chunkAt(guess) == c then guess else chunkFrom(c)
    if i >= chunkCount || chunkAt(i) != c then -1
    else
      val m = maskAt(i)
      val bit = 1L << (col & 63)
      if (m & bit) == 0L then -1 else baseAt(i) + java.lang.Long.bitCount(m & (bit - 1))

/**
 * The column bands of a row (#891). Band `col >> Shift` (an arithmetic shift, so negative columns
 * keep their order) holds the row's cells in `Width` consecutive columns: [[Band.Chunks]] of a
 * leaf's 64-column chunks. A write to a row copies one band, so its cost is bounded by `Width`
 * cells however wide the row is.
 */
private[xl] object Band:
  final val Shift = 8
  final val Width = 1 << Shift
  final val Chunks = Width >> 6

  /** The band of a column. */
  def of(col: Int): Int = col >> Shift

/**
 * A row whose cells span several column bands (#891): one [[RowLeaf]] per populated band, sorted by
 * band id. `ids(i)` is the band of `bands(i)`, whose cells all lie in that band; there are at least
 * two bands and none is empty. A row whose cells share one band is a bare [[RowLeaf]], never a
 * one-band `BandedRow`, so a row's form is a function of its columns alone.
 *
 * A single-cell write or removal copies the band it touches and the `bands` array (`ids` is shared
 * unless a band is added or dropped): a bounded number of cells however wide the row is. Band ids
 * are arbitrary `Int`s (columns off the grid are legal), so the index is a sorted array searched by
 * bisection rather than a bitmap.
 */
private[xl] final class BandedRow private (
  val ids: Array[Int],
  val bands: Array[RowLeaf],
  count: Int,
  valued: Int,
  minCol: Int,
  maxCol: Int,
  vMinCol: Int,
  vMaxCol: Int,
  tagMask: Int
) extends Entry(count, valued, minCol, maxCol, vMinCol, vMaxCol, tagMask):

  /**
   * The index of the first band numbered `id` or later (`ids.length` when there is none). A dense
   * row holds consecutive bands, so band `id` is first sought `id - ids(0)` places in; the sorted
   * index is bisected only when it is not there.
   */
  def indexFrom(id: Int): Int =
    val guess = id - ids(0)
    if guess >= 0 && guess < ids.length && ids(guess) == id then guess
    else BandedRow.indexFrom(ids, id)

  /**
   * The band holding column `col`'s cells; [[RowLeaf.Empty]] when the row has none there. Tries the
   * consecutive-bands position first, as [[indexFrom]] does, checking the band's own columns so a
   * hit reads nothing but the leaf it returns.
   */
  def bandOf(col: Int): RowLeaf =
    val id = col >> Band.Shift
    val guess = id - (minCol >> Band.Shift)
    val hit = if guess >= 0 && guess < bands.length then bands(guess) else RowLeaf.Empty
    if hit.count > 0 && (hit.minCol >> Band.Shift) == id then hit
    else
      val i = BandedRow.indexFrom(ids, id)
      if i < ids.length && ids(i) == id then bands(i) else RowLeaf.Empty

// Flat-array loops over the trie (house style, as in SccKernel/LongIntMap): local vars and while
// loops behind pure signatures; nothing mutable escapes.
@SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
private[xl] object BandedRow:

  /** The index of the first of the sorted `ids` that is `id` or later (`ids.length` if none). */
  def indexFrom(ids: Array[Int], id: Int): Int =
    var lo = 0
    var hi = ids.length
    while lo < hi do
      val mid = (lo + hi) >>> 1
      if ids(mid) < id then lo = mid + 1 else hi = mid
    lo

  /**
   * The row of these bands (two or more, ids strictly ascending, none empty; the arrays are owned
   * by the row from now on), its summaries computed from the bands'.
   */
  def of(ids: Array[Int], bands: Array[RowLeaf]): BandedRow =
    var count = 0
    var valued = 0
    var i = 0
    while i < bands.length do
      count += bands(i).count
      valued += bands(i).valued
      i += 1
    new BandedRow(
      ids,
      bands,
      count,
      valued,
      bands(0).minCol,
      bands(bands.length - 1).maxCol,
      vMinOf(bands),
      vMaxOf(bands),
      tagsOf(bands)
    )

  /** The least valued column over the bands; `Int.MaxValue` when none holds a value. */
  private def vMinOf(bands: Array[RowLeaf]): Int =
    var v = Int.MaxValue
    var i = 0
    while i < bands.length do
      if bands(i).vMinCol < v then v = bands(i).vMinCol
      i += 1
    v

  /** The greatest valued column over the bands; `Int.MinValue` when none holds a value. */
  private def vMaxOf(bands: Array[RowLeaf]): Int =
    var v = Int.MinValue
    var i = 0
    while i < bands.length do
      if bands(i).vMaxCol > v then v = bands(i).vMaxCol
      i += 1
    v

  private def tagsOf(bands: Array[RowLeaf]): Int =
    var t = 0
    var i = 0
    while i < bands.length do
      t |= bands(i).tagMask
      i += 1
    t

  /**
   * `row` with band `i` replaced by `band` (non-empty, the same band id). The summaries are derived
   * from the row's and the two bands' in constant time; the other bands are rescanned only when the
   * old band may have held the row's only instance of a valued extreme or of a tag (as
   * [[Group.withRow]] does for a block's rows).
   */
  def replaced(row: BandedRow, i: Int, band: RowLeaf): BandedRow =
    val old = row.bands(i)
    val bands = row.bands.clone()
    bands(i) = band
    val valued = row.valued - old.valued + band.valued
    val tagMask =
      if (old.tagMask & ~band.tagMask) == 0 then row.tagMask | band.tagMask else tagsOf(bands)
    val minOk = old.valued == 0 || old.vMinCol > row.vMinCol ||
      (band.valued > 0 && band.vMinCol <= old.vMinCol)
    val maxOk = old.valued == 0 || old.vMaxCol < row.vMaxCol ||
      (band.valued > 0 && band.vMaxCol >= old.vMaxCol)
    val vMin =
      if valued == 0 then Int.MaxValue
      else if minOk then math.min(row.vMinCol, band.vMinCol)
      else vMinOf(bands)
    val vMax =
      if valued == 0 then Int.MinValue
      else if maxOk then math.max(row.vMaxCol, band.vMaxCol)
      else vMaxOf(bands)
    new BandedRow(
      row.ids,
      bands,
      row.count - old.count + band.count,
      valued,
      if i == 0 then band.minCol else row.minCol,
      if i == bands.length - 1 then band.maxCol else row.maxCol,
      vMin,
      vMax,
      tagMask
    )

  /**
   * `row` with `band` (non-empty) added as band `id` at index `i`, where the row has none. A new
   * band can only extend the summaries.
   */
  def inserted(row: BandedRow, i: Int, id: Int, band: RowLeaf): BandedRow =
    val n = row.ids.length
    val ids = new Array[Int](n + 1)
    val bands = new Array[RowLeaf](n + 1)
    System.arraycopy(row.ids, 0, ids, 0, i)
    System.arraycopy(row.bands, 0, bands, 0, i)
    ids(i) = id
    bands(i) = band
    System.arraycopy(row.ids, i, ids, i + 1, n - i)
    System.arraycopy(row.bands, i, bands, i + 1, n - i)
    new BandedRow(
      ids,
      bands,
      row.count + band.count,
      row.valued + band.valued,
      math.min(row.minCol, band.minCol),
      math.max(row.maxCol, band.maxCol),
      math.min(row.vMinCol, band.vMinCol),
      math.max(row.vMaxCol, band.vMaxCol),
      row.tagMask | band.tagMask
    )

  /** Structural equality of two banded rows. */
  def same(a: BandedRow, b: BandedRow): Boolean =
    if a eq b then true
    else if a.count != b.count || !java.util.Arrays.equals(a.ids, b.ids) then false
    else
      var ok = true
      var i = 0
      while ok && i < a.bands.length do
        ok = RowLeaf.same(a.bands(i), b.bands(i))
        i += 1
      ok

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
    if leaf.count == 0 then single(col, meta, value, link)
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

  /** The leaf of one cell. */
  def single(col: Int, meta: Long, value: CellValue, link: String): RowLeaf =
    build(Array(col), Array(meta), Array(value), Array(link), 0, 1)

  /**
   * `leaf` without its cell at `col`, as [[remove]] gives it, by splicing: each array is allocated
   * once at its final size and the chunk index is patched (the chunk's mask loses the bit, a chunk
   * left empty is dropped, later bases move down by one). The bands of a [[BandedRow]] are removed
   * from this way, so a removal there costs one copy of the band.
   */
  def removeCell(leaf: RowLeaf, col: Int): RowLeaf =
    val p = leaf.rank(col)
    if p < 0 then leaf
    else if leaf.count == 1 then Empty
    else
      val n = leaf.count
      val bit = 1L << (col & 63)
      val chunks = leaf.chunkCount
      val i = leaf.chunkFrom(col >> 6)
      val emptied = leaf.maskAt(i) == bit
      val index = new Array[Long](2 * (if emptied then chunks - 1 else chunks))
      var j = 0
      var k = 0
      while k < chunks do
        if !(emptied && k == i) then
          val mask = if k == i then leaf.maskAt(k) & ~bit else leaf.maskAt(k)
          val base = if k > i then leaf.baseAt(k) - 1 else leaf.baseAt(k)
          index(2 * j) = mask
          index(2 * j + 1) = (leaf.chunkAt(k).toLong << 32) | base.toLong
          j += 1
        k += 1
      val m = new Array[Long](n - 1)
      val v = new Array[CellValue](n - 1)
      System.arraycopy(leaf.meta, 0, m, 0, p)
      System.arraycopy(leaf.values, 0, v, 0, p)
      System.arraycopy(leaf.meta, p + 1, m, p, n - 1 - p)
      System.arraycopy(leaf.values, p + 1, v, p, n - 1 - p)
      // links stay canonical: dropped when the removed cell held the leaf's last one
      val l =
        if leaf.links.length == 0 || (Meta.hasLink(leaf.meta(p)) && !hasAnyLink(m)) then NoLinks
        else
          val out = new Array[String](n - 1)
          System.arraycopy(leaf.links, 0, out, 0, p)
          System.arraycopy(leaf.links, p + 1, out, p, n - 1 - p)
          out
      fromIndex(index, m, v, l)

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
