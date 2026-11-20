package com.tjclp.xl.sheets.store

import com.tjclp.xl.cells.CellValue

/**
 * One row's cells as the trie holds them (#891): a [[BandedRow]] of one leaf per populated band
 * when the row's columns span more than one column band (see [[Band]]) and it holds more than
 * [[Band.FlatMax]] cells; otherwise a single [[RowLeaf]]. The form is a function of the row's
 * columns alone, so equal rows have equal shapes whatever their history.
 */
private[xl] type StoredRow = RowLeaf | BandedRow

/**
 * The row operations of the store, over either form. Every result is canonical: a leaf that comes
 * to hold more than [[Band.FlatMax]] cells over two or more bands becomes banded, and a banded row
 * left with one band, or with [[Band.FlatMax]] cells or fewer, becomes a leaf again. A single-cell
 * write or removal copies one band (at most [[Band.Width]] cells) and, in a banded row, the array
 * of band references; in a leaf spanning bands it copies at most [[Band.FlatMax]] cells. Converting
 * between the forms costs at most a band and a leaf of that size.
 */
// Flat-array loops over the trie (house style, as in SccKernel/LongIntMap): local vars and while
// loops behind pure signatures; nothing mutable escapes.
@SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
private[xl] object StoredRow:

  /**
   * The row holding cells `from until until` of the parallel arrays, whose columns must be strictly
   * ascending; see [[RowLeaf.build]] for `links`. The arrays are copied.
   */
  def build(
    cols: Array[Int],
    meta: Array[Long],
    values: Array[CellValue],
    links: Array[String],
    from: Int,
    until: Int
  ): StoredRow =
    if until <= from then RowLeaf.Empty
    else if until - from <= Band.FlatMax || Band.of(cols(from)) == Band.of(cols(until - 1)) then
      RowLeaf.build(cols, meta, values, links, from, until)
    else split(cols, meta, values, links, from, until)

  /**
   * One leaf per band of cells `from until until`, which must span two bands or more. The result is
   * canonical only when it holds more than [[Band.FlatMax]] cells; callers check.
   */
  private def split(
    cols: Array[Int],
    meta: Array[Long],
    values: Array[CellValue],
    links: Array[String],
    from: Int,
    until: Int
  ): BandedRow =
    var n = 1
    var i = from + 1
    while i < until do
      if Band.of(cols(i)) != Band.of(cols(i - 1)) then n += 1
      i += 1
    val ids = new Array[Int](n)
    val bands = new Array[RowLeaf](n)
    var b = 0
    var k = from
    while k < until do
      val id = Band.of(cols(k))
      var end = k + 1
      while end < until && Band.of(cols(end)) == id do end += 1
      ids(b) = id
      bands(b) = RowLeaf.build(cols, meta, values, links, k, end)
      b += 1
      k = end
    BandedRow.of(ids, bands)

  /** Whether a leaf's columns lie in one band (an empty leaf's do). */
  private def oneBand(leaf: RowLeaf): Boolean =
    leaf.count == 0 || Band.of(leaf.minCol) == Band.of(leaf.maxCol)

  /** The bands of a leaf as a banded row would hold them: (ids, leaves). */
  private def bandsOf(leaf: RowLeaf): (Array[Int], Array[RowLeaf]) =
    if oneBand(leaf) then (Array(Band.of(leaf.minCol)), Array(leaf))
    else
      val row = split(leaf.cols, leaf.meta, leaf.values, leaf.links, 0, leaf.count)
      (row.ids, row.bands)

  /** The canonical row of bands `ids(i) -> bands(i)` for `i < n` (ids ascending, none empty). */
  private def of(ids: Array[Int], bands: Array[RowLeaf], n: Int): StoredRow =
    if n == 0 then RowLeaf.Empty
    else if n == 1 then bands(0)
    else
      var count = 0
      var i = 0
      while i < n do
        count += bands(i).count
        i += 1
      if count <= Band.FlatMax then flatten(bands, n, count)
      else if n == ids.length then BandedRow.of(ids, bands)
      else BandedRow.of(java.util.Arrays.copyOf(ids, n), java.util.Arrays.copyOf(bands, n))

  /** The canonical row of a leaf's cells: banded when they are many and span two bands. */
  private def ofLeaf(leaf: RowLeaf): StoredRow =
    if leaf.count <= Band.FlatMax || oneBand(leaf) then leaf
    else split(leaf.cols, leaf.meta, leaf.values, leaf.links, 0, leaf.count)

  /** The canonical row of a banded row's cells: one leaf when they have become few. */
  private def ofBanded(row: BandedRow): StoredRow =
    if row.count <= Band.FlatMax then flatten(row.bands, row.bands.length, row.count) else row

  /** The cells of the first `n` of `bands` (`count` in all), in band order, as one leaf. */
  private def flatten(bands: Array[RowLeaf], n: Int, count: Int): RowLeaf =
    val cols = new Array[Int](count)
    val m = new Array[Long](count)
    val v = new Array[CellValue](count)
    val l = new Array[String](count)
    var k = 0
    var j = 0
    while j < n do
      val band = bands(j)
      val bandCols = band.cols
      System.arraycopy(bandCols, 0, cols, k, band.count)
      System.arraycopy(band.meta, 0, m, k, band.count)
      System.arraycopy(band.values, 0, v, k, band.count)
      var i = 0
      while i < band.count do
        l(k + i) = band.linkAt(i)
        i += 1
      k += band.count
      j += 1
    RowLeaf.build(cols, m, v, l, 0, count)

  /**
   * `row` with `f` applied to the leaf holding `col`'s cells: the row's leaf, or in a banded row
   * the leaf of `col`'s band (`f` sees [[RowLeaf.Empty]] when the row has no cell in that band, and
   * may return an empty leaf to drop the band). `row` itself when `f` returns its argument.
   */
  private inline def modifyBand(row: StoredRow, col: Int)(inline f: RowLeaf => RowLeaf): StoredRow =
    row match
      case leaf: RowLeaf => ofLeaf(f(leaf))
      case b: BandedRow =>
        val id = Band.of(col)
        val i = b.indexFrom(id)
        if i < b.ids.length && b.ids(i) == id then
          val old = b.bands(i)
          val next = f(old)
          if next eq old then b
          else if next.count == 0 then withoutBand(b, i)
          else ofBanded(BandedRow.replaced(b, i, next))
        else
          val made = f(RowLeaf.Empty)
          if made.count == 0 then b else BandedRow.inserted(b, i, id, made)

  /** `row` without band `i`: one leaf when one band, or few cells, are left. */
  private def withoutBand(row: BandedRow, i: Int): StoredRow =
    val n = row.ids.length
    if n == 2 then row.bands(1 - i)
    else
      val ids = new Array[Int](n - 1)
      val bands = new Array[RowLeaf](n - 1)
      System.arraycopy(row.ids, 0, ids, 0, i)
      System.arraycopy(row.bands, 0, bands, 0, i)
      System.arraycopy(row.ids, i + 1, ids, i, n - 1 - i)
      System.arraycopy(row.bands, i + 1, bands, i, n - 1 - i)
      of(ids, bands, n - 1)

  /** `row` with the cell at `col` set to (`meta`, `value`, `link`), added if absent. */
  def set(row: StoredRow, col: Int, meta: Long, value: CellValue, link: String): StoredRow =
    modifyBand(row, col)(leaf => RowLeaf.set(leaf, col, meta, value, link))

  /** `row` with `value` at `col`: an existing cell keeps its style and link. */
  def setValue(row: StoredRow, col: Int, value: CellValue): StoredRow =
    modifyBand(row, col)(leaf => RowLeaf.setValue(leaf, col, value))

  /** `row` without its cell at `col`; `row` itself when it has none there. */
  def remove(row: StoredRow, col: Int): StoredRow = row match
    case leaf: RowLeaf => RowLeaf.remove(leaf, col)
    case _: BandedRow => modifyBand(row, col)(leaf => RowLeaf.removeCell(leaf, col))

  /** `row` keeping only the cells whose column passes `keep`; `row` itself when all pass. */
  def retain(row: StoredRow, keep: Int => Boolean): StoredRow = row match
    case leaf: RowLeaf => RowLeaf.retain(leaf, keep)
    case b: BandedRow =>
      val n = b.bands.length
      val ids = new Array[Int](n)
      val bands = new Array[RowLeaf](n)
      var changed = false
      var k = 0
      var i = 0
      while i < n do
        val next = RowLeaf.retain(b.bands(i), keep)
        if !(next eq b.bands(i)) then changed = true
        if next.count > 0 then
          ids(k) = b.ids(i)
          bands(k) = next
          k += 1
        i += 1
      if changed then of(ids, bands, k) else b

  /** `row` with its columns moved by the shift rule of [[RowLeaf.shift]], re-banded. */
  def shift(row: StoredRow, at: Int, by: Int): StoredRow = row match
    case leaf: RowLeaf => ofLeaf(RowLeaf.shift(leaf, at, by))
    case b: BandedRow =>
      val n = b.count
      val outCols = new Array[Int](n)
      val m = new Array[Long](n)
      val v = new Array[CellValue](n)
      val l = new Array[String](n)
      var k = 0
      var j = 0
      while j < b.bands.length do
        val band = b.bands(j)
        val cols = band.cols
        var i = 0
        while i < cols.length do
          val dest = Shift.dest(cols(i), at, by)
          if dest != Shift.Dropped then
            outCols(k) = dest.toInt
            m(k) = band.meta(i)
            v(k) = band.values(i)
            l(k) = band.linkAt(i)
            k += 1
          i += 1
        j += 1
      build(outCols, m, v, l, 0, k)

  /**
   * `row` with `f` applied to every value whose tag is in `tagMask`; `row` itself when `f` returns
   * every such value unchanged (by reference). The columns, and so the bands, stay as they are.
   */
  def mapValues(row: StoredRow, tagMask: Int, f: CellValue => CellValue): StoredRow = row match
    case leaf: RowLeaf => RowLeaf.mapValues(leaf, tagMask, f)
    case b: BandedRow =>
      if (b.tagMask & tagMask) == 0 then b
      else
        var bands = b.bands
        var i = 0
        while i < b.bands.length do
          val next = RowLeaf.mapValues(b.bands(i), tagMask, f)
          if !(next eq b.bands(i)) then
            if bands eq b.bands then bands = b.bands.clone()
            bands(i) = next
          i += 1
        if bands eq b.bands then b else BandedRow.of(b.ids, bands)

  /**
   * The union of two rows, `b` winning where both hold a column (see [[RowLeaf.merge]] for
   * `valuesOnly`). Bands only one row holds are shared with it.
   */
  def merge(a: StoredRow, b: StoredRow, valuesOnly: Boolean): StoredRow =
    if a.count == 0 then b
    else if b.count == 0 then a
    else
      a match
        case la: RowLeaf =>
          b match
            case lb: RowLeaf => ofLeaf(RowLeaf.merge(la, lb, valuesOnly))
            case bb: BandedRow =>
              val (ids, bands) = bandsOf(la)
              mergeBands(ids, bands, bb.ids, bb.bands, valuesOnly)
        case ba: BandedRow =>
          b match
            case lb: RowLeaf =>
              val (ids, bands) = bandsOf(lb)
              mergeBands(ba.ids, ba.bands, ids, bands, valuesOnly)
            case bb: BandedRow => mergeBands(ba.ids, ba.bands, bb.ids, bb.bands, valuesOnly)

  private def mergeBands(
    ia: Array[Int],
    la: Array[RowLeaf],
    ib: Array[Int],
    lb: Array[RowLeaf],
    valuesOnly: Boolean
  ): StoredRow =
    val ids = new Array[Int](ia.length + ib.length)
    val bands = new Array[RowLeaf](ia.length + ib.length)
    var i = 0
    var j = 0
    var k = 0
    while i < ia.length || j < ib.length do
      if j >= ib.length || (i < ia.length && ia(i) < ib(j)) then
        ids(k) = ia(i); bands(k) = la(i); i += 1
      else if i >= ia.length || ib(j) < ia(i) then
        ids(k) = ib(j); bands(k) = lb(j); j += 1
      else
        ids(k) = ia(i); bands(k) = RowLeaf.merge(la(i), lb(j), valuesOnly); i += 1; j += 1
      k += 1
    of(ids, bands, k)

  /** Structural equality of two rows. */
  def same(a: StoredRow, b: StoredRow): Boolean = a match
    case la: RowLeaf =>
      b match
        case lb: RowLeaf => RowLeaf.same(la, lb)
        case _: BandedRow => false
    case ba: BandedRow =>
      b match
        case bb: BandedRow => BandedRow.same(ba, bb)
        case _: RowLeaf => false
