package com.tjclp.xl.sheets.store

/**
 * The row-radix bitmap trie under [[com.tjclp.xl.sheets.CellStore]] (#891).
 *
 * A row key is the row index read as an unsigned 32-bit number, held in a `Long`. A trie of height
 * `h` (1 to 7) consumes five key bits per level, the root at shift `5 * (h - 1)`. The entries at
 * shift 0 are slot groups (see [[Group]]): a [[Block]] when the group's cells share one 64-column
 * chunk, otherwise a [[Node]] of rows ([[StoredRow]]s). Height is minimal: `h` is the fewest levels
 * that cover the largest key, so a root taller than one level always has a child in a slot other
 * than 0. Height 0 is the empty trie, whose root is [[Node.Empty]]. Lookups are index arithmetic on
 * the key bits; nothing is hashed.
 */
// Flat-array loops over the trie (house style, as in SccKernel/LongIntMap): local vars and while
// loops behind pure signatures; nothing mutable escapes.
@SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
private[xl] object Trie:
  final val MaxHeight = 7
  final val UnsignedMask = 0xffffffffL

  /** The row key of a row index. */
  def key(row: Int): Long = row.toLong & UnsignedMask

  /** The fewest levels whose slots cover `key`. */
  def heightFor(key: Long): Int =
    var h = 1
    while h < MaxHeight && (key >>> (5 * h)) != 0L do h += 1
    h

  /** Whether a trie of height `h` can hold `key`. */
  def covers(h: Int, key: Long): Boolean = h >= MaxHeight || (key >>> (5 * h)) == 0L

  def slotIndex(bitmap: Int, slot: Int): Int = Integer.bitCount(bitmap & ((1 << slot) - 1))

  /**
   * The arrays holding the cell at (`key`, `col`), if it is stored: the row's [[RowLeaf]] (the leaf
   * of `col`'s band when the row is banded) or the [[Block]] of its group; `find(key, col)` then
   * locates the cell. [[RowLeaf.Empty]] when the row or the band cannot be present.
   */
  def lookup(root: Entry, height: Int, key: Long, col: Int): Cells =
    if height == 0 || !covers(height, key) then RowLeaf.Empty
    else
      var e = root
      var shift = 5 * (height - 1)
      var result: Cells = RowLeaf.Empty
      var searching = true
      while searching do
        e match
          case node: Node =>
            val bit = 1 << ((key >>> shift) & 31L).toInt
            if (node.bitmap & bit) == 0 then searching = false
            else
              e = node.children(Integer.bitCount(node.bitmap & (bit - 1)))
              shift -= 5
          case row: BandedRow =>
            result = row.bandOf(col)
            searching = false
          case c: Cells =>
            result = c
            searching = false
      result

  /** `root` (of height `from`) wrapped in slot-0 nodes up to height `to`. */
  def grow(root: Entry, from: Int, to: Int): Entry =
    if root.count == 0 then Node.Empty
    else
      var r = root
      var h = from
      while h < to do
        r = Node.of(1, Array[Entry](r))
        h += 1
      r

  /** The minimal-height form of a trie: strips slot-0-only roots; an empty trie has height 0. */
  def shrink(root: Entry, height: Int): (Entry, Int) =
    if root.count == 0 then (Node.Empty, 0)
    else
      var r = root
      var h = height
      var going = true
      while going && h > 1 do
        r match
          case n: Node if n.bitmap == 1 =>
            r = n.children(0)
            h -= 1
          case _ => going = false
      (r, h)

  def removeAt(children: Array[Entry], i: Int): Array[Entry] =
    val out = new Array[Entry](children.length - 1)
    System.arraycopy(children, 0, out, 0, i)
    System.arraycopy(children, i + 1, out, i, children.length - i - 1)
    out

  def insertAt(children: Array[Entry], i: Int, e: Entry): Array[Entry] =
    val out = new Array[Entry](children.length + 1)
    System.arraycopy(children, 0, out, 0, i)
    out(i) = e
    System.arraycopy(children, i, out, i + 1, children.length - i)
    out

  def replaceAt(children: Array[Entry], i: Int, e: Entry): Array[Entry] =
    val out = children.clone()
    out(i) = e
    out

  /**
   * `e` (at `shift`) with the row at `key` replaced by `f` of it (`f` sees [[RowLeaf.Empty]] when
   * the row is absent; returning an empty leaf removes the row). Returns `e` itself when `f`
   * returns its argument. The key must be covered by the entry.
   */
  def modify(e: Entry, shift: Int, key: Long, f: StoredRow => StoredRow): Entry = e match
    case b: Block =>
      val low = (key & 31L).toInt
      val present = (b.rowMask & (1 << low)) != 0
      val old = if present then Group.rowLeaf(b, b.rowIndex(low)) else RowLeaf.Empty
      val next = f(old)
      if next eq old then b else Group.withRow(b, low, old, next)
    case leaf: RowLeaf => f(leaf) // not reached: rows sit under shift-0 nodes
    case row: BandedRow => f(row) // not reached: rows sit under shift-0 nodes
    case node: Node =>
      val slot = ((key >>> shift) & 31L).toInt
      val bit = 1 << slot
      val i = Integer.bitCount(node.bitmap & (bit - 1))
      val result: Entry =
        if (node.bitmap & bit) != 0 then
          val child = node.children(i)
          val updated: Entry = child match
            case leaf: RowLeaf => f(leaf)
            case row: BandedRow => f(row)
            case other => modify(other, shift - 5, key, f)
          if updated eq child then node
          else if updated.count == 0 then
            if node.bitmap == bit then Node.Empty
            else Node.of(node.bitmap & ~bit, removeAt(node.children, i))
          else Node.of(node.bitmap, replaceAt(node.children, i, updated))
        else
          val created: Entry =
            if shift == 0 then f(RowLeaf.Empty) else modify(Node.Empty, shift - 5, key, f)
          if created.count == 0 then node
          else Node.of(node.bitmap | bit, insertAt(node.children, i, created))
      result match
        case n: Node if shift == 0 && !(n eq node) => Group.compact(n)
        case other => other

  /**
   * `e` (at `shift`, covering keys from `prefix`) with `f` applied to every row whose key lies in
   * `[lo, hi]`. A subtree or row for which `skip` holds is kept as it is, and so is every subtree
   * in which `f` returns each row unchanged; rows `f` empties are removed.
   */
  def mapRows(
    e: Entry,
    shift: Int,
    prefix: Long,
    lo: Long,
    hi: Long,
    skip: Entry => Boolean,
    f: StoredRow => StoredRow
  ): Entry =
    if e.count == 0 || skip(e) then e
    else if shift == 0 then
      val (lows, leaves, n) = Group.explode(e)
      var changed = false
      var i = 0
      while i < n do
        val k = prefix + lows(i)
        if k >= lo && k <= hi && !skip(leaves(i)) then
          val next = f(leaves(i))
          if !(next eq leaves(i)) then
            changed = true
            leaves(i) = next
        i += 1
      if !changed then e
      else
        // drop the rows f emptied, then rebuild the group canonically
        var k = 0
        var j = 0
        while j < n do
          if leaves(j).count > 0 then
            lows(k) = lows(j)
            leaves(k) = leaves(j)
            k += 1
          j += 1
        Group.of(lows, leaves, k)
    else
      e match
        case node: Node =>
          val n = node.children.length
          val out = new Array[Entry](n)
          var outBitmap = 0
          var k = 0
          var changed = false
          var bm = node.bitmap
          var i = 0
          while bm != 0 do
            val slot = Integer.numberOfTrailingZeros(bm)
            bm &= bm - 1
            val childLo = prefix + (slot.toLong << shift)
            val childHi = childLo + (1L << shift) - 1
            val child = node.children(i)
            val next =
              if childHi < lo || childLo > hi then child
              else mapRows(child, shift - 5, childLo, lo, hi, skip, f)
            if !(next eq child) then changed = true
            if next.count > 0 then
              out(k) = next
              outBitmap |= 1 << slot
              k += 1
            i += 1
          if !changed then node
          else if k == n then Node.of(outBitmap, out)
          else Node.of(outBitmap, java.util.Arrays.copyOf(out, k))
        case other => other // not reached: only nodes sit above shift 0

  /**
   * The smallest key in `[lo, hi]` under `e` holding a cell (a valued cell when `valuedOnly`), or
   * -1.
   */
  def firstKey(e: Entry, shift: Int, prefix: Long, lo: Long, hi: Long, valuedOnly: Boolean): Long =
    var found = -1L
    e match
      case b: Block =>
        var bm = b.rowMask
        var i = 0
        while found < 0 && bm != 0 do
          val low = Integer.numberOfTrailingZeros(bm)
          bm &= bm - 1
          val k = prefix + low
          if k > hi then bm = 0
          else if k >= lo && (!valuedOnly || Group.rowValued(b, i)) then found = k
          i += 1
      case node: Node =>
        var bm = node.bitmap
        var i = 0
        while found < 0 && bm != 0 do
          val slot = Integer.numberOfTrailingZeros(bm)
          bm &= bm - 1
          val childLo = prefix + (slot.toLong << shift)
          val childHi = childLo + (1L << shift) - 1
          val child = node.children(i)
          if childLo > hi then bm = 0
          else if childHi >= lo && (!valuedOnly || child.valued > 0) then
            child match
              case _: RowLeaf => found = childLo
              case _: BandedRow => found = childLo
              case other => found = firstKey(other, shift - 5, childLo, lo, hi, valuedOnly)
          i += 1
      case _: RowLeaf => ()
      case _: BandedRow => ()
    found

  /** The largest key in `[lo, hi]` under `e` holding a (valued) cell, or -1. */
  def lastKey(e: Entry, shift: Int, prefix: Long, lo: Long, hi: Long, valuedOnly: Boolean): Long =
    var found = -1L
    e match
      case b: Block =>
        var bm = b.rowMask
        var i = b.rows - 1
        while found < 0 && bm != 0 do
          val low = 31 - Integer.numberOfLeadingZeros(bm)
          bm &= ~(1 << low)
          val k = prefix + low
          if k < lo then bm = 0
          else if k <= hi && (!valuedOnly || Group.rowValued(b, i)) then found = k
          i -= 1
      case node: Node =>
        var bm = node.bitmap
        var i = node.children.length - 1
        while found < 0 && bm != 0 do
          val slot = 31 - Integer.numberOfLeadingZeros(bm)
          bm &= ~(1 << slot)
          val childLo = prefix + (slot.toLong << shift)
          val childHi = childLo + (1L << shift) - 1
          val child = node.children(i)
          if childHi < lo then bm = 0
          else if childLo <= hi && (!valuedOnly || child.valued > 0) then
            child match
              case _: RowLeaf => found = childLo
              case _: BandedRow => found = childLo
              case other => found = lastKey(other, shift - 5, childLo, lo, hi, valuedOnly)
          i -= 1
      case _: RowLeaf => ()
      case _: BandedRow => ()
    found

  /**
   * The trie holding `leaves(i)` at `keys(i)` for `i < n`; keys strictly ascending, leaves
   * non-empty. Returns the root and its (minimal) height.
   */
  def build(keys: Array[Long], leaves: Array[StoredRow], n: Int): (Entry, Int) =
    if n == 0 then (Node.Empty, 0)
    else
      val h = heightFor(keys(n - 1))
      (buildEntry(keys, leaves, 0, n, 5 * (h - 1)), h)

  private def buildEntry(
    keys: Array[Long],
    leaves: Array[StoredRow],
    from: Int,
    until: Int,
    shift: Int
  ): Entry =
    if shift == 0 then
      val n = until - from
      val lows = new Array[Int](n)
      val ls = new Array[StoredRow](n)
      var i = 0
      while i < n do
        lows(i) = (keys(from + i) & 31L).toInt
        ls(i) = leaves(from + i)
        i += 1
      Group.of(lows, ls, n)
    else
      val tmp = new Array[Entry](32)
      var bitmap = 0
      var k = 0
      var i = from
      while i < until do
        val slot = ((keys(i) >>> shift) & 31L).toInt
        var j = i + 1
        while j < until && ((keys(j) >>> shift) & 31L).toInt == slot do j += 1
        tmp(k) = buildEntry(keys, leaves, i, j, shift - 5)
        bitmap |= 1 << slot
        k += 1
        i = j
      Node.of(bitmap, java.util.Arrays.copyOf(tmp, k))

  /**
   * The trie holding the slot group `groups(i)` (a canonical shift-0 entry, non-empty) under the
   * group key `groupKeys(i)` (its rows' keys above their five low bits) for `i < n`; group keys
   * strictly ascending. Returns the root and its (minimal) height, as [[build]] does from rows.
   */
  def buildGroups(groupKeys: Array[Long], groups: Array[Entry], n: Int): (Entry, Int) =
    if n == 0 then (Node.Empty, 0)
    else
      val h = heightFor(groupKeys(n - 1) << 5)
      if h == 1 then (groups(0), 1)
      else (buildAbove(groupKeys, groups, 0, n, 5 * (h - 1)), h)

  /** The node at `shift` (5 or more) over the groups `from until until`, which share its prefix. */
  private def buildAbove(
    groupKeys: Array[Long],
    groups: Array[Entry],
    from: Int,
    until: Int,
    shift: Int
  ): Entry =
    val tmp = new Array[Entry](32)
    var bitmap = 0
    var k = 0
    var i = from
    while i < until do
      val slot = ((groupKeys(i) >>> (shift - 5)) & 31L).toInt
      var j = i + 1
      while j < until && ((groupKeys(j) >>> (shift - 5)) & 31L).toInt == slot do j += 1
      tmp(k) = if shift == 5 then groups(i) else buildAbove(groupKeys, groups, i, j, shift - 5)
      bitmap |= 1 << slot
      k += 1
      i = j
    Node.of(bitmap, java.util.Arrays.copyOf(tmp, k))

  /**
   * The union of two entries at the same `shift`; on a row both hold, the rows merge through
   * [[StoredRow.merge]] (`b` wins, or with `valuesOnly` `b`'s values over `a`'s styles and links).
   * Above shift 0, returns `a` itself when the union adds nothing to it.
   */
  def union(a: Entry, b: Entry, shift: Int, valuesOnly: Boolean): Entry =
    if b.count == 0 then a
    else if a.count == 0 then b
    else if shift == 0 then
      val (la, va, na) = Group.explode(a)
      val (lb, vb, nb) = Group.explode(b)
      val lows = new Array[Int](na + nb)
      val leaves = new Array[StoredRow](na + nb)
      var i = 0
      var j = 0
      var k = 0
      while i < na || j < nb do
        if j >= nb || (i < na && la(i) < lb(j)) then
          lows(k) = la(i); leaves(k) = va(i); i += 1
        else if i >= na || lb(j) < la(i) then
          lows(k) = lb(j); leaves(k) = vb(j); j += 1
        else
          lows(k) = la(i); leaves(k) = StoredRow.merge(va(i), vb(j), valuesOnly); i += 1; j += 1
        k += 1
      Group.of(lows, leaves, k)
    else
      (a, b) match
        case (na: Node, nb: Node) =>
          val bitmap = na.bitmap | nb.bitmap
          val out = new Array[Entry](Integer.bitCount(bitmap))
          var changed = bitmap != na.bitmap
          var bm = bitmap
          var k = 0
          while bm != 0 do
            val slot = Integer.numberOfTrailingZeros(bm)
            bm &= bm - 1
            val bit = 1 << slot
            val inA = (na.bitmap & bit) != 0
            val inB = (nb.bitmap & bit) != 0
            val child: Entry =
              if !inB then na.children(slotIndex(na.bitmap, slot))
              else if !inA then nb.children(slotIndex(nb.bitmap, slot))
              else
                union(
                  na.children(slotIndex(na.bitmap, slot)),
                  nb.children(slotIndex(nb.bitmap, slot)),
                  shift - 5,
                  valuesOnly
                )
            if inA && !(child eq na.children(slotIndex(na.bitmap, slot))) then changed = true
            out(k) = child
            k += 1
          if changed then Node.of(bitmap, out) else a
        case _ => a // not reached: only nodes sit above shift 0

  /**
   * Calls `f(key, row)` for every row under `e`, in ascending key order (block rows are copied).
   */
  def foreachRow(e: Entry, shift: Int, prefix: Long, f: (Long, StoredRow) => Unit): Unit =
    if shift == 0 then
      val (lows, leaves, n) = Group.explode(e)
      var i = 0
      while i < n do
        f(prefix + lows(i), leaves(i))
        i += 1
    else
      e match
        case node: Node =>
          var bm = node.bitmap
          var i = 0
          while bm != 0 do
            val slot = Integer.numberOfTrailingZeros(bm)
            bm &= bm - 1
            foreachRow(node.children(i), shift - 5, prefix + (slot.toLong << shift), f)
            i += 1
        case _ => ()

  /** Structural equality of two entries, short-circuiting on shared subtrees. */
  def same(a: Entry, b: Entry): Boolean =
    if a eq b then true
    else
      a match
        case na: Node =>
          b match
            case nb: Node =>
              if na.bitmap != nb.bitmap || na.count != nb.count || na.valued != nb.valued then false
              else
                var ok = true
                var i = 0
                while ok && i < na.children.length do
                  ok = same(na.children(i), nb.children(i))
                  i += 1
                ok
            case _ => false
        case ba: Block =>
          b match
            case bb: Block => Group.same(ba, bb)
            case _ => false
        case la: RowLeaf =>
          b match
            case lb: RowLeaf => RowLeaf.same(la, lb)
            case _ => false
        case ra: BandedRow =>
          b match
            case rb: BandedRow => BandedRow.same(ra, rb)
            case _ => false

/**
 * A stable sort of `Long` keys that reports the permutation, so equal keys keep their arrival order
 * (the builder's last-write-wins rule depends on it). A bottom-up merge sort over `Int` indices: no
 * boxing.
 */
// Flat-array loops over the trie (house style, as in SccKernel/LongIntMap): local vars and while
// loops behind pure signatures; nothing mutable escapes.
@SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
private[xl] object StableSort:
  def order(keys: Array[Long], n: Int): Array[Int] =
    var src = Array.tabulate(n)(identity)
    var dst = new Array[Int](n)
    var width = 1
    while width < n do
      var lo = 0
      while lo < n do
        val mid = math.min(lo + width, n)
        val hi = math.min(lo + 2 * width, n)
        var i = lo
        var j = mid
        var k = lo
        while k < hi do
          if j >= hi || (i < mid && keys(src(i)) <= keys(src(j))) then
            dst(k) = src(i)
            i += 1
          else
            dst(k) = src(j)
            j += 1
          k += 1
        lo += 2 * width
      val t = src
      src = dst
      dst = t
      width *= 2
    src
