package com.tjclp.xl.sheets.store

/**
 * The row-radix bitmap trie under [[com.tjclp.xl.sheets.CellStore]] (#891).
 *
 * A row key is the row index read as an unsigned 32-bit number, held in a `Long`. A trie of height
 * `h` (1 to 7) consumes five key bits per level, the root at shift `5 * (h - 1)`, the bottom node
 * at shift 0 holding [[RowLeaf]]s. Height is minimal: `h` is the fewest levels that cover the
 * largest key, so the root of a trie taller than one level always has a child in a slot other than
 * 0. Height 0 is the empty trie, whose root is [[Node.Empty]]. Lookups are index arithmetic on the
 * key bits; nothing is hashed.
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

  /** The leaf at `key` ([[RowLeaf.Empty]] when the row holds nothing). */
  def lookup(root: Node, height: Int, key: Long): RowLeaf =
    if height == 0 || !covers(height, key) then RowLeaf.Empty
    else
      var node = root
      var shift = 5 * (height - 1)
      var result: RowLeaf = RowLeaf.Empty
      var searching = true
      while searching do
        val slot = ((key >>> shift) & 31L).toInt
        val bit = 1 << slot
        if (node.bitmap & bit) == 0 then searching = false
        else
          node.children(Integer.bitCount(node.bitmap & (bit - 1))) match
            case leaf: RowLeaf =>
              result = leaf
              searching = false
            case child: Node =>
              node = child
              shift -= 5
      result

  /** `root` (of height `from`) wrapped in slot-0 nodes up to height `to`. */
  def grow(root: Node, from: Int, to: Int): Node =
    if root.count == 0 then Node.Empty
    else
      var r = root
      var h = from
      while h < to do
        r = Node.of(1, Array[Entry](r))
        h += 1
      r

  /** The minimal-height form of a trie: strips slot-0-only roots; an empty trie has height 0. */
  def shrink(root: Node, height: Int): (Node, Int) =
    if root.count == 0 then (Node.Empty, 0)
    else
      var r = root
      var h = height
      var going = true
      while going && h > 1 && r.bitmap == 1 do
        r.children(0) match
          case n: Node =>
            r = n
            h -= 1
          case _: RowLeaf => going = false
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
   * `node` (at `shift`) with the leaf at `key` replaced by `f` of it (`f` sees [[RowLeaf.Empty]]
   * when the row is absent; returning an empty leaf removes the row). Returns `node` itself when
   * `f` returns its argument. The key must be covered by the node.
   */
  def modify(node: Node, shift: Int, key: Long, f: RowLeaf => RowLeaf): Node =
    val slot = ((key >>> shift) & 31L).toInt
    val bit = 1 << slot
    val i = Integer.bitCount(node.bitmap & (bit - 1))
    if (node.bitmap & bit) != 0 then
      val child = node.children(i)
      val updated: Entry = child match
        case leaf: RowLeaf => f(leaf)
        case n: Node => modify(n, shift - 5, key, f)
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

  /**
   * `node` (at `shift`, covering keys from `prefix`) with `f` applied to every leaf whose key lies
   * in `[lo, hi]`. A subtree for which `skip` holds is kept as it is, and so is every subtree in
   * which `f` returns each leaf unchanged; leaves `f` empties are removed.
   */
  def mapRows(
    node: Node,
    shift: Int,
    prefix: Long,
    lo: Long,
    hi: Long,
    skip: Entry => Boolean,
    f: RowLeaf => RowLeaf
  ): Node =
    if node.count == 0 || skip(node) then node
    else
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
        val next: Entry =
          if childHi < lo || childLo > hi || skip(child) then child
          else
            child match
              case leaf: RowLeaf => f(leaf)
              case inner: Node => mapRows(inner, shift - 5, childLo, lo, hi, skip, f)
        if !(next eq child) then changed = true
        if next.count > 0 then
          out(k) = next
          outBitmap |= 1 << slot
          k += 1
        i += 1
      if !changed then node
      else if k == n then Node.of(outBitmap, out)
      else Node.of(outBitmap, java.util.Arrays.copyOf(out, k))

  /**
   * The smallest key in `[lo, hi]` under `node` holding a cell (a valued cell when `valuedOnly`),
   * or -1.
   */
  def firstKey(
    node: Node,
    shift: Int,
    prefix: Long,
    lo: Long,
    hi: Long,
    valuedOnly: Boolean
  ): Long =
    var bm = node.bitmap
    var i = 0
    var found = -1L
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
          case inner: Node => found = firstKey(inner, shift - 5, childLo, lo, hi, valuedOnly)
      i += 1
    found

  /** The largest key in `[lo, hi]` under `node` holding a (valued) cell, or -1. */
  def lastKey(node: Node, shift: Int, prefix: Long, lo: Long, hi: Long, valuedOnly: Boolean): Long =
    var bm = node.bitmap
    var i = node.children.length - 1
    var found = -1L
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
          case inner: Node => found = lastKey(inner, shift - 5, childLo, lo, hi, valuedOnly)
      i -= 1
    found

  /**
   * The trie holding `leaves(i)` at `keys(i)` for `i < n`; keys strictly ascending, leaves
   * non-empty. Returns the root and its (minimal) height.
   */
  def build(keys: Array[Long], leaves: Array[RowLeaf], n: Int): (Node, Int) =
    if n == 0 then (Node.Empty, 0)
    else
      val h = heightFor(keys(n - 1))
      (buildNode(keys, leaves, 0, n, 5 * (h - 1)), h)

  private def buildNode(
    keys: Array[Long],
    leaves: Array[RowLeaf],
    from: Int,
    until: Int,
    shift: Int
  ): Node =
    val tmp = new Array[Entry](32)
    var bitmap = 0
    var k = 0
    var i = from
    while i < until do
      val slot = ((keys(i) >>> shift) & 31L).toInt
      var j = i + 1
      while j < until && ((keys(j) >>> shift) & 31L).toInt == slot do j += 1
      tmp(k) = if shift == 0 then leaves(i) else buildNode(keys, leaves, i, j, shift - 5)
      bitmap |= 1 << slot
      k += 1
      i = j
    Node.of(bitmap, java.util.Arrays.copyOf(tmp, k))

  /**
   * The union of two nodes at the same `shift`; on a row both hold, the rows merge through
   * [[RowLeaf.merge]] (`b` wins, or with `valuesOnly` `b`'s values over `a`'s styles and links).
   * Returns `a` itself when the union adds nothing to it.
   */
  def union(a: Node, b: Node, shift: Int, valuesOnly: Boolean): Node =
    if b.count == 0 then a
    else if a.count == 0 then b
    else
      val bitmap = a.bitmap | b.bitmap
      val out = new Array[Entry](Integer.bitCount(bitmap))
      var changed = bitmap != a.bitmap
      var bm = bitmap
      var k = 0
      while bm != 0 do
        val slot = Integer.numberOfTrailingZeros(bm)
        bm &= bm - 1
        val bit = 1 << slot
        val inA = (a.bitmap & bit) != 0
        val inB = (b.bitmap & bit) != 0
        val child: Entry =
          if !inB then a.children(slotIndex(a.bitmap, slot))
          else if !inA then b.children(slotIndex(b.bitmap, slot))
          else
            val ca = a.children(slotIndex(a.bitmap, slot))
            val cb = b.children(slotIndex(b.bitmap, slot))
            (ca, cb) match
              case (la: RowLeaf, lb: RowLeaf) => RowLeaf.merge(la, lb, valuesOnly)
              case (na: Node, nb: Node) => union(na, nb, shift - 5, valuesOnly)
              case _ => ca // unreachable: siblings sit at one depth
        if inA && !(child eq a.children(slotIndex(a.bitmap, slot))) then changed = true
        out(k) = child
        k += 1
      if changed then Node.of(bitmap, out) else a

  /** Calls `f(key, leaf)` for every row under `node`, in ascending key order. */
  def foreachRow(node: Node, shift: Int, prefix: Long, f: (Long, RowLeaf) => Unit): Unit =
    var bm = node.bitmap
    var i = 0
    while bm != 0 do
      val slot = Integer.numberOfTrailingZeros(bm)
      bm &= bm - 1
      val childLo = prefix + (slot.toLong << shift)
      node.children(i) match
        case leaf: RowLeaf => f(childLo, leaf)
        case inner: Node => foreachRow(inner, shift - 5, childLo, f)
      i += 1

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
            case _: RowLeaf => false
        case la: RowLeaf =>
          b match
            case lb: RowLeaf => RowLeaf.same(la, lb)
            case _: Node => false

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
