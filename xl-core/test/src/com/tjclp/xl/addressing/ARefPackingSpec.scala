package com.tjclp.xl.addressing

import munit.FunSuite

/**
 * #891: `ARef` packs the column as `col ^ (row << 14)` in its low word. The packing must stay a
 * bijection on every `Int` pair (off-grid references are legal), and a `Long` hash of it must
 * spread the cells of a tall sheet.
 */
class ARefPackingSpec extends FunSuite:
  private val edge = Vector(
    0,
    1,
    2,
    16383,
    16384,
    16385,
    262143,
    262144,
    1048575,
    1048576,
    Int.MaxValue,
    Int.MinValue,
    -1,
    -2,
    -16384,
    1 << 17,
    (1 << 18) + 5,
    0x7fffc000,
    0x80004000
  )

  test("row and column round-trip, and distinct pairs stay distinct, at every edge") {
    val all = for r <- edge; c <- edge yield (c, r)
    val refs = all.map((c, r) => ARef.from0(c, r))
    all.zip(refs).foreach { case ((c, r), ref) =>
      assertEquals(ref.col.index0, c, s"col of ($c,$r)")
      assertEquals(ref.row.index0, r, s"row of ($c,$r)")
    }
    assertEquals(refs.toSet.size, all.distinct.size)
  }

  test("toA1, parse and shift at the corners of the grid") {
    assertEquals(ARef.from0(0, 0).toA1, "A1")
    assertEquals(ARef.from0(16383, 1048575).toA1, "XFD1048576")
    assertEquals(ARef.parse("XFD1048576"), Right(ARef.from0(16383, 1048575)))
    assertEquals(ARef.from0(16383, 1048575).shift(-16383, -1048575), ARef.from0(0, 0))
    assertEquals(ARef.from0(0, 0).shift(-1, -1).shift(1, 1), ARef.from0(0, 0))
  }

  test("hashes spread: a tall strip is collision-free and column A never shares a bucket") {
    val tall = for r <- 0 until 1048576 by 37; c <- 0 until 64 yield ARef.from0(c, r)
    assertEquals(tall.map(_.hashCode).distinct.size, tall.size)
    val columnA = (0 until 1048576).map(r => ARef.from0(0, r).hashCode)
    assertEquals(columnA.distinct.size, columnA.size)
  }
