package com.tjclp.xl.sheets

import scala.util.Random

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.styles.units.StyleId

/**
 * A seeded fuzz of [[CellStore]] against a plain `Map` oracle (#891): long random walks through
 * every write (`updated`, `updatedValue`, `removed`, `removedIn`, `putAll`, `withValues`,
 * `transformValues`, `shiftRows`, `shiftCols`) over small, adversarial domains, where
 * [[CellStoreLawSpec]] runs many short generated sequences. After every step the store must hold
 * its structural invariants, iterate exactly the oracle's cells in row-major order, report the
 * brute-force `usedRange` and `extent`, equal (with the same hash and height) a store rebuilt from
 * the same cells shuffled, answer point reads as the oracle does, and leave the previous version
 * untouched.
 *
 * The walks stay inside one or two slot groups for hundreds of steps, so they reach states the
 * generated laws seldom do: a block that loses the cell holding its extreme valued column or its
 * last tag, a group hovering at the block cell cap, a row crossing a 64-column chunk edge. Seeds
 * are fixed, so a failure names the seed, the step and the operation and replays exactly.
 */
// A fuzz loop: local vars and a while loop over the steps.
@SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
class CellStoreFuzzSpec extends munit.FunSuite:

  // 105,000 store operations, each followed by a full comparison with the oracle: about 7 s on
  // a quiet machine. The timeout is only a backstop against a hang on a loaded CI runner.
  override val munitTimeout = scala.concurrent.duration.Duration(5, "min")

  private def ord(a: (Int, Int), b: (Int, Int)): Boolean =
    val r = Integer.compareUnsigned(a._1, b._1)
    if r != 0 then r < 0 else a._2 < b._2

  private def dest(i: Int, at: Int, by: Int): Option[Int] =
    val il = i.toLong
    if by < 0 && il >= at.toLong + by && il < at.toLong then None
    else if il >= at.toLong then
      val d = il + by
      if d >= Int.MinValue && d <= Int.MaxValue then Some(d.toInt) else None
    else Some(i)

  private def brute(cells: Iterable[Cell]): Option[CellRange] =
    if cells.isEmpty then None
    else
      val rs = cells.map(_.ref.row.index0)
      val cs = cells.map(_.ref.col.index0)
      Some(CellRange(ARef.from0(cs.min, rs.min), ARef.from0(cs.max, rs.max)))

  private def run(seed: Long, steps: Int, rowsDom: Vector[Int], colsDom: Vector[Int]): Unit =
    val rnd = new Random(seed)
    def row() = rowsDom(rnd.nextInt(rowsDom.length))
    def col() = colsDom(rnd.nextInt(colsDom.length))
    def ref() = ARef.from0(col(), row())
    def value(): CellValue = rnd.nextInt(6) match
      case 0 | 1 => CellValue.Empty
      case 2 => CellValue.Text("t" + rnd.nextInt(3))
      case 3 => CellValue.Number(BigDecimal(rnd.nextInt(5)))
      case 4 => CellValue.Formula("A1", None)
      case _ => CellValue.Formula("A1", Some(CellValue.Number(BigDecimal(1))))
    def cell(): Cell =
      Cell(
        ref(),
        value(),
        if rnd.nextInt(4) == 0 then Some(StyleId(rnd.nextInt(3) - 1)) else None,
        if rnd.nextInt(6) == 0 then Some("l" + rnd.nextInt(2)) else None
      )
    var o = Map.empty[(Int, Int), Cell]
    var s = CellStore.empty
    def k(r: ARef) = (r.row.index0, r.col.index0)
    var i = 0
    while i < steps do
      val prevS = s
      val prevO = o
      val op = rnd.nextInt(22)
      val desc: String =
        if op < 8 then
          val c = cell(); o = o.updated(k(c.ref), c); s = s.updated(c); s"updated $c"
        else if op < 11 then
          val r = ref(); val v = value()
          o = o.updated(k(r), o.get(k(r)).fold(Cell(r, v))(_.withValue(v)))
          s = s.updatedValue(r, v); s"updatedValue ${k(r)} $v"
        else if op < 15 then
          val r = ref(); o = o - k(r); s = s.removed(r); s"removed ${k(r)}"
        else if op == 15 then
          val rg = CellRange(ref(), ref())
          o = o.filterNot((_, c) => rg.contains(c.ref)); s = s.removedIn(rg); s"removedIn $rg"
        else if op == 16 then
          val cs = List.fill(rnd.nextInt(40))(cell())
          o = cs.foldLeft(o)((m, c) => m.updated(k(c.ref), c)); s = s.putAll(cs); "putAll"
        else if op == 17 then
          val ps = List.fill(rnd.nextInt(40))(ref() -> value())
          o = ps.foldLeft(o) { case (m, (r, v)) =>
            m.updated(k(r), m.get(k(r)).fold(Cell(r, v))(_.withValue(v)))
          }
          s = s.withValues(IArray.from(ps.map(_._1)), IArray.from(ps.map(_._2))); "withValues"
        else if op == 18 then
          val mask = rnd.nextInt(1024)
          val f: CellValue => CellValue =
            if rnd.nextBoolean() then _ => CellValue.Empty
            else
              case CellValue.Formula(e, _, kd) => CellValue.Formula(e, None, kd)
              case v => v
          o = o.view.mapValues { c =>
            if (mask & CellStore.tagBit(CellStore.tagOf(c.value))) != 0 then c.withValue(f(c.value))
            else c
          }.toMap
          s = s.transformValues(mask)(f); s"transform $mask"
        else if op == 19 then
          val at = row(); val by = List(-33, -32, -3, -1, 1, 2, 31, 32, 33, 1000)(rnd.nextInt(10))
          o = o.values.flatMap { c =>
            dest(c.ref.row.index0, at, by).map { r =>
              val nr = ARef.from0(c.ref.col.index0, r); k(nr) -> c.copy(ref = nr)
            }
          }.toMap
          s = s.shiftRows(at, by); s"shiftRows $at $by"
        else if op == 20 then
          val at = col(); val by = List(-65, -64, -3, -1, 1, 2, 63, 64, 65)(rnd.nextInt(9))
          o = o.values.flatMap { c =>
            dest(c.ref.col.index0, at, by).map { cc =>
              val nr = ARef.from0(cc, c.ref.row.index0); k(nr) -> c.copy(ref = nr)
            }
          }.toMap
          s = s.shiftCols(at, by); s"shiftCols $at $by"
        else
          // clear every row from 32 up, so the trie shrinks
          val rg = CellRange(ARef.from0(Int.MinValue, 32), ARef.from0(Int.MaxValue, Int.MaxValue))
          o = o.filterNot((_, c) => rg.contains(c.ref)); s = s.removedIn(rg); "clear rows>=32"
      val expected = o.toList.sortWith((a, b) => ord(a._1, b._1)).map(_._2)
      val clue = s"seed=$seed step=$i op=$desc"
      assertEquals(s.checkInvariants, Right(()), clue)
      assertEquals(s.iterator.toList, expected, clue)
      assertEquals(s.size, o.size, clue)
      assertEquals(s.usedRange, brute(expected.filter(_.nonEmpty)), clue)
      assertEquals(s.extent, brute(expected), clue)
      val rebuilt = CellStore.from(Random(seed + i).shuffle(expected))
      assert(s == rebuilt, clue)
      assertEquals(s.hashCode, rebuilt.hashCode, clue)
      assertEquals(s.debugHeight, rebuilt.debugHeight, clue)
      // persistence: the previous version must be untouched
      assertEquals(prevS.size, prevO.size, clue)
      if i % 7 == 0 then
        assertEquals(
          prevS.iterator.toList,
          prevO.toList.sortWith((a, b) => ord(a._1, b._1)).map(_._2),
          clue + " (persistence)"
        )
      // point reads incl. misses
      var j = 0
      while j < 6 do
        val r = ref()
        assertEquals(s.get(r), o.get(k(r)), clue)
        j += 1
      i += 1

  test("fuzz: one narrow block group, cap boundary (rows 0..40, cols 0..9)") {
    (1L to 150L).foreach(seed => run(seed, 250, (0 to 40).toVector, (0 to 9).toVector))
  }

  test("fuzz: chunk edges 60..70 and negative cols, rows across groups/heights") {
    val rows = Vector(
      0,
      1,
      30,
      31,
      32,
      33,
      63,
      64,
      1023,
      1024,
      32767,
      32768,
      1048575,
      1048576,
      -1,
      -2,
      Int.MinValue,
      Int.MaxValue
    )
    val cols =
      Vector(-65, -64, -1, 0, 1, 62, 63, 64, 65, 127, 128, 16383, 16384, Int.MinValue, Int.MaxValue)
    (1L to 150L).foreach(seed => run(seed, 250, rows, cols))
  }

  test("fuzz: dense 32 rows x 8..9 cols hovering at MaxBlockCells") {
    (1L to 100L).foreach(seed => run(seed, 300, (0 to 31).toVector, (0 to 8).toVector))
  }
