package com.tjclp.xl.sheets

import java.time.LocalDateTime

import scala.collection.immutable.SortedMap

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.{Cell, CellError, CellValue, FormulaKind}
import com.tjclp.xl.richtext.RichText
import com.tjclp.xl.sheets.store.CellCursor
import com.tjclp.xl.styles.units.StyleId
import munit.ScalaCheckSuite
import org.scalacheck.{Gen, Prop}
import org.scalacheck.Prop.*

/**
 * Laws for [[CellStore]] against an immutable `SortedMap` oracle in the store's documented order
 * (#891): rows ascending as unsigned 32-bit numbers, then columns ascending as signed Ints.
 *
 * Every random operation sequence is checked step by step: reads, iteration order, bounds, range
 * walks, tags, structural invariants, canonical equality with a store rebuilt from the oracle, and
 * persistence (the previous store still matches the previous oracle, which catches writes into
 * shared arrays).
 */
// Cursor walks in the spec are while loops over the private[xl] cursor.
@SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
class CellStoreLawSpec extends ScalaCheckSuite:
  import CellStoreLawSpec.*

  override def scalaCheckTestParameters =
    super.scalaCheckTestParameters.withMinSuccessfulTests(500).withMaxSize(40)

  // ----- the oracle -----

  type Oracle = SortedMap[(Int, Int), Cell]

  def oracleOf(cells: Iterable[Cell]): Oracle =
    cells.foldLeft(SortedMap.empty[(Int, Int), Cell](using StoreOrder))((m, c) =>
      m.updated(key(c.ref), c)
    )

  def key(ref: ARef): (Int, Int) = (ref.row.index0, ref.col.index0)

  def destOf(i: Int, at: Int, by: Int): Option[Int] =
    val il = i.toLong
    if by < 0 && il >= at.toLong + by && il < at.toLong then None
    else if il >= at.toLong then
      val d = il + by
      if d >= Int.MinValue && d <= Int.MaxValue then Some(d.toInt) else None
    else Some(i)

  def applyOracle(o: Oracle, op: Op): Oracle = op match
    case Op.Update(c) => o.updated(key(c.ref), c)
    case Op.UpdateValue(ref, v) =>
      o.updated(key(ref), o.get(key(ref)).fold(Cell(ref, v))(_.withValue(v)))
    case Op.Remove(ref) => o - key(ref)
    case Op.RemoveIn(r) => o.filterNot((_, c) => r.contains(c.ref))
    case Op.PutAll(cells) => cells.foldLeft(o)((m, c) => m.updated(key(c.ref), c))
    case Op.WithValues(pairs) =>
      pairs.foldLeft(o) { case (m, (ref, v)) =>
        m.updated(key(ref), m.get(key(ref)).fold(Cell(ref, v))(_.withValue(v)))
      }
    case Op.Transform(mask, fn) =>
      o.transform { (_, c) =>
        if (mask & CellStore.tagBit(CellStore.tagOf(c.value))) != 0 then c.withValue(fn(c.value))
        else c
      }
    case Op.ShiftRows(at, by) =>
      oracleOf(o.values.flatMap { c =>
        destOf(c.ref.row.index0, at, by).map(r => c.copy(ref = ARef.from0(c.ref.col.index0, r)))
      })
    case Op.ShiftCols(at, by) =>
      oracleOf(o.values.flatMap { c =>
        destOf(c.ref.col.index0, at, by).map(cc => c.copy(ref = ARef.from0(cc, c.ref.row.index0)))
      })

  def applyStore(s: CellStore, op: Op): CellStore = op match
    case Op.Update(c) => s.updated(c)
    case Op.UpdateValue(ref, v) => s.updatedValue(ref, v)
    case Op.Remove(ref) => s.removed(ref)
    case Op.RemoveIn(r) => s.removedIn(r)
    case Op.PutAll(cells) => s.putAll(cells)
    case Op.WithValues(pairs) =>
      s.withValues(IArray.from(pairs.map(_._1)), IArray.from(pairs.map(_._2)))
    case Op.Transform(mask, fn) => s.transformValues(mask)(fn)
    case Op.ShiftRows(at, by) => s.shiftRows(at, by)
    case Op.ShiftCols(at, by) => s.shiftCols(at, by)

  // ----- brute-force answers -----

  /** `Sheet.usedRange`'s loop, verbatim in spirit: signed bounds of the non-Empty cells. */
  def bruteBox(cells: Iterable[Cell]): Option[CellRange] =
    if cells.isEmpty then None
    else
      val init = (Int.MaxValue, Int.MaxValue, Int.MinValue, Int.MinValue)
      val (r0, c0, r1, c1) = cells.foldLeft(init) { case ((r0, c0, r1, c1), c) =>
        val r = c.ref.row.index0
        val col = c.ref.col.index0
        (math.min(r0, r), math.min(c0, col), math.max(r1, r), math.max(c1, col))
      }
      Some(CellRange(ARef.from0(c0, r0), ARef.from0(c1, r1)))

  /** Every agreement between a store and its oracle, as one Prop. */
  def agrees(s: CellStore, o: Oracle, probes: List[ARef], ranges: List[CellRange]): Prop =
    val expected = o.values.toList
    val reads = (probes ++ expected.map(_.ref)).map { ref =>
      val oc = o.get(key(ref))
      (s.get(ref) == oc) :| s"get ${key(ref)}: ${s.get(ref)} vs $oc" &&
      (s.contains(ref) == oc.isDefined) :| s"contains ${key(ref)}" &&
      (s.valueAt(ref) == oc.fold(CellValue.Empty)(_.value)) :| s"valueAt ${key(ref)}" &&
      (s.styleAt(ref) == oc.flatMap(_.styleId)) :| s"styleAt ${key(ref)}" &&
      (s.hyperlinkAt(ref) == oc.flatMap(_.hyperlink)) :| s"hyperlinkAt ${key(ref)}"
    }
    val walks = ranges.map { r =>
      val want = expected.filter(c => r.contains(c.ref))
      (s.inRange(r).toList == want) :| s"inRange $r" &&
      (cursorCells(CellCursor.range(s, r)) == want) :| s"cursor range $r"
    }
    val tags = (0 until 10).toList.map { t =>
      val want = expected.exists(c => CellStore.tagOf(c.value) == t)
      (s.hasTag(CellStore.tagBit(t)) == want) :| s"hasTag $t"
    }
    val rebuilt = CellStore.from(expected)
    (s.checkInvariants == Right(())) :| s"invariants: ${s.checkInvariants} in $s" &&
    (s.size == o.size) :| s"size ${s.size} vs ${o.size}" &&
    (s.valuedCount == expected.count(_.nonEmpty)) :| "valuedCount" &&
    (s.isEmpty == o.isEmpty) :| "isEmpty" &&
    (s.iterator.toList == expected) :| s"iterator order: $s vs ${expected.map(c => key(c.ref))}" &&
    (s.refs.toList == expected.map(_.ref)) :| "refs order" &&
    (s.usedRange == bruteBox(expected.filter(_.nonEmpty))) :| s"usedRange ${s.usedRange}" &&
    (s.extent == bruteBox(expected)) :| s"extent ${s.extent}" &&
    (s == rebuilt) :| s"canonical: $s vs rebuilt $rebuilt" &&
    (s.hashCode == rebuilt.hashCode) :| "hashCode consistent" &&
    (s.debugHeight == rebuilt.debugHeight) :| "minimal height" &&
    Prop.all(reads*) && Prop.all(walks*) && Prop.all(tags*)

  def cursorCells(c: CellCursor): List[Cell] =
    val b = List.newBuilder[Cell]
    while c.next() do b += c.cell
    b.result()

  // ----- laws -----

  def sequenceAgrees(
    start: List[Cell],
    ops: List[Op],
    probes: List[ARef],
    ranges: List[CellRange]
  ): Prop =
    val s0 = CellStore.from(start)
    val o0 = oracleOf(start)
    val (_, _, props) = ops.foldLeft((s0, o0, List(agrees(s0, o0, probes, ranges)))) {
      case ((s, o, acc), op) =>
        val s2 = applyStore(s, op)
        val o2 = applyOracle(o, op)
        val persistent = (s.iterator.toList == o.values.toList) :| s"persistence after $op"
        (s2, o2, (agrees(s2, o2, probes, ranges) :| s"after $op") :: persistent :: acc)
    }
    Prop.all(props*)

  property("every operation sequence agrees with the SortedMap oracle, step by step") {
    forAll(genStart, Gen.listOf(genOp), Gen.listOfN(4, genRef), Gen.listOfN(2, genRange))(
      sequenceAgrees
    )
  }

  // The wide pools above rarely keep a slot group inside one chunk, so they seldom build a block
  // and almost never take a block's extreme valued column or last cell of a tag away. A block
  // keeps its summaries (valued bounds, tag mask) by shortcut on such writes; this domain is one
  // dense block and its neighbour, written and emptied cell by cell.
  property("a dense block agrees with the oracle under cell-by-cell writes and removals") {
    forAll(
      genBlockStart,
      Gen.choose(20, 60).flatMap(Gen.listOfN(_, genBlockOp)),
      Gen.listOfN(4, genBlockRef),
      Gen.listOfN(2, genBlockRange)
    )(sequenceAgrees)
  }

  property("from / newBuilder equal a fold of updated: sorted, unsorted and with duplicates") {
    forAll(Gen.listOf(genCell), Gen.choose(0L, Long.MaxValue)) { (cells, seed) =>
      val folded = cells.foldLeft(CellStore.empty)(_.updated(_))
      val sorted = oracleOf(cells).values.toList
      val shuffled = new scala.util.Random(seed).shuffle(cells)
      val shuffledFold = shuffled.foldLeft(CellStore.empty)(_.updated(_))
      val dup = cells ++ cells.reverse ++ cells // last occurrence of each ref = its last in `cells`
      val b = CellStore.newBuilder
      cells.foreach(b += _)
      val first = b.result()
      b.addAll(Nil)
      (CellStore.from(cells) == folded) :| "from unsorted" &&
      (CellStore.from(sorted) == folded) :| "from sorted" &&
      (CellStore.from(sorted).iterator.toList == sorted) :| "sorted round trip" &&
      (CellStore.from(shuffled) == shuffledFold) :| "from shuffled" &&
      (CellStore.from(dup) == folded) :| "duplicates: last write wins" &&
      (first == folded && b.result() == folded) :| "builder result is repeatable" &&
      (CellStore.of(cells*) == folded) :| "of"
    }
  }

  property("putAll equals a fold of updated, and shares an untouched store") {
    forAll(genStart, Gen.listOf(genCell)) { (start, more) =>
      val s = CellStore.from(start)
      (s.putAll(more) == more.foldLeft(s)(_.updated(_))) :| "putAll" &&
      (s.putAll(Nil) eq s) :| "empty putAll returns this"
    }
  }

  property("withValues and transformValues equal their maps over the oracle") {
    forAll(genStart, Gen.listOf(genRef.flatMap(r => genValue.map(r -> _))), genTagMask, genFn) {
      (start, pairs, mask, fn) =>
        val s = CellStore.from(start)
        val o = oracleOf(start)
        val wv = s.withValues(IArray.from(pairs.map(_._1)), IArray.from(pairs.map(_._2)))
        val tv = s.transformValues(mask)(fn._2)
        (wv.iterator.toList == applyOracle(
          o,
          Op.WithValues(pairs)
        ).values.toList) :| "withValues" &&
        (tv.iterator.toList == applyOracle(
          o,
          Op.Transform(mask, fn._2)
        ).values.toList) :| "transform" &&
        (s.transformValues(mask)(identity) eq s) :| "identity transform shares the store" &&
        (s.withValues(IArray(ARef.from0(1, 1)), IArray.empty[CellValue]) eq s) :| "length mismatch"
    }
  }

  property("shiftRows / shiftCols equal re-keying the oracle; delete = shift past the window") {
    forAll(genStart, genIndex, genBy) { (start, at, by) =>
      val s = CellStore.from(start)
      val o = oracleOf(start)
      (s.shiftRows(at, by).iterator.toList ==
        applyOracle(o, Op.ShiftRows(at, by)).values.toList) :| "rows" &&
      (s.shiftCols(at, by).iterator.toList ==
        applyOracle(o, Op.ShiftCols(at, by)).values.toList) :| "cols" &&
      (s.shiftRows(at, 0) eq s) :| "zero shift" &&
      (s.shiftRows(at, by).checkInvariants == Right(())) :| "rows invariants" &&
      (s.shiftCols(at, by).checkInvariants == Right(())) :| "cols invariants"
    }
  }

  property("equals holds exactly when the oracles are equal, whatever the build history") {
    forAll(
      Gen.listOf(genCell),
      Gen.listOf(genCell),
      Gen.oneOf(true, false),
      Gen.choose(0L, 1000L)
    ) { (a, b0, same, seed) =>
      // half the time b is a reshuffle of a plus redundant writes, so equal stores are common
      val b =
        if same then new scala.util.Random(seed).shuffle(oracleOf(a).values.toList) else b0
      val viaFold = b.foldLeft(CellStore.empty)(_.updated(_))
      val viaRemovals = (b ++ a)
        .foldLeft(CellStore.empty)(_.updated(_))
        .removedIn(
          CellRange(ARef.from0(Int.MinValue, Int.MinValue), ARef.from0(Int.MaxValue, Int.MaxValue))
        )
      val sa = CellStore.from(a)
      val eqOracle = oracleOf(a) == oracleOf(b)
      ((sa == viaFold) == eqOracle) :| "equals iff oracles equal" &&
      ((viaFold == sa) == eqOracle) :| "symmetric" &&
      (!eqOracle || sa.hashCode == viaFold.hashCode) :| "hashCode consistent" &&
      (viaRemovals == CellStore.empty) :| "removing everything gives empty"
    }
  }

  property("canonical shape: height is minimal and shrinks back after removals") {
    forAll(Gen.listOf(genCell)) { cells =>
      val s = CellStore.from(cells)
      val o = oracleOf(cells)
      val removedOneByOne = o.keys.foldLeft(s)((acc, k) => acc.removed(ARef.from0(k._2, k._1)))
      val maxKey = o.keys.map(k => Integer.toUnsignedLong(k._1)).maxOption
      val expectedHeight = maxKey.fold(0) { k =>
        Iterator.from(1).find(h => h >= 7 || (k >>> (5 * h)) == 0L).getOrElse(7)
      }
      (removedOneByOne == CellStore.empty) :| "empty after removing each" &&
      (removedOneByOne.debugHeight == 0) :| "height 0" &&
      (s.debugHeight == expectedHeight) :| s"height ${s.debugHeight} vs $expectedHeight" &&
      (s.removed(ARef.from0(123456, -77)) eq s) :| "removing an absent ref returns this"
    }
  }

  property("dense single row and single column shapes") {
    forAll(Gen.choose(1, 700), Gen.choose(-100, 100), Gen.oneOf(true, false)) {
      (n, offset, rowShape) =>
        val cells = (0 until n).map { i =>
          val ref = if rowShape then ARef.from0(offset + i, 3) else ARef.from0(5, offset + i)
          Cell(ref, CellValue.Number(BigDecimal(i)), if i % 3 == 0 then Some(StyleId(i)) else None)
        }
        val s = CellStore.from(cells.reverse)
        val o = oracleOf(cells)
        val mid = cells(n / 2).ref
        val col = CellCursor.column(s, 5, Int.MinValue, Int.MaxValue)
        val colWant = o.values.filter(_.ref.col.index0 == 5).toList
        (s.iterator.toList == o.values.toList) :| "order" &&
        (s.checkInvariants == Right(())) :| "invariants" &&
        (cursorCells(col) == colWant) :| "column cursor" &&
        (s.removed(mid).updated(cells(n / 2)) == s) :| "remove then restore" &&
        (s.shiftCols(offset + n / 2, -(n / 4) - 1).checkInvariants == Right(())) :| "shift"
    }
  }

  test("tall narrow data and narrow tables are stored as 32-row blocks") {
    import com.tjclp.xl.sheets.store.{BandedRow, Block, Entry, Node, RowLeaf}
    def shapes(e: Entry): (Int, Int) = e match
      case b: Block => (1, 0)
      case l: RowLeaf => (0, 1)
      case r: BandedRow => (0, 1)
      case n: Node => n.children.map(shapes).foldLeft((0, 0))((a, b) => (a._1 + b._1, a._2 + b._2))
    val column =
      CellStore.from((0 until 1000).map(r => Cell(ARef.from0(0, r), CellValue.Number(r))))
    val table = CellStore.from(
      for r <- 0 until 320; c <- 0 until 20 yield Cell(ARef.from0(c, r), CellValue.Number(c))
    )
    val wide = CellStore.from(
      for r <- 0 until 64; c <- 0 until 100 yield Cell(ARef.from0(c, r), CellValue.Number(c))
    )
    assertEquals(shapes(column.root), (32, 0)) // ceil(1000 / 32) blocks, no per-row leaves
    assertEquals(shapes(table.root), (0, 320)) // 20 columns: over the block cap, per-row leaves
    val narrow = CellStore.from(
      for r <- 0 until 320; c <- 0 until 8 yield Cell(ARef.from0(c, r), CellValue.Number(c))
    )
    assertEquals(shapes(narrow.root), (10, 0))
    // one more column in one row tips its group over the cap, and removing it tips it back
    val over = narrow.updated(Cell(ARef.from0(8, 40), CellValue.Number(1)))
    assertEquals(shapes(over.root), (9, 32))
    assertEquals(over.removed(ARef.from0(8, 40)), narrow)
    assertEquals(shapes(wide.root), (0, 64)) // rows wider than one chunk stay per-row leaves
    // one row breaking out of the chunk turns its group into leaves, and back on removal
    val broken = column.updated(Cell(ARef.from0(64, 5), CellValue.Text("x")))
    assertEquals(shapes(broken.root), (31, 32))
    assertEquals(broken.removed(ARef.from0(64, 5)), column)
    assertEquals(shapes(broken.removed(ARef.from0(64, 5)).root), (32, 0))
  }

  // ----- examples documenting the order and edge semantics -----

  test("iteration is row-major: unsigned rows (negative rows last), signed columns") {
    val refs = List(
      ARef.from0(0, -1),
      ARef.from0(5, Int.MinValue),
      ARef.from0(-3, 0),
      ARef.from0(2, 0),
      ARef.from0(70, 0),
      ARef.from0(0, 1048576),
      ARef.from0(Int.MaxValue, 1),
      ARef.from0(Int.MinValue, 1)
    )
    val s = CellStore.from(refs.map(r => Cell(r, CellValue.Text("x"))))
    assertEquals(
      s.refs.map(r => (r.row.index0, r.col.index0)).toList,
      List(
        (0, -3),
        (0, 2),
        (0, 70),
        (1, Int.MinValue),
        (1, Int.MaxValue),
        (1048576, 0),
        (Int.MinValue, 5),
        (-1, 0)
      )
    )
    // bounds are signed, like Sheet.usedRange: the negative row is the minimum
    assertEquals(
      s.usedRange,
      Some(CellRange(ARef.from0(Int.MinValue, Int.MinValue), ARef.from0(Int.MaxValue, 1048576)))
    )
    assertEquals(s.debugHeight, 7)
    assertEquals(
      s.removedIn(CellRange(ARef.from0(Int.MinValue, -5), ARef.from0(10, -1))).debugHeight,
      7
    )
    val lowRowsOnly = s
      .removedIn(CellRange(ARef.from0(Int.MinValue, 2), ARef.from0(Int.MaxValue, Int.MaxValue)))
      .removedIn(CellRange(ARef.from0(Int.MinValue, Int.MinValue), ARef.from0(Int.MaxValue, -1)))
    assertEquals(lowRowsOnly.debugHeight, 1)
    assertEquals(lowRowsOnly.size, 5)
  }

  test("style-only Empty cells are stored, count toward extent but not usedRange") {
    val s = CellStore.of(
      Cell(ARef.from0(0, 0), CellValue.Number(BigDecimal(1))),
      Cell(ARef.from0(9, 9), CellValue.Empty, Some(StyleId(-1)))
    )
    assertEquals(s.size, 2)
    assertEquals(s.valuedCount, 1)
    assertEquals(s.styleAt(ARef.from0(9, 9)), Some(StyleId(-1)))
    assertEquals(s.usedRange, Some(CellRange(ARef.from0(0, 0), ARef.from0(0, 0))))
    assertEquals(s.extent, Some(CellRange(ARef.from0(0, 0), ARef.from0(9, 9))))
  }

  test("updatedValue keeps style and link; on an absent ref it stores a plain (even Empty) cell") {
    val ref = ARef.from0(3, 4)
    val s = CellStore.of(Cell(ref, CellValue.Text("a"), Some(StyleId(Int.MinValue)), Some("")))
    assertEquals(
      s.updatedValue(ref, CellValue.Bool(true)).get(ref),
      Some(Cell(ref, CellValue.Bool(true), Some(StyleId(Int.MinValue)), Some("")))
    )
    assertEquals(
      CellStore.empty.updatedValue(ref, CellValue.Empty).get(ref),
      Some(Cell(ref, CellValue.Empty))
    )
  }

  test("value tags") {
    val dt: FormulaKind.DataTable =
      FormulaKind.DataTable(CellRange(ARef.from0(0, 0), ARef.from0(1, 1)), false, false, None, None)
    val tags = List(
      CellValue.Empty -> CellStore.TagEmpty,
      CellValue.Text("t") -> CellStore.TagText,
      CellValue.RichText(RichText.plain("r")) -> CellStore.TagRichText,
      CellValue.Number(BigDecimal(1)) -> CellStore.TagNumber,
      CellValue.Bool(false) -> CellStore.TagBool,
      CellValue.DateTime(LocalDateTime.of(2020, 1, 1, 0, 0)) -> CellStore.TagDateTime,
      CellValue.Formula("A1", None) -> CellStore.TagFormulaUncached,
      CellValue.Formula("A1", Some(CellValue.Number(BigDecimal(2)))) -> CellStore.TagFormulaCached,
      CellValue.dataTable(dt) -> CellStore.TagFormulaDataTable,
      CellValue.Error(CellError.NA) -> CellStore.TagError
    )
    tags.foreach((v, t) => assertEquals(CellStore.tagOf(v), t, v.toString))
    val s = CellStore.of(Cell(ARef.from0(0, 0), CellValue.Formula("A1", None)))
    assert(s.hasTag(CellStore.FormulaTags))
    assert(!s.hasTag(CellStore.tagBit(CellStore.TagNumber)))
  }

  test("toString is row-major and readable") {
    val s = CellStore.of(
      Cell(ARef.from0(1, 0), CellValue.Text("b"), Some(StyleId(2))),
      Cell(ARef.from0(0, 0), CellValue.Text("a"), None, Some("https://x")),
      Cell(ARef.from0(0, -1), CellValue.Empty)
    )
    assertEquals(
      s.toString,
      "CellStore(A1 -> Text(a) link=https://x, B1 -> Text(b) style=2, R-1C0 -> Empty)"
    )
  }

object CellStoreLawSpec:

  /** The store's order: rows as unsigned 32-bit numbers, then columns as signed Ints. */
  val StoreOrder: Ordering[(Int, Int)] = (a, b) =>
    val r = Integer.compareUnsigned(a._1, b._1)
    if r != 0 then r else Integer.compare(a._2, b._2)

  enum Op:
    case Update(cell: Cell)
    case UpdateValue(ref: ARef, value: CellValue)
    case Remove(ref: ARef)
    case RemoveIn(range: CellRange)
    case PutAll(cells: List[Cell])
    case WithValues(pairs: List[(ARef, CellValue)])
    case Transform(mask: Int, f: CellValue => CellValue)
    case ShiftRows(at: Int, by: Int)
    case ShiftCols(at: Int, by: Int)

  // Small pools of rows and columns at every boundary the layout cares about, so random refs
  // collide often (exercising replace/remove) and every trie height and chunk edge is reached.
  val rowPool: Vector[Int] = Vector(
    0,
    1,
    2,
    31,
    32,
    33,
    1023,
    1024,
    32767,
    32768,
    100000,
    1048575,
    1048576,
    Int.MaxValue,
    -1,
    -2,
    Int.MinValue
  )
  val colPool: Vector[Int] =
    Vector(
      0,
      1,
      62,
      63,
      64,
      65,
      127,
      128,
      129,
      500,
      16383,
      16384,
      -1,
      -64,
      -65,
      Int.MinValue,
      Int.MaxValue
    )

  val genRow: Gen[Int] = Gen.frequency(
    6 -> Gen.oneOf(rowPool),
    3 -> Gen.choose(0, 40),
    1 -> Gen.choose(Int.MinValue, Int.MaxValue)
  )
  val genCol: Gen[Int] = Gen.frequency(
    6 -> Gen.oneOf(colPool),
    3 -> Gen.choose(0, 200),
    1 -> Gen.choose(Int.MinValue, Int.MaxValue)
  )
  val genRef: Gen[ARef] = Gen.frequency(
    4 -> (for r <- genRow; c <- genCol yield ARef.from0(c, r)),
    // tall narrow data: lands in the multi-row blocks of the narrow start shapes
    1 -> (for r <- Gen.choose(0, 70); c <- Gen.choose(0, 9) yield ARef.from0(c, r))
  )

  val genRange: Gen[CellRange] = Gen.frequency(
    4 -> (for a <- genRef; b <- genRef yield CellRange(a, b)),
    1 -> Gen.const(
      CellRange(ARef.from0(Int.MinValue, Int.MinValue), ARef.from0(Int.MaxValue, Int.MaxValue))
    ),
    1 -> Gen.const(CellRange.empty)
  )

  private val dataTable: FormulaKind.DataTable =
    FormulaKind.DataTable(CellRange(ARef.from0(0, 0), ARef.from0(1, 1)), true, false, None, None)

  val genValue: Gen[CellValue] = Gen.frequency(
    3 -> Gen.const(CellValue.Empty),
    3 -> Gen.alphaNumStr.map(s => CellValue.Text(s.take(4))),
    3 -> Gen.choose(-1000, 1000).map(i => CellValue.Number(BigDecimal(i))),
    1 -> Gen.oneOf(true, false).map(CellValue.Bool.apply),
    1 -> Gen.const(CellValue.RichText(RichText.plain("rt"))),
    1 -> Gen.const(CellValue.DateTime(LocalDateTime.of(2024, 2, 29, 12, 0))),
    1 -> Gen.const(CellValue.Formula("SUM(A1:A2)", None)),
    1 -> Gen.const(CellValue.Formula("A1*2", Some(CellValue.Number(BigDecimal(4))))),
    1 -> Gen.const(CellValue.dataTable(dataTable)),
    1 -> Gen.const(CellValue.Error(CellError.Div0))
  )

  val genStyle: Gen[Option[StyleId]] = Gen.frequency(
    3 -> Gen.const(None),
    1 -> Gen.oneOf(0, 1, -1, Int.MinValue, Int.MaxValue, 0x7fff).map(i => Some(StyleId(i)))
  )

  val genLink: Gen[Option[String]] = Gen.frequency(
    4 -> Gen.const(None),
    1 -> Gen.oneOf("", "https://example.com/x", "Sheet1!A1").map(Some(_))
  )

  val genCell: Gen[Cell] =
    for ref <- genRef; v <- genValue; s <- genStyle; l <- genLink yield Cell(ref, v, s, l)

  /** A starting store's cells: random, or one of the dense shapes. */
  val genStart: Gen[List[Cell]] = Gen.frequency(
    4 -> Gen.listOf(genCell),
    1 -> Gen
      .choose(1, 300)
      .map(n => List.tabulate(n)(i => Cell(ARef.from0(i, 7), CellValue.Number(BigDecimal(i))))),
    1 -> Gen
      .choose(1, 300)
      .map(n => List.tabulate(n)(i => Cell(ARef.from0(3, i * 7), CellValue.Text("c")))),
    // a tall dense column: uniform blocks
    1 -> Gen
      .choose(1, 200)
      .map(n => List.tabulate(n)(i => Cell(ARef.from0(0, i), CellValue.Number(BigDecimal(i))))),
    // tall narrow data with holes, styles and links: per-row masks in one block
    1 -> (for
      n <- Gen.choose(1, 120)
      holes <- Gen.listOfN(n * 5, Gen.frequency(4 -> true, 1 -> false))
      style <- genStyle
      link <- genLink
    yield List.tabulate(n * 5)(k => (k / 5, k % 5)).zip(holes).collect { case ((r, c), true) =>
      Cell(
        ARef.from0(c, r),
        if c == 2 then CellValue.Empty else CellValue.Text(s"$r"),
        if r % 3 == 0 then style else None,
        if c == 4 then link else None
      )
    }),
    // groups at the block size cap: 32 rows of 8 columns (a block) or 9 (per-row leaves)
    1 -> Gen
      .oneOf(8, 9)
      .map(w =>
        List.tabulate(32 * w)(k => Cell(ARef.from0(k % w, k / w), CellValue.Number(BigDecimal(k))))
      ),
    // rows of differing widths sharing a block, one of them crossing into a second chunk
    1 -> Gen
      .choose(0, 31)
      .map(wide =>
        (0 until 32).toList.flatMap { r =>
          val width = if r == wide then 70 else r % 7 + 1
          (0 until width).map(c => Cell(ARef.from0(c, r), CellValue.Bool(c % 2 == 0)))
        }
      )
  )

  val genTagMask: Gen[Int] = Gen.oneOf(
    Gen.const(CellStore.AllTags),
    Gen.const(CellStore.FormulaTags),
    Gen.choose(0, 9).map(CellStore.tagBit),
    Gen.choose(0, CellStore.AllTags)
  )

  /** Value transforms with a label (the label keeps failure output readable). */
  private val inc: CellValue => CellValue =
    case CellValue.Number(n) => CellValue.Number(n + 1)
    case v => v
  private val uncache: CellValue => CellValue =
    case CellValue.Formula(e, _, k) => CellValue.Formula(e, None, k)
    case v => v
  private val toEmpty: CellValue => CellValue = _ => CellValue.Empty
  private val toText: CellValue => CellValue = _ => CellValue.Text("t")

  val genFn: Gen[(String, CellValue => CellValue)] =
    Gen.oneOf("inc" -> inc, "empty" -> toEmpty, "text" -> toText, "uncache" -> uncache)

  // ----- one dense block: rows 0..35 (one full slot group and the start of the next), 8 columns -----

  val genBlockRef: Gen[ARef] =
    for r <- Gen.choose(0, 35); c <- Gen.choose(0, 7) yield ARef.from0(c, r)

  val genBlockRange: Gen[CellRange] =
    for a <- genBlockRef; b <- genBlockRef yield CellRange(a, b)

  /** Mostly Empty, so valued bounds and tag bits come and go; every tag still appears. */
  val genBlockValue: Gen[CellValue] = Gen.frequency(
    6 -> Gen.const(CellValue.Empty),
    6 -> genValue
  )

  val genBlockCell: Gen[Cell] =
    for
      ref <- genBlockRef
      v <- genBlockValue
      s <- Gen.frequency(5 -> Gen.const(None), 1 -> Gen.const(Some(StyleId(1))))
      l <- Gen.frequency(9 -> Gen.const(None), 1 -> Gen.const(Some("Sheet1!A1")))
    yield Cell(ref, v, s, l)

  /** Every cell of the first `rows` rows over `cols` columns, each valued or style-only Empty. */
  val genBlockStart: Gen[List[Cell]] =
    for
      rows <- Gen.choose(1, 34)
      cols <- Gen.choose(1, 8)
      cells <- Gen.sequence[List[Cell], Cell](
        for r <- 0 until rows; c <- 0 until cols
        yield genBlockValue.map(v => Cell(ARef.from0(c, r), v, Some(StyleId(1)), None))
      )
    yield cells

  val genBlockOp: Gen[Op] = Gen.frequency(
    5 -> genBlockCell.map(Op.Update.apply),
    5 -> (for r <- genBlockRef; v <- genBlockValue yield Op.UpdateValue(r, v)),
    5 -> genBlockRef.map(Op.Remove.apply),
    1 -> genBlockRange.map(Op.RemoveIn.apply),
    1 -> Gen.listOfN(6, genBlockCell).map(Op.PutAll.apply),
    1 -> Gen
      .listOfN(6, genBlockRef.flatMap(r => genBlockValue.map(r -> _)))
      .map(Op.WithValues.apply),
    1 -> (for m <- genTagMask; f <- genFn yield Op.Transform(m, f._2))
  )

  val genIndex: Gen[Int] = Gen.oneOf(genRow, genCol, Gen.choose(-5, 70))
  val genBy: Gen[Int] =
    Gen.oneOf(-3, -1, 1, 2, 64, -64, -65, 1000000, Int.MaxValue, Int.MinValue + 1, Int.MinValue, 0)

  val genOp: Gen[Op] = Gen.frequency(
    6 -> genCell.map(Op.Update.apply),
    2 -> (for r <- genRef; v <- genValue yield Op.UpdateValue(r, v)),
    3 -> genRef.map(Op.Remove.apply),
    1 -> genRange.map(Op.RemoveIn.apply),
    1 -> Gen.listOf(genCell).map(cs => Op.PutAll(cs.take(12))),
    1 -> Gen
      .listOf(genRef.flatMap(r => genValue.map(r -> _)))
      .map(ps => Op.WithValues(ps.take(12))),
    1 -> (for m <- genTagMask; f <- genFn yield Op.Transform(m, f._2)),
    1 -> (for a <- genIndex; b <- genBy yield Op.ShiftRows(a, b)),
    1 -> (for a <- genIndex; b <- genBy yield Op.ShiftCols(a, b))
  )
