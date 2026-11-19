package com.tjclp.xl.sheets

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.sheets.store.{Band, BandedRow, Entry, Node, RowLeaf}
import com.tjclp.xl.styles.units.StyleId
import munit.ScalaCheckSuite
import org.scalacheck.{Gen, Prop}
import org.scalacheck.Prop.*

/**
 * Laws for the rows of a [[CellStore]] that span several column bands (#891), against the same
 * `SortedMap` oracle and step-by-step agreement as [[CellStoreLawSpec]], whose pools seldom put two
 * populated bands in one row and almost never empty one.
 *
 * The domain here is a handful of rows a few hundred to a few thousand columns wide, with every
 * column drawn from the edges of the bands around zero (negative band ids included) and the ends of
 * the `Int` range. Writes and removals create and destroy bands, take a band's extreme valued
 * column or last tag away, collapse a banded row back to a single leaf and to nothing, and shift
 * columns across band edges.
 */
class CellStoreWideRowLawSpec extends ScalaCheckSuite:
  import CellStoreLawSpec.{Op, genBlockValue, genFn, genTagMask}
  import CellStoreWideRowLawSpec.*

  override def scalaCheckTestParameters =
    super.scalaCheckTestParameters.withMinSuccessfulTests(200).withMaxSize(40)

  /**
   * The oracle laws of [[CellStoreLawSpec]], reused as they are (its tests are not re-run here).
   */
  private val laws = new CellStoreLawSpec

  property("dense and holed multi-band rows agree with the oracle, step by step") {
    forAll(
      genWideStart,
      Gen.choose(10, 30).flatMap(Gen.listOfN(_, genWideOp)),
      Gen.listOfN(4, genWideRef),
      Gen.listOfN(3, genWideRange)
    )(laws.sequenceAgrees)
  }

  property("a few cells per band: bands come and go with every write and removal") {
    forAll(
      Gen.choose(0, 40).flatMap(Gen.listOfN(_, genSparseCell)),
      Gen.choose(30, 80).flatMap(Gen.listOfN(_, genSparseOp)),
      Gen.listOfN(4, genSparseRef),
      Gen.listOfN(3, genWideRange)
    )(laws.sequenceAgrees)
  }

  property("a multi-band row emptied cell by cell stays canonical down to nothing, and back") {
    forAll(genWideRow(3), Gen.choose(0L, Long.MaxValue)) { (cells, seed) =>
      val rnd = new scala.util.Random(seed)
      val full = CellStore.from(cells)
      val removal = rnd.shuffle(cells)
      // every prefix of the removal order, checked against a store built from what is left
      val (emptied, down) =
        removal.tails.toList.drop(1).zip(removal).foldLeft((full, List.empty[Prop])) {
          case ((s, acc), (left, gone)) =>
            val next = s.removed(gone.ref)
            val want = CellStore.from(rnd.shuffle(left))
            val ok = (next.checkInvariants == Right(())) :| s"invariants with ${left.size} left" &&
              (next == want) :| s"canonical with ${left.size} left" &&
              (next.usedRange == want.usedRange && next.extent == want.extent) :| "bounds" &&
              (next.valuedCount == left.count(_.nonEmpty)) :| "valued"
            (next, ok :: acc)
        }
      val restored = rnd.shuffle(cells).foldLeft(emptied)(_.updated(_))
      Prop.all(down*) &&
      (emptied == CellStore.empty) :| "emptied" &&
      (restored == full && restored.hashCode == full.hashCode) :| "restored in another order" &&
      (restored.checkInvariants == Right(())) :| "restored invariants"
    }
  }

  property("removing whole bands leaves the row a store built from the rest would hold") {
    forAll(genWideRow(0), Gen.listOfN(6, genBandRange)) { (cells, ranges) =>
      val (_, _, props) = ranges.foldLeft((CellStore.from(cells), cells, List.empty[Prop])) {
        case ((s, left, acc), r) =>
          val next = s.removedIn(r)
          val rest = left.filterNot(c => r.contains(c.ref))
          val ok = (next == CellStore.from(rest)) :| s"removedIn $r" &&
            (next.checkInvariants == Right(())) :| s"invariants after $r" &&
            (next.iterator.toList == rest) :| s"cells after $r"
          (next, rest, ok :: acc)
      }
      Prop.all(props*)
    }
  }

  property("shiftCols across band edges equals re-keying, at every edge and distance") {
    forAll(genWideRow(1), Gen.oneOf(edgeCols), genColShift) { (cells, at, by) =>
      val s = CellStore.from(cells)
      val shifted = s.shiftCols(at, by)
      val want = cells.flatMap { c =>
        laws
          .destOf(c.ref.col.index0, at, by)
          .map(d => c.copy(ref = ARef.from0(d, c.ref.row.index0)))
      }
      (shifted.iterator.toList == want) :| s"cells after shiftCols($at, $by)" &&
      (shifted == CellStore.from(want)) :| "canonical" &&
      (shifted.checkInvariants == Right(())) :| s"invariants: ${shifted.checkInvariants}"
    }
  }

  property("union paths (putAll, withValues) over multi-band rows equal the folds") {
    forAll(genWideRow(0), Gen.listOfN(40, genSparseCell), genWideRow(0)) { (cells, more, other) =>
      val s = CellStore.from(cells)
      val all = more ++ other
      val put = s.putAll(all)
      val vals = s.withValues(IArray.from(all.map(_.ref)), IArray.from(all.map(_.value)))
      (put == all.foldLeft(s)(_.updated(_))) :| "putAll" &&
      (vals == all.foldLeft(s)((acc, c) => acc.updatedValue(c.ref, c.value))) :| "withValues" &&
      (put.checkInvariants == Right(()) && vals.checkInvariants == Right(())) :| "invariants" &&
      (s.putAll(cells) == s) :| "putting a row's own cells changes nothing"
    }
  }

  property("transformValues over multi-band rows keeps bands and shares untouched stores") {
    forAll(genWideRow(0), genTagMask, genFn) { (cells, mask, fn) =>
      val s = CellStore.from(cells)
      val t = s.transformValues(mask)(fn._2)
      val want = cells.map { c =>
        if (mask & CellStore.tagBit(CellStore.tagOf(c.value))) != 0 then c.withValue(fn._2(c.value))
        else c
      }
      (t == CellStore.from(want)) :| s"transform ${fn._1}" &&
      (t.checkInvariants == Right(())) :| "invariants" &&
      (s.transformValues(mask)(identity) eq s) :| "identity transform shares the store"
    }
  }

  test("a row is banded exactly when its columns span two bands, whatever its history") {
    def rows(e: Entry): List[Either[RowLeaf, BandedRow]] = e match
      case l: RowLeaf => List(Left(l))
      case r: BandedRow => List(Right(r))
      case n: Node => n.children.toList.flatMap(rows)
      case _ => Nil
    def bandsOf(s: CellStore): List[Int] = rows(s.root).map(_.fold(_ => 1, _.bands.length))
    def cell(col: Int) = Cell(ARef.from0(col, 0), CellValue.Number(BigDecimal(col)))
    // 20 rows over the block cap, so every row is its own entry
    val filler = for r <- 1 until 20; c <- 0 until 20 yield Cell(ARef.from0(c, r), CellValue.Empty)
    val oneBand = CellStore.from(filler ++ (0 until W).map(cell))
    assertEquals(bandsOf(oneBand), List.fill(20)(1), "a full band is one leaf")
    val two = oneBand.updated(cell(W))
    assertEquals(bandsOf(two).take(1), List(2), "the first column of the next band adds a band")
    assertEquals(two.removed(ARef.from0(W, 0)), oneBand)
    assertEquals(bandsOf(two.removed(ARef.from0(W, 0))), List.fill(20)(1), "and collapses back")
    val negative = oneBand.updatedValue(ARef.from0(-1, 0), CellValue.Text("n"))
    assertEquals(bandsOf(negative).take(1), List(2), "column -1 is band -1, not band 0")
    assertEquals(negative.removedIn(CellRange(ARef.from0(0, 0), ARef.from0(W - 1, 0))).size, 381)
    val sparse = CellStore.from(filler ++ List(cell(Int.MinValue), cell(0), cell(Int.MaxValue)))
    assertEquals(bandsOf(sparse).take(1), List(3), "bands at the ends of the Int range")
    val wide = CellStore.from(filler ++ (0 until 16384).map(cell))
    assertEquals(bandsOf(wide).take(1), List(16384 / W))
    // emptying all but one band, by any route, leaves that band's leaf
    val lastBand = CellRange(ARef.from0(0, 0), ARef.from0(16383 - W, 0))
    val viaRange = wide.removedIn(lastBand)
    val viaCells = (0 until 16384 - W).foldLeft(wide)((s, c) => s.removed(ARef.from0(c, 0)))
    val direct = CellStore.from(filler ++ (16384 - W until 16384).map(cell))
    assertEquals(bandsOf(direct), List.fill(20)(1))
    List(viaRange, viaCells).foreach { s =>
      assert(s == direct, "one band left")
      assertEquals(bandsOf(s), List.fill(20)(1))
      assertEquals(s.checkInvariants, Right(()))
    }
    // a column delete that leaves one band, then moves it back across every band edge
    val row = CellStore.from((0 until 16384).map(cell))
    val viaShift = row.shiftCols(16384 - W, -(16384 - W)).shiftCols(0, 16384 - W)
    assert(viaShift == CellStore.from((16384 - W until 16384).map(cell)), "shifted")
    assertEquals(viaShift.checkInvariants, Right(()))
  }

  test("writes that change nothing return the same store") {
    val cells =
      (0 until 3 * W).map(c => Cell(ARef.from0(c - W, 0), CellValue.Number(BigDecimal(c))))
    val s = CellStore.from(cells)
    assert(s.removed(ARef.from0(5 * W, 0)) eq s, "absent band")
    assert(s.removed(ARef.from0(-W - 1, 0)) eq s, "absent column next to a band")
    assert(
      s.removedIn(CellRange(ARef.from0(2 * W, 0), ARef.from0(9 * W, 0))) eq s,
      "range past the row"
    )
    assert(s.shiftCols(2 * W, 7) eq s, "shift past the row")
  }

object CellStoreWideRowLawSpec:
  import CellStoreLawSpec.{Op, genBlockValue, genFn, genTagMask}

  /** The band width under test (columns per band). */
  val W: Int = Band.Width

  private val bandIds: Vector[Int] = Vector(-2, -1, 0, 1, 2, 3, 4)

  /** Both sides of every band edge around zero, and of the chunk edge inside each band. */
  val edgeCols: Vector[Int] =
    (for k <- bandIds; d <- Vector(-1, 0, 1, 63, 64, W / 2) yield k * W + d).distinct

  /** Band ids at the ends of the `Int` range, and the grid's last column. */
  val farCols: Vector[Int] = Vector(
    Int.MinValue,
    Int.MinValue + 1,
    Int.MinValue + W,
    Int.MaxValue,
    Int.MaxValue - 1,
    Int.MaxValue - W,
    16383,
    16384
  )

  val genWideCol: Gen[Int] = Gen.frequency(
    6 -> Gen.oneOf(edgeCols),
    4 -> Gen.choose(-2 * W, 5 * W),
    1 -> Gen.oneOf(farCols)
  )

  /** Rows 0 and 1 share a slot group and take most writes; the rest reach other trie paths. */
  val genWideRowIndex: Gen[Int] = Gen.frequency(
    6 -> Gen.const(0),
    2 -> Gen.const(1),
    1 -> Gen.oneOf(3, 33, 1048576, -1)
  )

  val genWideRef: Gen[ARef] =
    for r <- genWideRowIndex; c <- genWideCol yield ARef.from0(c, r)

  private val genStyle: Gen[Option[StyleId]] =
    Gen.frequency(5 -> Gen.const(None), 1 -> Gen.oneOf(1, -1).map(i => Some(StyleId(i))))
  private val genLink: Gen[Option[String]] =
    Gen.frequency(12 -> Gen.const(None), 1 -> Gen.const(Some("Sheet1!A1")))

  val genWideCell: Gen[Cell] =
    for ref <- genWideRef; v <- genBlockValue; s <- genStyle; l <- genLink yield Cell(ref, v, s, l)

  /**
   * One row's cells in column order: a run of columns starting next to a band edge, between half a
   * band and four bands wide, dense or holed, mostly Empty so valued bounds and tags come and go.
   */
  def genWideRow(row: Int): Gen[List[Cell]] =
    for
      first <- Gen.oneOf(-W - 2, -W, -1, 0, 1, W - 1, W, W + 3)
      width <- Gen.frequency(
        3 -> Gen.oneOf(W - 1, W, W + 1, 2 * W, 2 * W + 1, 3 * W + 5),
        2 -> Gen.choose(W / 2, 4 * W)
      )
      keep <- Gen.oneOf(1, 3, 10) // keeps every cell, or drops one column in 3 or in 10
      holes <- Gen.listOfN(width, Gen.choose(0, 29))
      values <- Gen.listOfN(width, genBlockValue)
      style <- genStyle
      link <- genLink
    yield holes.zip(values).zipWithIndex.collect {
      case ((h, v), i) if keep == 1 || h % keep != 0 =>
        Cell(
          ARef.from0(first + i, row),
          v,
          if i % 5 == 0 then style else None,
          if i % 97 == 0 then link else None
        )
    }

  /** One or two wide rows (in one slot group), sometimes with a narrow neighbour. */
  val genWideStart: Gen[List[Cell]] =
    for
      a <- genWideRow(0)
      b <- Gen.frequency(2 -> Gen.const(Nil), 1 -> genWideRow(1))
      narrow <- Gen.oneOf(true, false)
    yield a ++ b ++ (if narrow then List(Cell(ARef.from0(2, 3), CellValue.Text("n"))) else Nil)

  /** A window whose column bounds sit on band edges (rows: one row, or every row). */
  val genBandRange: Gen[CellRange] =
    for
      c0 <- Gen.oneOf(edgeCols)
      c1 <- Gen.frequency(4 -> Gen.oneOf(edgeCols), 1 -> Gen.oneOf(Int.MinValue, Int.MaxValue))
      allRows <- Gen.oneOf(true, false)
      row <- genWideRowIndex
    yield
      val lo = math.min(c0, c1)
      val hi = math.max(c0, c1)
      if allRows then CellRange(ARef.from0(lo, 0), ARef.from0(hi, Int.MaxValue))
      else CellRange(ARef.from0(lo, row), ARef.from0(hi, row))

  val genWideRange: Gen[CellRange] = Gen.frequency(
    3 -> genBandRange,
    1 -> (for a <- genWideRef; b <- genWideRef yield CellRange(a, b))
  )

  /** Shift distances of one column, one chunk, one band and one band give or take a column. */
  val genColShift: Gen[Int] = Gen
    .oneOf(1, 2, 63, 64, 65, W - 1, W, W + 1, 3 * W, Int.MaxValue)
    .flatMap(d => Gen.oneOf(d, -d))

  val genWideOp: Gen[Op] = Gen.frequency(
    5 -> genWideCell.map(Op.Update.apply),
    4 -> (for r <- genWideRef; v <- genBlockValue yield Op.UpdateValue(r, v)),
    6 -> genWideRef.map(Op.Remove.apply),
    3 -> genBandRange.map(Op.RemoveIn.apply),
    1 -> Gen.listOfN(8, genWideCell).map(Op.PutAll.apply),
    1 -> Gen
      .listOfN(8, genWideRef.flatMap(r => genBlockValue.map(r -> _)))
      .map(Op.WithValues.apply),
    1 -> (for m <- genTagMask; f <- genFn yield Op.Transform(m, f._2)),
    2 -> (for at <- Gen.oneOf(edgeCols); by <- genColShift yield Op.ShiftCols(at, by)),
    1 -> (for at <- Gen.oneOf(0, 1, 2); by <- Gen.oneOf(1, -1, 32) yield Op.ShiftRows(at, by))
  )

  // ----- a few cells per band: every column an edge column, two rows of one slot group -----

  val genSparseRef: Gen[ARef] =
    for
      r <- Gen.frequency(5 -> Gen.const(0), 1 -> Gen.const(1))
      c <- Gen.frequency(12 -> Gen.oneOf(edgeCols), 1 -> Gen.oneOf(farCols))
    yield ARef.from0(c, r)

  val genSparseCell: Gen[Cell] =
    for ref <- genSparseRef; v <- genBlockValue; s <- genStyle; l <- genLink
    yield Cell(ref, v, s, l)

  val genSparseOp: Gen[Op] = Gen.frequency(
    6 -> genSparseCell.map(Op.Update.apply),
    4 -> (for r <- genSparseRef; v <- genBlockValue yield Op.UpdateValue(r, v)),
    8 -> genSparseRef.map(Op.Remove.apply),
    1 -> genBandRange.map(Op.RemoveIn.apply),
    1 -> Gen.listOfN(5, genSparseCell).map(Op.PutAll.apply),
    1 -> (for m <- genTagMask; f <- genFn yield Op.Transform(m, f._2)),
    1 -> (for at <- Gen.oneOf(edgeCols); by <- genColShift yield Op.ShiftCols(at, by))
  )
