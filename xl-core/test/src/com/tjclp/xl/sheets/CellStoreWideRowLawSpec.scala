package com.tjclp.xl.sheets

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.{Cell, CellValue, Hyperlink}
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
 *
 * A row is banded only above [[Band.FlatMax]] cells, so a second domain sits on that count: rows
 * spanning several bands that hover one cell either side of it and cross it by every write path.
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
      Gen.choose(0, 160).flatMap(Gen.listOfN(_, genSparseCell)),
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

  property("rows hovering at the cell-count threshold convert both ways, by every write path") {
    forAll(
      genThresholdStart,
      Gen.choose(40, 80).flatMap(Gen.listOfN(_, genThresholdOp)),
      Gen.listOfN(4, genThresholdRef),
      Gen.listOfN(3, genWideRange)
    )(laws.sequenceAgrees)
  }

  property("crossing the threshold by one cell changes the form and nothing else") {
    forAll(Gen.oneOf(T - 1, T, T + 1), Gen.choose(0L, Long.MaxValue)) { (n, seed) =>
      val rnd = new scala.util.Random(seed)
      // n cells over three bands, then single-cell writes and removals walking the count
      // n -> T + 2 -> T - 2 -> n in a random column order
      val pool = rnd.shuffle(thresholdCols ++ toggleCols)
      def cell(c: Int) = Cell(ARef.from0(c, 0), CellValue.Number(BigDecimal(c)), Some(StyleId(1)))
      val start = pool.take(n)
      val up = pool.slice(n, T + 2)
      val down = rnd.shuffle(start ++ up).take(4)
      val steps: List[(String, CellStore => CellStore, Set[Int] => Set[Int])] =
        up.map(c => (s"put $c", (s: CellStore) => s.updated(cell(c)), (cs: Set[Int]) => cs + c)) ++
          down.map(c =>
            (s"remove $c", (s: CellStore) => s.removed(ARef.from0(c, 0)), (cs: Set[Int]) => cs - c)
          ) ++
          down.map(c =>
            (
              s"putValue $c",
              (s: CellStore) => s.updatedValue(ARef.from0(c, 0), CellValue.Text("v")),
              (cs: Set[Int]) => cs + c
            )
          )
      val s0 = CellStore.from(start.map(cell))
      val (_, _, props) =
        steps.foldLeft((s0, start.toSet, List(formOk(s0, start.toSet, "start")))) {
          case ((s, cols, acc), (label, f, g)) =>
            val next = f(s)
            val nextCols = g(cols)
            (next, nextCols, formOk(next, nextCols, label) :: acc)
        }
      Prop.all(props*)
    }
  }

  property("removedIn, putAll and withValues land on the right form at the threshold") {
    forAll(Gen.choose(T - 2, T + 3), Gen.choose(0, 5), Gen.choose(0L, Long.MaxValue)) {
      (n, k, seed) =>
        val rnd = new scala.util.Random(seed)
        val pool = rnd.shuffle(thresholdCols ++ toggleCols)
        def cell(c: Int) = Cell(ARef.from0(c, 0), CellValue.Number(BigDecimal(c)))
        val have = pool.take(n)
        val extra = pool.slice(n, n + k)
        val s = CellStore.from(have.map(cell))
        val put = s.putAll(extra.map(cell))
        val vals = s.withValues(
          IArray.from(extra.map(c => ARef.from0(c, 0))),
          IArray.from(extra.map(_ => CellValue.Empty: CellValue))
        )
        val sorted = (have ++ extra).sorted
        val cut = sorted.drop(k).headOption.getOrElse(0)
        val range = CellRange(ARef.from0(Int.MinValue, 0), ARef.from0(cut, 0))
        val removed = put.removedIn(range)
        val emptied = put.transformValues(CellStore.AllTags)(_ => CellValue.Empty)
        formOk(s, have.toSet, "start") &&
        formOk(put, sorted.toSet, s"putAll of $k") &&
        formOk(vals, sorted.toSet, s"withValues of $k") &&
        formOk(removed, sorted.filter(_ > cut).toSet, s"removedIn up to $cut") &&
        formOk(emptied, sorted.toSet, "transform to Empty keeps the count and the form") &&
        (emptied.valuedCount == 0) :| "emptied" &&
        (formsOf(emptied) == formsOf(put)) :| "transformValues does not convert"
    }
  }

  property("shiftCols moving a row into one band collapses it, and moving it out re-bands it") {
    forAll(Gen.choose(T - 2, T + 40), Gen.choose(1, T - 3)) { (n, before) =>
      // n consecutive columns with `before` of them below the edge of band 1
      val cells = (0 until n).map { i =>
        Cell(ARef.from0(W - before + i, 0), CellValue.Number(BigDecimal(i)), None)
      }.toList
      val cols = cells.map(_.ref.col.index0).toSet
      val s = CellStore.from(cells)
      val in = s.shiftCols(W - before, before) // every column moves up into band 1
      val out = in.shiftCols(W, -before) // and back: the vacated window is empty
      formOk(s, cols, "start") &&
      formOk(in, cols.map(_ + before), "shifted into one band") &&
      (formsOf(in) == List(1)) :| "one leaf" &&
      formOk(out, cols, "shifted back") &&
      (out == s && out.hashCode == s.hashCode) :| "round trip"
    }
  }

  test("a row is banded exactly when its columns span two bands, whatever its history") {
    def bandsOf(s: CellStore): List[Int] = formsOf(s)
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
    val ends = List(cell(Int.MinValue), cell(0), cell(Int.MaxValue))
    val sparse = CellStore.from(filler ++ ends)
    assertEquals(bandsOf(sparse).take(1), List(1), "a few cells over three bands stay one leaf")
    val few = (1 to T - 3).map(i => cell(i * 3))
    val atMax = CellStore.from(filler ++ ends ++ few)
    assertEquals(atMax.size - filler.size, T)
    assertEquals(bandsOf(atMax).take(1), List(1), "exactly FlatMax cells over bands: one leaf")
    val over = atMax.updated(cell(1))
    assertEquals(bandsOf(over).take(1), List(3), "one more: bands at the ends of the Int range")
    assertEquals(over.removed(ARef.from0(1, 0)), atMax)
    assertEquals(bandsOf(over.removed(ARef.from0(0, 0))).take(1), List(1), "and back to a leaf")
    val overEmpty = over.transformValues(CellStore.AllTags)(_ => CellValue.Empty)
    assertEquals(
      bandsOf(overEmpty).take(1),
      List(3),
      "emptying values keeps the cells, so the form"
    )
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
  private val genLink: Gen[Option[Hyperlink]] =
    Gen.frequency(12 -> Gen.const(None), 1 -> Gen.const(Some(Hyperlink("Sheet1!A1"))))

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

  // ----- the cell-count threshold: rows over three bands, one cell either side of FlatMax -----

  /** The most cells a row spanning several bands holds as one leaf. */
  val T: Int = Band.FlatMax

  /** `T - 2` columns spread evenly over bands 0 to 2. */
  val thresholdCols: List[Int] =
    val step = (2 * W + W / 2) / T + 1
    List.tabulate(T - 2)(i => 5 + i * step)

  /**
   * Columns outside [[thresholdCols]] that writes and removals toggle, so the count moves between
   * `T - 2` and `T + 4`: inside each populated band, either side of a band edge, in a band of their
   * own above the row and in a negative band below it.
   */
  val toggleCols: List[Int] = List(-3, 6, W - 1, W, 2 * W + 2, 4 * W + 7)

  /** The number of leaves each row of `s` is stored as (1 for a row that is a single leaf). */
  def formsOf(s: CellStore): List[Int] =
    def rows(e: Entry): List[Int] = e match
      case _: RowLeaf => List(1)
      case r: BandedRow => List(r.bands.length)
      case n: Node => n.children.toList.flatMap(rows)
      case _ => List(1) // a block's rows are single leaves; these domains build one row per group
    rows(s.root)

  /**
   * A one-row store holding exactly `cols` has the form the rule gives it, the invariants hold, and
   * it equals the store built from those columns' cells in one go.
   */
  def formOk(s: CellStore, cols: Set[Int], label: String): Prop =
    val bands = cols.map(_ >> Band.Shift).size
    val want = if cols.isEmpty then Nil else List(if cols.size > T && bands > 1 then bands else 1)
    val rebuilt = CellStore.from(s.iterator.toList.reverse)
    (s.checkInvariants == Right(())) :| s"$label: invariants ${s.checkInvariants}" &&
    (s.refs.map(_.col.index0).toSet == cols && s.size == cols.size) :| s"$label: columns" &&
    (formsOf(s) == want) :| s"$label: ${cols.size} cells stored as ${formsOf(s)}, want $want" &&
    (s == rebuilt && s.hashCode == rebuilt.hashCode) :| s"$label: canonical"

  val genThresholdRef: Gen[ARef] = Gen
    .frequency(8 -> Gen.oneOf(toggleCols), 2 -> Gen.oneOf(thresholdCols), 1 -> Gen.oneOf(edgeCols))
    .map(c => ARef.from0(c, 0))

  private val genToggleRef: Gen[ARef] = Gen.oneOf(toggleCols).map(c => ARef.from0(c, 0))

  val genThresholdCell: Gen[Cell] =
    for ref <- genToggleRef; v <- genBlockValue; s <- genStyle; l <- genLink
    yield Cell(ref, v, s, l)

  /** Every base column and a random half of the toggles: `T - 2` to `T + 4` cells. */
  val genThresholdStart: Gen[List[Cell]] =
    for
      on <- Gen.listOfN(toggleCols.size, Gen.oneOf(true, false))
      values <- Gen.listOfN(T + 4, genBlockValue)
    yield
      val cols = (thresholdCols ++ toggleCols.zip(on).collect { case (c, true) => c }).sorted
      cols.zip(values).map((c, v) => Cell(ARef.from0(c, 0), v))

  val genThresholdOp: Gen[Op] = Gen.frequency(
    6 -> genThresholdCell.map(Op.Update.apply),
    3 -> (for r <- genToggleRef; v <- genBlockValue yield Op.UpdateValue(r, v)),
    8 -> genToggleRef.map(Op.Remove.apply),
    1 -> (for r <- genThresholdRef; v <- genBlockValue yield Op.UpdateValue(r, v)),
    1 -> Gen.oneOf(thresholdCols).map(c => Op.Remove(ARef.from0(c, 0))),
    2 -> (for a <- Gen.oneOf(toggleCols); b <- Gen.oneOf(0, 1, W)
    yield Op.RemoveIn(CellRange(ARef.from0(a, 0), ARef.from0(a + b, 0)))),
    2 -> Gen.choose(1, 4).flatMap(Gen.listOfN(_, genThresholdCell)).map(Op.PutAll.apply),
    2 -> Gen
      .choose(1, 4)
      .flatMap(Gen.listOfN(_, genToggleRef.flatMap(r => genBlockValue.map(r -> _))))
      .map(Op.WithValues.apply),
    1 -> (for m <- genTagMask; f <- genFn yield Op.Transform(m, f._2)),
    1 -> (for at <- Gen.oneOf(edgeCols); by <- Gen.oneOf(1, -1, 2, -2) yield Op.ShiftCols(at, by))
  )

  // ----- a few cells per band: every column an edge column, two rows of one slot group -----

  /**
   * Fourteen columns in each band around zero: about as many cells as the threshold when half are
   * set.
   */
  val sparseCols: Vector[Int] =
    (for
      k <- bandIds
      d <- Vector(-2, -1, 0, 1, 2, 31, 63, 64, 65, 127, 128, 129, 191, 192)
    yield k * W + d).distinct

  val genSparseRef: Gen[ARef] =
    for
      r <- Gen.frequency(5 -> Gen.const(0), 1 -> Gen.const(1))
      c <- Gen.frequency(12 -> Gen.oneOf(sparseCols), 1 -> Gen.oneOf(farCols))
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
