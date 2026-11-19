package com.tjclp.xl.sheets

import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.sheets.store.{Band, Group}
import com.tjclp.xl.styles.units.StyleId
import munit.ScalaCheckSuite
import org.scalacheck.{Gen, Prop}
import org.scalacheck.Prop.*

/**
 * The seams of [[CellStore.Builder]] (#891), which streams the cells that arrive in row-major order
 * one slot group at a time and buffers the ones that arrive out of order: where a group closes,
 * where a group stops being a block, where a cell arrives before the last one streamed (and sorted
 * cells go on arriving after it), where a hyperlink first appears, and `result()` taken in the
 * middle of a row and built upon.
 *
 * The law is the one [[CellStoreLawSpec]] states for the builder (a build equals a fold of
 * `updated`), checked here after every prefix of the input, in arrival orders made to cross each
 * seam, together with the store's structural invariants.
 */
class CellStoreBuilderSpec extends ScalaCheckSuite:

  private def assertSeam(cells: Seq[Cell])(using munit.Location): Unit =
    assertEquals(CellStoreBuilderSpec.prefixFailure(cells), None)

  import CellStoreBuilderSpec.*

  override def scalaCheckTestParameters =
    super.scalaCheckTestParameters.withMinSuccessfulTests(200).withMaxSize(60)

  property("sorted, then unsorted from any point on, with repeats: every prefix equals the fold") {
    forAll(genArrival)(agreesAtEveryPrefix)
  }

  property("a result taken anywhere does not disturb the builder") {
    forAll(genArrival, Gen.choose(0, 400)) { (cells, at) =>
      val cut = math.min(at, cells.length)
      val b = CellStore.newBuilder
      b.addAll(cells.take(cut))
      val early = b.result()
      val again = b.result()
      b.addAll(cells.drop(cut))
      (early == fold(cells.take(cut))) :| "the early result" &&
      (again == early) :| "result is repeatable" &&
      (b.result() == fold(cells)) :| "the builder kept accumulating" &&
      (early == fold(cells.take(cut))) :| "the early result is untouched by later adds"
    }
  }

  test("a group closes at each 32-row edge, at every trie height and off the grid") {
    val rows = List(
      0,
      1,
      30,
      31,
      32,
      33,
      63,
      64,
      1022,
      1023,
      1024,
      1025,
      32767,
      32768,
      1048575,
      1048576,
      Int.MaxValue - 1,
      Int.MaxValue,
      Int.MinValue,
      Int.MinValue + 31,
      Int.MinValue + 32,
      -33,
      -32,
      -1
    )
    val cells = for
      row <- rows
      col <- List(-1, 0, 3)
    yield number(row, col)
    assertSeam(cells)
    // only off-grid rows, only the last group, only the first
    assertSeam(cells.filter(_.ref.row.index0 < 0))
    assertSeam(cells.takeRight(3))
    assertSeam(cells.take(3))
  }

  test("a group is a block up to the cell cap and rows one cell past it") {
    val cap = Group.MaxBlockCells
    def group(count: Int, firstRow: Int): List[Cell] =
      List.tabulate(count)(i => number(firstRow + i / 8, i % 8))
    assertEquals(cap, 256, "the fixture below fills 32 rows of 8 columns")
    assertSeam(group(cap, 0))
    assertSeam(group(cap - 1, 0))
    // 257 cells in one group: 32 rows of 8 and a ninth column in the last row
    assertSeam(group(cap, 0) :+ number(31, 8))
    // a full block, a group one over the cap, a full block
    assertSeam(group(cap, 0) ++ group(cap, 32) ++ List(number(63, 8)) ++ group(cap, 64))
  }

  test("a group whose cells leave its 64-column chunk is rows, wherever the stray cell arrives") {
    val base = List.tabulate(64)(i => number(64 + i / 4, 60 + i % 4))
    // every cell in chunk 0: a block
    assertSeam(base)
    // the stray cell first, in the middle and last in the group
    assertSeam(number(64, 0 - 1) :: base)
    assertSeam(base.take(30) ++ List(number(64 + 7, 64)) ++ base.drop(30))
    assertSeam(base :+ number(64 + 15, 64))
    // a whole group in another chunk, between two groups in chunk 0
    assertSeam(
      List.tabulate(8)(i => number(i, i)) ++
        List.tabulate(8)(i => number(32 + i, 64 + i)) ++
        List.tabulate(8)(i => number(64 + i, i))
    )
  }

  test("the first hyperlink may arrive late: in a block, in rows, after the scratch has grown") {
    def linked(c: Cell): Cell = c.withHyperlink("https://example.com/" + c.ref.col.index0)
    val block = List.tabulate(40)(i => number(i / 4, i % 4))
    val rows = List.tabulate(40)(i => number(32 + i / 4, 62 + i % 4))
    val both = block ++ rows
    both.indices.foreach(i => assertSeam(both.updated(i, linked(both(i)))))
    // links in one group only: the group after it must not inherit them
    assertSeam(block.map(linked) ++ rows)
    assertSeam(block ++ rows.map(linked) ++ List(number(64, 0), number(65, 1)))
    // the last cell gains, and then loses, a link by an in-place repeat
    assertSeam(block :+ linked(block.last))
    assertSeam(block.init ++ List(linked(block.last), block.last))
  }

  test("a repeat of the last reference replaces it in place, across every seam") {
    val a = number(31, 5)
    val b = Cell(a.ref, CellValue.Text("again"), Some(StyleId(2)), Some("l"))
    val c = Cell(a.ref, CellValue.Empty)
    assertSeam(List(a, b))
    assertSeam(List(a, b, c, number(32, 0), number(32, 0).withValue(CellValue.Bool(true))))
    // a repeat after the input stopped being sorted, and of a streamed cell from the tail
    assertSeam(List(number(5, 5), a, number(0, 0), number(0, 0).withValue(CellValue.Text("x")), b))
  }

  test("the first out-of-order cell: same row, an earlier row, an earlier group, a repeat") {
    val sorted = List.tabulate(120)(i => number(i / 3, i % 3))
    val last = sorted.last.ref
    val strays = List(
      number(last.row.index0, last.col.index0 - 1),
      number(last.row.index0, 0),
      number(last.row.index0 - 1, 2),
      number(last.row.index0 - 1, 70),
      number(0, 0),
      number(0, 1).withValue(CellValue.Text("rewritten")),
      number(33, 500),
      number(-1, 0)
    )
    strays.foreach { stray =>
      // a stray, then more sorted cells, then cells that rewrite streamed ones
      assertSeam(sorted ++ List(stray))
      assertSeam(sorted ++ List(stray, number(40, 0), number(41, 0), number(2, 2)))
      assertSeam(sorted.take(1) ++ List(stray) ++ sorted.drop(1))
    }
    // strays between runs of sorted cells: a streamed cell rewritten from the tail, the tail cell
    // rewritten from the tail, and the last streamed cell repeated after a stray
    assertSeam(
      List(
        number(0, 0),
        number(0, 5),
        number(0, 2),
        number(1, 0),
        number(0, 5).withValue(text),
        number(1, 1),
        number(0, 5),
        number(0, 2).withValue(text),
        number(1, 1).withValue(text),
        number(70, 0),
        number(33, 1),
        number(70, 0).withValue(text),
        number(70, 1)
      )
    )
    // unsorted from the second cell on, and a fully reversed input
    assertSeam(sorted.reverse)
    assertSeam(sorted.head :: sorted.tail.reverse)
  }

  test("banded rows: more than FlatMax cells over several bands, streamed and from the tail") {
    val width = Band.FlatMax + 6
    def wide(row: Int): List[Cell] = List.tabulate(width)(i => number(row, i * 9 - 40))
    val cells = wide(0) ++ wide(1) ++ List(number(2, 0)) ++ wide(31) ++ wide(32) ++ wide(-1)
    assertEquals(CellStore.from(cells).checkInvariants, Right(()))
    assertEquals(CellStore.from(cells), fold(cells))
    assertEquals(CellStore.from(cells.reverse), fold(cells.reverse))
    // a row one cell short of banding, and one that reaches it by its last cell
    assertSeam(wide(5).init.drop(5) ++ wide(6).drop(5))
    val b = CellStore.newBuilder
    b.addAll(cells.take(width / 2))
    assertEquals(b.result(), fold(cells.take(width / 2)), "a result in the middle of a wide row")
    b.addAll(cells.drop(width / 2))
    assertEquals(b.result(), fold(cells))
  }

object CellStoreBuilderSpec:

  private val text: CellValue = CellValue.Text("rewritten")

  def number(row: Int, col: Int): Cell =
    Cell(ARef.from0(col, row), CellValue.Number(BigDecimal(row.toLong * 31 + col)))

  def fold(cells: Seq[Cell]): CellStore = cells.foldLeft(CellStore.empty)(_.updated(_))

  /**
   * One builder fed `cells` in order; after every cell its result must equal the fold of the cells
   * so far, hold the store's invariants and iterate what the fold iterates.
   */
  def prefixFailure(cells: Seq[Cell]): Option[String] =
    val b = CellStore.newBuilder
    cells.indices.iterator
      .flatMap { i =>
        b += cells(i)
        val built = b.result()
        val expected = fold(cells.take(i + 1))
        val clue = s"after ${i + 1} of ${cells.length} cells (last ${cells(i).ref.toA1})"
        built.checkInvariants match
          case Left(broken) => Some(s"$clue: $broken")
          case Right(()) =>
            if built != expected then Some(s"$clue: not the fold\n built $built\n fold  $expected")
            else if built.iterator.toList != expected.iterator.toList then Some(s"$clue: iteration")
            else if built.hashCode != expected.hashCode then Some(s"$clue: hashCode")
            else None
      }
      .nextOption()

  def agreesAtEveryPrefix(cells: List[Cell]): Prop =
    val failure = prefixFailure(cells)
    failure.isEmpty :| failure.getOrElse("")

  private val genValue: Gen[CellValue] = Gen.frequency(
    4 -> Gen.choose(0, 9).map(i => CellValue.Number(BigDecimal(i))),
    2 -> Gen.const(CellValue.Empty),
    1 -> Gen.const(CellValue.Text("t")),
    1 -> Gen.const(CellValue.Formula("A1", None))
  )

  /** Rows either side of slot-group and height edges, a few of them off the grid. */
  private val genRow: Gen[Int] = Gen.frequency(
    12 -> Gen.choose(0, 70),
    2 -> Gen.oneOf(1023, 1024, 1025, 32768, 1048576),
    1 -> Gen.oneOf(-1, -32, -33, Int.MinValue, Int.MaxValue)
  )

  /** Columns in one chunk mostly (so groups are blocks), with strays over chunk and band edges. */
  private val genCol: Gen[Int] = Gen.frequency(
    14 -> Gen.choose(0, 9),
    2 -> Gen.oneOf(62, 63, 64, 65, 255, 256, 300),
    1 -> Gen.oneOf(-1, -64, -65, 16383, Int.MinValue, Int.MaxValue)
  )

  private val genCell: Gen[Cell] =
    for
      row <- genRow
      col <- genCol
      value <- genValue
      style <- Gen.frequency(
        5 -> Gen.const(None),
        1 -> Gen.choose(-1, 2).map(i => Some(StyleId(i)))
      )
      link <- Gen.frequency(9 -> Gen.const(None), 1 -> Gen.oneOf("a", "b").map(Some(_)))
    yield Cell(ARef.from0(col, row), value, style, link)

  /** A wide row's worth of cells, so some rows are banded. */
  private val genWideRow: Gen[List[Cell]] =
    for
      row <- genRow
      count <- Gen.choose(Band.FlatMax - 2, Band.FlatMax + 8)
      step <- Gen.oneOf(1, 5, 9)
    yield List.tabulate(count)(i => number(row, i * step))

  /**
   * An arrival order: the cells in store order up to a cut (often the whole list, often none of
   * it), the rest shuffled, and some cells repeated with other content (next to the original, and
   * at the end).
   */
  val genArrival: Gen[List[Cell]] =
    for
      cells <- Gen.listOf(genCell)
      wide <- Gen.frequency(4 -> Gen.const(Nil), 1 -> genWideRow)
      all = cells ++ wide
      distinct = all.groupBy(c => (c.ref.row.index0, c.ref.col.index0)).values.map(_.last).toList
      sorted = distinct.sortBy(c => (c.ref.row.index0, c.ref.col.index0))(using
        CellStoreLawSpec.StoreOrder
      )
      cut <- Gen.frequency(
        3 -> Gen.const(sorted.length),
        1 -> Gen.const(0),
        4 -> Gen.choose(0, sorted.length)
      )
      seed <- Gen.choose(0L, Long.MaxValue)
      rest = new scala.util.Random(seed).shuffle(sorted.drop(cut))
      repeats <- Gen.listOf(Gen.zip(Gen.choose(0, math.max(0, sorted.length - 1)), genValue))
      order = sorted.take(cut) ++ rest
      withRepeats = repeats.take(6).foldLeft(order) { case (acc, (at, value)) =>
        if acc.isEmpty then acc
        else
          val i = at % acc.length
          val again = acc(i).withValue(value)
          if at % 2 == 0 then acc.take(i + 1) ++ (again :: acc.drop(i + 1)) else acc :+ again
      }
    yield withRepeats
