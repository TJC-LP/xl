package com.tjclp.xl.ooxml

import java.lang.management.ManagementFactory

import munit.FunSuite

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.richtext.RichText
import com.tjclp.xl.sheets.{CellStore, Sheet}
import com.tjclp.xl.styles.units.StyleId

/**
 * #891: a generated shared-string table is numbered by walking each sheet's cells row-major, and
 * the store already iterates that way, so [[SharedStrings.fromSheets]] streams the store's cells
 * into the table instead of copying every text cell into an array and sorting it.
 *
 * The two orders are not the same everywhere: the table's is signed on rows, the store's unsigned,
 * so a sheet with a negative row (off the grid, but the writer emits it) keeps the copy and sort.
 * Pinned here: the decision, that both paths number the table as the sort always has on either side
 * of it, and that the streamed path holds nothing per text cell (cells pulled per entry, and bytes
 * allocated per cell at two sizes; counts, not clocks).
 */
class SharedStringsStreamSpec extends FunSuite:
  import SharedStringsStreamSpec.*

  test("the store's order is the table's exactly when no row is negative") {
    assert(SharedStrings.storeIsRowMajor(sheetOf(Nil)), "an empty sheet")
    assert(SharedStrings.storeIsRowMajor(sheetOf(onGrid)))
    // past the grid's last row and column, and a negative column: the same order in both
    assert(SharedStrings.storeIsRowMajor(sheetOf(onGrid ++ pastGrid)))
    assert(SharedStrings.storeIsRowMajor(sheetOf(onGrid ++ pastGrid ++ negativeCols)))
    // one negative row, with a value or style-only: the store walks it last, the table first
    assert(!SharedStrings.storeIsRowMajor(sheetOf(onGrid ++ negativeRows)))
    assert(!SharedStrings.storeIsRowMajor(sheetOf(onGrid :+ Cell(ARef.from0(0, -1)))))
    assert(
      !SharedStrings.storeIsRowMajor(
        sheetOf(onGrid :+ Cell(ARef.from0(0, Int.MinValue), CellValue.Empty, Some(StyleId(1))))
      )
    )
  }

  test("without a negative row the store's own walk is the sorted one") {
    for cells <- List(Nil, onGrid, onGrid ++ pastGrid, onGrid ++ pastGrid ++ negativeCols) do
      val sheet = sheetOf(cells)
      assertEquals(
        sheet.cells.iterator.filter(isText).toList,
        SharedStrings.sortedTextCells(sheet).toList
      )
      assertEquals(
        SharedStrings.entriesOf(SharedStrings.rowMajorCells(sheet), false).toList,
        SharedStrings.entriesOf(SharedStrings.sortedTextCells(sheet).iterator, false).toList
      )
  }

  test("with a negative row the walks differ, and the table keeps the sorted numbering") {
    val sheet = sheetOf(onGrid ++ pastGrid ++ negativeCols ++ negativeRows)
    val sorted = SharedStrings.sortedTextCells(sheet).toList
    assertNotEquals(sheet.cells.iterator.filter(isText).toList, sorted)
    assertEquals(sorted.headOption.map(_.ref), Some(ARef.from0(2, Int.MinValue)))
    assertEquals(SharedStrings.rowMajorCells(sheet).toList, sorted)
    val table = SharedStrings.fromSheets(List(sheet), escapeFormulas = false)
    assertEquals(table.strings.headOption, Some(Left("min"): SSTEntry))
  }

  test("fromSheets numbers the table as the copy-and-sort did, sheet after sheet") {
    val sheets = List(
      sheetOf(onGrid),
      sheetOf(Nil),
      sheetOf(onGrid ++ pastGrid ++ negativeCols ++ negativeRows),
      sheetOf(onGrid.reverse.map(c => c.copy(ref = ARef.from0(c.ref.row.index0, c.ref.col.index0))))
    )
    for escape <- List(false, true) do
      val expected = SharedStrings.fromEntries(
        sheets.flatMap(s =>
          SharedStrings.entriesOf(SharedStrings.sortedTextCells(s).iterator, escape)
        )
      )
      val actual = SharedStrings.fromSheets(sheets, escape)
      assertEquals(actual.strings, expected.strings)
      assertEquals(actual.totalCount, expected.totalCount)
      assertEquals(actual.indexMap, expected.indexMap)
      assertEquals(actual.richIndexMap, expected.richIndexMap)
    // "=cmd" is escaped under the flag, and deduplicated after escaping
    assert(SharedStrings.fromSheets(sheets, true).indexMap.contains("'=cmd"))
    assert(SharedStrings.fromSheets(sheets, false).indexMap.contains("=cmd"))
  }

  test("entries are taken one cell at a time: nothing is read ahead, at any size") {
    for n <- List(1000, 64000) do
      var pulled = 0
      val cells = Iterator.range(0, n).map { i =>
        pulled += 1
        // two numbers, then a text cell
        if i % 3 == 2 then Cell(ARef.from0(i % 7, i / 7), CellValue.Text(s"t$i"))
        else Cell(ARef.from0(i % 7, i / 7), CellValue.Number(BigDecimal(i)))
      }
      val entries = SharedStrings.entriesOf(cells, escapeFormulas = false)
      assertEquals(pulled, 0, s"n=$n: nothing before the first entry is asked for")
      assertEquals(entries.next(), Left("t2"): SSTEntry)
      assertEquals(pulled, 3, s"n=$n: the cells up to the first text cell")
      assertEquals(entries.next(), Left("t5"): SSTEntry)
      assertEquals(pulled, 6, s"n=$n")
      assertEquals(entries.size, n / 3 - 2)
      assertEquals(pulled, n)
  }

  test("building the table allocates per text cell what the walk does, not a copy of it") {
    // 1,000 distinct strings over n text cells. The copy-and-sort allocates 109 bytes per cell
    // (the cell, its place in a growing array and in the sorted one, its entry and a list cell
    // for that); the streamed walk allocates the cell and its entry, 49 to 51 bytes, at either
    // size. The bound sits between the two (a leaner copy-and-sort still takes 88): an array or
    // a list of the occurrences cannot come back under it.
    val strings = Array.tabulate(1000)(i => CellValue.Text(s"s$i"))
    val perCell = List(100000, 400000).map { n =>
      val sheet = sheetOf(
        Iterator.range(0, n).map(i => Cell(ARef.from0(i % 10, i / 10), strings(i % 1000)))
      )
      // warm: the measured call compiles nothing and loads no class
      for _ <- 0 until 3 do SharedStrings.fromSheets(List(sheet), escapeFormulas = false)
      val (table, bytes) = allocated(SharedStrings.fromSheets(List(sheet), escapeFormulas = false))
      assertEquals(table.totalCount, n)
      assertEquals(table.strings.size, 1000)
      n -> bytes.toDouble / n
    }
    perCell.foreach { (n, bytes) =>
      assert(bytes < 70.0, s"n=$n: $bytes bytes allocated per text cell (all: $perCell)")
    }
  }

object SharedStringsStreamSpec:

  private def text(col: Int, row: Int, s: String): Cell =
    Cell(ARef.from0(col, row), CellValue.Text(s))

  /** Text (with repeats), rich text, a formula-looking string, numbers and a style-only cell. */
  private val onGrid: List[Cell] = List(
    text(3, 0, "d"),
    text(0, 0, "a"),
    text(1, 2, "a"),
    text(0, 2, "=cmd"),
    Cell(ARef.from0(2, 1), CellValue.RichText(RichText.plain("rich"))),
    Cell(ARef.from0(0, 1), CellValue.Number(BigDecimal(1))),
    Cell(ARef.from0(5, 5), CellValue.Empty, Some(StyleId(2))),
    text(16383, 1048575, "corner"),
    text(300, 40, "wide"),
    text(1, 40, "a")
  ) ++ List.tabulate(80)(i => text(i * 5, 41, s"w${i % 9}"))

  private val pastGrid: List[Cell] = List(
    text(16384, 1048576, "past"),
    text(0, Int.MaxValue, "last row"),
    text(Int.MaxValue, 7, "last col")
  )

  private val negativeCols: List[Cell] =
    List(text(-1, 0, "left"), text(Int.MinValue, 2, "leftmost"))

  private val negativeRows: List[Cell] =
    List(text(0, -1, "above"), text(2, Int.MinValue, "min"), text(-3, -2, "above left"))

  private def isText(cell: Cell): Boolean = cell.value match
    case _: CellValue.Text | _: CellValue.RichText => true
    case _ => false

  private def sheetOf(cells: IterableOnce[Cell]): Sheet =
    Sheet(SheetName.unsafe("S")).copy(cells = CellStore.from(cells))

  /** The bytes this thread allocates while `f` runs: the JVM's own count, exact per thread. */
  private def allocated[A](f: => A): (A, Long) =
    ManagementFactory.getThreadMXBean match
      case bean: com.sun.management.ThreadMXBean =>
        val id = Thread.currentThread.threadId
        val before = bean.getThreadAllocatedBytes(id)
        val a = f
        (a, bean.getThreadAllocatedBytes(id) - before)
      case _ => (f, 0L)
