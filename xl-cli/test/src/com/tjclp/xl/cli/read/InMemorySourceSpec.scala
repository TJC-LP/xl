package com.tjclp.xl.cli.read

import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{CellRange, Column, Row}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.contract.OutputMode
import com.tjclp.xl.macros.ref
import com.tjclp.xl.sheets.{ColumnProperties, RowProperties}

/**
 * The loaded-workbook source's indexes and its window projection. `MergeIndex` keeps two buckets —
 * ranges up to [[InMemorySource.MergeIndex.tallRows]] high indexed by row, taller ones scanned
 * directly — and a cell must be found in either; `search --json` is where the answer reaches a
 * user. `rows` streams a window a chunk of rows at a time, so a full column folds in bounded
 * memory.
 */
class InMemorySourceSpec extends CatsEffectSuite:

  private val tallMerge = CellRange(ref"A1", ref"A100")
  private val shortMerge = CellRange(ref"B1", ref"C1")

  test("MergeIndex.at: a cell in a tall merge, in an indexed merge, and in none") {
    val index = InMemorySource.MergeIndex(Vector(tallMerge, shortMerge))
    assertEquals(index.at(ref"A50"), Some(tallMerge))
    assertEquals(index.at(ref"A1"), Some(tallMerge))
    assertEquals(index.at(ref"A100"), Some(tallMerge))
    assertEquals(index.at(ref"C1"), Some(shortMerge))
    assertEquals(index.at(ref"B2"), None)
    assertEquals(index.at(ref"A101"), None)
    assertEquals(index.tallRanges, Vector(tallMerge))
    assertEquals(index.indexedRows, 1)
  }

  test("MergeIndex buckets on height > tallRows: 65 rows are tall, 64 are indexed") {
    assertEquals(InMemorySource.MergeIndex.tallRows, 64)
    val sixtyFive = CellRange(ref"A1", ref"A65")
    val sixtyFour = CellRange(ref"A1", ref"A64")
    val tall = InMemorySource.MergeIndex(Vector(sixtyFive))
    assertEquals(tall.tallRanges, Vector(sixtyFive))
    assertEquals(tall.indexedRows, 0)
    assertEquals(tall.at(ref"A30"), Some(sixtyFive))
    assertEquals(tall.at(ref"A65"), Some(sixtyFive))
    val indexed = InMemorySource.MergeIndex(Vector(sixtyFour))
    assertEquals(indexed.tallRanges, Vector.empty)
    assertEquals(indexed.indexedRows, 64)
    assertEquals(indexed.at(ref"A30"), Some(sixtyFour))
    assertEquals(indexed.at(ref"A65"), None)
  }

  test(
    "search --json: a cell inside a >64-row merge carries mergedInto in memory, null under --stream"
  ) {
    val sheet = Sheet("Data")
      .put(ref"A1", "Total")
      .put(ref"B1", "wide")
      .put(ref"D1", "free")
      .put(ref"A50", "inner")
      .merge(tallMerge)
      .merge(shortMerge)
    ReadTestKit.withTempWorkbook(Workbook(Vector(sheet))) { path =>
      val query = ReadQuery.Search("Total|wide|free|inner", 50, None, exactTotal = false)
      def mergedInto(json: String): Map[String, ujson.Value] =
        ujson.read(json)("matches").arr.map(m => m("ref").str -> m("mergedInto")).toMap
      for
        // the loaded workbook is the file's, so both sources see the same cells
        loaded <- ReadTestKit.excel.read(path)
        memory <- ReadTestKit.inMemory(loaded, None, query, OutputMode.Json).map(ReadTestKit.text)
        stream <- ReadTestKit.streaming(path, None, query, OutputMode.Json).map(ReadTestKit.text)
      yield
        val expected: Map[String, ujson.Value] = Map(
          "A1" -> ujson.Str("A1:A100"),
          "A50" -> ujson.Str("A1:A100"),
          "B1" -> ujson.Str("B1:C1"),
          "D1" -> ujson.Null
        )
        assertEquals(mergedInto(memory), expected)
        // the streaming reader never parses <mergeCells>: unknown, never guessed
        assertEquals(
          mergedInto(stream),
          expected.map { (ref, _) => ref -> (ujson.Null: ujson.Value) }
        )
    }
  }

  test("rows: a 200k-row window streams in bounded chunks, never one Vector of every row") {
    val sheet = Sheet("Big")
      .put(ref"A1", CellValue.Number(1))
      .put(ref"A200000", CellValue.Number(2))
    InMemorySource
      .rows(sheet, CellRange(ref"A1", ref"A200000"))
      .chunks
      .map(_.size)
      .compile
      .fold((0, 0)) { case ((total, widest), n) => (total + n, math.max(widest, n)) }
      .map { (total, widest) =>
        assertEquals(total, 200000)
        assert(widest <= InMemorySource.rowChunk, s"a chunk of $widest rows")
        assert(widest > 1, "rows are chunked, not emitted one at a time")
      }
  }

  test("stats over a full column (A1:A1048576) folds through the lazy rows and completes") {
    val sheet = Sheet("Data")
      .put(ref"A1", CellValue.Number(10))
      .put(ref"A3", CellValue.Number(5))
      .put(ref"B1", CellValue.Number(100))
    ReadTestKit
      .inMemory(Workbook(Vector(sheet)), Some("Data"), ReadQuery.Stats("A1:A1048576"))
      .map(ReadTestKit.text)
      .map(out => assertEquals(out, "count: 2, sum: 15.00, min: 5.00, max: 10.00, mean: 7.50"))
  }

  test("hiddenLines: the sheet's hidden rows and columns, clipped to the window") {
    val sheet = Sheet("H")
      .put(ref"A1", CellValue.Number(1))
      .setRowProperties(Row.from0(2), RowProperties(hidden = true))
      .setRowProperties(Row.from0(40), RowProperties(hidden = true))
      .setColumnProperties(Column.from0(1), ColumnProperties(hidden = true))
    val window = CellRange(ref"A1", ref"C5")
    assertEquals(InMemorySource.hiddenLines(sheet, window), (Set(2), Set(1)))
    assertEquals(
      InMemorySource.hiddenLines(sheet, CellRange(ref"D10", ref"D20")),
      (Set.empty[Int], Set.empty[Int])
    )
    // and the window's records carry them
    InMemorySource.rows(sheet, window).compile.toVector.map { rows =>
      val hidden = rows.map(_.map(_.hidden))
      assertEquals(hidden.lift(2), Some(Vector(Some(true), Some(true), Some(true))))
      assertEquals(hidden.lift(0), Some(Vector(Some(false), Some(true), Some(false))))
    }
  }
