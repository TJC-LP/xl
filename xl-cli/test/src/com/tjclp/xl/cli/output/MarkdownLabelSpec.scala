package com.tjclp.xl.cli.output

import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.ViewFormat
import com.tjclp.xl.cli.read.ReadTestKit
import com.tjclp.xl.macros.ref

/** GH-641: the markdown row-label column is as wide as the widest row number in the window. */
class MarkdownLabelSpec extends CatsEffectSuite:

  private def book(rows: Int): Workbook =
    val sheet = (1 to rows).foldLeft(Sheet("Data")) { (s, r) =>
      s.put(ARef.from0(0, r - 1), CellValue.Text("x"))
    }
    Workbook(Vector(sheet))

  test("a window past row 9 pads every label to two characters; a short window keeps one") {
    for
      wide <- ReadTestKit.readText(book(12), Some("Data"), ReadTestKit.view(Some("A1:A12")))
      narrow <- ReadTestKit.readText(book(3), Some("Data"), ReadTestKit.view(Some("A1:A3")))
    yield
      val wideLines = wide.linesIterator.toVector
      assertEquals(wideLines(0), "|    | A   |")
      assertEquals(wideLines(1), "|----|-----|")
      assertEquals(wideLines(2), "| 1  | x   |")
      assertEquals(wideLines(11), "| 10 | x   |")
      assertEquals(wideLines(13), "| 12 | x   |")
      // every data row starts its first cell at the same column
      assertEquals(wideLines.drop(2).map(_.indexOf("| x")).distinct, Vector(5))
      val narrowLines = narrow.linesIterator.toVector
      assertEquals(narrowLines(0), "|   | A   |")
      assertEquals(narrowLines(1), "|---|-----|")
      assertEquals(narrowLines(2), "| 1 | x   |")
  }

  test("GH-641: a whole-column span is clamped to the used rows (totalRows = 12, not 1,048,576)") {
    val wb = book(12)
    val json = ReadTestKit.view(Some("A:A"), format = ViewFormat.Json, limit = 5)
    ReadTestKit.withTempWorkbook(wb) { path =>
      for
        memory <- ReadTestKit.inMemory(wb, Some("Data"), json)
        streamed <- ReadTestKit.streaming(path, Some("Data"), json)
        all <- ReadTestKit.readText(wb, Some("Data"), ReadTestKit.view(Some("A:A"), limit = 0))
        rows <- ReadTestKit.readText(wb, Some("Data"), ReadTestKit.view(Some("3:3"), limit = 0))
        beyond <- ReadTestKit.readText(wb, Some("Data"), ReadTestKit.view(Some("Z:Z"), limit = 0))
      yield
        val doc = ujson.read(ReadTestKit.text(memory))
        assertEquals(doc("truncated"), ujson.True)
        assertEquals(doc("totalRows").num.toInt, 12)
        assertEquals(doc("range").str, "A1:A5")
        assertEquals(ReadTestKit.text(streamed), ReadTestKit.text(memory))
        // --limit 0 renders the 12 used rows, not the axis
        assertEquals(all.linesIterator.size, 14, all)
        assert(!all.contains("… showing"), all)
        // a whole-row span keeps its row and takes the used columns (A alone here)
        assertEquals(rows.linesIterator.toVector.drop(2), Vector("| 3 | x   |"))
        // a column outside the used range renders empty over the used rows
        assertEquals(beyond.linesIterator.size, 14, beyond)
        assert(beyond.linesIterator.drop(2).forall(_.startsWith("| ")), beyond)
    }
  }

  test("the streaming engine renders the same labels") {
    val wb = book(11)
    ReadTestKit.withTempWorkbook(wb) { path =>
      for
        memory <- ReadTestKit.inMemory(wb, Some("Data"), ReadTestKit.view(Some("A1:A11")))
        streamed <- ReadTestKit.streaming(path, Some("Data"), ReadTestKit.view(Some("A1:A11")))
      yield assertEquals(ReadTestKit.text(streamed), ReadTestKit.text(memory))
    }
  }
