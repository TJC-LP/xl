package com.tjclp.xl.cli.output

import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.CellValue
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

  test("the streaming engine renders the same labels") {
    val wb = book(11)
    ReadTestKit.withTempWorkbook(wb) { path =>
      for
        memory <- ReadTestKit.inMemory(wb, Some("Data"), ReadTestKit.view(Some("A1:A11")))
        streamed <- ReadTestKit.streaming(path, Some("Data"), ReadTestKit.view(Some("A1:A11")))
      yield assertEquals(ReadTestKit.text(streamed), ReadTestKit.text(memory))
    }
  }
