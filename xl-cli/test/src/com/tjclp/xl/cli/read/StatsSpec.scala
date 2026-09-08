package com.tjclp.xl.cli.read

import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.contract.{OutputMode, WarningCode}
import com.tjclp.xl.macros.ref

/**
 * `stats` over ranges that used to fail or misparse (GH-641): a range holding no numbers is a
 * zero-count result with a `NO_NUMERIC_VALUES` warning, never `OTHER`; whole-column and whole-row
 * spans — bare or sheet-qualified — are accepted like `view`'s.
 */
class StatsSpec extends CatsEffectSuite:

  private val sheet = Sheet("Data")
    .put(ref"A1", CellValue.Text("Item"))
    .put(ref"A2", CellValue.Text("Widget"))
    .put(ref"A3", CellValue.Text("Gadget"))
    .put(ref"B1", CellValue.Number(10))
    .put(ref"B2", CellValue.Number(20))
    .put(ref"B3", CellValue.Number(BigDecimal("12.5")))
    .put(ref"C2", CellValue.Number(7))

  private val wb = Workbook(Vector(sheet, Sheet("Other").put(ref"A1", CellValue.Number(99))))

  test("GH-641: a non-numeric range is zero-count stats, exit 0, with NO_NUMERIC_VALUES") {
    for
      text <- ReadTestKit.inMemory(wb, Some("Data"), ReadQuery.Stats("A1:A3"))
      json <- ReadTestKit.inMemory(wb, Some("Data"), ReadQuery.Stats("A1:A3"), OutputMode.Json)
    yield
      assert(text.ok, text.error.toString)
      assertEquals(text.exitCode.code, 0)
      assertEquals(
        ReadTestKit.text(text),
        "count: 0, sum: 0.00, min: n/a, max: n/a, mean: n/a"
      )
      assertEquals(text.warnings.map(_.code), Vector(WarningCode.NO_NUMERIC_VALUES))
      assertEquals(text.warnings.map(_.message), Vector("No numeric values in range A1:A3"))
      val doc = ujson.read(ReadTestKit.text(json))
      assertEquals(doc("count").num.toInt, 0)
      assertEquals(doc("sum").num, 0.0)
      assertEquals(doc("min"), ujson.Null)
      assertEquals(doc("max"), ujson.Null)
      assertEquals(doc("mean"), ujson.Null)
      assertEquals(json.warnings.map(_.code), Vector(WarningCode.NO_NUMERIC_VALUES))
  }

  test("GH-641: whole-column spans are accepted, bare and qualified, and labelled as spelled") {
    for
      bare <- ReadTestKit.inMemory(wb, Some("Data"), ReadQuery.Stats("B:B"))
      wide <- ReadTestKit.inMemory(wb, Some("Data"), ReadQuery.Stats("B:C"), OutputMode.Json)
      qualified <- ReadTestKit.inMemory(wb, None, ReadQuery.Stats("Other!A:A"))
      quoted <- ReadTestKit.inMemory(wb, None, ReadQuery.Stats("'Other'!A:A"), OutputMode.Json)
      rows <- ReadTestKit.inMemory(wb, Some("Data"), ReadQuery.Stats("2:2"), OutputMode.Json)
    yield
      assertEquals(
        ReadTestKit.text(bare),
        "count: 3, sum: 42.50, min: 10.00, max: 20.00, mean: 14.17"
      )
      assertEquals(bare.warnings, Vector.empty)
      val wideDoc = ujson.read(ReadTestKit.text(wide))
      assertEquals(wideDoc("range").str, "B:C")
      assertEquals(wideDoc("count").num.toInt, 4)
      assertEquals(
        ReadTestKit.text(qualified),
        "count: 1, sum: 99.00, min: 99.00, max: 99.00, mean: 99.00"
      )
      val quotedDoc = ujson.read(ReadTestKit.text(quoted))
      assertEquals(quotedDoc("sheet").str, "Other")
      assertEquals(quotedDoc("range").str, "A:A")
      val rowDoc = ujson.read(ReadTestKit.text(rows))
      assertEquals(rowDoc("range").str, "2:2")
      assertEquals(rowDoc("count").num.toInt, 2)
  }

  test("GH-641: a whole-column span folds the used rows only, labelled as spelled") {
    // 3 used rows: the fold sees rows 1-3 of B, and the label is still the span asked for
    val tall = Sheet("T").put(ref"B2", CellValue.Number(1)).put(ref"B3", CellValue.Number(2))
    ReadTestKit
      .inMemory(Workbook(Vector(tall)), Some("T"), ReadQuery.Stats("B:B"), OutputMode.Json)
      .map { outcome =>
        val doc = ujson.read(ReadTestKit.text(outcome))
        assertEquals(doc("range").str, "B:B")
        assertEquals(doc("count").num.toInt, 2)
        assertEquals(doc("sum").num.toInt, 3)
      }
  }

  test("GH-641: a malformed ref is still INVALID_REFERENCE with the parser's text") {
    ReadTestKit.inMemory(wb, Some("Data"), ReadQuery.Stats("B1:")).map { outcome =>
      assertEquals(outcome.error.map(_.code), Some("INVALID_REFERENCE"))
      assertEquals(outcome.exitCode.code, 3)
    }
  }

  test("GH-641: the streaming source answers whole-column stats identically") {
    ReadTestKit.withTempWorkbook(wb) { path =>
      for
        memory <- ReadTestKit.inMemory(wb, Some("Data"), ReadQuery.Stats("B:B"), OutputMode.Json)
        streamed <- ReadTestKit.streaming(
          path,
          Some("Data"),
          ReadQuery.Stats("B:B"),
          OutputMode.Json
        )
        empty <- ReadTestKit.streaming(path, Some("Data"), ReadQuery.Stats("A1:A3"))
      yield
        assertEquals(ReadTestKit.text(streamed), ReadTestKit.text(memory))
        assert(empty.ok, empty.error.toString)
        assertEquals(empty.warnings.map(_.code), Vector(WarningCode.NO_NUMERIC_VALUES))
    }
  }
