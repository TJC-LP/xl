package com.tjclp.xl.cli

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.{Workbook, Sheet, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.contract.OutputMode
import com.tjclp.xl.cli.read.{ReadQuery, ReadTestKit}
import com.tjclp.xl.macros.ref

/**
 * The streaming source renders through the same renderers as the loaded workbook (W2.4): the
 * outputs are identical, the streaming `view --format json` included (the typed `{sheet, range,
 * rows}` shape — Changed from the bare array of strings it printed before).
 */
class StreamingReadSpec extends CatsEffectSuite:

  private val sheet = Sheet("Test")
    .put(ref"A1", CellValue.Text("Hello"))
    .put(ref"A2", CellValue.Number(BigDecimal(42)))
    .put(ref"A3", CellValue.Text("World"))
  private val wb = Workbook(Vector(sheet))

  test("streaming view csv matches in-memory output without extra trailing newline") {
    ReadTestKit.withTempWorkbook(wb) { path =>
      val query = ReadTestKit.view(Some("A1:A3"), ViewFormat.Csv, limit = 100)
      for
        inMemory <- ReadTestKit.inMemory(wb, Some("Test"), query).map(ReadTestKit.text)
        streaming <- ReadTestKit.streaming(path, Some("Test"), query).map(ReadTestKit.text)
      yield
        assertEquals(streaming, inMemory)
        assert(
          !streaming.endsWith("\n"),
          s"streaming CSV should not end with newline: '$streaming'"
        )
    }
  }

  test("streaming view json is the typed in-memory shape, byte for byte") {
    ReadTestKit.withTempWorkbook(wb) { path =>
      val query = ReadTestKit.view(Some("A1:B3"), ViewFormat.Json, limit = 100)
      for
        inMemory <- ReadTestKit.inMemory(wb, Some("Test"), query).map(ReadTestKit.text)
        streaming <- ReadTestKit.streaming(path, Some("Test"), query).map(ReadTestKit.text)
      yield
        assertEquals(streaming, inMemory)
        assert(streaming.startsWith("{\n  \"sheet\": \"Test\","), streaming)
        assert(
          streaming.contains(
            """{"ref": "A2", "type": "number", "value": 42, "formatted": "42"}"""
          ),
          streaming
        )
    }
  }

  test("streaming view markdown matches the in-memory table, labels included") {
    ReadTestKit.withTempWorkbook(wb) { path =>
      val query = ReadTestKit.view(Some("A1:A3"), ViewFormat.Markdown, limit = 100)
      for
        inMemory <- ReadTestKit.inMemory(wb, Some("Test"), query).map(ReadTestKit.text)
        streaming <- ReadTestKit.streaming(path, Some("Test"), query).map(ReadTestKit.text)
      yield
        assertEquals(streaming, inMemory)
        assert(streaming.startsWith("|   | A     |"), streaming)
    }
  }

  test("streaming search, stats and cell agree with the in-memory verbs") {
    ReadTestKit.withTempWorkbook(wb) { path =>
      val queries: Vector[(Option[String], ReadQuery)] = Vector(
        (None, ReadQuery.Search("o", 50, None)),
        (Some("Test"), ReadQuery.Stats("A1:A3")),
        (Some("Test"), ReadQuery.Cell("A2", noStyle = false))
      )
      queries.foldLeft(IO.unit) { case (acc, (flag, query)) =>
        acc *> (for
          inMemory <- ReadTestKit.inMemory(wb, flag, query, OutputMode.Json)
          streaming <- ReadTestKit.streaming(path, flag, query, OutputMode.Json)
        yield
          val memoryData = ujson.read(ReadTestKit.text(inMemory))
          val streamData = ujson.read(ReadTestKit.text(streaming))
          query match
            case _: ReadQuery.Cell =>
              assertEquals(memoryData("dependents"), ujson.Arr())
              assertEquals(streamData("dependents"), ujson.Null)
              memoryData.obj.remove("dependents")
              streamData.obj.remove("dependents")
              assertEquals(streamData, memoryData)
            case _ => assertEquals(streamData, memoryData))
      }
    }
  }

  test("streaming --skip-hidden is reported as ignored, not silently accepted") {
    ReadTestKit.withTempWorkbook(wb) { path =>
      ReadTestKit
        .streaming(path, Some("Test"), ReadTestKit.view(Some("A1:A3"), skipHidden = true))
        .map { outcome =>
          assertEquals(outcome.warnings.map(_.code), Vector("FLAG_IGNORED"))
          assert(outcome.warnings.exists(_.message.contains("--skip-hidden")))
        }
    }
  }
