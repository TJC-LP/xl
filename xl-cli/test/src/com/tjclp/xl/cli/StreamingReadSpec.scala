package com.tjclp.xl.cli

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.{Workbook, Sheet, given}
import com.tjclp.xl.cells.{CellValue, Comment}
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
              // the graph and hidden fields belong to capabilities the reader lacks: null, not a guess
              assertEquals(memoryData("dependencies"), ujson.Arr())
              assertEquals(memoryData("dependents"), ujson.Arr())
              assertEquals(memoryData("hidden"), ujson.Bool(false))
              assertEquals(streamData("dependencies"), ujson.Null)
              assertEquals(streamData("dependents"), ujson.Null)
              assertEquals(streamData("hidden"), ujson.Null)
              Vector("dependencies", "dependents", "hidden").foreach { key =>
                memoryData.obj.remove(key)
                streamData.obj.remove(key)
              }
              assertEquals(streamData, memoryData)
            case _: ReadQuery.Search =>
              streamData("matches").arr.foreach(m => assertEquals(m("hidden"), ujson.Null))
              streamData("matches").arr.foreach(_.obj.remove("hidden"))
              memoryData("matches").arr.foreach(_.obj.remove("hidden"))
              assertEquals(streamData, memoryData)
            case _ => assertEquals(streamData, memoryData))
      }
    }
  }

  test("streaming cell on a formula: both graph lists say they are not available") {
    val formulas = Sheet("Calc")
      .put(ref"A1", CellValue.Number(BigDecimal(2)))
      .put(ref"A2", CellValue.Number(BigDecimal(3)))
      .put(ref"A3", CellValue.Formula("SUM(A1:A2)", Some(CellValue.Number(BigDecimal(5)))))
    ReadTestKit.withTempWorkbook(Workbook(Vector(formulas))) { path =>
      val query = ReadQuery.Cell("A3", noStyle = false)
      for
        inMemory <- ReadTestKit.inMemory(Workbook(Vector(formulas)), Some("Calc"), query)
        streaming <- ReadTestKit.streaming(path, Some("Calc"), query)
      yield
        val loaded = ReadTestKit.text(inMemory)
        val streamed = ReadTestKit.text(streaming)
        assert(loaded.contains("Dependencies: A1, A2"), loaded)
        assert(streamed.contains("Dependencies: (not available in streaming mode)"), streamed)
        assert(streamed.contains("Dependents: (not available in streaming mode)"), streamed)
        // and never the reader's token list (`A1, A1:A2, A2`) that pre-0.21.0 streaming printed
        val dependenciesLine = streamed.linesIterator.find(_.startsWith("Dependencies:"))
        assertEquals(dependenciesLine, Some("Dependencies: (not available in streaming mode)"))
    }
  }

  test("streaming cell on a commented cell prints the comment the in-memory reader prints") {
    val commented = Sheet("Notes")
      .put(ref"A1", CellValue.Number(BigDecimal(5)))
      .comment(ref"A1", Comment.plainText("input", Some("qa")))
      .comment(ref"B2", Comment.plainText("no author", None))
    val book = Workbook(Vector(commented))
    ReadTestKit.withTempWorkbook(book) { path =>
      Vector("A1", "B2").foldLeft(IO.unit) { (acc, cellRef) =>
        acc *> (for
          inMemory <- ReadTestKit.inMemory(book, Some("Notes"), ReadQuery.Cell(cellRef, false))
          streaming <- ReadTestKit.streaming(path, Some("Notes"), ReadQuery.Cell(cellRef, false))
          memoryJson <- ReadTestKit
            .inMemory(book, Some("Notes"), ReadQuery.Cell(cellRef, false), OutputMode.Json)
          streamJson <- ReadTestKit
            .streaming(path, Some("Notes"), ReadQuery.Cell(cellRef, false), OutputMode.Json)
        yield
          val loaded = ReadTestKit.text(inMemory)
          val streamed = ReadTestKit.text(streaming)
          val commentLine = loaded.linesIterator.find(_.startsWith("Comment:"))
          assert(commentLine.isDefined, loaded)
          assertEquals(streamed.linesIterator.find(_.startsWith("Comment:")), commentLine)
          assertEquals(
            ujson.read(ReadTestKit.text(streamJson))("comment"),
            ujson.read(ReadTestKit.text(memoryJson))("comment")
          ))
      } *> ReadTestKit.streaming(path, Some("Notes"), ReadQuery.Cell("A1", false)).map { out =>
        // the writer's "qa:" prefix run is stripped on both paths, the author reported separately
        assert(
          ReadTestKit.text(out).contains("Comment: \"input\" (Author: qa)"),
          ReadTestKit.text(out)
        )
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
