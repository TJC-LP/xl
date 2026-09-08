package com.tjclp.xl.cli

import cats.effect.IO
import cats.syntax.all.*
import munit.CatsEffectSuite

import com.tjclp.xl.{Workbook, Sheet, given}
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cells.{CellValue, Comment}
import com.tjclp.xl.cli.contract.OutputMode
import com.tjclp.xl.cli.read.{InMemorySource, ReadQuery, ReadTestKit, SheetSource}
import com.tjclp.xl.macros.ref
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.numfmt.NumFmt

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
        (None, ReadQuery.Search("o", 50, None, exactTotal = false)),
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

  /** Both sources' payload text for `query`, asserted byte-equal; the in-memory text returned. */
  private def agree(
    loaded: Workbook,
    path: java.nio.file.Path,
    sheet: String,
    query: ReadQuery,
    mode: OutputMode = OutputMode.Text
  ): IO[String] =
    for
      memory <- ReadTestKit.inMemory(loaded, Some(sheet), query, mode).map(ReadTestKit.text)
      stream <- ReadTestKit.streaming(path, Some(sheet), query, mode).map(ReadTestKit.text)
    yield
      assertEquals(stream, memory, s"$query differs between the sources")
      memory

  test(
    "openpyxl's package-absolute worksheet Targets: the default window is one box from both sources"
  ) {
    // A1:B2 hold values; D5 is styled but empty. The stored-cell box, A1:D5, is the <dimension> the
    // writer records — and no scan of the non-empty cells (A1:B2) could recover it, so the
    // streaming source must find the element behind `Target="/xl/worksheets/sheet1.xml"`.
    val values = Sheet("Data").put(ref"A1", "h1").put(ref"B1", "h2").put(ref"A2", 1).put(ref"B2", 2)
    val data = values.styleAt("D5", CellStyle.default.withNumFmt(NumFmt.Percent)).getOrElse(values)
    ReadTestKit.withTempWorkbook(Workbook(Vector(data))) { path =>
      for
        _ <- ReadTestKit.rewriteZip(path)(ReadTestKit.openpyxlTargets)
        rels <- ReadTestKit.zipEntry(path, "xl/_rels/workbook.xml.rels")
        _ = assert(rels.contains("Target=\"/xl/worksheets/sheet1.xml\""), rels)
        loaded <- ReadTestKit.excel.read(path)
        streamed <- SheetSource
          .streaming(path, ReadTestKit.excel)
          .flatMap(_.usedRange(SheetName.unsafe("Data")))
        _ = assertEquals(streamed.map(_.toA1), Some("A1:D5"))
        _ = assertEquals(streamed, loaded.sheets.headOption.flatMap(InMemorySource.dimension))
        markdown <- agree(loaded, path, "Data", ReadTestKit.view(None))
        json <- agree(loaded, path, "Data", ReadTestKit.view(None, ViewFormat.Json))
        _ <- agree(loaded, path, "Data", ReadTestKit.view(None, ViewFormat.Csv))
        filtered <- agree(loaded, path, "Data", ReadTestKit.filter("A IS EMPTY"))
        _ <- agree(
          loaded,
          path,
          "Data",
          ReadTestKit.filter("A IS EMPTY", format = FilterFormat.Json)
        )
        _ <- agree(loaded, path, "Data", ReadTestKit.filter("A IS EMPTY"), OutputMode.Json)
      yield
        // the window really is the stored-cell box: column D and row 5 are addressed
        assertEquals(ujson.read(json)("range").str, "A1:D5")
        assert(markdown.linesIterator.exists(_.startsWith("| 5 ")), markdown)
        // rows 3, 4 and 5 have an empty A; a scan window (A1:B2) would have matched none
        assert(!filtered.startsWith("No rows matched"), filtered)
        assert(filtered.linesIterator.exists(_.startsWith("| 5 ")), filtered)
    }
  }

  test("Excel's <dimension ref=\"A1\"/> on an empty sheet: (empty sheet) from both sources") {
    // Excel writes `A1` for an empty sheet and `C3` for a sheet whose only cell is C3: a one-cell
    // <dimension> cannot say which, so the streaming source scans instead of trusting it
    val one = Sheet("One").put(ref"C3", "only")
    ReadTestKit.withTempWorkbook(Workbook(Vector(Sheet("Blank"), one))) { path =>
      for
        _ <- ReadTestKit.rewriteZip(path)(ReadTestKit.excelEmptySheetDimension)
        xml <- ReadTestKit.zipEntry(path, "xl/worksheets/sheet1.xml")
        _ = assert(xml.contains("<dimension ref=\"A1\"/>"), xml)
        meta <- ReadTestKit.excel.readMetadata(path)
        _ = assertEquals(
          meta.sheets.map(_.dimension.map(_.toA1)),
          Vector(Some("A1:A1"), Some("C3:C3"))
        )
        loaded <- ReadTestKit.excel.read(path)
        source <- SheetSource.streaming(path, ReadTestKit.excel)
        blank <- source.usedRange(SheetName.unsafe("Blank"))
        _ = assertEquals(blank, None)
        oneUsed <- source.usedRange(SheetName.unsafe("One"))
        _ = assertEquals(oneUsed.map(_.toA1), Some("C3:C3"))
        empty <- agree(loaded, path, "Blank", ReadTestKit.view(None))
        _ = assertEquals(empty, "(empty sheet)")
        emptyJson <- agree(loaded, path, "Blank", ReadTestKit.view(None, ViewFormat.Json))
        _ = assertEquals(ujson.read(emptyJson)("range"), ujson.Null)
        _ <- agree(loaded, path, "Blank", ReadTestKit.view(None, ViewFormat.Csv))
        noRows <- agree(loaded, path, "Blank", ReadTestKit.filter("A IS EMPTY"))
        _ = assertEquals(noRows, "No rows matched.")
        cell <- agree(loaded, path, "One", ReadTestKit.view(None))
        _ = assert(cell.contains("only"), cell)
        _ <- agree(loaded, path, "One", ReadTestKit.view(None, ViewFormat.Json))
      yield ()
    }
  }
