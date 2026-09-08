package com.tjclp.xl.cli

import java.nio.file.Files

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.{Sheet, Workbook, given}
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.contract.{Outcome, WarningCode}
import com.tjclp.xl.cli.raster.BatikRasterizer
import com.tjclp.xl.cli.read.{ReadQuery, ReadTestKit}

/**
 * GH-351: view/search must report when --limit clips output.
 *
 * Pins the truncation marker semantics per format:
 *   - markdown: trailer line after the table ("… showing X of Y rows")
 *   - csv/svg: stdout stays byte-identical, notice goes to the warnings (stderr in text mode)
 *   - html: notice as a warning plus a trailing HTML comment on stdout
 *   - json: top-level "truncated"/"totalRows" fields (only when clipped) — the streaming source too
 *     (W2.4: one renderer for both sources)
 *   - raster (png/jpeg/webp/pdf): notice appended to the "Exported:" status line
 *   - search: total match count + trailer when the hit list is clipped, in both sources
 *   - --limit 0 means "no limit"
 */
class ViewTruncationSpec extends CatsEffectSuite:

  /** Sheet with n numeric rows in column A (A1=1 .. An=n). */
  private def sheetWithRows(n: Int): Sheet =
    (1 to n).foldLeft(Sheet("Data")) { (s, i) =>
      s.put(ARef.from0(0, i - 1), CellValue.Number(BigDecimal(i)))
    }

  private def wbWithRows(n: Int): Workbook = Workbook(Vector(sheetWithRows(n)))

  private def runView(wb: Workbook, range: String, limit: Int, format: ViewFormat): IO[Outcome] =
    ReadTestKit.inMemory(wb, Some("Data"), ReadTestKit.view(Some(range), format, limit = limit))

  private def truncationWarnings(outcome: Outcome): Vector[String] =
    outcome.warnings.filter(_.code == WarningCode.TRUNCATED).map(_.message)

  // ========== view: markdown ==========

  test("view markdown: clipped output appends trailer with correct counts") {
    runView(wbWithRows(100), "A1:A100", 50, ViewFormat.Markdown).map(ReadTestKit.text).map { out =>
      assert(out.contains("… showing 50 of 100 rows"), s"missing trailer:\n$out")
      assert(out.contains("--limit 0 = no limit"), s"missing unlimited hint:\n$out")
      assert(out.contains("| 50"), "row 50 should be rendered")
      assert(!out.contains("| 51"), "row 51 should be clipped")
      // Trailer must be the last line, outside the table body
      val lines = out.linesIterator.toVector
      assert(
        lines.lastOption.exists(_.startsWith("… showing")),
        s"trailer should be last line, got: ${lines.lastOption}"
      )
      assert(
        lines.dropRight(1).forall(l => !l.startsWith("… showing")),
        "trailer must not appear inside the table body"
      )
    }
  }

  test("view markdown: unclipped output has no trailer") {
    runView(wbWithRows(30), "A1:A30", 50, ViewFormat.Markdown).map(ReadTestKit.text).map { out =>
      assert(!out.contains("… showing"), s"unexpected trailer:\n$out")
    }
  }

  test("view markdown: exact-limit output has no trailer") {
    runView(wbWithRows(50), "A1:A50", 50, ViewFormat.Markdown).map(ReadTestKit.text).map { out =>
      assert(!out.contains("… showing"), s"unexpected trailer:\n$out")
    }
  }

  test("view markdown: --limit 0 returns everything with no trailer") {
    runView(wbWithRows(100), "A1:A100", 0, ViewFormat.Markdown).map(ReadTestKit.text).map { out =>
      assert(out.contains("| 100"), "row 100 should be rendered with --limit 0")
      assert(!out.contains("… showing"), s"unexpected trailer:\n$out")
    }
  }

  // ========== view: paging (W2.4) ==========

  test("view --offset pages the range and says which rows are shown") {
    val wb = wbWithRows(10)
    ReadTestKit
      .inMemory(wb, Some("Data"), ReadTestKit.view(Some("A1:A10"), limit = 3, offset = 4))
      .map { outcome =>
        val out = ReadTestKit.text(outcome)
        assert(out.contains("| 5 "), s"first shown row must be 5:\n$out")
        assert(out.contains("| 7 "), s"last shown row must be 7:\n$out")
        assert(!out.contains("| 4 "), s"row 4 must be skipped:\n$out")
        assert(!out.contains("| 8 "), s"row 8 must be clipped:\n$out")
        assert(out.contains("… showing rows 5–7 of 10"), s"paging trailer missing:\n$out")
      }
  }

  test("view --offset beyond the range is a usage error") {
    ReadTestKit
      .inMemory(wbWithRows(3), Some("Data"), ReadTestKit.view(Some("A1:A3"), offset = 3))
      .map { outcome =>
        assertEquals(outcome.error.map(_.code), Some("USAGE"))
        assertEquals(outcome.exitCode.code, 2)
      }
  }

  test("view --max-cols clips columns and json reports totalCols") {
    val sheet = (0 until 5).foldLeft(Sheet("Data")) { (s, col) =>
      s.put(ARef.from0(col, 0), CellValue.Number(BigDecimal(col)))
    }
    val wb = Workbook(Vector(sheet))
    for
      json <- ReadTestKit.readText(
        wb,
        Some("Data"),
        ReadTestKit.view(None, ViewFormat.Json, maxCols = 2)
      )
      md <- ReadTestKit.inMemory(wb, Some("Data"), ReadTestKit.view(None, maxCols = 2))
    yield
      assert(json.contains("\"range\": \"A1:B1\""), json)
      assert(json.contains("\"truncated\": true"), json)
      assert(json.contains("\"totalCols\": 5"), json)
      assert(!json.contains("totalRows"), json)
      val text = ReadTestKit.text(md)
      assert(text.contains("… showing 2 of 5 columns"), text)
  }

  test("view with no range shows the used range") {
    val sheet = Sheet("Data")
      .put(ARef.from0(1, 1), CellValue.Text("b2"))
      .put(ARef.from0(3, 4), CellValue.Text("d5"))
    val wb = Workbook(Vector(sheet))
    for
      json <- ReadTestKit.readText(wb, Some("Data"), ReadTestKit.view(None, ViewFormat.Json))
      empty <- ReadTestKit.readText(
        Workbook(Vector(Sheet("Data"))),
        Some("Data"),
        ReadTestKit.view(None, ViewFormat.Json)
      )
    yield
      assert(json.contains("\"range\": \"B2:D5\""), json)
      assertEquals(ujson.read(empty)("range"), ujson.Null)
      assertEquals(ujson.read(empty)("rows").arr.size, 0)
  }

  // ========== view: csv ==========

  test("view csv: stdout stays byte-identical, notice goes to the warnings") {
    val wb = wbWithRows(100)
    for
      clipped <- runView(wb, "A1:A100", 50, ViewFormat.Csv)
      exact <- runView(wb, "A1:A50", 50, ViewFormat.Csv)
    yield
      assertEquals(
        ReadTestKit.text(clipped),
        ReadTestKit.text(exact),
        "clipped CSV stdout must equal the unclipped 50-row render"
      )
      assert(!ReadTestKit.text(clipped).contains("showing"), "notice must not leak into CSV stdout")
      assert(
        truncationWarnings(clipped).exists(_.contains("… showing 50 of 100 rows")),
        s"expected truncation warning, got: ${clipped.warnings}"
      )
      assertEquals(exact.warnings, Vector.empty)
  }

  // ========== view: json ==========

  test("view json: clipped output carries truncated/totalRows fields") {
    runView(wbWithRows(100), "A1:A100", 50, ViewFormat.Json).map(ReadTestKit.text).map { out =>
      assert(out.contains("\"truncated\": true"), s"missing truncated field:\n$out")
      assert(out.contains("\"totalRows\": 100"), s"missing totalRows field:\n$out")
      assert(out.contains("\"range\": \"A1:A50\""), "range should reflect emitted rows")
    }
  }

  test("view json: unclipped output has no truncation fields") {
    runView(wbWithRows(30), "A1:A30", 50, ViewFormat.Json).map(ReadTestKit.text).map { out =>
      assert(!out.contains("truncated"), s"unexpected truncated field:\n$out")
      assert(!out.contains("totalRows"), s"unexpected totalRows field:\n$out")
    }
  }

  test("view json: header-row records mode also carries truncation fields") {
    val sheet = (1 to 100).foldLeft(Sheet("Data").put(ARef.from0(0, 0), CellValue.Text("Col"))) {
      (s, i) => s.put(ARef.from0(0, i), CellValue.Number(BigDecimal(i)))
    }
    val wb = Workbook(Vector(sheet))
    ReadTestKit
      .readText(
        wb,
        Some("Data"),
        ReadTestKit.view(Some("A1:A101"), ViewFormat.Json, limit = 50, headerRow = Some(1))
      )
      .map { out =>
        assert(out.contains("\"records\""), s"expected records mode:\n$out")
        assert(out.contains("\"truncated\": true"), s"missing truncated field:\n$out")
        assert(out.contains("\"totalRows\": 101"), s"missing totalRows field:\n$out")
      }
  }

  test("view json: a header row outside the window still keys the records") {
    val sheet = (1 to 10).foldLeft(Sheet("Data").put(ARef.from0(0, 0), CellValue.Text("Col"))) {
      (s, i) => s.put(ARef.from0(0, i), CellValue.Number(BigDecimal(i)))
    }
    val wb = Workbook(Vector(sheet))
    ReadTestKit
      .readText(
        wb,
        Some("Data"),
        ReadTestKit.view(Some("A5:A7"), ViewFormat.Json, headerRow = Some(1))
      )
      .map { out =>
        assert(out.contains("""{"Col": 4}"""), s"header outside the window must still key:\n$out")
        assert(out.contains("""{"Col": 6}"""), out)
      }
  }

  test("view json: --header-row below 1 is a usage error, not a silent fallback to letters") {
    val wb = Workbook(Vector(Sheet("Data").put(ARef.from0(0, 0), CellValue.Text("Col"))))
    Vector(0, -3).foldLeft(IO.unit) { (acc, row) =>
      acc *> ReadTestKit
        .inMemory(
          wb,
          Some("Data"),
          ReadTestKit.view(Some("A1:A2"), ViewFormat.Json, headerRow = Some(row))
        )
        .map { outcome =>
          assertEquals(outcome.exitCode.code, 2, outcome.toString)
          assertEquals(outcome.error.map(_.code), Some("USAGE"))
          assert(outcome.error.exists(_.message.contains("--header-row must be 1 or more")))
        }
    }
  }

  // ========== view: html ==========

  test("view html: clipped output appends HTML comment trailer and warns") {
    runView(wbWithRows(100), "A1:A100", 50, ViewFormat.Html).map { outcome =>
      val out = ReadTestKit.text(outcome)
      assert(
        out.trim.endsWith(
          "<!-- … showing 50 of 100 rows (use --limit to raise; --limit 0 = no limit) -->"
        ),
        s"missing trailing HTML comment:\n${out.takeRight(200)}"
      )
      assert(
        truncationWarnings(outcome).exists(_.contains("… showing 50 of 100 rows")),
        s"expected truncation warning, got: ${outcome.warnings}"
      )
    }
  }

  test("view html: unclipped output has no comment or warning") {
    runView(wbWithRows(30), "A1:A30", 50, ViewFormat.Html).map { outcome =>
      assert(!ReadTestKit.text(outcome).contains("… showing"), "unexpected marker")
      assertEquals(outcome.warnings, Vector.empty)
    }
  }

  // ========== view: svg ==========

  test("view svg: stdout stays byte-identical, notice goes to the warnings") {
    val wb = wbWithRows(100)
    for
      clipped <- runView(wb, "A1:A100", 50, ViewFormat.Svg)
      exact <- runView(wb, "A1:A50", 50, ViewFormat.Svg)
    yield
      assertEquals(
        ReadTestKit.text(clipped),
        ReadTestKit.text(exact),
        "clipped SVG stdout must equal the unclipped 50-row render"
      )
      assert(!ReadTestKit.text(clipped).contains("showing"), "notice must not leak into SVG stdout")
      assert(
        truncationWarnings(clipped).exists(_.contains("… showing 50 of 100 rows")),
        s"expected truncation warning, got: ${clipped.warnings}"
      )
      assertEquals(exact.warnings, Vector.empty)
  }

  // ========== view: raster ==========

  test("view png: clipped export appends notice to the Exported status line") {
    BatikRasterizer.isAvailable.flatMap { batikAvailable =>
      assume(batikAvailable, "AWT not available - skipping raster truncation test")
      val tempFile = Files.createTempFile("xl-cli-truncation-", ".png")
      tempFile.toFile.deleteOnExit()
      ReadTestKit
        .readText(
          wbWithRows(100),
          Some("Data"),
          ReadTestKit
            .view(Some("A1:A100"), ViewFormat.Png, limit = 5, rasterOutput = Some(tempFile))
        )
        .map { out =>
          assert(out.contains("Exported:"), s"expected export status line:\n$out")
          assert(out.contains("… showing 5 of 100 rows"), s"missing raster notice:\n$out")
          assert(Files.size(tempFile) > 0, "PNG should have content")
        }
        .guarantee(IO(Files.deleteIfExists(tempFile)).void)
    }
  }

  // ========== search ==========

  private def search(wb: Workbook, limit: Int, exactTotal: Boolean = false): IO[String] =
    ReadTestKit.readText(wb, Some("Data"), ReadQuery.Search("\\d", limit, None, exactTotal))

  test("search: clipped hit list reports the total as a lower bound and a trailer (GH-637)") {
    search(wbWithRows(100), 10).map { out =>
      assert(out.contains("Found at least 11 matches"), s"expected a lower bound:\n$out")
      assert(out.contains("… showing first 10 matches; more exist"), s"missing trailer:\n$out")
      assert(out.contains("--total for the exact count"), s"missing --total hint:\n$out")
    }
  }

  test("search --total: clipped hit list reports the exact total and trailer") {
    search(wbWithRows(100), 10, exactTotal = true).map { out =>
      assert(out.contains("Found 100 matches"), s"expected true total count:\n$out")
      assert(out.contains("… showing 10 of 100 matches"), s"missing trailer:\n$out")
    }
  }

  test("search: unclipped hit list has no trailer") {
    search(wbWithRows(10), 50).map { out =>
      assert(out.contains("Found 10 matches"), s"expected count:\n$out")
      assert(!out.contains("… showing"), s"unexpected trailer:\n$out")
    }
  }

  test("search: --limit 0 returns all matches with no trailer") {
    search(wbWithRows(100), 0).map { out =>
      assert(out.contains("Found 100 matches"), s"expected all matches:\n$out")
      assert(out.contains("Data!A100"), "last match should be present with --limit 0")
      assert(!out.contains("… showing"), s"unexpected trailer:\n$out")
    }
  }

  test("search: matches are listed in row-major order") {
    val sheet = Sheet("Data")
      .put(ARef.from0(2, 0), CellValue.Number(BigDecimal(3)))
      .put(ARef.from0(0, 1), CellValue.Number(BigDecimal(4)))
      .put(ARef.from0(0, 0), CellValue.Number(BigDecimal(1)))
      .put(ARef.from0(1, 0), CellValue.Number(BigDecimal(2)))
    search(Workbook(Vector(sheet)), 0).map { out =>
      val refs = out.linesIterator.collect {
        case l if l.contains("Data!") => l.split("\\|")(1).trim
      }
      assertEquals(refs.toVector, Vector("Data!A1", "Data!B1", "Data!C1", "Data!A2"))
    }
  }

  // ========== streaming parity ==========

  test("streaming view markdown: clipped output appends the same trailer") {
    ReadTestKit.withTempWorkbook(wbWithRows(100)) { path =>
      ReadTestKit
        .streaming(path, Some("Data"), ReadTestKit.view(Some("A1:A100"), limit = 50))
        .map(ReadTestKit.text)
        .map(out => assert(out.contains("… showing 50 of 100 rows"), s"missing trailer:\n$out"))
    }
  }

  test("streaming view markdown: --limit 0 returns everything with no trailer") {
    ReadTestKit.withTempWorkbook(wbWithRows(100)) { path =>
      ReadTestKit
        .streaming(path, Some("Data"), ReadTestKit.view(Some("A1:A100"), limit = 0))
        .map(ReadTestKit.text)
        .map { out =>
          assert(out.contains("| 100 |"), "row 100 should be rendered with --limit 0")
          assert(!out.contains("… showing"), s"unexpected trailer:\n$out")
        }
    }
  }

  test("streaming view csv: stdout stays byte-identical, notice goes to the warnings") {
    ReadTestKit.withTempWorkbook(wbWithRows(100)) { path =>
      for
        clipped <- ReadTestKit.streaming(
          path,
          Some("Data"),
          ReadTestKit.view(Some("A1:A100"), ViewFormat.Csv, limit = 50)
        )
        exact <- ReadTestKit.streaming(
          path,
          Some("Data"),
          ReadTestKit.view(Some("A1:A50"), ViewFormat.Csv, limit = 50)
        )
      yield
        assertEquals(ReadTestKit.text(clipped), ReadTestKit.text(exact))
        assert(!ReadTestKit.text(clipped).contains("showing"))
        assert(truncationWarnings(clipped).exists(_.contains("… showing 50 of 100 rows")))
        assertEquals(exact.warnings, Vector.empty)
    }
  }

  test("streaming view json: the typed shape with in-band truncation fields, no warning") {
    ReadTestKit.withTempWorkbook(wbWithRows(100)) { path =>
      for
        streamed <- ReadTestKit.streaming(
          path,
          Some("Data"),
          ReadTestKit.view(Some("A1:A100"), ViewFormat.Json, limit = 50)
        )
        memory <- runView(wbWithRows(100), "A1:A100", 50, ViewFormat.Json)
      yield
        assertEquals(ReadTestKit.text(streamed), ReadTestKit.text(memory))
        assert(ReadTestKit.text(streamed).contains("\"truncated\": true"))
        assertEquals(streamed.warnings, Vector.empty)
    }
  }

  test("streaming search: the same lower bound, exact total and trailers as the in-memory search") {
    ReadTestKit.withTempWorkbook(wbWithRows(100)) { path =>
      for
        clipped <- ReadTestKit
          .streaming(path, Some("Data"), ReadQuery.Search("\\d", 10, None, exactTotal = false))
          .map(ReadTestKit.text)
        exact <- ReadTestKit
          .streaming(path, Some("Data"), ReadQuery.Search("\\d", 10, None, exactTotal = true))
          .map(ReadTestKit.text)
        all <- ReadTestKit
          .streaming(path, Some("Data"), ReadQuery.Search("\\d", 0, None, exactTotal = false))
          .map(ReadTestKit.text)
      yield
        assert(clipped.contains("Found at least 11 matches"), s"expected a lower bound:\n$clipped")
        assert(
          clipped.contains("… showing first 10 matches; more exist"),
          s"missing trailer:\n$clipped"
        )
        assert(exact.contains("Found 100 matches"), s"expected the true total:\n$exact")
        assert(exact.contains("… showing 10 of 100 matches"), s"missing trailer:\n$exact")
        assert(all.contains("Found 100 matches"), s"expected all matches:\n$all")
        assert(!all.contains("… showing"), s"unexpected clip notice:\n$all")
    }
  }
