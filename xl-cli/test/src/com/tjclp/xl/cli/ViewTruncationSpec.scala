package com.tjclp.xl.cli

import java.io.{ByteArrayOutputStream, PrintStream}
import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.{Sheet, Workbook, given}
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.commands.{ReadCommands, StreamingReadCommands}
import com.tjclp.xl.io.ExcelIO

/**
 * GH-351: view/search must report when --limit clips output.
 *
 * Pins the truncation marker semantics per format:
 *   - markdown: trailer line after the table ("… showing X of Y rows")
 *   - csv: stdout stays byte-identical, notice goes to stderr
 *   - json: top-level "truncated"/"totalRows" fields (only when clipped)
 *   - search: total match count + trailer when the hit list is clipped
 *   - --limit 0 means "no limit"
 */
class ViewTruncationSpec extends CatsEffectSuite:

  /** Sheet with n numeric rows in column A (A1=1 .. An=n). */
  private def sheetWithRows(n: Int): Sheet =
    (1 to n).foldLeft(Sheet("Data")) { (s, i) =>
      s.put(ARef.from0(0, i - 1), CellValue.Number(BigDecimal(i)))
    }

  private def wbWithRows(n: Int): Workbook = Workbook(Vector(sheetWithRows(n)))

  private def runView(
    wb: Workbook,
    range: String,
    limit: Int,
    format: ViewFormat
  ): IO[String] =
    ReadCommands.view(
      wb,
      wb.sheets.headOption,
      range,
      showFormulas = false,
      evalFormulas = false,
      strict = false,
      limit = limit,
      format = format,
      printScale = false,
      showGridlines = false,
      showLabels = false,
      dpi = 96,
      quality = 90,
      rasterOutput = None,
      skipEmpty = false,
      headerRow = None
    )

  /** Capture System.err produced while running the given IO. */
  private def captureStderr[A](io: IO[A]): IO[(A, String)] =
    IO.blocking {
      val baos = new ByteArrayOutputStream()
      val prevErr = System.err
      System.setErr(new PrintStream(baos, true, "UTF-8"))
      try
        import cats.effect.unsafe.implicits.global
        val result = io.unsafeRunSync()
        (result, new String(baos.toByteArray, StandardCharsets.UTF_8))
      finally System.setErr(prevErr)
    }

  // ========== view: markdown ==========

  test("view markdown: clipped output appends trailer with correct counts") {
    runView(wbWithRows(100), "A1:A100", 50, ViewFormat.Markdown).map { out =>
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
    runView(wbWithRows(30), "A1:A30", 50, ViewFormat.Markdown).map { out =>
      assert(!out.contains("… showing"), s"unexpected trailer:\n$out")
    }
  }

  test("view markdown: exact-limit output has no trailer") {
    runView(wbWithRows(50), "A1:A50", 50, ViewFormat.Markdown).map { out =>
      assert(!out.contains("… showing"), s"unexpected trailer:\n$out")
    }
  }

  test("view markdown: --limit 0 returns everything with no trailer") {
    runView(wbWithRows(100), "A1:A100", 0, ViewFormat.Markdown).map { out =>
      assert(out.contains("| 100"), "row 100 should be rendered with --limit 0")
      assert(!out.contains("… showing"), s"unexpected trailer:\n$out")
    }
  }

  // ========== view: csv ==========

  test("view csv: stdout stays byte-identical, notice goes to stderr") {
    val wb = wbWithRows(100)
    val io =
      for
        clipped <- runView(wb, "A1:A100", 50, ViewFormat.Csv)
        exact <- runView(wb, "A1:A50", 50, ViewFormat.Csv)
      yield (clipped, exact)
    captureStderr(io).map { case ((clipped, exact), stderr) =>
      assertEquals(clipped, exact, "clipped CSV stdout must equal the unclipped 50-row render")
      assert(!clipped.contains("showing"), "notice must not leak into CSV stdout")
      assert(
        stderr.contains("… showing 50 of 100 rows"),
        s"expected truncation notice on stderr, got: '$stderr'"
      )
    }
  }

  test("view csv: no stderr notice when not clipped") {
    captureStderr(runView(wbWithRows(30), "A1:A30", 50, ViewFormat.Csv)).map { case (_, stderr) =>
      assert(!stderr.contains("showing"), s"unexpected stderr notice: '$stderr'")
    }
  }

  // ========== view: json ==========

  test("view json: clipped output carries truncated/totalRows fields") {
    runView(wbWithRows(100), "A1:A100", 50, ViewFormat.Json).map { out =>
      assert(out.contains("\"truncated\": true"), s"missing truncated field:\n$out")
      assert(out.contains("\"totalRows\": 100"), s"missing totalRows field:\n$out")
      assert(out.contains("\"range\": \"A1:A50\""), "range should reflect emitted rows")
    }
  }

  test("view json: unclipped output has no truncation fields") {
    runView(wbWithRows(30), "A1:A30", 50, ViewFormat.Json).map { out =>
      assert(!out.contains("truncated"), s"unexpected truncated field:\n$out")
      assert(!out.contains("totalRows"), s"unexpected totalRows field:\n$out")
    }
  }

  test("view json: header-row records mode also carries truncation fields") {
    val sheet = (1 to 100).foldLeft(Sheet("Data").put(ARef.from0(0, 0), CellValue.Text("Col"))) {
      (s, i) => s.put(ARef.from0(0, i), CellValue.Number(BigDecimal(i)))
    }
    val wb = Workbook(Vector(sheet))
    ReadCommands
      .view(
        wb,
        wb.sheets.headOption,
        "A1:A101",
        showFormulas = false,
        evalFormulas = false,
        strict = false,
        limit = 50,
        format = ViewFormat.Json,
        printScale = false,
        showGridlines = false,
        showLabels = false,
        dpi = 96,
        quality = 90,
        rasterOutput = None,
        skipEmpty = false,
        headerRow = Some(1)
      )
      .map { out =>
        assert(out.contains("\"records\""), s"expected records mode:\n$out")
        assert(out.contains("\"truncated\": true"), s"missing truncated field:\n$out")
        assert(out.contains("\"totalRows\": 101"), s"missing totalRows field:\n$out")
      }
  }

  // ========== search ==========

  test("search: clipped hit list reports total count and trailer") {
    val wb = wbWithRows(100) // every cell matches \d
    ReadCommands.search(wb, wb.sheets.headOption, "\\d", limit = 10, sheetsFilter = None).map {
      out =>
        assert(out.contains("Found 100 matches"), s"expected true total count:\n$out")
        assert(out.contains("… showing 10 of 100 matches"), s"missing trailer:\n$out")
    }
  }

  test("search: unclipped hit list has no trailer") {
    val wb = wbWithRows(10)
    ReadCommands.search(wb, wb.sheets.headOption, "\\d", limit = 50, sheetsFilter = None).map {
      out =>
        assert(out.contains("Found 10 matches"), s"expected count:\n$out")
        assert(!out.contains("… showing"), s"unexpected trailer:\n$out")
    }
  }

  test("search: --limit 0 returns all matches with no trailer") {
    val wb = wbWithRows(100)
    ReadCommands.search(wb, wb.sheets.headOption, "\\d", limit = 0, sheetsFilter = None).map {
      out =>
        assert(out.contains("Found 100 matches"), s"expected all matches:\n$out")
        assert(out.contains("Data!A100"), "last match should be present with --limit 0")
        assert(!out.contains("… showing"), s"unexpected trailer:\n$out")
    }
  }

  // ========== streaming parity ==========

  private def withTempWorkbook[A](wb: Workbook)(test: Path => IO[A]): IO[A] =
    IO.blocking {
      val tempFile = Files.createTempFile("xl-cli-truncation-", ".xlsx")
      tempFile.toFile.deleteOnExit()
      tempFile
    }.flatMap { tempFile =>
      ExcelIO.instance[IO].write(wb, tempFile) *> test(tempFile)
    }

  test("streaming view markdown: clipped output appends the same trailer") {
    withTempWorkbook(wbWithRows(100)) { path =>
      StreamingReadCommands
        .view(
          path,
          Some("Data"),
          "A1:A100",
          showFormulas = false,
          limit = 50,
          format = ViewFormat.Markdown,
          showLabels = true,
          skipEmpty = false,
          headerRow = None
        )
        .map { out =>
          assert(out.contains("… showing 50 of 100 rows"), s"missing trailer:\n$out")
        }
    }
  }

  test("streaming view markdown: --limit 0 returns everything with no trailer") {
    withTempWorkbook(wbWithRows(100)) { path =>
      StreamingReadCommands
        .view(
          path,
          Some("Data"),
          "A1:A100",
          showFormulas = false,
          limit = 0,
          format = ViewFormat.Markdown,
          showLabels = true,
          skipEmpty = false,
          headerRow = None
        )
        .map { out =>
          assert(out.contains("| 100 |"), "row 100 should be rendered with --limit 0")
          assert(!out.contains("… showing"), s"unexpected trailer:\n$out")
        }
    }
  }

  test("streaming search: hitting the limit flags a possible clip") {
    withTempWorkbook(wbWithRows(100)) { path =>
      StreamingReadCommands
        .search(path, Some("Data"), "\\d", limit = 10, sheetsFilter = None)
        .map { out =>
          assert(out.contains("Found 10 matches"), s"expected clipped count:\n$out")
          assert(
            out.contains("… showing first 10 matches"),
            s"missing streaming clip notice:\n$out"
          )
        }
    }
  }

  test("streaming search: --limit 0 scans everything with no clip notice") {
    withTempWorkbook(wbWithRows(100)) { path =>
      StreamingReadCommands
        .search(path, Some("Data"), "\\d", limit = 0, sheetsFilter = None)
        .map { out =>
          assert(out.contains("Found 100 matches"), s"expected all matches:\n$out")
          assert(!out.contains("… showing"), s"unexpected clip notice:\n$out")
        }
    }
  }
