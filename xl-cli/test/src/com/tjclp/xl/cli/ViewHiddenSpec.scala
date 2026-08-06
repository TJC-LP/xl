package com.tjclp.xl.cli

import java.io.{ByteArrayOutputStream, PrintStream}
import java.nio.charset.StandardCharsets
import java.nio.file.Files

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.{Sheet, Workbook, given}
import com.tjclp.xl.addressing.{ARef, Column, Row}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.commands.{ReadCommands, StreamingReadCommands}
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.sheets.{ColumnProperties, RowProperties}

/**
 * GH-474: `view` must not silently elide hidden rows/columns from an explicitly requested range.
 *
 * The asymmetry that burned agents: `search` finds H!C5 and `cell C5` reads it, but `view A1:E10`
 * rendered neither the cell nor any marker — so "view shows empty, search finds a value" read as
 * corruption. Contract pinned here:
 *   - hidden rows/columns inside an explicitly requested range are INCLUDED by default
 *   - every format carries a marker: markdown trailer, CSV stderr note, JSON structured fields
 *   - `--skip-hidden` restores the elision, and then says what it elided
 */
class ViewHiddenSpec extends CatsEffectSuite:

  /** 5x5 grid A1:E5 where every cell holds "<A1-ref>", with row 5 and column C hidden. */
  private def hiddenSheet: Sheet =
    val filled = (0 until 5).foldLeft(Sheet("H")) { (s, col) =>
      (0 until 5).foldLeft(s) { (s2, row) =>
        val ref = ARef.from0(col, row)
        s2.put(ref, CellValue.Text(ref.toA1))
      }
    }
    filled
      .setRowProperties(Row.from0(4), RowProperties(hidden = true))
      .setColumnProperties(Column.from0(2), ColumnProperties(hidden = true))

  private def wb: Workbook = Workbook(Vector(hiddenSheet))

  private def runView(
    range: String,
    format: ViewFormat,
    skipHidden: Boolean = false,
    showLabels: Boolean = false
  ): IO[String] =
    val book = wb
    ReadCommands.view(
      book,
      book.sheets.headOption,
      range,
      showFormulas = false,
      evalFormulas = false,
      strict = false,
      limit = 100,
      format = format,
      printScale = false,
      showGridlines = false,
      showLabels = showLabels,
      dpi = 96,
      quality = 90,
      rasterOutput = None,
      skipEmpty = false,
      headerRow = None,
      skipHidden = skipHidden
    )

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

  // ========== markdown ==========

  test("GH-474: markdown view includes the hidden row and column by default") {
    runView("A1:E5", ViewFormat.Markdown).map { out =>
      assert(out.contains("C5"), s"hidden-row/hidden-column cell C5 must be rendered:\n$out")
      assert(out.contains("| C "), s"hidden column C header must be rendered:\n$out")
      assert(out.contains("A5"), s"hidden row 5 must be rendered:\n$out")
    }
  }

  test("GH-474: markdown view marks the hidden row/column in a trailer") {
    runView("A1:E5", ViewFormat.Markdown).map { out =>
      val lines = out.linesIterator.toVector
      assert(
        lines.lastOption.exists(_.startsWith("note: ")),
        s"hidden notice must be the trailing line, got: ${lines.lastOption}"
      )
      assert(out.contains("hidden row(s) 5"), s"notice must name row 5:\n$out")
      assert(out.contains("column(s) C"), s"notice must name column C:\n$out")
      assert(out.contains("--skip-hidden"), s"notice must name the opt-out flag:\n$out")
    }
  }

  test("GH-474: markdown view with no hidden rows/columns has no trailer") {
    runView("A1:B4", ViewFormat.Markdown).map { out =>
      assert(!out.contains("note: "), s"unexpected notice:\n$out")
    }
  }

  test("GH-474: --skip-hidden restores elision AND says what it elided") {
    runView("A1:E5", ViewFormat.Markdown, skipHidden = true).map { out =>
      assert(!out.contains("C5"), s"--skip-hidden must omit hidden cells:\n$out")
      assert(!out.contains("| C "), s"--skip-hidden must omit the hidden column:\n$out")
      assert(out.contains("omitted hidden"), s"elision marker missing:\n$out")
      assert(out.contains("row(s) 5"), s"elision marker must name row 5:\n$out")
      assert(out.contains("column(s) C"), s"elision marker must name column C:\n$out")
    }
  }

  // ========== csv ==========

  test("GH-474: csv view includes hidden cells; the note goes to stderr only") {
    captureStderr(runView("A1:E5", ViewFormat.Csv, showLabels = true)).map { case (out, err) =>
      assert(out.contains("C5"), s"hidden cell C5 must appear in CSV:\n$out")
      assert(out.linesIterator.next().contains("C"), s"hidden column label must appear:\n$out")
      assert(!out.contains("note:"), s"CSV stdout must stay machine-parseable:\n$out")
      assert(err.contains("hidden row(s) 5"), s"stderr note missing row 5: $err")
      assert(err.contains("column(s) C"), s"stderr note missing column C: $err")
    }
  }

  test("GH-474: csv view with --skip-hidden elides and notes on stderr") {
    captureStderr(runView("A1:E5", ViewFormat.Csv, skipHidden = true)).map { case (out, err) =>
      assert(!out.contains("C5"), s"--skip-hidden must omit hidden cells:\n$out")
      assert(err.contains("omitted hidden"), s"stderr elision note missing: $err")
    }
  }

  // ========== json ==========

  test("GH-474: json view never drops an addressed cell and flags the hidden lines") {
    runView("A1:E5", ViewFormat.Json).map { out =>
      assert(out.contains("\"C5\""), s"JSON must carry the addressed hidden cell C5:\n$out")
      assert(out.contains("\"hiddenRows\": [5]"), s"JSON must flag hidden row 5:\n$out")
      assert(out.contains("\"hiddenCols\": [\"C\"]"), s"JSON must flag hidden column C:\n$out")
    }
  }

  test("GH-474: json view with no hidden lines omits the hidden fields") {
    runView("A1:B4", ViewFormat.Json).map { out =>
      assert(!out.contains("hiddenRows"), s"unexpected hiddenRows field:\n$out")
      assert(!out.contains("hiddenCols"), s"unexpected hiddenCols field:\n$out")
    }
  }

  test("GH-474: json --skip-hidden drops the cells but still reports what it dropped") {
    runView("A1:E5", ViewFormat.Json, skipHidden = true).map { out =>
      assert(!out.contains("\"C5\""), s"--skip-hidden must omit hidden cells:\n$out")
      assert(out.contains("\"hiddenRows\": [5]"), s"JSON must still flag hidden row 5:\n$out")
      assert(
        out.contains("\"hiddenCols\": [\"C\"]"),
        s"JSON must still flag hidden column C:\n$out"
      )
    }
  }

  // ========== streaming (--stream) ==========

  /** Write the hidden-line fixture to a temp .xlsx and run the streaming view over it. */
  private def runStreamingView(skipHidden: Boolean): IO[(String, String)] =
    IO.blocking {
      val tmp = Files.createTempFile("xl-view-hidden-stream-", ".xlsx")
      tmp.toFile.deleteOnExit()
      tmp
    }.flatMap { tmp =>
      ExcelIO.instance[IO].write(wb, tmp) *>
        captureStderr(
          StreamingReadCommands.view(
            tmp,
            Some("H"),
            "A1:E5",
            showFormulas = false,
            limit = 100,
            format = ViewFormat.Csv,
            showLabels = false,
            skipEmpty = false,
            headerRow = None,
            skipHidden = skipHidden
          )
        )
    }

  test("GH-474: --stream view with --skip-hidden is not a silent no-op") {
    runStreamingView(skipHidden = true).map { case (out, err) =>
      // Streaming does not read row/column properties, so it renders everything...
      assert(out.contains("C5"), s"streaming renders every addressed cell:\n$out")
      // ...but it must SAY so rather than accept the flag and ignore it.
      assert(err.nonEmpty, "streaming --skip-hidden must not be silent")
      assert(err.contains("--skip-hidden"), s"note must name the flag: '$err'")
      assert(err.contains("--stream"), s"note must name the mode: '$err'")
    }
  }

  test("GH-474: --stream view without --skip-hidden stays quiet") {
    runStreamingView(skipHidden = false).map { case (_, err) =>
      assertEquals(err, "", s"unexpected stderr: '$err'")
    }
  }
