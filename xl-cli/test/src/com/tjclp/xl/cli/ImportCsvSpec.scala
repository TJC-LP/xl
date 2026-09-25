package com.tjclp.xl.cli

import java.nio.file.{Files, Path}

import cats.effect.{IO, Resource}
import munit.CatsEffectSuite

import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.commands.ImportCommands
import com.tjclp.xl.cli.contract.{CliHarness, TestFixtures}
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.WriterConfig
import com.tjclp.xl.styles.numfmt.NumFmt
import com.tjclp.xl.workbooks.Workbook

/**
 * `import` through the in-process harness (GH-667): a CSV column typed as dates carries the date
 * format `put` gives the same text, so `view` shows the date and not its serial. Type inference
 * itself is CsvParserSpec's; this suite pins what reaches the written cells.
 */
class ImportCsvSpec extends CatsEffectSuite:

  private val fixtures = ResourceSuiteLocalFixture(
    "import-csv-fixtures",
    Resource.make(TestFixtures.materialize)(TestFixtures.delete)
  )

  override def munitFixtures = List(fixtures)

  private def file(name: String): String = fixtures().resolve(name).toString

  /**
   * `weaver-cli/d.csv` from the 0.23.0 dogfood: a header row, a blank amount, a negative, exponent
   * notation, a quoted comma and a blank flag around an ISO date column.
   */
  private val ledgerCsv = Seq(
    "name,amount,when,flag",
    "Alpha,1200.5,2026-01-15,true",
    "Beta,,2026-02-01,false",
    "Gamma,-30,2026-03-10,true",
    "\"Delta, Inc\",4.5e3,2026-04-01,"
  ).mkString("\n")

  private def writeCsv(name: String): IO[Path] =
    IO.blocking {
      val p = fixtures().resolve(name)
      Files.writeString(p, ledgerCsv)
      p
    }

  /** The number format of the cell's style in the written book, None for an unstyled cell. */
  private def numFmtAt(path: String, sheetName: String, ref: ARef): IO[Option[NumFmt]] =
    ExcelIO.instance[IO].read(Path.of(path)).map { wb =>
      wb.sheets.find(_.name.value == sheetName).flatMap { sheet =>
        sheet.cells.get(ref).flatMap(_.styleId).flatMap(sheet.styleRegistry.get).map(_.numFmt)
      }
    }

  private val dateText = """\d{1,2}/\d{1,2}/\d{2,4}""".r

  test("GH-667: import gives an ISO date column the date format `put` gives the same text") {
    val out = file("import-dates.xlsx")
    val viaPut = file("put-date.xlsx")
    for
      csv <- writeCsv("d.csv")
      run <- CliHarness.run("-f", file("single.xlsx"), "-o", out, "import", csv.toString)
      when <- numFmtAt(out, "Sheet1", ref"C2")
      lastWhen <- numFmtAt(out, "Sheet1", ref"C5")
      amount <- numFmtAt(out, "Sheet1", ref"B2")
      header <- numFmtAt(out, "Sheet1", ref"C1")
      view <- CliHarness.run("-f", out, "view", "C1:C5")
      put <- CliHarness.run("-f", file("single.xlsx"), "-o", viaPut, "put", "C2", "2026-01-15")
      putView <- CliHarness.run("-f", viaPut, "view", "C2:C2")
    yield
      assertEquals(run.exit, 0, run.stderr)
      assertEquals(when, Some(NumFmt.Date), "the `when` column carries the date format")
      assertEquals(lastWhen, Some(NumFmt.Date), "every date in the column is formatted")
      assertEquals(amount, None, "a plain number is written unstyled, as before")
      assertEquals(header, None, "the text header of a date column is written unstyled")
      assertEquals(view.exit, 0, view.stderr)
      assert(!view.stdout.contains("46037"), s"the date's serial must not show:\n${view.stdout}")
      assertEquals(put.exit, 0, put.stderr)
      val expected = dateText.findFirstIn(putView.stdout).getOrElse(fail(putView.stdout))
      assert(view.stdout.contains(expected), s"expected $expected in:\n${view.stdout}")
  }

  test("GH-667: import --new-sheet formats the date column the same way") {
    val out = file("import-dates-new-sheet.xlsx")
    for
      csv <- writeCsv("d2.csv")
      run <- CliHarness.run(
        "-f",
        file("simple.xlsx"),
        "-o",
        out,
        "import",
        csv.toString,
        "--new-sheet",
        "Ledger"
      )
      when <- numFmtAt(out, "Ledger", ref"C3")
      name <- numFmtAt(out, "Ledger", ref"A3")
    yield
      assertEquals(run.exit, 0, run.stderr)
      assertEquals(when, Some(NumFmt.Date))
      assertEquals(name, None)
  }

  test("GH-667: --no-type-inference keeps the dates as text with no format") {
    val out = file("import-dates-text.xlsx")
    for
      csv <- writeCsv("d3.csv")
      run <- CliHarness.run(
        "-f",
        file("single.xlsx"),
        "-o",
        out,
        "import",
        csv.toString,
        "--no-type-inference"
      )
      when <- numFmtAt(out, "Sheet1", ref"C2")
      cell <- CliHarness.run("-f", out, "cell", "C2")
    yield
      assertEquals(run.exit, 0, run.stderr)
      assertEquals(when, None)
      assert(cell.stdout.contains("2026-01-15"), cell.stdout)
  }

  // ===== GH-675: the O(1) path — `import --stream --new-sheet` into a NEW workbook =====
  //
  // `importCsv` takes the true streaming branch only for an EMPTY workbook (no sheets to
  // preserve). The `xl` verb always loads `-f` (a write verb requires it, and every readable book
  // has a sheet), so these call the command directly, as StreamingWriteSpec does.

  /** `import <csv> --new-sheet <name>` into a workbook with no sheets; the command's message. */
  private def importFresh(
    csv: Path,
    out: String,
    sheet: String,
    stream: Boolean,
    noTypeInference: Boolean = false
  ): IO[String] =
    ImportCommands.importCsv(
      Workbook(Vector.empty),
      None,
      csv.toString,
      None,
      delimiter = ',',
      skipHeader = false,
      encoding = "UTF-8",
      newSheetName = Some(sheet),
      noTypeInference = noTypeInference,
      Path.of(out),
      WriterConfig.default,
      stream = stream
    )

  /** Every cell of `sheetName` → (value, resolved numFmt): what the reader hands a consumer. */
  private def cellsOf(path: String, sheetName: String) =
    ExcelIO.instance[IO].read(Path.of(path)).map { wb =>
      wb.sheets.find(_.name.value == sheetName).fold(Map.empty) { sheet =>
        sheet.cells.map { (ref, cell) =>
          ref -> (cell.value, cell.styleId.flatMap(sheet.styleRegistry.get).map(_.numFmt))
        }
      }
    }

  test("GH-675: import --stream --new-sheet into a new workbook formats the date column") {
    val out = file("import-stream-new.xlsx")
    val viaPut = file("put-date-stream.xlsx")
    for
      csv <- writeCsv("d-stream.csv")
      message <- importFresh(csv, out, "Ledger", stream = true)
      when <- numFmtAt(out, "Ledger", ref"C2")
      lastWhen <- numFmtAt(out, "Ledger", ref"C5")
      header <- numFmtAt(out, "Ledger", ref"C1")
      amount <- numFmtAt(out, "Ledger", ref"B2")
      view <- CliHarness.run("-f", out, "view", "C1:C5")
      put <- CliHarness.run("-f", file("single.xlsx"), "-o", viaPut, "put", "C2", "2026-01-15")
      putView <- CliHarness.run("-f", viaPut, "view", "C2:C2")
    yield
      assert(message.contains("Streamed:"), s"the O(1) path must run:\n$message")
      assertEquals(when, Some(NumFmt.Date), "the `when` column carries the date format")
      assertEquals(lastWhen, Some(NumFmt.Date), "every date in the column is formatted")
      assertEquals(header, None, "the text header is written unstyled")
      assertEquals(amount, None, "a plain number is written unstyled")
      assertEquals(view.exit, 0, view.stderr)
      assert(!view.stdout.contains("46037"), s"the date's serial must not show:\n${view.stdout}")
      assertEquals(put.exit, 0, put.stderr)
      val expected = dateText.findFirstIn(putView.stdout).getOrElse(fail(putView.stdout))
      assert(view.stdout.contains(expected), s"expected $expected in:\n${view.stdout}")
  }

  test("GH-675: stream and in-memory import of the same CSV agree on every value's format") {
    val streamed = file("import-parity-stream.xlsx")
    val inMemory = file("import-parity-memory.xlsx")
    for
      csv <- writeCsv("d-parity.csv")
      s <- importFresh(csv, streamed, "L", stream = true)
      m <- importFresh(csv, inMemory, "L", stream = false)
      streamedCells <- cellsOf(streamed, "L")
      memoryCells <- cellsOf(inMemory, "L")
      streamedView <- CliHarness.run("-f", streamed, "view", "C1:C5")
      memoryView <- CliHarness.run("-f", inMemory, "view", "C1:C5")
    yield
      assert(s.contains("Streamed:"), s)
      assert(m.contains("Imported:"), m)
      val refs = (streamedCells.keySet ++ memoryCells.keySet).toVector.sortBy(r =>
        (r.row.index0, r.col.index0)
      )
      refs.foreach { ref =>
        assertEquals(
          streamedCells.get(ref).flatMap(_._2),
          memoryCells.get(ref).flatMap(_._2),
          s"numFmt at ${ref.toA1}"
        )
      }
      val dates = memoryCells.collect { case (ref, (_, Some(NumFmt.Date))) => ref }
      assertEquals(dates.size, 4, "the four dates are the date-formatted cells")
      dates.foreach(ref =>
        assertEquals(streamedCells.get(ref).map(_._1), memoryCells.get(ref).map(_._1), ref.toA1)
      )
      assertEquals(streamedView.exit, 0, streamedView.stderr)
      assertEquals(streamedView.stdout, memoryView.stdout, "the displayed dates are the same")
  }

  test("GH-675: --stream --no-type-inference keeps the dates as text with no style") {
    val out = file("import-stream-text.xlsx")
    for
      csv <- writeCsv("d-stream-text.csv")
      message <- importFresh(csv, out, "Ledger", stream = true, noTypeInference = true)
      cells <- cellsOf(out, "Ledger")
    yield
      assert(message.contains("Streamed:"), message)
      assertEquals(cells.get(ref"C2").map(_._1), Some(CellValue.Text("2026-01-15")))
      assertEquals(cells.values.flatMap(_._2).toVector, Vector.empty, "no cell is styled")
  }

  test("GH-681: import --stream types a CSV by column, exactly as import does") {
    // a mostly-numeric column holding one `true`: the column model makes it Number and keeps the
    // odd value as text, where a per-value model would read it as a boolean (the docs' claim)
    val csv = (1 to 9).map(_.toString).mkString("n\n", "\n", "\ntrue\n")
    def cells(path: String) =
      ExcelIO.instance[IO].read(Path.of(path)).map { wb =>
        wb.sheets.find(_.name.value == "Csv").fold(Vector.empty) { sheet =>
          (0 to 10).toVector.map { r =>
            sheet.cells.get(ARef.from0(0, r)).map(c => (c.value, c.styleId))
          }
        }
      }
    for
      path <- IO.blocking(Files.writeString(fixtures().resolve("column.csv"), csv))
      base = List("-f", file("simple.xlsx"), "import", path.toString, "--new-sheet", "Csv")
      plain <- CliHarness.run(base ++ List("-o", file("import-column.xlsx"))*)
      streamed <- CliHarness.run(
        base ++ List("-o", file("import-column-stream.xlsx"), "--stream")*
      )
      plainCells <- cells(file("import-column.xlsx"))
      streamedCells <- cells(file("import-column-stream.xlsx"))
    yield
      assertEquals(plain.exit, 0, plain.stderr)
      assertEquals(streamed.exit, 0, streamed.stderr)
      assertEquals(streamedCells, plainCells)
      assertEquals(
        plainCells.lift(10).flatten.map(_._1),
        Some(com.tjclp.xl.cells.CellValue.Text("true"))
      )
  }
