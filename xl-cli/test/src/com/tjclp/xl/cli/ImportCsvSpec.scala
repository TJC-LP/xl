package com.tjclp.xl.cli

import java.nio.file.{Files, Path}

import cats.effect.{IO, Resource}
import munit.CatsEffectSuite

import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cli.contract.{CliHarness, TestFixtures}
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.macros.ref
import com.tjclp.xl.styles.numfmt.NumFmt

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
