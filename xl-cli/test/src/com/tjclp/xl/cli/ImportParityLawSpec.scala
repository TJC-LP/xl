package com.tjclp.xl.cli

import java.nio.file.{Files, Path}
import java.time.LocalDate
import java.time.format.DateTimeFormatter

import cats.effect.IO
import cats.effect.unsafe.implicits.global
import munit.{FunSuite, ScalaCheckSuite}
import org.scalacheck.{Gen, Prop}

import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.commands.ImportCommands
import com.tjclp.xl.cli.contract.CliHarness
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.ooxml.WriterConfig
import com.tjclp.xl.styles.numfmt.NumFmt
import com.tjclp.xl.workbooks.Workbook

/**
 * GH-675's parity law: `import --new-sheet` of one CSV into a workbook with no sheets writes the
 * same cells whether it streams (the O(1) branch, `writeStreamStyledWithAutoDetect`) or loads in
 * memory — every cell's value and resolved number format agree, and `view` prints the same table.
 *
 * The domain is well-formed columns: a header row, a column of ISO dates (1900-03-01 to 9999-12-31,
 * past the 1900 leap-year bug), and optionally an integer column and a plain-text column, in any
 * order. EXCLUDED: a column mixing dates with text, or holding an invalid date such as 2023-02-29.
 * Those diverge by design of the two inference strategies, not by this change — the in-memory path
 * samples the whole column (header included) and degrades it to text, while the streaming path
 * types each cell on its own.
 */
class ImportParityLawSpec extends FunSuite with ScalaCheckSuite:

  override def scalaCheckTestParameters =
    super.scalaCheckTestParameters.withMinSuccessfulTests(15)

  private val minDate = LocalDate.of(1900, 3, 1)
  private val maxDate = LocalDate.of(9999, 12, 31)

  private val genDate: Gen[String] =
    Gen
      .frequency(
        4 -> Gen.choose(LocalDate.of(1990, 1, 1).toEpochDay, LocalDate.of(2040, 12, 31).toEpochDay),
        1 -> Gen.choose(minDate.toEpochDay, maxDate.toEpochDay),
        1 -> Gen.oneOf(minDate.toEpochDay, maxDate.toEpochDay, LocalDate.of(2024, 2, 29).toEpochDay)
      )
      .map(d => LocalDate.ofEpochDay(d).format(DateTimeFormatter.ISO_LOCAL_DATE))

  private val genInt: Gen[String] = Gen.choose(-100000, 100000).map(_.toString)

  /** Letters only, `x`-prefixed: never a number, boolean, date or quoted field. */
  private val genText: Gen[String] =
    Gen.alphaStr.map(s => "x" + s.take(10))

  /** (header, cell generator) per column; the date column is always present. */
  private val genColumns: Gen[Vector[(String, Gen[String])]] =
    for
      withInt <- Gen.oneOf(true, false)
      withText <- Gen.oneOf(true, false)
      cols = Vector(Some("when" -> genDate), Option.when(withInt)("qty" -> genInt)).flatten ++
        Option.when(withText)("label" -> genText)
      ordered <- Gen.oneOf(cols.permutations.toVector)
    yield ordered

  private val genCsv: Gen[(Int, Int, String)] =
    for
      cols <- genColumns
      n <- Gen.choose(1, 8)
      rows <- Gen.listOfN(n, Gen.sequence[Vector[String], String](cols.map(_._2)))
    yield
      val lines = cols.map(_._1).mkString(",") +: rows.map(_.mkString(","))
      (cols.size, n + 1, lines.mkString("\n"))

  private def importFresh(csv: Path, out: Path, stream: Boolean): IO[String] =
    ImportCommands.importCsv(
      Workbook(Vector.empty),
      None,
      csv.toString,
      None,
      delimiter = ',',
      skipHeader = false,
      encoding = "UTF-8",
      newSheetName = Some("L"),
      noTypeInference = false,
      out,
      WriterConfig.default,
      stream = stream
    )

  private def cellsOf(path: Path): IO[Map[ARef, (CellValue, Option[NumFmt])]] =
    ExcelIO.instance[IO].read(path).map { wb =>
      wb.sheets.find(_.name.value == "L").fold(Map.empty) { sheet =>
        sheet.cells.map { (ref, cell) =>
          ref -> (cell.value, cell.styleId.flatMap(sheet.styleRegistry.get).map(_.numFmt))
        }
      }
    }

  property("GH-675: stream and in-memory import agree on every cell and on `view`") {
    Prop.forAll(genCsv) { (width, height, text) =>
      val dir = Files.createTempDirectory("import-parity")
      val csv = Files.writeString(dir.resolve("in.csv"), text)
      val streamed = dir.resolve("stream.xlsx")
      val inMemory = dir.resolve("memory.xlsx")
      val range = s"A1:${ARef.from1(width, height).toA1}"
      val io =
        for
          s <- importFresh(csv, streamed, stream = true)
          m <- importFresh(csv, inMemory, stream = false)
          streamedCells <- cellsOf(streamed)
          memoryCells <- cellsOf(inMemory)
          streamedView <- CliHarness.run("-f", streamed.toString, "view", range)
          memoryView <- CliHarness.run("-f", inMemory.toString, "view", range)
        yield
          assert(s.contains("Streamed:"), s"the O(1) path must run:\n$s")
          assert(m.contains("Imported:"), m)
          assertEquals(streamedCells, memoryCells, s"cells differ for CSV:\n$text")
          assert(
            memoryCells.values.exists(_._2.contains(NumFmt.Date)),
            s"the date column is date-formatted:\n$text"
          )
          assertEquals(streamedView.exit, 0, streamedView.stderr)
          assertEquals(streamedView.stdout, memoryView.stdout, s"view differs for CSV:\n$text")
      try
        io.unsafeRunSync()
        true
      finally
        Files.walk(dir).sorted(java.util.Comparator.reverseOrder()).forEach(p => Files.delete(p))
    }
  }
