package com.tjclp.xl.cli.batch

import java.nio.file.{Files, Path}

import cats.effect.{IO, Resource}
import cats.syntax.all.*
import munit.CatsEffectSuite

import com.tjclp.xl.{Sheet, Workbook}
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.contract.{CliHarness, CliRun}
import com.tjclp.xl.extensions.style
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.sheets.syntax.*
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.font.Font
import com.tjclp.xl.styles.numfmt.NumFmt

/**
 * ADR-017 invariant 2 meets invariant 4: under `--stream`, a formatted `put`/`putf` must land on
 * exactly the style the in-memory path produces — an Explicit format replaces the numFmt on the
 * cell's existing style (font, fill, borders kept), an Inferred hint applies only onto General.
 * Every case runs both paths through real files and compares them.
 */
class BatchStreamFormatSpec extends CatsEffectSuite:

  private val percentCode = """0.0%_);\(0.0%\)"""

  /** What a cell's style contributes here: its number format and whether its font is bold. */
  private final case class Facts(value: Option[CellValue], numFmt: Option[NumFmt], bold: Boolean)

  private def bold(numFmt: NumFmt): CellStyle =
    CellStyle.default.withFont(Font.default.withBold()).withNumFmt(numFmt)

  /** Row 1 and row 2, four bold cells each, alternating a Custom percent code and Currency. */
  private def fixture: Workbook =
    val refs = for row <- 0 to 1; col <- 0 to 3 yield ARef.from0(col, row)
    val data = refs.zipWithIndex.foldLeft(Sheet("Data")) { case (sheet, (ref, i)) =>
      val numFmt = if i % 2 == 0 then NumFmt.Custom(percentCode) else NumFmt.Currency
      sheet.put(ref, CellValue.Number(BigDecimal(i + 1))).style(ref, bold(numFmt))
    }
    Workbook(Vector(data, Sheet("Other")))

  private val dir = ResourceSuiteLocalFixture(
    "stream-format-dir",
    Resource.make(IO.blocking(Files.createTempDirectory("xl-stream-format-"))) { d =>
      IO.blocking {
        val entries = Files.list(d)
        try entries.forEach(p => Files.deleteIfExists(p))
        finally entries.close()
        Files.deleteIfExists(d)
      }.void
    }
  )

  override def munitFixtures = List(dir)

  private def input(name: String): IO[Path] =
    val path = dir().resolve(s"$name-in.xlsx")
    ExcelIO.instance[IO].write(fixture, path).as(path)

  private def batch(in: Path, out: Path, json: String, stream: Boolean): IO[CliRun] =
    val streamFlag = if stream then List("--stream") else Nil
    CliHarness.run(
      List("-f", in.toString, "-s", "Data", "-o", out.toString) ++ streamFlag ++ List("batch", "-"),
      json
    )

  private def facts(path: Path, refs: Vector[ARef]): IO[Vector[Facts]] =
    ExcelIO.instance[IO].read(path).map { wb =>
      val data = wb.sheets.find(_.name.value == "Data").getOrElse(fail("no Data sheet"))
      refs.map { ref =>
        Facts(
          data.cells.get(ref).map(_.value),
          data.getCellStyle(ref).map(_.numFmt),
          data.getCellStyle(ref).exists(_.font.bold)
        )
      }
    }

  /** Run the same batch through `--stream` and in memory; return both runs' facts for `refs`. */
  private def bothPaths(
    name: String,
    json: String,
    refs: Vector[ARef]
  ): IO[(Vector[Facts], Vector[Facts])] =
    for
      in <- input(name)
      streamed = dir().resolve(s"$name-stream.xlsx")
      memory = dir().resolve(s"$name-memory.xlsx")
      s <- batch(in, streamed, json, stream = true)
      m <- batch(in, memory, json, stream = false)
      _ = assertEquals(s.exit, 0, s"stream: ${s.stderr}")
      _ = assertEquals(m.exit, 0, s"memory: ${m.stderr}")
      fs <- facts(streamed, refs)
      fm <- facts(memory, refs)
    yield (fs, fm)

  private val a1 = ARef.from0(0, 0)
  private val b1 = ARef.from0(1, 0)
  private val c1 = ARef.from0(2, 0)
  private val d1 = ARef.from0(3, 0)
  private val a2 = ARef.from0(0, 1)
  private val b2 = ARef.from0(1, 1)
  private val c2 = ARef.from0(2, 1)
  private val d2 = ARef.from0(3, 1)

  test(
    "--stream explicit format keeps the font, replaces the numFmt, and equals the in-memory result"
  ) {
    val json = """[{"op":"put","ref":"A1","value":1234,"format":"percent"}]"""
    bothPaths("explicit", json, Vector(a1)).map { (streamed, memory) =>
      assertEquals(streamed, memory)
      assertEquals(streamed.map(_.numFmt), Vector(Some(NumFmt.Percent)))
      assertEquals(streamed.map(_.bold), Vector(true), "the explicit format touches only numFmt")
      assertEquals(streamed.map(_.value), Vector(Some(CellValue.Number(BigDecimal(1234)))))
    }
  }

  test("--stream inferred date onto Currency keeps Currency and the font, like in memory") {
    val json = """[{"op":"put","ref":"B1","value":"2025-01-15"}]"""
    bothPaths("inferred", json, Vector(b1)).map { (streamed, memory) =>
      assertEquals(streamed.map(f => (f.numFmt, f.bold)), memory.map(f => (f.numFmt, f.bold)))
      assertEquals(streamed.map(_.numFmt), Vector(Some(NumFmt.Currency)))
      assertEquals(streamed.map(_.bold), Vector(true))
    }
  }

  test("--stream putf with a format keeps the font on single, dragged and listed formula cells") {
    val json =
      """[{"op":"putf","ref":"B1","value":"=1+1","format":"#,##0.0"},""" +
        """{"op":"putf","ref":"C1:D1","value":"=2","from":"C1","format":"0.0"},""" +
        """{"op":"putf","ref":"A2:B2","values":["=3","=4"],"format":"0.00"}]"""
    bothPaths("putf", json, Vector(b1, c1, d1, a2, b2)).map { (streamed, memory) =>
      assertEquals(streamed.map(f => (f.numFmt, f.bold)), memory.map(f => (f.numFmt, f.bold)))
      assertEquals(
        streamed.map(_.numFmt),
        Vector(
          Some(NumFmt.Custom("#,##0.0")),
          Some(NumFmt.Custom("0.0")),
          Some(NumFmt.Custom("0.0")),
          Some(NumFmt.Custom("0.00")),
          Some(NumFmt.Custom("0.00"))
        )
      )
      assert(streamed.forall(_.bold), s"bold lost under --stream: $streamed")
    }
  }

  test(
    "--stream values array: an explicit op format replaces, inferred element hints merge only onto General"
  ) {
    val json =
      """[{"op":"put","ref":"C2:D2","values":[5,6],"format":"percent"},""" +
        """{"op":"put","ref":"A2:B2","values":["$7","2025-01-15"]}]"""
    bothPaths("values", json, Vector(c2, d2, a2, b2)).map { (streamed, memory) =>
      assertEquals(streamed.map(f => (f.numFmt, f.bold)), memory.map(f => (f.numFmt, f.bold)))
      assertEquals(
        streamed.map(_.numFmt),
        Vector(
          Some(NumFmt.Percent),
          Some(NumFmt.Percent),
          Some(NumFmt.Custom(percentCode)),
          Some(NumFmt.Currency)
        )
      )
      assert(streamed.forall(_.bold), s"bold lost under --stream: $streamed")
    }
  }

  test("--stream put verb: a detected format stays an inferred hint and keeps the cell's style") {
    for
      in <- input("verb")
      out = dir().resolve("verb-out.xlsx")
      run <- CliHarness.run(
        "-f",
        in.toString,
        "-s",
        "Data",
        "-o",
        out.toString,
        "--stream",
        "put",
        "B1",
        "2025-01-15"
      )
      fs <- facts(out, Vector(b1))
    yield
      assertEquals(run.exit, 0, run.stderr)
      assertEquals(fs.map(_.numFmt), Vector(Some(NumFmt.Currency)))
      assertEquals(fs.map(_.bold), Vector(true))
  }
