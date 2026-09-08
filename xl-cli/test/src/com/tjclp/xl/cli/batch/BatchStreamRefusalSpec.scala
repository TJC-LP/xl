package com.tjclp.xl.cli.batch

import java.nio.file.{Files, Path}

import cats.effect.{IO, Resource}
import munit.CatsEffectSuite

import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.contract.{CliHarness, CliRun, TestFixtures}
import com.tjclp.xl.io.ExcelIO

/**
 * ADR-017 invariant 2: `--stream` honours an op with identical semantics or refuses it by index
 * before any byte is written (`UNSUPPORTED_IN_STREAM`, exit 2). It never degrades.
 */
class BatchStreamRefusalSpec extends CatsEffectSuite:

  private val fixtures = ResourceSuiteLocalFixture(
    "stream-refusal-fixtures",
    Resource.make(TestFixtures.materialize)(TestFixtures.delete)
  )

  override def munitFixtures = List(fixtures)

  private def input: String = fixtures().resolve("simple.xlsx").toString

  private def fresh(name: String): Path = fixtures().resolve(name)

  private def streamBatch(
    out: Path,
    json: String,
    sheet: Option[String] = Some("Data")
  ): IO[CliRun] =
    val sheetArgs = sheet.toList.flatMap(s => List("-s", s))
    CliHarness.run(
      List("-f", input) ++ sheetArgs ++ List("-o", out.toString, "--stream", "batch", "-"),
      json
    )

  private def assertRefused(run: CliRun, out: Path, code: String): Unit =
    assertEquals(run.exit, 2, run.stderr)
    assertEquals(run.stdout, "", "nothing on stdout when refused")
    assert(run.stderr.startsWith("Error: "), run.stderr)
    assert(run.stderr.contains(s"  code: $code"), run.stderr)
    assert(!Files.exists(out), s"refusal must write nothing, but $out exists")

  test("--stream refuses non-streamable ops by index before writing anything") {
    val out = fresh("refuse-index.xlsx")
    val json =
      """[{"op":"put","ref":"A5","value":1},{"op":"freeze","ref":"B2"},{"op":"clear","range":"A1:B2"}]"""
    streamBatch(out, json).map { run =>
      assertRefused(run, out, "UNSUPPORTED_IN_STREAM")
      assert(run.stderr.contains("[2 freeze, 3 clear]"), run.stderr)
      assert(run.stderr.contains("not supported in streaming mode"), run.stderr)
      assert(run.stderr.contains("hint: drop --stream to apply them in memory"), run.stderr)
    }
  }

  test("--stream refuses an op scoped to a sheet other than the streamed worksheet") {
    val out = fresh("refuse-scope.xlsx")
    val json =
      """[{"op":"put","ref":"A5","value":1},{"op":"put","sheet":"Summary","ref":"A1","value":2}]"""
    streamBatch(out, json).map { run =>
      assertRefused(run, out, "UNSUPPORTED_IN_STREAM")
      assert(run.stderr.contains("[2 put]"), run.stderr)
      assert(run.stderr.contains("Summary") && run.stderr.contains("Data"), run.stderr)
    }
  }

  test("--stream refuses a ref qualified with another sheet") {
    val out = fresh("refuse-qualified.xlsx")
    streamBatch(out, """[{"op":"style","range":"Summary!A1:B1","bold":true}]""").map { run =>
      assertRefused(run, out, "UNSUPPORTED_IN_STREAM")
      assert(run.stderr.contains("[1 style]"), run.stderr)
    }
  }

  test("--stream honours a qualified ref that names the streamed worksheet") {
    val out = fresh("honour-qualified.xlsx")
    streamBatch(
      out,
      """[{"op":"put","ref":"Data!A7","value":"q"},{"op":"put","sheet":"Data","ref":"A8","value":"s"}]"""
    )
      .flatMap { run =>
        assertEquals(run.exit, 0, run.stderr)
        ExcelIO.instance[IO].read(out).map { wb =>
          val data = wb.sheets.find(_.name.value == "Data").getOrElse(fail("no Data sheet"))
          assertEquals(data.cells.get(ARef.from0(0, 6)).map(_.value), Some(CellValue.Text("q")))
          assertEquals(data.cells.get(ARef.from0(0, 7)).map(_.value), Some(CellValue.Text("s")))
        }
      }
  }

  test("--stream with only streamable ops still writes") {
    val out = fresh("streamable.xlsx")
    val json =
      """[{"op":"put","ref":"A5","value":"s"},{"op":"style","range":"A5","bold":true},""" +
        """{"op":"merge","range":"A6:B6"},{"op":"colwidth","col":"A","width":20},{"op":"row-hide","row":9}]"""
    streamBatch(out, json).flatMap { run =>
      assertEquals(run.exit, 0, run.stderr)
      assert(run.stdout.contains("Applied 5 operations (streaming)"), run.stdout)
      assert(Files.exists(out))
      ExcelIO.instance[IO].read(out).map { wb =>
        val data = wb.sheets.find(_.name.value == "Data").getOrElse(fail("no Data sheet"))
        assertEquals(data.cells.get(ARef.from0(0, 4)).map(_.value), Some(CellValue.Text("s")))
      }
    }
  }

  test("--stream honours dragging putf with the same shifted formulas as the in-memory path") {
    val streamed = fresh("drag-stream.xlsx")
    val inMemory = fresh("drag-memory.xlsx")
    val json = """[{"op":"putf","ref":"D1:D3","value":"=B1*2","from":"D1"}]"""
    def formulas(path: Path): IO[Vector[Option[String]]] =
      ExcelIO.instance[IO].read(path).map { wb =>
        val data = wb.sheets.find(_.name.value == "Data").getOrElse(fail("no Data sheet"))
        (0 to 2).toVector.map { row =>
          data.cells.get(ARef.from0(3, row)).map(_.value).collect {
            case CellValue.Formula(f, _, _) => f
          }
        }
      }
    for
      s <- streamBatch(streamed, json)
      m <- CliHarness.run(
        List("-f", input, "-s", "Data", "-o", inMemory.toString, "batch", "-"),
        json
      )
      fs <- formulas(streamed)
      fm <- formulas(inMemory)
    yield
      assertEquals(s.exit, 0, s.stderr)
      assertEquals(m.exit, 0, m.stderr)
      assertEquals(fs, Vector(Some("B1*2"), Some("B2*2"), Some("B3*2")))
      assertEquals(fs, fm, "streaming must shift exactly like the in-memory path")
  }

  test("an unknown op under --stream is BATCH_OP_UNKNOWN (exit 2) and writes nothing") {
    val out = fresh("unknown.xlsx")
    streamBatch(out, """[{"op":"fill","range":"A1:A3","value":1}]""").map { run =>
      assertRefused(run, out, "BATCH_OP_UNKNOWN")
      assert(run.stderr.contains("Unknown operation 'fill'"), run.stderr)
    }
  }

  test("non-streamable ops are refused before the sheet is resolved") {
    // Multi-sheet file, no -s: the streamability refusal comes first, not SHEET_REQUIRED
    val out = fresh("before-resolve.xlsx")
    streamBatch(out, """[{"op":"freeze","ref":"B2"}]""", sheet = None).map { run =>
      assertRefused(run, out, "UNSUPPORTED_IN_STREAM")
      assert(run.stderr.contains("[1 freeze]"), run.stderr)
    }
  }

  // ---------------------------------------------------------------------------------------------
  // The streamed worksheet without -s (ADR-017 §2.5 for a batch): the ops themselves say where they
  // land, and the writer honours one unanimous sheet or refuses by index before writing anything.
  // ---------------------------------------------------------------------------------------------

  private def memoryBatch(out: Path, json: String): IO[CliRun] =
    CliHarness.run(List("-f", input, "-o", out.toString, "batch", "-"), json)

  private def assertFailed(run: CliRun, out: Path, code: String): Unit =
    assertEquals(run.exit, 3, run.stderr)
    assertEquals(run.stdout, "", "nothing on stdout when the batch fails")
    assert(run.stderr.contains(s"  code: $code"), run.stderr)
    assert(!Files.exists(out), s"a failed batch must write nothing, but $out exists")

  test("--stream without -s honours the one sheet every op names, by key or qualifier") {
    val out = fresh("derive-sheet.xlsx")
    val json =
      """[{"op":"put","sheet":"Data","ref":"A9","value":"k"},{"op":"put","ref":"Data!A10","value":"q"}]"""
    streamBatch(out, json, sheet = None).flatMap { run =>
      assertEquals(run.exit, 0, run.stderr)
      assert(run.stdout.contains("Applied 2 operations (streaming)"), run.stdout)
      ExcelIO.instance[IO].read(out).map { wb =>
        val data = wb.sheets.find(_.name.value == "Data").getOrElse(fail("no Data sheet"))
        assertEquals(data.cells.get(ARef.from0(0, 8)).map(_.value), Some(CellValue.Text("k")))
        assertEquals(data.cells.get(ARef.from0(0, 9)).map(_.value), Some(CellValue.Text("q")))
      }
    }
  }

  test("--stream without -s refuses ops naming different sheets, by index, before writing") {
    val out = fresh("derive-disagree.xlsx")
    val json =
      """[{"op":"put","sheet":"Data","ref":"A9","value":1},{"op":"put","sheet":"Summary","ref":"A1","value":2}]"""
    streamBatch(out, json, sheet = None).map { run =>
      assertRefused(run, out, "UNSUPPORTED_IN_STREAM")
      assert(run.stderr.contains("ops [1 put, 2 put] target different sheets"), run.stderr)
      assert(run.stderr.contains("Data") && run.stderr.contains("Summary"), run.stderr)
      assert(run.stderr.contains("pass -s to choose it"), run.stderr)
    }
  }

  test(
    "--stream without -s: an unscoped op on a multi-sheet book fails at its index, as in memory"
  ) {
    val streamed = fresh("derive-unscoped-stream.xlsx")
    val memory = fresh("derive-unscoped-memory.xlsx")
    // Op 1 names its sheet; op 2 names none and the book has two — in memory that is
    // BATCH_OP_FAILED wrapping SHEET_REQUIRED at op 2, so --stream must say the same, not degrade
    val json =
      """[{"op":"put","sheet":"Data","ref":"A9","value":1},{"op":"merge","range":"A1:B1"}]"""
    for
      s <- streamBatch(streamed, json, sheet = None)
      m <- memoryBatch(memory, json)
    yield
      assertFailed(s, streamed, "BATCH_OP_FAILED")
      assertFailed(m, memory, "BATCH_OP_FAILED")
      assert(s.stderr.contains("Object 2 (merge): batch merge requires --sheet"), s.stderr)
      assertEquals(s.stderr.linesIterator.next(), m.stderr.linesIterator.next(), "same first line")
  }

  test(
    "--stream without -s: a sheet key naming a missing sheet fails at its index with candidates"
  ) {
    val streamed = fresh("derive-missing-stream.xlsx")
    val memory = fresh("derive-missing-memory.xlsx")
    val json = """[{"op":"put","sheet":"Dat","ref":"A9","value":1}]"""
    for
      s <- streamBatch(streamed, json, sheet = None)
      m <- memoryBatch(memory, json)
    yield
      assertFailed(s, streamed, "BATCH_OP_FAILED")
      assertFailed(m, memory, "BATCH_OP_FAILED")
      assert(s.stderr.contains("Object 1 (put): Sheet not found: Dat. Available: "), s.stderr)
      assert(s.stderr.contains("did you mean: Data"), s.stderr)
      assertEquals(s.stderr.linesIterator.next(), m.stderr.linesIterator.next(), "same first line")
  }

  test(
    "--stream with -s still refuses an op scoped elsewhere even when the ops agree among themselves"
  ) {
    val out = fresh("flag-wins.xlsx")
    val json = """[{"op":"put","sheet":"Summary","ref":"A1","value":1}]"""
    streamBatch(out, json).map { run =>
      assertRefused(run, out, "UNSUPPORTED_IN_STREAM")
      assert(run.stderr.contains("[1 put]"), run.stderr)
    }
  }

  // ---------------------------------------------------------------------------------------------
  // Apply-time failures: the same BATCH_OP_FAILED shape as in memory — `Object N (op): …`, exit 3,
  // `location.opIndex` — and nothing written.
  // ---------------------------------------------------------------------------------------------

  test(
    "--stream: an apply-time failure is BATCH_OP_FAILED naming the op, with opIndex, nothing written"
  ) {
    val streamed = fresh("apply-fail-stream.xlsx")
    val memory = fresh("apply-fail-memory.xlsx")
    val json = """[{"op":"put","ref":"A5","value":1},{"op":"merge","range":"nonsense"}]"""
    for
      s <- streamBatch(streamed, json)
      m <- CliHarness.run(
        List("-f", input, "-s", "Data", "-o", memory.toString, "batch", "-"),
        json
      )
      j <- CliHarness.run(
        List(
          "--json",
          "-f",
          input,
          "-s",
          "Data",
          "-o",
          streamed.toString,
          "--stream",
          "batch",
          "-"
        ),
        json
      )
    yield
      assertFailed(s, streamed, "BATCH_OP_FAILED")
      assertFailed(m, memory, "BATCH_OP_FAILED")
      assert(s.stderr.startsWith("Error: Object 2 (merge): "), s.stderr)
      assert(m.stderr.startsWith("Error: Object 2 (merge): "), m.stderr)
      assert(s.stderr.contains("nonsense"), s.stderr)
      assertEquals(j.exit, 3, j.stderr)
      val error = ujson.read(j.stdout)("error")
      assertEquals(error("code").str, "BATCH_OP_FAILED")
      assertEquals(error("location")("opIndex").num.toInt, 2)
      assert(error("message").str.startsWith("Object 2 (merge): "), error("message").str)
      assert(!Files.exists(streamed), "a failed --json run must write nothing either")
  }

  test("--stream parse warnings reach the run's sink: stderr line in text, envelope under --json") {
    val textOut = fresh("warn-text.xlsx")
    val jsonOut = fresh("warn-json.xlsx")
    val json = """[{"op":"put","ref":"A5","value":1,"bogus":true}]"""
    for
      t <- streamBatch(textOut, json)
      j <- CliHarness.run(
        List("--json", "-f", input, "-s", "Data", "-o", jsonOut.toString, "--stream", "batch", "-"),
        json
      )
    yield
      assertEquals(t.exit, 0, t.stderr)
      assert(
        t.stderr.contains(
          "Warning[UNKNOWN_PROPERTY]: Object 1 (put): unknown properties ignored: bogus"
        ),
        t.stderr
      )
      assertEquals(j.exit, 0, j.stderr)
      assertEquals(j.stderr, "")
      val warnings = ujson.read(j.stdout)("warnings").arr
      assertEquals(warnings.map(_("code").str), Vector("UNKNOWN_PROPERTY").toBuffer)
      assertEquals(warnings(0)("location")("opIndex").num.toInt, 1)
  }
