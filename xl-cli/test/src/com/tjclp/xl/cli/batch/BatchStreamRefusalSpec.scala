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
