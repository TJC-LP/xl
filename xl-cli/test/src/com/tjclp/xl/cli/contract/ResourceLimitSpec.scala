package com.tjclp.xl.cli.contract

import java.nio.file.{Files, Path}

import scala.jdk.CollectionConverters.*

import cats.effect.{IO, Resource}
import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.{Cli, CliIO, MemoryGuard}
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.macros.ref

/**
 * GH-636: memory exhaustion is a typed failure — exit 3, `RESOURCE_LIMIT`, the heap in the message,
 * `--stream` and `-Xmx` in the hint, the envelope under `--json` — never a raw
 * `java.lang.OutOfMemoryError` on stderr with exit 1 and an empty stdout.
 *
 * cats-effect treats an `OutOfMemoryError` as fatal (it halts the runtime), so it cannot be
 * injected into the program as an `IO` error; the seam is [[MemoryGuard.blocking]], the same thunk
 * guard the in-memory load runs under, here as the run's `stdin` so the real program — parser,
 * handler, classification, rendering — carries the failure to the streams.
 */
class ResourceLimitSpec extends CatsEffectSuite:

  private val dir = ResourceSuiteLocalFixture(
    "resource-limit-fixtures",
    Resource.make(IO.blocking(Files.createTempDirectory("xl-resource-limit-")))(d =>
      IO.blocking {
        Files.list(d).iterator.asScala.toVector.foreach(Files.deleteIfExists)
        Files.deleteIfExists(d)
        ()
      }
    )
  )

  override def munitFixtures = List(dir)

  /** One sheet (so no `-s` is needed), a constant and a formula over it. */
  private def fixture(name: String): IO[Path] =
    val path = dir().resolve(name)
    IO.blocking(Files.exists(path)).flatMap {
      case true => IO.pure(path)
      case false =>
        val sheet = Sheet("Data").put(ref"A1", 10).put(ref"B1", CellValue.Formula("A1*2", None))
        ExcelIO.instance[IO].write(Workbook(Vector(sheet)), path).as(path)
    }

  /** Pin the heap the guard sizes `path` against, around `io`. */
  private def withHeap[A](path: Path, heap: Long)(io: IO[A]): IO[A] =
    IO(MemoryGuard.Seams.heap.updateAndGet(_ + (path -> heap))).void
      .bracket(_ => io)(_ => IO(MemoryGuard.Seams.heap.updateAndGet(_ - path)).void)

  /**
   * Pin a fault for a stage around `io` — stage-wide, since a write's recalculation runs against
   * the staging file, never the `-o` path; the harness lock keeps other runs out meanwhile.
   */
  private def withFault[A](stage: String, fault: Throwable)(io: IO[A]): IO[A] =
    val key = (Option.empty[Path], stage)
    IO(MemoryGuard.Seams.faults.updateAndGet(_ + (key -> fault))).void
      .bracket(_ => io)(_ => IO(MemoryGuard.Seams.faults.updateAndGet(_ - key)).void)

  /** A heap between the fixture's bands: the load may not fit — the warning band. */
  private def doubtfulHeap(path: Path): IO[Long] =
    MemoryGuard.footprint(path).map(fp => (fp.atLeast + fp.upTo) / 2)

  private def pressureWarnings(json: String): Vector[ujson.Value] =
    ujson.read(json)("warnings").arr.toVector.filter(_("code").str == WarningCode.MEMORY_PRESSURE)

  private def run(args: List[String], stdin: IO[String]): IO[(Int, String, String)] =
    CliIO.capturing("").flatMap { (base, collect) =>
      Cli.run(args, base.copy(stdin = stdin)).flatMap { code =>
        collect.map((out, err) => (code.code, out, err))
      }
    }

  private val exhausted: IO[String] =
    MemoryGuard.blocking[String](throw new OutOfMemoryError("Garbage-collected heap size exceeded"))

  test("text: exit 3, stdout empty, Error + code RESOURCE_LIMIT + the --stream/-Xmx hint") {
    run(List("batch", "--dry-run", "-"), exhausted).map { (exit, out, err) =>
      assertEquals(exit, 3, err)
      assertEquals(out, "", "stdout must be empty on failure")
      val lines = err.split("\n", -1).toVector
      assertEquals(lines.headOption, Some(s"Error: ${MemoryGuard.exhausted.message}"), err)
      assert(lines.contains(s"  code: ${ErrorCode.RESOURCE_LIMIT}"), err)
      assert(lines.contains(s"  hint: ${MemoryGuard.hint}"), err)
      assert(err.contains("--stream") && err.contains("-Xmx"), err)
      assert(!err.contains("java.lang.OutOfMemoryError"), s"no raw error may leak:\n$err")
    }
  }

  test("--json: one schema-valid envelope, ok:false, exitCode 3, error.code RESOURCE_LIMIT") {
    run(List("--json", "batch", "--dry-run", "-"), exhausted).map { (exit, out, err) =>
      assertEquals(exit, 3, out + err)
      val envelope = ujson.read(out)
      EnvelopeSchema.assertValid(envelope)
      assertEquals(envelope("ok"), ujson.False)
      assertEquals(envelope("exitCode").num.toInt, 3)
      assertEquals(envelope("verb").str, "batch")
      assertEquals(envelope("data"), ujson.Null)
      val error = envelope("error")
      assertEquals(error("code").str, ErrorCode.RESOURCE_LIMIT)
      assertEquals(error("message").str, MemoryGuard.exhausted.message)
      assertEquals(error("hint").str, MemoryGuard.hint)
      assert(
        error("message").str.contains(MemoryGuard.human(MemoryGuard.maxHeapBytes)),
        s"the message names the heap: ${error("message").str}"
      )
      assertEquals(err, s"Error: ${MemoryGuard.exhausted.message}\n")
    }
  }

  test("a different Error subtype is still INTERNAL, exit 3 — only memory exhaustion is typed") {
    run(List("--json", "batch", "--dry-run", "-"), IO.raiseError(new AssertionError("boom"))).map {
      (exit, out, _) =>
        assertEquals(exit, 3, out)
        val envelope = ujson.read(out)
        EnvelopeSchema.assertValid(envelope)
        assertEquals(envelope("error")("code").str, ErrorCode.INTERNAL)
        assertEquals(envelope("error")("message").str, "boom")
    }
  }

  test("--max-size below 0 is a usage error (exit 2): it is not an unlimited load") {
    fixture("book.xlsx").flatMap { path =>
      for
        equals <- CliHarness.run("-f", path.toString, "--max-size=-1", "sheets")
        spaced <- CliHarness.run("-f", path.toString, "--max-size", "-1", "sheets")
      yield
        assertEquals(equals.exit, 2, equals.stderr)
        assertEquals(equals.stdout, "")
        assert(equals.stderr.contains("  code: USAGE"), equals.stderr)
        assert(equals.stderr.contains("--max-size must be 0 or more"), equals.stderr)
        assertEquals(spaced.exit, 2, spaced.stderr)
        assert(spaced.stderr.contains("  code: USAGE"), spaced.stderr)
    }
  }

  test("MEMORY_PRESSURE: a read under a lifted --max-size that may not fit warns and proceeds") {
    fixture("book.xlsx").flatMap { path =>
      doubtfulHeap(path).flatMap { heap =>
        withHeap(path, heap) {
          for
            text <- CliHarness.run("-f", path.toString, "--max-size", "0", "view", "A1:B1")
            json <- CliHarness.run(
              "--json",
              "-f",
              path.toString,
              "--max-size",
              "0",
              "view",
              "A1:B1"
            )
            silent <- CliHarness.run("-f", path.toString, "view", "A1:B1")
          yield
            assertEquals(text.exit, 0, text.stderr)
            assert(text.stdout.contains("10"), text.stdout)
            assert(
              text.stderr.contains(
                s"Warning[${WarningCode.MEMORY_PRESSURE}]: $path may not fit in memory"
              ),
              text.stderr
            )
            assert(text.stderr.contains("--stream") && text.stderr.contains("-Xmx"), text.stderr)
            assertEquals(json.exit, 0, json.stdout)
            val envelope = ujson.read(json.stdout)
            EnvelopeSchema.assertValid(envelope)
            assertEquals(envelope("ok"), ujson.True)
            val pressure = pressureWarnings(json.stdout)
            assertEquals(pressure.size, 1, json.stdout)
            assertEquals(pressure.headOption.map(_("location")("file").str), Some(path.toString))
            assertEquals(json.stderr, "", "under --json a warning belongs to the envelope")
            // at the default limit the guard is inactive: the same pinned heap, no warning
            assertEquals(silent.exit, 0)
            assert(!silent.stderr.contains("MEMORY_PRESSURE"), silent.stderr)
        }
      }
    }
  }

  test("MEMORY_PRESSURE: a write under a lifted --max-size warns, recalculates and writes") {
    fixture("book.xlsx").flatMap { path =>
      val textOut = dir().resolve("pressure-text.xlsx")
      val jsonOut = dir().resolve("pressure-json.xlsx")
      doubtfulHeap(path).flatMap { heap =>
        withHeap(path, heap) {
          for
            text <- CliHarness.run(
              "-f",
              path.toString,
              "--max-size",
              "0",
              "-o",
              textOut.toString,
              "put",
              "A1",
              "21"
            )
            json <- CliHarness.run(
              "--json",
              "-f",
              path.toString,
              "--max-size",
              "0",
              "-o",
              jsonOut.toString,
              "put",
              "A1",
              "21"
            )
          yield
            assertEquals(text.exit, 0, text.stderr)
            assert(text.stderr.contains(s"Warning[${WarningCode.MEMORY_PRESSURE}]:"), text.stderr)
            assert(Files.exists(textOut), "the write lands")
            assertEquals(json.exit, 0, json.stdout)
            val envelope = ujson.read(json.stdout)
            EnvelopeSchema.assertValid(envelope)
            assertEquals(envelope("data")("written"), ujson.True)
            assertEquals(pressureWarnings(json.stdout).size, 1, json.stdout)
            assert(Files.exists(jsonOut))
        }
      }
    }
  }

  test("a write whose recalculation exhausts the heap is RESOURCE_LIMIT, exit 3, nothing written") {
    fixture("book.xlsx").flatMap { path =>
      val out = dir().resolve("recalc-oom.xlsx")
      withFault("recalc", new OutOfMemoryError("Garbage-collected heap size exceeded")) {
        for
          json <- CliHarness.run(
            "--json",
            "-f",
            path.toString,
            "-o",
            out.toString,
            "put",
            "A1",
            "21"
          )
          text <- CliHarness.run("-f", path.toString, "-o", out.toString, "put", "A1", "21")
        yield
          assertEquals(json.exit, 3, json.stdout + json.stderr)
          val envelope = ujson.read(json.stdout)
          EnvelopeSchema.assertValid(envelope)
          assertEquals(envelope("ok"), ujson.False)
          assertEquals(envelope("verb").str, "put")
          assertEquals(envelope("error")("code").str, ErrorCode.RESOURCE_LIMIT)
          assertEquals(envelope("error")("message").str, MemoryGuard.exhausted.message)
          assertEquals(envelope("error")("hint").str, MemoryGuard.hint)
          assertEquals(text.exit, 3, text.stderr)
          assertEquals(text.stdout, "")
          assert(text.stderr.contains(s"  code: ${ErrorCode.RESOURCE_LIMIT}"), text.stderr)
          assert(!text.stderr.contains("java.lang.OutOfMemoryError"), text.stderr)
          assert(!Files.exists(out), "a failed write must not create the output")
      }
    }
  }

  test("diff sizes both books together: one MEMORY_PRESSURE naming both files, none per file") {
    for
      a <- fixture("book.xlsx")
      b <- fixture("book-copy.xlsx")
      fp <- MemoryGuard.footprint(a)
      // each book alone fits under its upper estimate; the two together do not
      json <- withHeap(a, fp.atLeast + fp.upTo) {
        CliHarness.run("--json", "-f", a.toString, "--max-size", "0", "diff", "-g", b.toString)
      }
    yield
      assertEquals(json.exit, 0, json.stdout + json.stderr)
      val pressure = pressureWarnings(json.stdout)
      assertEquals(pressure.size, 1, json.stdout)
      val message = pressure.headOption.map(_("message").str).getOrElse("")
      assert(message.contains(a.toString) && message.contains(b.toString), message)
  }

  test("RESOURCE_LIMIT is published: in ErrorCode.cli, exit 3, in schema --json's errorCodes") {
    assert(ErrorCode.cli.contains(ErrorCode.RESOURCE_LIMIT))
    assert(WarningCode.all.contains(WarningCode.MEMORY_PRESSURE))
    assertEquals(ExitCodes.forCode(ErrorCode.RESOURCE_LIMIT), ExitCodes.failed)
    val codes = Schema
      .json("test")("errorCodes")
      .arr
      .map(e => (e("code").str, e("exit").num.toInt))
    assert(codes.contains((ErrorCode.RESOURCE_LIMIT, 3)), codes.toString)
  }
