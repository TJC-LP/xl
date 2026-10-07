package com.tjclp.xl.cli.raster

import java.io.ByteArrayInputStream
import java.nio.charset.StandardCharsets
import java.nio.file.Files

import scala.concurrent.duration.*

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.cli.contract.{CliError, ErrorCode}

/**
 * GH-690: the subprocess layer's bounds, run against `/bin/sh` in this JVM: stderr's tail, a
 * deadline, a typed spawn failure, the message without a dangling colon, and the chain's write to
 * an output path it cannot replace.
 */
class PipedBackendSpec extends CatsEffectSuite:

  private def failure(io: IO[Unit]): IO[Throwable] =
    io.attempt.map(_.swap.getOrElse(fail("expected a failure")))

  test("stderr is kept to its tail: a 2 MB stream costs StderrLimit bytes of message") {
    val script = "head -c 2000000 /dev/zero | tr '\\0' a >&2; echo TAIL-MARK >&2; exit 1"
    failure(PipedBackend.run("shim", "sh", List("-c", script), Array.emptyByteArray)).map {
      case RasterError.ConversionFailed("shim", stderr, 1) =>
        assert(stderr.startsWith("…"), stderr.take(20))
        assert(stderr.strip.endsWith("TAIL-MARK"), stderr.takeRight(20))
        assert(stderr.getBytes(StandardCharsets.UTF_8).length <= PipedBackend.StderrLimit + 3)
      case other => fail(s"expected ConversionFailed, got $other")
    }
  }

  test("tail: verbatim under the limit; a cut never splits a UTF-8 character") {
    def tailOf(text: String, limit: Int): String =
      PipedBackend.tail(new ByteArrayInputStream(text.getBytes(StandardCharsets.UTF_8)), limit)
    assertEquals(tailOf("short", 16), "short")
    assertEquals(tailOf("", 16), "")
    // 2-byte characters against an odd limit: the cut lands mid-character
    val cut = tailOf("é" * 3000, 4095)
    assert(cut.startsWith("…"))
    assert(cut.drop(1).forall(_ == 'é'), cut.take(10))
    assertEquals(cut.length - 1, 2047)
    // a stream longer than one read chunk, past the limit several times over
    val digits = (0 until 50000).map(i => (i % 10).toString).mkString
    assertEquals(tailOf(digits, 100), "…" + digits.takeRight(100))
  }

  test("a run past its deadline is TimedOut and the child is stopped") {
    val started = System.nanoTime()
    failure(
      PipedBackend.run("sleeper", "sh", List("-c", "sleep 60"), Array.emptyByteArray, 1.second)
    ).map { error =>
      assertEquals(error, RasterError.TimedOut("sleeper", 1.second))
      assert((System.nanoTime() - started).nanos < 20.seconds, "the timeout did not stop the run")
      val cli = CliError.fromThrowable(error)
      assertEquals(cli.code, ErrorCode.RASTERIZER_UNAVAILABLE)
      assertEquals(cli.message, "sleeper did not finish within 1 s and was stopped")
    }
  }

  test("a command that cannot be started is RasterizerNotFound, not a raw IOException") {
    failure(
      PipedBackend.run("ghost", "xl-gh690-no-such-binary", Nil, Array.emptyByteArray)
    ).map { error =>
      assert(error.isInstanceOf[RasterError.RasterizerNotFound], error.toString)
      val cli = CliError.fromThrowable(error)
      assertEquals(cli.code, ErrorCode.RASTERIZER_UNAVAILABLE)
      assert(cli.message.contains("`xl-gh690-no-such-binary` could not be started"), cli.message)
    }
  }

  test("ConversionFailed with an empty stderr ends at the exit code") {
    assertEquals(
      RasterError.ConversionFailed("rsvg-convert", "  \n", 1).message,
      "rsvg-convert conversion failed (exit 1)"
    )
    assertEquals(
      RasterError.ConversionFailed("rsvg-convert", "bad svg\n", 2).message,
      "rsvg-convert conversion failed (exit 2): bad svg"
    )
  }

  test("an output path the converted image cannot replace is IO_WRITE, the old path untouched") {
    val dir = Files.createTempDirectory("xl-raster-outdir-")
    val target = Files.createDirectory(dir.resolve("out.png"))
    val blocker = Files.writeString(target.resolve("keep"), "x")
    val svg =
      """<svg xmlns="http://www.w3.org/2000/svg" width="4" height="4"><rect width="4" height="4"/></svg>"""
    failure(
      RasterizerChain.convert(svg, target, RasterFormat.Png, 96, Some("batik")).void
    ).map { error =>
      assert(error.isInstanceOf[RasterError.OutputFailed], error.toString)
      assertEquals(CliError.fromThrowable(error).code, ErrorCode.IO_WRITE)
      assert(Files.exists(blocker))
      val left = Files.list(dir)
      try assertEquals(left.count(), 1L, "the staging file is removed")
      finally left.close()
    }.guarantee(IO.blocking {
      Files.deleteIfExists(blocker)
      Files.deleteIfExists(target)
      Files.deleteIfExists(dir)
    }.void)
  }
