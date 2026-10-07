package com.tjclp.xl.cli.raster

import java.io.ByteArrayInputStream
import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.nio.file.attribute.PosixFilePermissions

import scala.concurrent.duration.*

import cats.effect.IO
import cats.syntax.all.*
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
    // forced, and through the default chain, which must not retry: another backend cannot
    // repair the destination, and retrying ended in RASTERIZER_UNAVAILABLE
    List(Some("batik"), None)
      .traverse_ { preferred =>
        failure(RasterizerChain.convert(svg, target, RasterFormat.Png, 96, preferred).void).map {
          error =>
            assert(error.isInstanceOf[RasterError.OutputFailed], s"$preferred: $error")
            assertEquals(CliError.fromThrowable(error).code, ErrorCode.IO_WRITE)
            assert(Files.exists(blocker))
            val left = Files.list(dir)
            try assertEquals(left.count(), 1L, "the staging file is removed")
            finally left.close()
        }
      }
      .guarantee(IO.blocking {
        Files.deleteIfExists(blocker)
        Files.deleteIfExists(target)
        Files.deleteIfExists(dir)
      }.void)
  }

  private val tinySvg =
    """<svg xmlns="http://www.w3.org/2000/svg" width="4" height="4"><rect width="4" height="4"/></svg>"""

  private def withDir[A](body: Path => IO[A]): IO[A] =
    IO.blocking(Files.createTempDirectory("xl-raster-perm-")).flatMap { dir =>
      body(dir).guarantee(IO.blocking {
        val entries = Files.list(dir)
        try entries.forEach(p => Files.deleteIfExists(p))
        finally entries.close()
        Files.deleteIfExists(dir)
      }.void)
    }

  private def mode(p: Path): String =
    PosixFilePermissions.toString(Files.getPosixFilePermissions(p))

  test("replacing an owner-only output keeps it owner-only; a new output gets the default mode") {
    withDir { dir =>
      val existing = dir.resolve("private.png")
      Files.writeString(existing, "OLD")
      Files.setPosixFilePermissions(existing, PosixFilePermissions.fromString("rw-------"))
      val fresh = dir.resolve("fresh.png")
      val reference = Files.createFile(dir.resolve("reference")) // the umask's mode for a new file
      for
        _ <- RasterizerChain.convert(tinySvg, existing, RasterFormat.Png, 96, Some("batik"))
        _ <- RasterizerChain.convert(tinySvg, fresh, RasterFormat.Png, 96, Some("batik"))
      yield
        assertEquals(mode(existing), "rw-------")
        assert(Files.size(existing) > 3, "replaced with the image")
        assertEquals(mode(fresh), mode(reference))
    }
  }

  test("a long output name still converts: the staging name does not copy it") {
    withDir { dir =>
      val out = dir.resolve(("n" * 220) + ".png")
      RasterizerChain.convert(tinySvg, out, RasterFormat.Png, 96, Some("batik")).map { used =>
        assertEquals(used, "Batik")
        assert(Files.size(out) > 0)
      }
    }
  }

  /** A child that ignores SIGTERM and never reads stdin. */
  private val stubborn = List("-c", "trap '' TERM; while :; do :; done")

  test("a deadline holds against a child that ignores SIGTERM while the stdin write is blocked") {
    val started = System.nanoTime()
    // 4 MB, far past any pipe buffer: the write blocks for as long as the child lives
    val input = Array.fill[Byte](4 << 20)('x')
    failure(PipedBackend.run("stubborn", "sh", stubborn, input, 1.second)).map { error =>
      assertEquals(error, RasterError.TimedOut("stubborn", 1.second))
      assert((System.nanoTime() - started).nanos < 20.seconds, "the escalation never fired")
    }
  }

  test("a probe that ignores SIGTERM is false at its deadline, not a hang") {
    val started = System.nanoTime()
    PipedBackend.probe("sh", stubborn, 1.second).map { ok =>
      assert(!ok)
      assert((System.nanoTime() - started).nanos < 20.seconds, "the probe outlived its deadline")
    }
  }

  test("kill stops the child even when descendant discovery throws") {
    IO.blocking {
      val process = new java.lang.ProcessBuilder("sleep", "60").start()
      PipedBackend.kill(process, _ => throw new RuntimeException("Operation not permitted"))
      assert(!process.isAlive, "the direct child must be stopped")
    }
  }

  test("capture: stdout on exit 0, None on a failure or a missing binary") {
    for
      ok <- PipedBackend.capture("sh", List("-c", "echo version-1; exit 0"))
      failed <- PipedBackend.capture("sh", List("-c", "echo nope; exit 2"))
      missing <- PipedBackend.capture("xl-gh690-no-such-binary", Nil)
    yield
      assertEquals(ok, Some("version-1\n"))
      assertEquals(failed, None)
      assertEquals(missing, None)
  }
