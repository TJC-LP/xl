package com.tjclp.xl.cli.raster

import java.io.ByteArrayInputStream
import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.nio.file.attribute.{BasicFileAttributes, PosixFilePermissions}

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
            assert(error.isInstanceOf[RasterError.UnsupportedOutput], s"$preferred: $error")
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

  test("replacing an existing output writes the same file: hard links and symlinks follow") {
    withDir { dir =>
      val existing = Files.write(dir.resolve("out.png"), "old".getBytes(StandardCharsets.UTF_8))
      val link = Files.createLink(dir.resolve("hard.png"), existing)
      val symlink = Files.createSymbolicLink(dir.resolve("sym.png"), existing)
      def key(p: Path) = Files.readAttributes(p, classOf[BasicFileAttributes]).fileKey
      val before = key(existing)
      RasterizerChain.convert(tinySvg, symlink, RasterFormat.Png, 96, Some("batik")).map { _ =>
        assert(Files.isSymbolicLink(symlink), "the symlink must stay a symlink")
        assertEquals(key(existing), before)
        assert(Files.size(existing) > 3, "replaced with the image")
        assertEquals(Files.readAllBytes(link).toSeq, Files.readAllBytes(existing).toSeq)
        assertEquals(Files.list(dir).toArray.count(_.toString.contains(".xl-raster-")), 0)
      }
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

  /** Whether `command` is installed: the process tests need a few POSIX tools. */
  private def onPath(command: String): Boolean =
    new java.lang.ProcessBuilder("sh", "-c", s"command -v $command").start().waitFor() == 0

  /** Wait (up to 10 s) for `path` to appear: the script's signal that its trap is installed. */
  private def awaitFile(path: Path): Unit =
    val deadline = System.nanoTime() + 10.seconds.toNanos
    while !Files.exists(path) && System.nanoTime() < deadline do Thread.sleep(10)
    assert(Files.exists(path), s"$path never appeared")

  /** The pids of the processes whose command line contains `pattern`, or "" for none. */
  private def pgrep(pattern: String): String =
    val process = new java.lang.ProcessBuilder("pgrep", "-f", pattern).start()
    new String(process.getInputStream.readAllBytes(), StandardCharsets.UTF_8).strip

  test("kill also stops a helper the child's SIGTERM handler starts during the grace period") {
    assume(onPath("pgrep"), "needs pgrep")
    val marker = "31.6901"
    withDir { dir =>
      val ready = dir.resolve("ready")
      val script =
        s"trap 'sleep $marker & while :; do :; done' TERM; touch '$ready'; while :; do :; done"
      IO.blocking {
        val process = new java.lang.ProcessBuilder("sh", "-c", script).start()
        awaitFile(ready) // the trap is installed
        PipedBackend.kill(process)
        assert(!process.isAlive, "the child must be stopped")
        assertEquals(pgrep(s"sleep $marker"), "", "the grace-period helper outlived the kill")
      }
    }
  }

  test("a descendant that keeps the pipes open after the child exits cannot stall a deadline") {
    assume(onPath("pgrep"), "needs pgrep")
    val marker = "31.6903"
    val started = System.nanoTime()
    // the child outlives the start of the pipe reads: a child gone before they start has its pipes
    // drained and closed by the JDK, and the orphan would hold nothing
    PipedBackend
      .probe("sh", List("-c", s"sleep $marker & sleep 0.5; exit 0"), 2.seconds)
      .map { ok =>
        assert(!ok, "the orphan holds the pipes, so the probe must reach its deadline")
        assert((System.nanoTime() - started).nanos < 6.seconds, "teardown waited on the orphan")
        // recorded while the child ran, so killed at release though the child had exited
        assertEquals(pgrep(s"sleep $marker"), "", "the orphan outlived the probe")
      }
  }

  test("kill stops a helper a SIGTERM handler starts even when the child then exits") {
    assume(onPath("pgrep"), "needs pgrep")
    val marker = "31.6902"
    withDir { dir =>
      val ready = dir.resolve("ready")
      val script =
        s"trap 'sleep $marker & sleep 0.3; exit 0' TERM; touch '$ready'; while :; do :; done"
      IO.blocking {
        val process = new java.lang.ProcessBuilder("sh", "-c", script).start()
        awaitFile(ready) // the trap is installed
        PipedBackend.kill(process)
        assert(!process.isAlive, "the child must be stopped")
        assertEquals(pgrep(s"sleep $marker"), "", "the helper outlived the kill")
      }
    }
  }

  private def bytes(text: String): Array[Byte] = text.getBytes(StandardCharsets.UTF_8)

  private def text(p: Path): String = new String(Files.readAllBytes(p), StandardCharsets.UTF_8)

  private def listing(dir: Path): List[Path] =
    Files.list(dir).toArray.toList.collect { case p: Path => p }

  test("publish: a copy that fails partway writes the previous output back") {
    withDir { dir =>
      IO.blocking {
        val target = Files.write(dir.resolve("out.png"), bytes("previous image"))
        val staging = Files.write(dir.resolve(".xl-raster-staged.png"), bytes("new, larger image"))
        val error = intercept[java.io.IOException] {
          RasterizerChain.publish(
            staging,
            target,
            (from, out) =>
              if from == staging then
                out.write(bytes("new"))
                throw new java.io.IOException("No space left on device")
              else Files.copy(from, out): Unit
          )
        }
        assertEquals(error.getMessage, "No space left on device")
        assertEquals(text(target), "previous image")
        assertEquals(listing(dir).toSet, Set(target, staging), "the backup is removed")
      }
    }
  }

  test(
    "publish: when the write-back fails too, the backup is kept and named, even for a long name"
  ) {
    withDir { dir =>
      IO.blocking {
        // 224 characters: a recovery name built from the output's would pass the 255-byte limit
        val target = Files.write(dir.resolve(("n" * 220) + ".png"), bytes("previous image"))
        val staging = Files.write(dir.resolve(".xl-raster-staged.png"), bytes("new image"))
        val error = intercept[java.io.IOException] {
          RasterizerChain.publish(
            staging,
            target,
            (_, _) => throw new java.io.IOException("No space left on device")
          )
        }
        val kept = listing(dir).filter(_.getFileName.toString.endsWith(".orig"))
        assertEquals(kept.size, 1, listing(dir).toString)
        kept.foreach { backup =>
          assert(error.getMessage.contains(backup.toString), error.getMessage)
          assertEquals(text(backup), "previous image")
        }
      }
    }
  }

  test("publish: a read-only output is reported intact, with no backup left behind") {
    withDir { dir =>
      IO.blocking {
        val target = Files.write(dir.resolve("out.png"), bytes("previous image"))
        Files.setPosixFilePermissions(target, PosixFilePermissions.fromString("r--r--r--"))
        val staging = Files.write(dir.resolve(".xl-raster-staged.png"), bytes("new image"))
        intercept[java.nio.file.AccessDeniedException](RasterizerChain.publish(staging, target))
        assertEquals(text(target), "previous image")
        assertEquals(listing(dir).toSet, Set(target, staging), "no backup of an untouched output")
      }
    }
  }

  test("a write-only output is overwritten: publishing needs no read access") {
    withDir { dir =>
      val target = Files.write(dir.resolve("out.png"), bytes("previous image"))
      Files.setPosixFilePermissions(target, PosixFilePermissions.fromString("-w-------"))
      RasterizerChain.convert(tinySvg, target, RasterFormat.Png, 96, Some("batik")).map { _ =>
        Files.setPosixFilePermissions(target, PosixFilePermissions.fromString("rw-------"))
        assert(Files.size(target) > 14, "replaced with the image")
        assertEquals(listing(dir), List(target), "no staging or backup left behind")
      }
    }
  }

  test("a FIFO output is IO_WRITE before any conversion, not a hang") {
    assume(onPath("mkfifo"), "needs mkfifo")
    withDir { dir =>
      val fifo = dir.resolve("out.png")
      IO.blocking(new java.lang.ProcessBuilder("mkfifo", fifo.toString).start().waitFor()) >>
        failure(
          RasterizerChain.convert(tinySvg, fifo, RasterFormat.Png, 96, Some("batik")).void
        ).timeout(10.seconds).map { error =>
          assert(error.isInstanceOf[RasterError.UnsupportedOutput], error.toString)
          assert(error.getMessage.contains("not a regular file"), error.getMessage)
          val cli = CliError.fromThrowable(error)
          assertEquals(cli.code, ErrorCode.IO_WRITE)
          assert(cli.hint.exists(_.contains("/dev/stdout")), cli.hint.toString)
          assertEquals(listing(dir), List(fifo), "no staging file left behind")
        }
    }
  }

  test("a dangling output symlink is written through: the link stays, its referent is created") {
    withDir { dir =>
      val referent = dir.resolve("made.png")
      val link = Files.createSymbolicLink(dir.resolve("out.png"), referent)
      RasterizerChain.convert(tinySvg, link, RasterFormat.Png, 96, Some("batik")).map { _ =>
        assert(Files.isSymbolicLink(link), "the symlink must stay a symlink")
        assert(Files.size(referent) > 0, "the referent holds the image")
        assertEquals(listing(dir).toSet, Set(link, referent))
      }
    }
  }

  test("publish: a failed write through a dangling symlink removes the referent, keeps the link") {
    withDir { dir =>
      IO.blocking {
        val referent = dir.resolve("made.png")
        val link = Files.createSymbolicLink(dir.resolve("out.png"), referent)
        val staging = Files.write(dir.resolve(".xl-raster-staged.png"), bytes("new image"))
        intercept[java.io.IOException] {
          RasterizerChain.publish(
            staging,
            link,
            (_, out) =>
              out.write(bytes("new"))
              throw new java.io.IOException("No space left on device")
          )
        }
        assert(Files.isSymbolicLink(link), "the symlink must stay a symlink")
        assert(!Files.exists(referent), "the partial referent is removed")
        assertEquals(listing(dir).toSet, Set(link, staging))
      }
    }
  }

  test("a root output path is IO_WRITE, not an internal error") {
    failure(
      RasterizerChain.convert(tinySvg, Path.of("/"), RasterFormat.Png, 96, Some("batik")).void
    ).map { error =>
      assert(error.isInstanceOf[RasterError.UnsupportedOutput], error.toString)
      assertEquals(CliError.fromThrowable(error).code, ErrorCode.IO_WRITE)
    }
  }

  test("a new output in a missing or read-only directory is IO_WRITE before any backend runs") {
    withDir { dir =>
      val locked = Files.createDirectory(dir.resolve("locked"))
      Files.setPosixFilePermissions(locked, PosixFilePermissions.fromString("r-x------"))
      // the default chain: each backend's own failure used to end in "no rasterizer available"
      List(dir.resolve("missing").resolve("out.png"), locked.resolve("out.png"))
        .traverse_ { out =>
          failure(RasterizerChain.convert(tinySvg, out, RasterFormat.Png, 96, None).void).map {
            error =>
              assert(error.isInstanceOf[RasterError.OutputFailed], s"$out: $error")
              assertEquals(CliError.fromThrowable(error).code, ErrorCode.IO_WRITE)
          }
        }
        .guarantee(IO.blocking {
          Files.setPosixFilePermissions(locked, PosixFilePermissions.fromString("rwx------"))
        }.void)
    }
  }
