package com.tjclp.xl.cli.raster

import java.io.{IOException, OutputStream}
import java.nio.charset.StandardCharsets

import scala.jdk.CollectionConverters.*

import cats.effect.{IO, Resource}
import cats.syntax.all.*

/**
 * A rasterizer backend run as a child process that reads the SVG on stdin (rsvg-convert, cairosvg,
 * ImageMagick).
 *
 * The stdin write and the stdout and stderr drains run concurrently (GH-673): written in sequence,
 * a backend that exits before consuming its input — a bad install, a crash, a refused option — made
 * the write raise `IOException` ("Stream closed", "Broken pipe"), which escaped as `INTERNAL` and
 * never showed the backend's stderr, the one thing the user can act on. Concurrent drains also keep
 * a chatty backend from filling a pipe while the write waits on it.
 *
 * The child is driven through `java.lang.Process` rather than fs2's process streams because fs2
 * closes stdin in a bracket finalizer. An SVG that fits the JDK's 8 KiB stdin buffer fails at the
 * flush and stays buffered, so that close flushes it again and fails a second time; cats-effect
 * reports a finalizer that fails after its body failed, printing a stack trace ahead of the typed
 * error. Here both failures are caught where they happen.
 */
private[raster] object PipedBackend:

  /**
   * Run `command args` with `input` on stdin. A non-zero exit is `ConversionFailed(name, stderr,
   * exit)` whether or not the backend read its input; a zero exit that left the input unread is a
   * failure too, since the output cannot be trusted.
   */
  def run(name: String, command: String, args: List[String], input: Array[Byte]): IO[Unit] =
    exchange(command :: args, input).flatMap { outcome =>
      (outcome.exit, outcome.unread) match
        case (0, None) => IO.unit
        case (exit @ 0, Some(closed)) =>
          val reason = Option(closed.getMessage).getOrElse(closed.toString)
          val detail = if outcome.stderr.isBlank then "" else s"; ${outcome.stderr}"
          IO.raiseError(
            RasterError.ConversionFailed(
              name,
              s"exited before reading the whole SVG ($reason)$detail",
              exit
            )
          )
        case (exit, _) => IO.raiseError(RasterError.ConversionFailed(name, outcome.stderr, exit))
    }

  /**
   * A probe conversion: true when `command args` read all of `input`, exited 0 and wrote something
   * to stdout. Any failure — a missing binary, an early exit, a closed pipe — is false, never an
   * error or a stack trace (ImageMagick's delegate check, GH-673).
   */
  def succeeds(command: String, args: List[String], input: Array[Byte]): IO[Boolean] =
    exchange(command :: args, input)
      .map(o => o.exit == 0 && o.unread.isEmpty && o.stdoutBytes > 0)
      .handleError(_ => false)

  /** What a child did with its input: exit code, the write's failure, stderr, stdout's size. */
  private final case class Outcome(
    exit: Int,
    unread: Option[IOException],
    stderr: String,
    stdoutBytes: Long
  )

  /** Spawn `command`, write `input` while draining stderr and stdout, and wait for the exit. */
  private def exchange(command: List[String], input: Array[Byte]): IO[Outcome] =
    spawn(command).use { process =>
      val destroy = IO.blocking(process.destroy())
      val write = IO.blocking(writeAndClose(process.getOutputStream, input)).cancelable(destroy)
      val stderr = IO
        .blocking(new String(process.getErrorStream.readAllBytes(), StandardCharsets.UTF_8))
        .cancelable(destroy)
      val stdout =
        IO.blocking(process.getInputStream.transferTo(OutputStream.nullOutputStream()))
          .cancelable(destroy)
      (write, stderr, stdout).parTupled.flatMap { (unread, err, bytes) =>
        IO.interruptible(process.waitFor()).map(exit => Outcome(exit, unread, err, bytes))
      }
    }

  /** The child, destroyed on release if it is still running (cancellation, a failed drain). */
  private def spawn(command: List[String]): Resource[IO, Process] =
    Resource.make(IO.blocking(new java.lang.ProcessBuilder(command.asJava).start())) { process =>
      IO.blocking {
        if process.isAlive then
          process.destroy()
          process.waitFor()
        ()
      }
    }

  /**
   * Write `input` and close the stream, answering the first `IOException` of either. A failed flush
   * leaves the bytes buffered, so the close retries it and fails too; both are expected from a
   * backend that exited early, and neither may escape as an error.
   */
  private def writeAndClose(stdin: OutputStream, input: Array[Byte]): Option[IOException] =
    val written =
      try
        stdin.write(input)
        stdin.flush()
        None
      catch case e: IOException => Some(e)
    val closed =
      try
        stdin.close()
        None
      catch case e: IOException => Some(e)
    written.orElse(closed)
