package com.tjclp.xl.cli.raster

import java.io.{IOException, InputStream, OutputStream}
import java.nio.charset.StandardCharsets
import java.util.concurrent.TimeUnit

import scala.annotation.tailrec
import scala.concurrent.duration.*
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
 *
 * GH-690: the message keeps only the tail of stderr ([[StderrLimit]]), a child that cannot be
 * started is [[RasterError.RasterizerNotFound]] rather than a raw `IOException`, and every exchange
 * has a deadline after which the child and its descendants are killed.
 */
private[raster] object PipedBackend:

  /** A conversion's deadline: a large sheet renders in seconds, a hung backend never finishes. */
  val ConversionTimeout: FiniteDuration = 5.minutes

  /** An availability probe's deadline (`--version`, ImageMagick's 1x1 conversion). */
  val ProbeTimeout: FiniteDuration = 30.seconds

  /** The stderr kept for a failure's message: its last bytes, where backends put the error. */
  val StderrLimit: Int = 4096

  /**
   * Run `command args` with `input` on stdin. A non-zero exit is `ConversionFailed(name, stderr,
   * exit)` whether or not the backend read its input; a zero exit that left the input unread is a
   * failure too, since the output cannot be trusted. Past `timeout` the child is killed and the run
   * is [[RasterError.TimedOut]].
   */
  def run(
    name: String,
    command: String,
    args: List[String],
    input: Array[Byte],
    timeout: FiniteDuration = ConversionTimeout
  ): IO[Unit] =
    exchange(name, command :: args, input)
      .timeoutTo(timeout, IO.raiseError(RasterError.TimedOut(name, timeout)))
      .flatMap { outcome =>
        (outcome.exit, outcome.unread) match
          case (0, None) => IO.unit
          case (exit @ 0, Some(_)) =>
            // the JDK's own wording ("Stream closed", "Broken pipe") says nothing the user can act on
            val detail = if outcome.stderr.isBlank then "" else s"; ${outcome.stderr}"
            IO.raiseError(
              RasterError.ConversionFailed(
                name,
                s"exited before reading the whole SVG$detail",
                exit
              )
            )
          case (exit, _) =>
            IO.raiseError(RasterError.ConversionFailed(name, outcome.stderr, exit))
      }

  /**
   * A probe conversion: true when `command args` exited 0 within [[ProbeTimeout]], wrote something
   * to stdout, and its stdin write did not fail. An input that fits the pipe buffer is written
   * whether or not the child reads it, so this does not prove the child read it; the non-empty
   * stdout is the evidence it converted. Any failure — a missing binary, an early exit, a closed
   * pipe, a hang — is false, never an error or a stack trace (ImageMagick's delegate check,
   * GH-673).
   */
  def succeeds(command: String, args: List[String], input: Array[Byte]): IO[Boolean] =
    exchange(command, command :: args, input)
      .map(o => o.exit == 0 && o.unread.isEmpty && o.stdoutBytes > 0)
      .timeoutTo(ProbeTimeout, IO.pure(false))
      .handleError(_ => false)

  /**
   * What a child did with its input: exit code, the write's failure, stderr's tail, stdout's size.
   */
  private final case class Outcome(
    exit: Int,
    unread: Option[IOException],
    stderr: String,
    stdoutBytes: Long
  )

  /** Spawn `command`, write `input` while draining stderr and stdout, and wait for the exit. */
  private def exchange(name: String, command: List[String], input: Array[Byte]): IO[Outcome] =
    spawn(name, command).use { process =>
      val stop = IO.blocking(kill(process))
      val write = IO.blocking(writeAndClose(process.getOutputStream, input)).cancelable(stop)
      val stderr = IO.blocking(tail(process.getErrorStream, StderrLimit)).cancelable(stop)
      val stdout =
        IO.blocking(process.getInputStream.transferTo(OutputStream.nullOutputStream()))
          .cancelable(stop)
      (write, stderr, stdout).parTupled.flatMap { (unread, err, bytes) =>
        IO.interruptible(process.waitFor()).map(exit => Outcome(exit, unread, err, bytes))
      }
    }

  /**
   * The child, killed on release if it is still running (cancellation, a timeout, a failed drain).
   * A command that cannot be started — not installed, gone since the availability probe, not
   * executable — is [[RasterError.RasterizerNotFound]], the code the probe itself would have given.
   */
  private def spawn(name: String, command: List[String]): Resource[IO, Process] =
    Resource.make(
      IO.blocking(new java.lang.ProcessBuilder(command.asJava).start()).adaptError {
        case e: IOException =>
          val program = command.headOption.getOrElse(name)
          val reason = Option(e.getMessage).getOrElse(e.toString)
          RasterError.RasterizerNotFound(name, s"`$program` could not be started: $reason.")
      }
    )(process => IO.blocking(kill(process)))

  /**
   * Stop `process` and the processes it started (`python3 -m cairosvg`, ImageMagick's delegates): a
   * polite destroy, then a forcible one for whatever outlives a short grace period.
   */
  private def kill(process: Process): Unit =
    if process.isAlive then
      val descendants = process.descendants().iterator().asScala.toList
      descendants.foreach(_.destroy())
      process.destroy()
      if !process.waitFor(2, TimeUnit.SECONDS) then process.destroyForcibly().waitFor()
      descendants.foreach(_.destroyForcibly())

  /**
   * The last `limit` bytes of `in`, read to its end, as UTF-8; a cut tail starts with `…` and drops
   * the partial character it was cut through. A backend that writes megabytes of stderr costs
   * `limit` bytes, not the whole stream (GH-690).
   */
  private[raster] def tail(in: InputStream, limit: Int): String =
    val ring = new Array[Byte](limit)
    val chunk = new Array[Byte](8192)
    @tailrec
    def fill(total: Long): Long =
      val n = in.read(chunk)
      if n < 0 then total
      else
        // only the chunk's last `limit` bytes can survive; copy them in at most two runs
        val from = math.max(0, n - limit)
        val count = n - from
        val at = ((total + from) % limit).toInt
        val first = math.min(count, limit - at)
        System.arraycopy(chunk, from, ring, at, first)
        System.arraycopy(chunk, from + first, ring, 0, count - first)
        fill(total + n)
    val total = fill(0L)
    if total <= limit then new String(ring, 0, total.toInt, StandardCharsets.UTF_8)
    else
      val start = (total % limit).toInt
      val bytes = ring.drop(start) ++ ring.take(start)
      val whole = bytes.dropWhile(b => (b & 0xc0) == 0x80) // UTF-8 continuation bytes
      "…" + new String(whole, StandardCharsets.UTF_8)

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
