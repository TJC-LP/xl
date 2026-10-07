package com.tjclp.xl.cli.raster

import java.io.{IOException, InputStream, OutputStream}
import java.nio.charset.StandardCharsets
import java.util.concurrent.{ConcurrentHashMap, ExecutorService, Executors, TimeUnit}

import scala.annotation.tailrec
import scala.concurrent.duration.*
import scala.jdk.CollectionConverters.*

import cats.effect.{Deferred, IO, Resource}
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
 * — conversions and availability probes alike — has a deadline after which the child and its
 * descendants are killed with a bounded escalation ([[kill]]). The pipe I/O runs on daemon threads
 * that a deadline abandons rather than joins ([[onPipeThread]]): a descendant that outlives the
 * child keeps its pipes open, and no read, write or close on them may hold up the teardown.
 *
 * Batik renders in this JVM, not through here, so no deadline bounds it.
 */
private[raster] object PipedBackend:

  /** A conversion's deadline: a large sheet renders in seconds, a hung backend never finishes. */
  val ConversionTimeout: FiniteDuration = 5.minutes

  /** An availability probe's deadline (`--version`, ImageMagick's 1x1 conversion). */
  val ProbeTimeout: FiniteDuration = 30.seconds

  /** The stderr kept for a failure's message: its last bytes, where backends put the error. */
  val StderrLimit: Int = 4096

  /** The stdout a [[capture]] keeps: its first bytes (a version line, a delegate list). */
  private[raster] val CaptureLimit: Int = 1 << 20

  /** How long a child gets to exit after the polite signal before it is killed outright. */
  private val GraceMillis: Long = 2000

  /** How often a running child's descendants are recorded ([[recordDescendants]]). */
  private val TrackMillis: Long = 100

  /** How often [[kill]] looks for new descendants during the grace period. */
  private val PollMillis: Long = 25

  /**
   * How long the pipes may stay open after the child exits before their reads are abandoned: a
   * helper it started may hold them, and the child's exit (and its output file) is the outcome.
   */
  private val AfterExit: FiniteDuration = 1.second

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
    exchange(name, command :: args, input, keepStdout = false)
      .timeoutTo(timeout, IO.raiseError(RasterError.TimedOut(name, timeout)))
      .flatMap { outcome =>
        (outcome.exit, outcome.inputWritten) match
          case (0, true) => IO.unit
          case (exit @ 0, false) =>
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
    exchange(command, command :: args, input, keepStdout = false)
      .map(o => o.exit == 0 && o.inputWritten && o.stdoutBytes > 0)
      .timeoutTo(ProbeTimeout, IO.pure(false))
      .handleError(_ => false)

  /**
   * An availability probe (`rsvg-convert --version`): true when `command args` exits 0 within
   * `timeout`. A missing binary, a failure or a hang is false; a hung probe is killed (GH-690).
   */
  def probe(
    command: String,
    args: List[String],
    timeout: FiniteDuration = ProbeTimeout
  ): IO[Boolean] =
    capture(command, args, timeout).map(_.isDefined)

  /**
   * `command args`'s stdout (its first [[CaptureLimit]] bytes) when it exits 0 within `timeout`;
   * None for a missing binary, a failure or a hang, which is killed (GH-690).
   */
  def capture(
    command: String,
    args: List[String],
    timeout: FiniteDuration = ProbeTimeout
  ): IO[Option[String]] =
    exchange(command, command :: args, Array.emptyByteArray, keepStdout = true)
      .map(o => Option.when(o.exit == 0)(o.stdout))
      .timeoutTo(timeout, IO.pure(None))
      .handleError(_ => None)

  /**
   * What a child did with its input: exit code, whether the input write completed successfully,
   * stderr's tail, stdout's size and (when kept) its head. A failed or still-pending write cannot
   * establish success, even when the child's output pipes must be abandoned.
   */
  private final case class Outcome(
    exit: Int,
    inputWritten: Boolean,
    stderr: String,
    stdoutBytes: Long,
    stdout: String
  )

  /**
   * Spawn `command`, write `input` while draining stderr and stdout, and wait for the exit. A child
   * that exits while a helper it started still holds its pipes is not kept waiting on them: after
   * [[AfterExit]] the reads are abandoned and the outcome keeps its exit code and the input write's
   * status, without stderr's text. Only a completed successful write can make exit 0 a success.
   */
  private def exchange(
    name: String,
    command: List[String],
    input: Array[Byte],
    keepStdout: Boolean
  ): IO[Outcome] =
    spawn(name, command).use { process =>
      Deferred[IO, Boolean].flatMap { written =>
        val write = onPipeThread(writeAndClose(process.getOutputStream, input))
          .map(_.isEmpty)
          .flatTap(written.complete(_).void)
        val stderr = onPipeThread(tail(process.getErrorStream, StderrLimit))
        val stdout =
          onPipeThread(drain(process.getInputStream, if keepStdout then CaptureLimit else 0))
        val exited = IO.fromCompletableFuture(IO(process.onExit())) >> IO.sleep(AfterExit)
        (write, stderr, stdout).parTupled.race(exited).flatMap {
          case Left((inputWritten, err, (bytes, out))) =>
            IO.interruptible(process.waitFor())
              .map(exit => Outcome(exit, inputWritten, err, bytes, out))
          case Right(_) =>
            written.tryGet.flatMap { inputWritten =>
              IO(Outcome(process.exitValue, inputWritten.contains(true), HeldPipes, 0, ""))
            }
        }
      }
    }

  /** The stderr an [[Outcome]] carries when its pipes were abandoned after the exit. */
  private val HeldPipes = "its output pipes were still held open by a process it started"

  /** Daemon threads for the pipe I/O, so an abandoned one never keeps the JVM alive. */
  private lazy val pipeThreads: ExecutorService = Executors.newCachedThreadPool { runnable =>
    val thread = new Thread(runnable, "xl-raster-pipe")
    thread.setDaemon(true)
    thread
  }

  /**
   * `work` on a pipe thread. Canceling abandons it instead of waiting: a read or write on a pipe a
   * descendant still holds open blocks until that descendant exits, and a deadline must not
   * (GH-690). The thread ends on its own once the pipe closes.
   */
  private def onPipeThread[A](work: => A): IO[A] =
    IO.async[A] { callback =>
      IO {
        pipeThreads.execute(() => callback(Either.catchNonFatal(work)))
        Some(IO.unit)
      }
    }

  /**
   * The child, killed on release if it is still running (cancellation, a timeout, a failed drain),
   * its streams closed once it is gone — on a pipe thread, since closing a stream waits for a read
   * still blocked on it. Its descendants are recorded while it runs ([[recordDescendants]]) and any
   * still alive at release are killed too: once the child exits they are no longer its descendants,
   * and one holding the pipes would otherwise outlive the deadline (GH-690). That includes a helper
   * a backend meant to leave running: none of the converters daemonizes one. A command that cannot
   * be started — not installed, gone since the availability probe, not executable — is
   * [[RasterError.RasterizerNotFound]], the code the probe itself would have given.
   */
  private def spawn(name: String, command: List[String]): Resource[IO, Process] =
    val start = IO.blocking(new java.lang.ProcessBuilder(command.asJava).start()).adaptError {
      case e: IOException =>
        val program = command.headOption.getOrElse(name)
        val reason = Option(e.getMessage).getOrElse(e.toString)
        RasterError.RasterizerNotFound(name, s"`$program` could not be started: $reason.")
    }
    Resource
      .make(start.map(_ -> ConcurrentHashMap.newKeySet[ProcessHandle]()).flatTap {
        (process, seen) => IO(pipeThreads.execute(() => recordDescendants(process, seen)))
      }) { (process, seen) =>
        IO.blocking {
          kill(process)
          seen.asScala.filter(_.isAlive).foreach(_.destroyForcibly())
          pipeThreads.execute(() => closeQuietly(process))
        }
      }
      .map(_._1)

  /**
   * Add `process`'s descendants to `seen` every [[TrackMillis]] until it exits. Best-effort: a
   * helper started and orphaned within one interval is missed (the JDK cannot start a child in its
   * own process group).
   */
  @tailrec
  private def recordDescendants(process: Process, seen: java.util.Set[ProcessHandle]): Unit =
    seen.addAll(descendantsOf(process).asJava)
    if !process.waitFor(TrackMillis, TimeUnit.MILLISECONDS) then recordDescendants(process, seen)

  /** Best-effort discovery of a child's descendants: a sandbox may refuse the enumeration. */
  private def descendantsOf(process: Process): List[ProcessHandle] =
    try process.descendants().iterator().asScala.toList
    catch case _: RuntimeException => Nil

  /**
   * Stop `process` and the processes it started (`python3 -m cairosvg`, ImageMagick's delegates): a
   * polite signal, a bounded wait, then a forcible kill. Every step signals through
   * [[ProcessHandle]], which closes none of the child's streams: `Process.destroy` closes stdin
   * first, and that close waits on the lock a stdin write blocked on an unread pipe holds, so the
   * escalation would never be reached (GH-690). The descendants are looked up every [[PollMillis]]
   * through the grace period, so a helper the child's SIGTERM handler starts is stopped too, even
   * when the child then exits: a child that has exited no longer parents its helpers, so one
   * started in the last poll interval before the child exits can still escape (the JDK cannot start
   * a child in its own process group). Discovery is best-effort and cannot keep the child itself
   * from being stopped.
   */
  private[raster] def kill(
    process: Process,
    discover: Process => List[ProcessHandle] = descendantsOf
  ): Unit =
    def found(): List[ProcessHandle] =
      try discover(process)
      catch case _: RuntimeException => Nil
    if process.isAlive then
      val handle = process.toHandle
      val first = found()
      first.foreach(_.destroy())
      handle.destroy(): Unit
      val deadline = System.nanoTime() + GraceMillis * 1000000L
      // every descendant seen through the grace period, and whether the child exited within it
      @tailrec
      def watch(seen: Set[ProcessHandle]): (Set[ProcessHandle], Boolean) =
        val now = seen ++ found()
        if process.waitFor(PollMillis, TimeUnit.MILLISECONDS) then (now, true)
        else if System.nanoTime() >= deadline then (now, false)
        else watch(now)
      val (seen, exited) = watch(first.toSet)
      seen.foreach(_.destroyForcibly())
      if !exited then
        handle.destroyForcibly(): Unit
        process.waitFor(GraceMillis, TimeUnit.MILLISECONDS): Unit

  /** Close the child's three streams, ignoring the errors a dead child's pipes may raise. */
  private def closeQuietly(process: Process): Unit =
    List[java.io.Closeable](process.getOutputStream, process.getInputStream, process.getErrorStream)
      .foreach(stream =>
        try stream.close()
        catch case _: IOException => ()
      )

  /** Read `in` to its end: its size, and its first `keep` bytes as UTF-8. */
  private[raster] def drain(in: InputStream, keep: Int): (Long, String) =
    if keep <= 0 then (in.transferTo(OutputStream.nullOutputStream()), "")
    else
      val head = in.readNBytes(keep)
      val rest = in.transferTo(OutputStream.nullOutputStream())
      (head.length + rest, new String(head, StandardCharsets.UTF_8))

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
