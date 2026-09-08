package com.tjclp.xl.cli

import java.io.{FileDescriptor, FileOutputStream, IOException, OutputStream}
import java.nio.charset.StandardCharsets

import scala.util.control.NoStackTrace

import cats.effect.{IO, Ref}
import cats.syntax.all.*

/**
 * The CLI's standard streams as a value.
 *
 * Every byte an agent sees from `xl` — results, errors, help — leaves through `out`, `write` or
 * `err`, and `batch -` reads its JSON through `stdin`. Threading this record through the program
 * instead of calling `IO.println` / `System.err.println` directly is what lets the contract suite
 * run the real command tree in-process and pin stdout, stderr and the exit code separately
 * (`CliHarness` under xl-cli/test). [[CliIO.system]] is the process; [[CliIO.capturing]] is memory.
 *
 * `out` prints a line (a newline appended); `write` puts text on stdout exactly as given — the
 * channel of a table written as it streams (GH-635), whose fragments end where the rows end, not at
 * lines. Handler-level `System.err.println` calls (batch parse warnings in WriteCommands, the
 * numFmt hint in StyleBuilder) and `import-md -`'s `System.in` read are not routed here yet; the
 * harness covers them with a `System.setErr` bracket. The read commands' notices (truncation,
 * hidden lines, a failed `--eval`) go through the run's warning sink instead.
 */
final case class CliIO(
  out: String => IO[Unit],
  err: String => IO[Unit],
  stdin: IO[String],
  write: String => IO[Unit] = text => IO.blocking(System.out.print(text))
)

object CliIO:

  /**
   * stdout's reader went away — the pipe is closed (`xl … | head`), the stream shut — so a run that
   * was writing rows ends quietly with exit 0, as a tool killed by SIGPIPE would. Raised by
   * [[CliIO.system]]'s `write`.
   */
  object StdoutClosed extends RuntimeException("stdout closed") with NoStackTrace

  /**
   * stdout refused the bytes for another reason — no space left, an I/O error on the file it is
   * redirected to — which is a failure to report (`IO_WRITE`, exit 3), never a truncated file with
   * exit 0. Raised by [[CliIO.system]]'s `write`.
   */
  final class StdoutFailed(cause: IOException)
      extends RuntimeException(
        s"cannot write to stdout: ${Option(cause.getMessage).getOrElse(cause.toString)}",
        cause
      )
      with NoStackTrace

  /** The descriptor itself, unbuffered: every write is one syscall and its error is ours to see. */
  private lazy val rawStdout: OutputStream = new FileOutputStream(FileDescriptor.out)

  /** The errno spellings of a reader that is gone, across the JDK's platforms. */
  private def closedPipe(e: IOException): Boolean =
    val message = Option(e.getMessage).getOrElse("").toLowerCase
    message.contains("broken pipe") || message.contains("stream closed") ||
    message.contains("pipe is being closed") || message.contains("pipe has been ended")

  /**
   * The process streams. Each line write resolves `System.out` / `System.err` at execution time
   * (which is what `IO.println` does), so a test that swaps the JVM streams with `System.setOut`
   * observes the program's output — InPlaceSpec and the harness rely on that. `stdin` drains
   * standard input to a String exactly as `batch -` always has. `write` goes to the descriptor
   * behind stdout through a stream of our own, because a `PrintStream` swallows the `IOException`
   * of a failed write and leaves one sticky flag that cannot tell a closed pipe from a full disk:
   * here a closed pipe is [[StdoutClosed]] and anything else [[StdoutFailed]]. (A `System.setOut`
   * redirection does not see these bytes; the harness supplies its own `write`.)
   */
  val system: CliIO = CliIO(
    out = line => IO.println(line),
    err = line => IO.blocking(System.err.println(line)),
    stdin = IO.blocking(scala.io.Source.stdin.mkString),
    write = text =>
      IO.blocking {
        try rawStdout.write(text.getBytes(StandardCharsets.UTF_8))
        catch
          case e: IOException if closedPipe(e) => throw StdoutClosed
          case e: IOException => throw new StdoutFailed(e)
      }
  )

  /**
   * In-memory streams: `stdin` is the given text; the returned collector yields everything written
   * so far as `(stdout, stderr)`, each printed line terminated the way the console would show it
   * and every `write` as given. The sinks accumulate immutable vectors rather than a mutable
   * builder so a contended `Ref.update` retry cannot append a piece twice.
   */
  def capturing(stdin: String): IO[(CliIO, IO[(String, String)])] =
    (Ref.of[IO, Vector[String]](Vector.empty), Ref.of[IO, Vector[String]](Vector.empty)).mapN {
      (outRef, errRef) =>
        val io = CliIO(
          out = line => outRef.update(_ :+ (line + System.lineSeparator())),
          err = line => errRef.update(_ :+ (line + System.lineSeparator())),
          stdin = IO.pure(stdin),
          write = text => outRef.update(_ :+ text)
        )
        val collect = (outRef.get, errRef.get).mapN((out, err) => (out.mkString, err.mkString))
        (io, collect)
    }
