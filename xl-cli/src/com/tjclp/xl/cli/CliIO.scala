package com.tjclp.xl.cli

import cats.effect.{IO, Ref}
import cats.syntax.all.*

/**
 * The CLI's standard streams as a value.
 *
 * Every byte an agent sees from `xl` — results, errors, help — leaves through `out` or `err`, and
 * `batch -` reads its JSON through `stdin`. Threading this record through the program instead of
 * calling `IO.println` / `System.err.println` directly is what lets the contract suite run the real
 * command tree in-process and pin stdout, stderr and the exit code separately (`CliHarness` under
 * xl-cli/test). [[CliIO.system]] is the process; [[CliIO.capturing]] is memory.
 *
 * Handler-level `System.err.println` calls (batch parse warnings in WriteCommands, the numFmt hint
 * in StyleBuilder, truncation notices in the read commands) and `import-md -`'s `System.in` read
 * are not routed here yet; the harness covers them with a `System.setErr` bracket.
 */
final case class CliIO(out: String => IO[Unit], err: String => IO[Unit], stdin: IO[String])

object CliIO:

  /**
   * The process streams. Each write resolves `System.out` / `System.err` at execution time (which
   * is what `IO.println` does), so a test that swaps the JVM streams with `System.setOut` observes
   * the program's output — InPlaceSpec and the harness rely on that. `stdin` drains standard input
   * to a String exactly as `batch -` always has.
   */
  val system: CliIO = CliIO(
    out = line => IO.println(line),
    err = line => IO.blocking(System.err.println(line)),
    stdin = IO.blocking(scala.io.Source.stdin.mkString)
  )

  /**
   * In-memory streams: `stdin` is the given text; the returned collector yields everything written
   * so far as `(stdout, stderr)`, each line terminated the way the console would show it. The sinks
   * accumulate immutable vectors rather than a mutable builder so a contended `Ref.update` retry
   * cannot append a line twice.
   */
  def capturing(stdin: String): IO[(CliIO, IO[(String, String)])] =
    (Ref.of[IO, Vector[String]](Vector.empty), Ref.of[IO, Vector[String]](Vector.empty)).mapN {
      (outRef, errRef) =>
        val io = CliIO(
          out = line => outRef.update(_ :+ line),
          err = line => errRef.update(_ :+ line),
          stdin = IO.pure(stdin)
        )
        val collect = (outRef.get, errRef.get).mapN((out, err) => (render(out), render(err)))
        (io, collect)
    }

  private def render(lines: Vector[String]): String =
    lines.map(_ + System.lineSeparator()).mkString
