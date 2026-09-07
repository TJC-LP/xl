package com.tjclp.xl.cli.contract

import java.io.{ByteArrayOutputStream, PrintStream}
import java.nio.charset.StandardCharsets

import cats.effect.IO
import cats.effect.std.Semaphore
import cats.effect.unsafe.implicits.global

import com.tjclp.xl.cli.{Cli, CliIO}

/** One in-process `xl` invocation as a shell — or an agent's tool call — observes it. */
final case class CliRun(exit: Int, stdout: String, stderr: String)

/**
 * Runs the real command tree in-process and captures exit code, stdout and stderr separately.
 *
 * The program runs through [[Cli.run]] with the production sinks ([[CliIO.system]], which resolve
 * `System.out` / `System.err` per write) while both JVM streams are redirected into buffers, and
 * with `stdin` replaced by the given text. Redirecting the JVM streams — rather than only handing
 * the program in-memory sinks — is what also captures the handler-level `System.err.println` calls
 * (batch parse warnings, numFmt hints, truncation notices) in their real order relative to the
 * program's own writes. Because those streams are process-global, every invocation takes a JVM-wide
 * lock: suites using the harness never race each other, whatever MUnit's parallelism.
 *
 * An exception the program lets escape propagates out of the returned IO (the binary would print a
 * stack trace and exit 1): a crash is a test failure here, never a pinned contract.
 */
object CliHarness:

  /** JVM streams are global: one invocation at a time across every suite in this test JVM. */
  private lazy val lock: Semaphore[IO] = Semaphore[IO](1).unsafeRunSync()

  /** Run with an empty stdin. */
  def run(args: String*): IO[CliRun] = run(args.toList, "")

  def run(args: List[String], stdin: String): IO[CliRun] =
    lock.permit.use { _ =>
      IO.blocking((new ByteArrayOutputStream(), new ByteArrayOutputStream())).flatMap {
        (outBuf, errBuf) =>
          val out = new PrintStream(outBuf, true, StandardCharsets.UTF_8)
          val err = new PrintStream(errBuf, true, StandardCharsets.UTF_8)
          val io = CliIO.system.copy(stdin = IO.pure(stdin))
          redirected(out, err)(Cli.run(args, io)).map { code =>
            CliRun(
              code.code,
              outBuf.toString(StandardCharsets.UTF_8),
              errBuf.toString(StandardCharsets.UTF_8)
            )
          }
      }
    }

  /** Swap the JVM streams for the duration of `fa`, restoring them however `fa` ends. */
  private def redirected[A](out: PrintStream, err: PrintStream)(fa: IO[A]): IO[A] =
    IO.blocking {
      val previous = (System.out, System.err)
      System.setOut(out)
      System.setErr(err)
      previous
    }.bracket(_ => fa) { (previousOut, previousErr) =>
      IO.blocking {
        out.flush()
        err.flush()
        System.setOut(previousOut)
        System.setErr(previousErr)
      }
    }
