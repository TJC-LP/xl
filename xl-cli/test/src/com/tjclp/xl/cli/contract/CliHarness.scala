package com.tjclp.xl.cli.contract

import java.io.{ByteArrayOutputStream, PrintStream}
import java.nio.charset.StandardCharsets
import java.util.Locale

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
 * (batch parse warnings, numFmt hints) in their real order relative to the program's own writes.
 * Because those streams are process-global, every invocation takes a JVM-wide lock: suites using
 * the harness never race each other, whatever MUnit's parallelism.
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
          // the streamed tables' `write` goes to the descriptor in production; here to the buffer
          val io =
            CliIO.system.copy(stdin = IO.pure(stdin), write = text => IO.blocking(out.print(text)))
          redirected(out, err)(Cli.run(args, io)).map { code =>
            CliRun(
              code.code,
              outBuf.toString(StandardCharsets.UTF_8),
              errBuf.toString(StandardCharsets.UTF_8)
            )
          }
      }
    }

  /**
   * Swap the JVM streams and pin the default locale to `Locale.US` for the duration of `fa`,
   * restoring both however `fa` ends. The locale pin makes the goldens machine-independent: `stats`
   * renders through `f"$x%.2f"`, which formats with the JVM default locale, so a
   * `-Duser.language=de` JVM would print `42,50` for the same file.
   */
  private def redirected[A](out: PrintStream, err: PrintStream)(fa: IO[A]): IO[A] =
    IO.blocking {
      val previous = Ambient.capture()
      System.setOut(out)
      System.setErr(err)
      Locale.setDefault(Locale.US)
      previous
    }.bracket(_ => fa) { previous =>
      IO.blocking {
        out.flush()
        err.flush()
        previous.restore()
      }
    }

  /**
   * The process-global state a run replaces: both JVM streams and every default-locale category.
   */
  private final case class Ambient(
    out: PrintStream,
    err: PrintStream,
    locale: Locale,
    formatLocale: Locale,
    displayLocale: Locale
  ):
    def restore(): Unit =
      System.setOut(out)
      System.setErr(err)
      Locale.setDefault(locale)
      Locale.setDefault(Locale.Category.FORMAT, formatLocale)
      Locale.setDefault(Locale.Category.DISPLAY, displayLocale)

  private object Ambient:
    def capture(): Ambient =
      Ambient(
        System.out,
        System.err,
        Locale.getDefault,
        Locale.getDefault(Locale.Category.FORMAT),
        Locale.getDefault(Locale.Category.DISPLAY)
      )
