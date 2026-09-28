package com.tjclp.xl.cli.contract

import java.nio.file.NoSuchFileException

import scala.util.control.NoStackTrace

import cats.effect.ExitCode
import org.xml.sax.SAXParseException

import com.tjclp.xl.cli.{CliIO, MemoryGuard, StrictFailure}
import com.tjclp.xl.cli.raster.RasterError
import com.tjclp.xl.error.{XLError, XLException}

/** Where a failure happened, as far as the raising site knows. Every field is optional. */
final case class Location(
  file: Option[String],
  sheet: Option[String],
  ref: Option[String],
  opIndex: Option[Int]
) derives CanEqual

object Location:
  val none: Location = Location(None, None, None, None)
  def file(path: String): Location = none.copy(file = Some(path))

/**
 * A CLI failure as the agent sees it (ADR-017 §2.3): a stable `code` (an `XLError.code` for domain
 * failures, an [[ErrorCode]] constant for CLI-only ones), the human message — today's exact text —
 * and the optional `hint`, `candidates` ("did you mean"), `location` and domain `cause`. The exit
 * code follows from the code alone ([[ExitCodes.forCode]]), so the table cannot drift per site.
 */
final case class CliError(
  code: String,
  message: String,
  hint: Option[String] = None,
  candidates: Vector[String] = Vector.empty,
  location: Option[Location] = None,
  cause: Option[XLError] = None
) derives CanEqual:
  def exitCode: ExitCode = ExitCodes.forCode(code)

/**
 * Carrier through `IO`. `getMessage` is `error.message`, so every `getMessage.contains(...)`
 * assertion written against the plain-`Exception` raise sites keeps holding; stack-trace free
 * because nothing here is a defect to debug.
 */
final class CliException(val error: CliError) extends Exception(error.message) with NoStackTrace

/**
 * A run that COMPLETED with findings, carried through `IO` like [[CliException]]: the report is the
 * payload, the verdict is `error` (a signal code, exit 1 — `AUDIT_FINDINGS`). The runner turns it
 * into [[Outcome.signal]], so text mode prints the report with no `Error:` line and `--json` keeps
 * it as `data`. Stack-trace free: nothing here is a defect to debug.
 */
final class CliSignal(val payload: Payload, val error: CliError)
    extends Exception(error.message)
    with NoStackTrace

object CliError:

  /**
   * The projection from the domain: `code = e.code`, `hint = e.hint`, `candidates = e.candidates`.
   * A caller holding the workbook adds `Suggest.closest` candidates for a `SheetNotFound` itself.
   */
  def fromXLError(e: XLError, at: Option[Location]): CliError =
    CliError(e.code, e.message, e.hint, e.candidates, at, Some(e))

  /**
   * Classify anything a handler can raise: a [[CliException]] is its error; a [[StrictFailure]] is
   * the `RECALC_GATE` (exit 1) carrying the summary as its message; an `XLException` projects its
   * `XLError`; a raster export with no backend to run it (none installed, or the `--rasterizer`
   * asked for is missing) is `RASTERIZER_UNAVAILABLE` with the chain's own install hints as the
   * message, and so is a forced backend that cannot write the format or that ran and failed (its
   * exit and stderr are the message, GH-673); a raster backend whose scratch SVG could not be
   * created (resvg, in `XL_SPILL_DIR` or `java.io.tmpdir`) is `IO_WRITE` with a hint naming the
   * directory and the lever, in the shape of the streaming writer's spill failure; a
   * `NoSuchFileException` is `IO_READ`; an `OutOfMemoryError` is the `RESOURCE_LIMIT` failure
   * [[com.tjclp.xl.cli.MemoryGuard.exhausted]] (GH-636 — the classification of a heap that ran out;
   * note that cats-effect halts on a fatal error before any handler runs, so the error only ever
   * arrives here after `MemoryGuard.blocking` caught it inside its thunk); a `SAXParseException` —
   * a part the streaming reader could not parse — is `IO_READ` with the parser's message and
   * position, as the loaded reader's `ParseError` is (GH-635; the streaming source adds the file
   * and sheet); a stdout that refused a streamed table's bytes
   * ([[com.tjclp.xl.cli.CliIO.StdoutFailed]]) is `IO_WRITE`; everything else — every other `Error`
   * included — is `INTERNAL` with its message (falling back to `toString` when the message is null
   * — an un-migrated `new Exception(msg)` still yields a well-formed diagnostic).
   */
  def fromThrowable(t: Throwable): CliError = t match
    case e: CliException => e.error
    case s: CliSignal => s.error
    case s: StrictFailure => CliError(ErrorCode.RECALC_GATE, s.summary)
    case x: XLException => fromXLError(x.error, None)
    case r: RasterError.NoRasterizerAvailable =>
      CliError(ErrorCode.RASTERIZER_UNAVAILABLE, r.message)
    case r: RasterError.RasterizerNotFound => CliError(ErrorCode.RASTERIZER_UNAVAILABLE, r.message)
    case f: RasterError.FormatNotSupported =>
      CliError(
        ErrorCode.RASTERIZER_UNAVAILABLE,
        f.message,
        hint = Some(
          s"pick a backend that writes ${f.format.extension} with --rasterizer " +
            "(imagemagick writes png, jpeg, webp and pdf); `xl rasterizers` lists the backends " +
            "on this machine"
        )
      )
    case c: RasterError.ConversionFailed =>
      CliError(
        ErrorCode.RASTERIZER_UNAVAILABLE,
        c.message,
        hint = Some(
          s"${c.rasterizer} ran and failed; retry with another --rasterizer " +
            "(`xl rasterizers` lists them) and report the backend's message"
        )
      )
    case s: RasterError.ScratchFileFailed =>
      CliError(
        ErrorCode.IO_WRITE,
        s"cannot write the ${s.rasterizer} scratch SVG: ${messageOf(s.cause)}",
        hint = Some(
          s"check that ${MemoryGuard.spillWhere(s.spillDir)} exists, is writable and has room"
        )
      )
    case n: NoSuchFileException =>
      CliError(ErrorCode.IO_READ, s"No such file: ${Option(n.getFile).getOrElse(messageOf(n))}")
    case _: OutOfMemoryError => MemoryGuard.exhausted
    case sax: SAXParseException => CliError(ErrorCode.IO_READ, parseFailure(sax))
    case out: CliIO.StdoutFailed =>
      CliError(
        ErrorCode.IO_WRITE,
        messageOf(out),
        hint = Some("check the device or file stdout is redirected to")
      )
    case other => CliError(ErrorCode.INTERNAL, messageOf(other))

  /** A SAX parser's rejection with its position: `Parse error: <message> (line L, column C)`. */
  def parseFailure(sax: SAXParseException): String =
    s"Parse error: ${messageOf(sax)} (line ${sax.getLineNumber}, column ${sax.getColumnNumber})"

  /** The command line is wrong (exit 2). */
  def usage(message: String, hint: Option[String]): CliError =
    CliError(ErrorCode.USAGE, message, hint)

  /** Null-safe message: `getMessage`, else the exception's `toString`. */
  def messageOf(t: Throwable): String = Option(t.getMessage).getOrElse(t.toString)
