package com.tjclp.xl.cli.contract

import cats.effect.IO

import com.tjclp.xl.cli.CliIO
import com.tjclp.xl.error.XLError

/**
 * The text form of errors and warnings on stderr (ADR-017 §2.3). Results go to stdout; everything
 * here goes to stderr, so an agent that captures stdout never has to parse an error out of a
 * result.
 *
 * {{{
 * Error: <message>                 today's exact first line
 *   code: <CODE>                   stable, one of ErrorCode.all
 *   did you mean: <a>, <b>         only when candidates is non-empty
 *   hint: <text>                   only when the error has one
 * }}}
 *
 * The format itself lives in xl-core as `XLError.renderDiagnostic` (GH-589), shared with the
 * scripting prelude's `exitMessage`/`orExit`: a failing script prints the same bytes as `xl`.
 */
object Diagnostics:

  def render(err: CliError): String =
    XLError.renderDiagnostic(err.code, err.message, err.hint, err.candidates)

  def report(err: CliError, io: CliIO): IO[Unit] = io.err(render(err))

  def renderWarning(warning: Warning): String =
    s"Warning[${warning.code}]: ${warning.message}"

  def warn(warning: Warning, io: CliIO): IO[Unit] = io.err(renderWarning(warning))
