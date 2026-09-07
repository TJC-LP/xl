package com.tjclp.xl.cli.contract

import cats.effect.IO

import com.tjclp.xl.cli.CliIO

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
 */
object Diagnostics:

  def render(err: CliError): String =
    val head = List(s"Error: ${err.message}", s"  code: ${err.code}")
    val suggestions =
      if err.candidates.isEmpty then Nil
      else List(s"  did you mean: ${err.candidates.mkString(", ")}")
    val hint = err.hint.toList.map(text => s"  hint: $text")
    (head ++ suggestions ++ hint).mkString("\n")

  def report(err: CliError, io: CliIO): IO[Unit] = io.err(render(err))

  def renderWarning(warning: Warning): String =
    s"Warning[${warning.code}]: ${warning.message}"

  def warn(warning: Warning, io: CliIO): IO[Unit] = io.err(renderWarning(warning))
