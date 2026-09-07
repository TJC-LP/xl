package com.tjclp.xl.cli.contract

import cats.effect.ExitCode

/**
 * Everything a run produced, before anything is printed (ADR-017 §2.4). [[Render]] turns it into
 * the bytes of either mode; `runStagedOutput`'s commit rule reads `outputComplete` exactly as it
 * read the field of the same name on the old `CommandOutcome`.
 *
 * Three shapes, built through the constructors on the companion so the exit code always follows the
 * error code table ([[ExitCodes.forCode]]) and `ok ⇔ error.isEmpty` holds:
 *   - [[Outcome.ok]]: a payload, no error, exit 0
 *   - [[Outcome.failed]]: an error, no payload, exit 2 or 3 — nothing was written
 *   - [[Outcome.signal]]: BOTH — findings or a failed gate (exit 1) whose report is the payload:
 *     `diff` differs, `lint` findings, a write's `--strict` summary
 *
 * @param verb
 *   the subcommand path joined by a space (`"sheets hide"`); best-effort for a usage error raised
 *   before dispatch
 */
final case class Outcome(
  verb: String,
  payload: Option[Payload],
  warnings: Vector[Warning],
  error: Option[CliError],
  exitCode: ExitCode,
  outputComplete: Boolean
) derives CanEqual:

  def ok: Boolean = error.isEmpty

  /** The staging step's verdict: what the payload's `saved`/`written` now say. */
  def committed(target: Option[String]): Outcome =
    copy(payload = payload.map(Payload.committed(_, target)))

object Outcome:

  def ok(verb: String, payload: Payload, warnings: Vector[Warning] = Vector.empty): Outcome =
    Outcome(verb, Some(payload), warnings, None, ExitCodes.ok, outputComplete = true)

  def failed(verb: String, error: CliError, warnings: Vector[Warning] = Vector.empty): Outcome =
    Outcome(verb, None, warnings, Some(error), error.exitCode, outputComplete = false)

  def signal(
    verb: String,
    payload: Payload,
    error: CliError,
    warnings: Vector[Warning] = Vector.empty
  ): Outcome =
    Outcome(verb, Some(payload), warnings, Some(error), error.exitCode, outputComplete = true)
