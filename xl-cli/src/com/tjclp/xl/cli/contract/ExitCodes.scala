package com.tjclp.xl.cli.contract

import cats.effect.ExitCode

/**
 * The exit-code table (ADR-017 §2.3, invariant 5: exit 1 is never a failure).
 *
 * {{{
 * 0  ok
 * 1  completed with findings or a failed gate (diff differs, lint findings, --strict) — file written as requested (-o), never with -i
 * 2  usage — the command line is wrong; nothing read, nothing written
 * 3  failed — the operation could not complete; nothing written
 * }}}
 */
object ExitCodes:
  val ok: ExitCode = ExitCode.Success
  val signal: ExitCode = ExitCode(1)
  val usage: ExitCode = ExitCode(2)
  val failed: ExitCode = ExitCode(3)

  private val signalCodes: Set[String] = Set(
    ErrorCode.RECALC_GATE,
    ErrorCode.DIFFERENCES_FOUND,
    ErrorCode.LINT_FINDINGS,
    ErrorCode.AUDIT_FINDINGS
  )

  private val usageCodes: Set[String] = Set(
    ErrorCode.USAGE,
    ErrorCode.UNKNOWN_VERB,
    ErrorCode.OUTPUT_REQUIRED,
    ErrorCode.UNSUPPORTED_IN_STREAM,
    ErrorCode.BATCH_JSON_INVALID,
    ErrorCode.BATCH_OP_UNKNOWN,
    ErrorCode.BATCH_OP_INVALID
  )

  /** Findings and gates exit 1, command-line mistakes 2, everything else (every domain code) 3. */
  def forCode(code: String): ExitCode =
    if signalCodes.contains(code) then signal
    else if usageCodes.contains(code) then usage
    else failed
