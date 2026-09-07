package com.tjclp.xl.cli.contract

import com.tjclp.xl.error.XLError

/**
 * The CLI-only error codes (ADR-017 §2.3). Domain failures carry `XLError.code`; these name the
 * failures only the command line can have. [[all]] is the complete vocabulary an agent can meet in
 * a `code:` line — every CLI code plus every `XLError` case code, each exactly once.
 */
object ErrorCode:
  val USAGE: String = "USAGE"
  val UNKNOWN_VERB: String = "UNKNOWN_VERB"
  val OUTPUT_REQUIRED: String = "OUTPUT_REQUIRED"
  val UNSUPPORTED_IN_STREAM: String = "UNSUPPORTED_IN_STREAM"
  val BATCH_JSON_INVALID: String = "BATCH_JSON_INVALID"
  val BATCH_OP_UNKNOWN: String = "BATCH_OP_UNKNOWN"
  val BATCH_OP_INVALID: String = "BATCH_OP_INVALID"
  val BATCH_OP_FAILED: String = "BATCH_OP_FAILED"
  val RASTERIZER_UNAVAILABLE: String = "RASTERIZER_UNAVAILABLE"
  val IO_READ: String = "IO_READ"
  val IO_WRITE: String = "IO_WRITE"
  val RECALC_GATE: String = "RECALC_GATE"
  val DIFFERENCES_FOUND: String = "DIFFERENCES_FOUND"
  val LINT_FINDINGS: String = "LINT_FINDINGS"
  val AUDIT_FINDINGS: String = "AUDIT_FINDINGS"
  val INTERNAL: String = "INTERNAL"

  /** The CLI-only codes, in the order above. */
  val cli: Vector[String] = Vector(
    USAGE,
    UNKNOWN_VERB,
    OUTPUT_REQUIRED,
    UNSUPPORTED_IN_STREAM,
    BATCH_JSON_INVALID,
    BATCH_OP_UNKNOWN,
    BATCH_OP_INVALID,
    BATCH_OP_FAILED,
    RASTERIZER_UNAVAILABLE,
    IO_READ,
    IO_WRITE,
    RECALC_GATE,
    DIFFERENCES_FOUND,
    LINT_FINDINGS,
    AUDIT_FINDINGS,
    INTERNAL
  )

  /** Every code a `code:` line can carry: the CLI codes, then `XLError.codes`. Unique by test. */
  val all: Vector[String] = (cli ++ XLError.codes).distinct
