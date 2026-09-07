package com.tjclp.xl.cli.contract

/**
 * A non-fatal condition a run wants the caller to know about. Rendered on stderr as
 * `Warning[<CODE>]: <message>` (see [[Diagnostics.renderWarning]]); the `--json` envelope lists
 * them under `warnings`.
 */
final case class Warning(code: String, message: String, location: Option[Location] = None)
    derives CanEqual

/** The warning vocabulary (ADR-017 §2.3). */
object WarningCode:
  /**
   * An `XlsxReader.Warning` raised while reading the input (missing styles part, malformed
   * drawing).
   */
  val READER_WARNING: String = "READER_WARNING"
  val TRUNCATED: String = "TRUNCATED"
  val HIDDEN_OMITTED: String = "HIDDEN_OMITTED"
  val UNKNOWN_PROPERTY: String = "UNKNOWN_PROPERTY"
  val FORMAT_HINT_IGNORED: String = "FORMAT_HINT_IGNORED"
  val STREAM_BACKEND_ONLY: String = "STREAM_BACKEND_ONLY"
  val RECALC_ERRORS: String = "RECALC_ERRORS"
  val SHEET_AUTOSELECTED: String = "SHEET_AUTOSELECTED"
  val FLAG_IGNORED: String = "FLAG_IGNORED"

  val all: Vector[String] = Vector(
    READER_WARNING,
    TRUNCATED,
    HIDDEN_OMITTED,
    UNKNOWN_PROPERTY,
    FORMAT_HINT_IGNORED,
    STREAM_BACKEND_ONLY,
    RECALC_ERRORS,
    SHEET_AUTOSELECTED,
    FLAG_IGNORED
  )
