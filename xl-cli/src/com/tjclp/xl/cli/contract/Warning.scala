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

  /** `view --eval` without `--strict` could not evaluate the range; cached values were rendered. */
  val EVAL_FAILED: String = "EVAL_FAILED"

  /**
   * GH-636: an in-memory load under a lifted `--max-size` whose estimated footprint may exceed the
   * heap — the upper-bound estimate does, the lower-bound one does not. The load proceeds; if it
   * exhausts the heap it fails typed (`RESOURCE_LIMIT`) with the same `--stream`/`-Xmx` hint.
   */
  val MEMORY_PRESSURE: String = "MEMORY_PRESSURE"

  /**
   * GH-641: `stats` over a range holding no numbers. The run succeeds with zero-count statistics —
   * an empty result is a result, as for `search` and `filter` — and this names the likely mistake
   * (a text column) out of band.
   */
  val NO_NUMERIC_VALUES: String = "NO_NUMERIC_VALUES"

  /**
   * GH-628: a `putf` drag, batch `putf … from`, `fill` or `copy` wrote `#REF!` for a reference the
   * shift carried off the grid (before row 1 or column A, past XFD1048576). Excel writes the
   * `#REF!` silently; for an agent it is almost always a mistake, so every such cell is listed —
   * and `--strict` fails on it.
   */
  val OFF_GRID_REF: String = "OFF_GRID_REF"

  val all: Vector[String] = Vector(
    READER_WARNING,
    TRUNCATED,
    HIDDEN_OMITTED,
    UNKNOWN_PROPERTY,
    FORMAT_HINT_IGNORED,
    STREAM_BACKEND_ONLY,
    RECALC_ERRORS,
    SHEET_AUTOSELECTED,
    FLAG_IGNORED,
    EVAL_FAILED,
    MEMORY_PRESSURE,
    NO_NUMERIC_VALUES,
    OFF_GRID_REF
  )
