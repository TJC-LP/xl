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

  /**
   * `view --eval` could not evaluate some formulas of the window's closure (without `--strict`, or
   * with it on a raster format, which never gates): the failing cells and the formulas blocked
   * behind them show the file's values, the rest is live. For a picture (html, svg, raster) the
   * closure includes what the window's conditional formatting reads.
   */
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

  /**
   * PR #659 review: `lint` found hygiene-tier findings only (shared-string orphans, unreferenced
   * parts) — the file opens intact, the exit is 0, and `--strict` would have made it 1.
   */
  val LINT_HYGIENE: String = "LINT_HYGIENE"

  /**
   * GH-497: an html/svg/raster `view` could not paint a conditional-format rule that applies to the
   * window — a kind xl does not evaluate yet (icon sets, above/below average, duplicate/unique
   * values, Excel 2010+ data bars), a rule formula that failed, or an evaluation that failed
   * outright — and the picture is drawn without that rule (`not rendered`); or a formula rule
   * failed at some cells only, and is painted wherever it evaluated (`partly rendered`).
   * Informational: never gates, `--strict` included.
   */
  val CF_NOT_RENDERED: String = "CF_NOT_RENDERED"

  /**
   * GH-714: an in-memory `putf` or batch `putf` left a plain formula whose value as a plain cell —
   * Excel's legacy implicit intersection — differs from its value as an array formula
   * (`=SUM(A1:A10*B1:B10)` in row 5 is `A5*B5`; `=A1:A10*2` cannot spill). Names the cells, both
   * values and the remedy: SUMPRODUCT, `putf --array`, or an explicit `@`. Informational: never
   * gates, `--strict` included.
   */
  val IMPLICIT_INTERSECTION: String = "IMPLICIT_INTERSECTION"

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
    OFF_GRID_REF,
    LINT_HYGIENE,
    CF_NOT_RENDERED,
    IMPLICIT_INTERSECTION
  )
