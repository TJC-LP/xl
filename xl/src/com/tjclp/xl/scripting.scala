package com.tjclp.xl

/**
 * Scripting prelude: everything a script needs in one import.
 *
 * {{{
 * //> using dep com.tjclp::xl:0.20.0
 * import com.tjclp.xl.scripting.{*, given}
 *
 * val wb = Excel.read("input.xlsx")
 * val updated = wb.update("Sheet1", _.put(ref"A1", "Hello")).unsafe
 * Excel.write(updated, "output.xlsx")
 * }}}
 *
 * Bundles the core API, DSL operators, compile-time literals, formula evaluation, the sync `Excel`
 * facade, streaming `ExcelIO`, and the `.unsafe` boundary. Importing this prelude is the explicit
 * opt-in to script mode; `import com.tjclp.xl.{*, given}` remains the 100% pure alternative (no
 * `.unsafe` in scope).
 *
 * Use EITHER this import OR `com.tjclp.xl.{*, given}` — never both in one file, as the overlapping
 * forwarders become ambiguous.
 */
object scripting:
  // Core domain model + String parsing helpers (asCell, asRange, asSheetName).
  // Extension methods on opaque types (ARef.toA1/col/row/shift, Column.toLetter) resolve via
  // the companion implicit scope and need no export — see the note in api.scala.
  export com.tjclp.xl.api.*

  // DSL operators, codec/optics/patch/sheet syntax, style DSL, Easy Mode string extensions,
  // rich text, display interpolator, compile-time literals (ref/col/fx/money/percent/date).
  // The low-priority `default` FormulaDisplayStrategy (raw formula text) is excluded in favor
  // of the evaluator-backed `evaluating` strategy below: the LowPriority inheritance trick does
  // not survive export forwarding (both would sit at the same depth here and be ambiguous), so
  // the prelude makes the choice explicitly — scripts display computed, NumFmt-formatted values.
  export com.tjclp.xl.syntax.*
  export com.tjclp.xl.syntax.{default as _, given}

  // Formula system: FormulaParser, Sheet/Workbook evaluator extensions, DependencyGraph, Clock,
  // plus the `evaluating` display given (wildcard exports skip givens, so it needs the explicit
  // given selector — without it, excel"" interpolation prints raw formula text).
  // (formulaExports.given also carries an inherited `default` forwarder — EvaluatingFormulaDisplay
  // extends the LowPriority trait — so the exclusion is needed on this hop too.)
  export com.tjclp.xl.formulaExports.*
  export com.tjclp.xl.formulaExports.{default as _, given}

  // Sync read/write/modify facade + streaming IO escape hatch
  export com.tjclp.xl.io.Excel
  export com.tjclp.xl.io.ExcelIO
  export com.tjclp.xl.io.RowData // streaming row type (readStream/writeStream)

  // Recalculating writes (GH-360 writeRecalculated, GH-589 writeChecked): recalculate (all, or
  // only the uncached cells), write, return the RecalcResult. Defined in this aggregate module
  // (the only one seeing both the evaluator and the sync facade) as extensions on Excel.type; the
  // wildcard export puts them in scope here.
  export com.tjclp.xl.io.ExcelRecalc.*

  // The one sanctioned unwrap: .unsafe / .getOrElse on XLResult
  export com.tjclp.xl.unsafe.*

  // ADR-017 §2.12 (W2.1): the evaluator-backed FormulaSupport behind wb.edit / sheet.edit, stated
  // explicitly — wildcard exports skip givens, and a LowPriority default would not survive the
  // export hop (see the display-strategy note above). `Edit` and its targets come through api.*.
  given com.tjclp.xl.ops.FormulaSupport = com.tjclp.xl.formula.eval.EvalFormulaSupport

  // ===== GH-589 (W2.8) scripting completions — one block, integrate as a unit =====
  // Lightweight metadata returned by Excel.readMetadata: xl-ooxml types, not part of the core api.
  export com.tjclp.xl.ooxml.metadata.{LightMetadata, SheetInfo}

  /**
   * Unwrap an `XLResult` at the script edge or end the run: on `Left` the [[exitMessage]] goes to
   * stderr and the process exits with status 1. The script-shaped twin of `.unsafe` — a failed step
   * reports the error's code, hint and candidates instead of a stack trace, the way `xl` itself
   * fails.
   *
   * {{{
   * val wb = orExit(Workbook.named("Data", "Summary"))
   * val sheet = orExit(wb("Summary"))
   * }}}
   */
  def orExit[A](result: com.tjclp.xl.error.XLResult[A]): A = result match
    case Right(value) => value
    case Left(err) =>
      System.err.println(exitMessage(err))
      sys.exit(1)

  /**
   * What [[orExit]] prints: `error: <message>` and `code: <CODE>`, then a `hint:` line and a
   * `did you mean:` line when the error carries them (ADR-017 §2.7) — the CLI's text envelope, so a
   * script and `xl` fail the same way on the same error.
   */
  def exitMessage(err: com.tjclp.xl.error.XLError): String =
    val lines =
      Vector(s"error: ${err.message}", s"code: ${err.code}") ++
        err.hint.map(h => s"hint: $h") ++
        Option.when(err.candidates.nonEmpty)(s"did you mean: ${err.candidates.mkString(", ")}")
    lines.mkString("\n")
  // ===== end GH-589 block =====

  // Script-only sugar: total smart detection of currency/percent/date/number/boolean from raw
  // strings ("$1,234.56".toFormatted → Currency). Kept out of the pure core import — heuristics
  // are script/CLI territory; the underlying FormattedParsers.detect is available everywhere.
  extension (s: String)
    def toFormatted: com.tjclp.xl.formatted.Formatted =
      com.tjclp.xl.formatted.FormattedParsers.detect(s)

  // Note: java.time types (LocalDate, LocalDateTime) cannot be re-exported — Scala's export
  // forwards only the type for Java classes, not the statics (LocalDate.of would not resolve).
  // Scripts import java.time directly when needed.
