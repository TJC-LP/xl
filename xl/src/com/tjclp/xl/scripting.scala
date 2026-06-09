package com.tjclp.xl

/**
 * Scripting prelude: everything a script needs in one import.
 *
 * {{{
 * //> using dep com.tjclp::xl:0.11.0
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
  // rich text, display interpolator, compile-time literals (ref/col/fx/money/percent/date)
  export com.tjclp.xl.syntax.*
  export com.tjclp.xl.syntax.given

  // Formula system: FormulaParser, Sheet/Workbook evaluator extensions, DependencyGraph, Clock
  export com.tjclp.xl.formulaExports.*

  // Sync read/write/modify facade + streaming IO escape hatch
  export com.tjclp.xl.io.Excel
  export com.tjclp.xl.io.ExcelIO
  export com.tjclp.xl.io.RowData // streaming row type (readStream/writeStream)

  // The one sanctioned unwrap: .unsafe / .getOrElse on XLResult
  export com.tjclp.xl.unsafe.*

  // Script-only sugar: total smart detection of currency/percent/date/number/boolean from raw
  // strings ("$1,234.56".toFormatted → Currency). Kept out of the pure core import — heuristics
  // are script/CLI territory; the underlying FormattedParsers.detect is available everywhere.
  extension (s: String)
    def toFormatted: com.tjclp.xl.formatted.Formatted =
      com.tjclp.xl.formatted.FormattedParsers.detect(s)

  // Note: java.time types (LocalDate, LocalDateTime) cannot be re-exported — Scala's export
  // forwards only the type for Java classes, not the statics (LocalDate.of would not resolve).
  // Scripts import java.time directly when needed.
