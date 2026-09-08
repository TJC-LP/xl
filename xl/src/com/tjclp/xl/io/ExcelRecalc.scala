package com.tjclp.xl.io

import com.tjclp.xl.error.{XLException, XLResult}
import com.tjclp.xl.formula.eval.{RecalcOptions, RecalcResult}
import com.tjclp.xl.formula.eval.WorkbookEvaluator.*
import com.tjclp.xl.formula.{Clock, Rng}
import com.tjclp.xl.workbooks.Workbook

/**
 * Recalculating write for the sync `Excel` facade (GH-360).
 *
 * `Excel.write` serializes formulas with whatever cached value they carry — a freshly built `fx"…"`
 * cell has none, so downstream cached-value consumers (openpyxl `data_only`, pandas, previewers,
 * Excel before its first recalc) see blanks. `Excel.writeRecalculated` closes that footgun: one
 * call recalculates the whole workbook (dependency-ordered, cross-sheet aware) and writes the
 * cached result.
 *
 * The workbook is written even when some formulas fail — errors are data conditions, reported in
 * the returned [[com.tjclp.xl.formula.eval.RecalcResult]]; failed cells stay uncached and Excel
 * recalculates them on open. GH-344: a formula that COMPUTES an Excel error value (#DIV/0!, #N/A,
 * ...) is a successful evaluation — the error value caches and writes as a `t="e"` cell; inspect
 * `result.excelErrors` for those. Scripts that must not proceed on partial results use
 * `result.toEither` / `result.errors` (and `result.excelErrors` to also reject error cells):
 *
 * {{{
 * import com.tjclp.xl.scripting.{*, given}
 *
 * val result = Excel.writeRecalculated(wb, "out.xlsx")
 * if !result.isClean then result.errors.foreach(e => println(e.render))
 * result.excelErrors.foreach((sheet, ref, err) => println(s"$${sheet.value}!$${ref.toA1}: $${err.toExcel}"))
 * }}}
 *
 * `Excel.writeChecked` (GH-589) is the lighter sibling for a freshly built or edited model: it
 * computes ONLY the formulas that have no cached value (`recalculateUncached`, ADR-017 §2.8) and
 * writes — every cache the book already carried, Excel's or another engine's, is written byte for
 * byte (the GH-468 doctrine), so nothing an agent authored ships blank and nothing it did not touch
 * moves. Reach for `writeRecalculated` when the caches themselves must be recomputed. `Excel.write`
 * keeps its 0.19 semantics: whatever cache a cell carries, nothing more.
 *
 * Lives in the aggregate `xl` module — the only module that sees both the evaluator and the sync IO
 * facade (xl-cats-effect cannot depend on xl-evaluator). Explicit overloads instead of default
 * parameters: defaulted extension methods do not survive the prelude's wildcard export (see the
 * WorkbookEvaluator note).
 */
object ExcelRecalc:

  extension (excel: Excel.type)

    /**
     * Fill in every UNCACHED formula (`recalculateUncached(RecalcOptions.default)`), write the
     * workbook to `path`, and return the [[com.tjclp.xl.formula.eval.RecalcResult]]. Cached cells
     * are never recomputed. Writes even when an uncached cell fails — errors are data; the failed
     * cell stays uncached and Excel computes it on open. Inspect `result.errors` /
     * `result.isClean`.
     */
    def writeChecked(workbook: Workbook, path: String): RecalcResult =
      excel.writeChecked(workbook, path, RecalcOptions.default)

    /**
     * [[writeChecked]] with explicit [[com.tjclp.xl.formula.eval.RecalcOptions]] — `clock` and
     * `rng` apply (deterministic TODAY/NOW, reproducible RAND); `iterative` and `parallelism` do
     * not, since only the dependency-ordered uncached pass runs (uncached cycle members are
     * reported, not iterated).
     */
    @annotation.targetName("writeCheckedWithOptions")
    def writeChecked(workbook: Workbook, path: String, options: RecalcOptions): RecalcResult =
      val result = workbook.recalculateUncached(options)
      Excel.write(result.workbook, path)
      result

    /**
     * `XLResult[Workbook]` overload: unwraps at the IO edge (throws
     * [[com.tjclp.xl.error.XLException]] on `Left`, before anything lands on disk), then
     * [[writeChecked]].
     */
    @annotation.targetName("writeCheckedResult")
    def writeChecked(result: XLResult[Workbook], path: String): RecalcResult =
      excel.writeChecked(result, path, RecalcOptions.default)

    /** `XLResult[Workbook]` overload with explicit options. */
    @annotation.targetName("writeCheckedResultWithOptions")
    def writeChecked(
      result: XLResult[Workbook],
      path: String,
      options: RecalcOptions
    ): RecalcResult =
      excel.writeChecked(result.fold(e => throw XLException(e), identity), path, options)

    /**
     * Recalculate every formula (system clock), write the cached workbook to `path`, and return the
     * [[com.tjclp.xl.formula.eval.RecalcResult]]. Writes even when formulas fail — inspect
     * `result.errors` / `result.isClean`.
     */
    def writeRecalculated(workbook: Workbook, path: String): RecalcResult =
      excel.writeRecalculated(workbook, path, Clock.system)

    /**
     * Recalculate with an explicit [[com.tjclp.xl.formula.Clock]] (deterministic TODAY/NOW), write
     * the cached workbook, and return the result.
     */
    @annotation.targetName("writeRecalculatedWithClock")
    def writeRecalculated(workbook: Workbook, path: String, clock: Clock): RecalcResult =
      val result = workbook.recalculate(clock)
      Excel.write(result.workbook, path)
      result

    /**
     * Recalculate with an explicit clock and randomness source (GH-115: `Rng.seeded` makes
     * RAND/RANDBETWEEN reproducible), write the cached workbook, and return the result.
     */
    @annotation.targetName("writeRecalculatedWithRng")
    def writeRecalculated(workbook: Workbook, path: String, clock: Clock, rng: Rng): RecalcResult =
      val result = workbook.recalculate(clock, rng)
      Excel.write(result.workbook, path)
      result

    /**
     * Recalculate with ONE [[com.tjclp.xl.formula.eval.RecalcOptions]] record (ADR-017 §2.8 —
     * clock, rng, iterative mode, parallelism, table seeding), write the cached workbook, and
     * return the result. `RecalcOptions.default` reproduces the two-argument form exactly.
     */
    @annotation.targetName("writeRecalculatedWithOptions")
    def writeRecalculated(workbook: Workbook, path: String, options: RecalcOptions): RecalcResult =
      val result = workbook.recalculate(options)
      Excel.write(result.workbook, path)
      result

    /**
     * Convenience overload mirroring `Excel.write(result, path)`: unwraps an `XLResult[Workbook]`
     * at the IO edge (throws [[com.tjclp.xl.error.XLException]] on `Left`), then recalculates and
     * writes with the system clock.
     */
    @annotation.targetName("writeRecalculatedResult")
    def writeRecalculated(result: XLResult[Workbook], path: String): RecalcResult =
      excel.writeRecalculated(result, path, Clock.system)

    /** `XLResult[Workbook]` overload with an explicit clock. */
    @annotation.targetName("writeRecalculatedResultWithClock")
    def writeRecalculated(result: XLResult[Workbook], path: String, clock: Clock): RecalcResult =
      excel.writeRecalculated(result.fold(e => throw XLException(e), identity), path, clock)
