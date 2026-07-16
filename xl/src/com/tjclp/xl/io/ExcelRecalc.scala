package com.tjclp.xl.io

import com.tjclp.xl.error.{XLException, XLResult}
import com.tjclp.xl.formula.eval.RecalcResult
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
 * recalculates them on open. Scripts that must not proceed on partial results use `result.toEither`
 * / `result.errors`:
 *
 * {{{
 * import com.tjclp.xl.scripting.{*, given}
 *
 * val result = Excel.writeRecalculated(wb, "out.xlsx")
 * if !result.isClean then result.errors.foreach(e => println(e.render))
 * }}}
 *
 * Lives in the aggregate `xl` module — the only module that sees both the evaluator and the sync IO
 * facade (xl-cats-effect cannot depend on xl-evaluator). Explicit overloads instead of default
 * parameters: defaulted extension methods do not survive the prelude's wildcard export (see the
 * WorkbookEvaluator note).
 */
object ExcelRecalc:

  extension (excel: Excel.type)

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
