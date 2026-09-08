package com.tjclp.xl.cli.helpers

import scala.annotation.tailrec

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row}
import com.tjclp.xl.cells.{CellValue, FormulaKind}
import com.tjclp.xl.formula.SheetEvaluator
import com.tjclp.xl.formula.eval.DependentRecalculation.*
import com.tjclp.xl.formula.eval.EvalFormulaSupport
import com.tjclp.xl.sheets.Sheet

/**
 * The `copy` verb and the `copy` batch op share this forwarder onto `Sheet.copyRange` /
 * `copyRangeFrom` (ADR-017 §2.12, W2.1). The pure semantics live in xl-core — source cells
 * snapshotted before any write (an overlapping same-sheet copy reads the pre-copy state),
 * cross-sheet targets, relative references shifted through the evaluator's `FormulaSupport`
 * (unparseable text copies as written), styles travelling with their cells, `valuesOnly` pasting
 * cached values. What stays here is what needs the evaluator: a `valuesOnly` paste of an UNCACHED
 * formula evaluates it against the source sheet, copied formulas are cached against the finished
 * target sheet (a fixed-point pass, so lookups over freshly copied siblings see their caches), and
 * the target sheet's transitive dependents are recalculated unless `recalcDependents` is false
 * (GH-468 `--no-recalc`, which leaves every other cache as it was). `cacheCopiedFormulas = false`
 * lets a caller own the complete post-copy evaluation and its diagnostics.
 */
object CopyOps:

  /**
   * Validate that `source` and `target` have identical dimensions.
   *
   * Returns `Left` with a user-facing message when they differ, `Right(())` when they match.
   * Extracted so both CLI and batch paths produce the same error text.
   */
  def validateDimensions(source: CellRange, target: CellRange): Either[String, Unit] =
    val srcRowCount = Row.index0(source.rowEnd) - Row.index0(source.rowStart) + 1
    val srcColCount = Column.index0(source.colEnd) - Column.index0(source.colStart) + 1
    val tgtRowCount = Row.index0(target.rowEnd) - Row.index0(target.rowStart) + 1
    val tgtColCount = Column.index0(target.colEnd) - Column.index0(target.colStart) + 1
    if srcRowCount == tgtRowCount && srcColCount == tgtColCount then Right(())
    else
      Left(
        s"Source (${srcRowCount}x${srcColCount}) and target (${tgtRowCount}x${tgtColCount}) dimensions mismatch. " +
          "Use a single target cell to auto-expand, or specify matching range."
      )

  /**
   * Copy `sourceRange` cells from `sourceSheet` into `targetSheet` at `targetRange` (same
   * dimensions — the caller validates with [[validateDimensions]]; the same sheet when the names
   * agree). Empty source cells leave their target unchanged. `Left` only when the formula support
   * refuses a shift, which the evaluator's never does; the CLI paths map it to a domain error.
   */
  def copyRange(
    wb: Workbook,
    sourceSheet: Sheet,
    sourceRange: CellRange,
    targetSheet: Sheet,
    targetRange: CellRange,
    valuesOnly: Boolean,
    recalcDependents: Boolean = true,
    cacheCopiedFormulas: Boolean = true
  ): XLResult[Workbook] =
    val colDelta = Column.index0(targetRange.start.col) - Column.index0(sourceRange.start.col)
    val rowDelta = Row.index0(targetRange.start.row) - Row.index0(sourceRange.start.row)
    def targetOf(src: ARef): ARef =
      ARef.from0(Column.index0(src.col) + colDelta, Row.index0(src.row) + rowDelta)

    // The source formula cells that paste as formulas (values-only: as their value). A data-table
    // record always pastes its cached constant (GH-430) and needs no evaluator phase.
    val liveFormulas: Vector[(ARef, String, Option[CellValue])] =
      sourceRange.cells.toVector.flatMap { src =>
        sourceSheet.cells.get(src).map(_.value) match
          case Some(CellValue.Formula(_, _, _: FormulaKind.DataTable)) => None
          case Some(CellValue.Formula(expr, cached, _)) => Some((src, expr, cached))
          case _ => None
      }

    val copied: XLResult[Sheet] =
      if sourceSheet.name == targetSheet.name then
        sourceSheet.copyRange(sourceRange, targetRange, valuesOnly)(using EvalFormulaSupport)
      else
        targetSheet.copyRangeFrom(sourceSheet, sourceRange, targetRange, valuesOnly)(using
          EvalFormulaSupport
        )

    copied.map { pure =>
      val materialized =
        if valuesOnly then
          // Sheet.copyRange pastes Empty for an uncached formula; the CLI evaluates it against the
          // source sheet first and pastes the value when the evaluation succeeds.
          liveFormulas.foldLeft(pure) {
            case (s, (src, expr, None)) =>
              SheetEvaluator
                .evaluateFormula(sourceSheet)(
                  s"=$expr",
                  workbook = Some(wb),
                  currentCell = Some(src)
                )
                .fold(_ => s, value => s.put(targetOf(src), value))
            case (s, _) => s
          }
        else pure
      val cached =
        if valuesOnly || !cacheCopiedFormulas || liveFormulas.isEmpty then materialized
        else populateFormulaCaches(materialized, wb, liveFormulas.map((src, _, _) => targetOf(src)))
      val wbWithCopied = wb.put(cached)
      if recalcDependents then
        wbWithCopied.recalculateDependents(cached.name, targetRange.cells.toSet)
      else wbWithCopied
    }

  /**
   * Populate caches for the copied formulas against the final target-sheet state.
   *
   * Some evaluator paths (notably lookup/match functions over ranges) only consult formula caches
   * rather than recursively evaluating uncached sibling formulas. We therefore iterate until the
   * copied formula caches reach a fixed point, rebuilding workbook context from the current sheet
   * state on each step so same-sheet-qualified refs like `Sheet1!A1:B2` also see freshly-populated
   * sibling caches.
   *
   * The pass count is bounded by the number of copied formulas. In the worst acyclic case, each
   * pass resolves one more layer of dependencies. Cyclic or otherwise unevaluable formulas remain
   * uncached (`None`), matching the prior best-effort behavior.
   */
  private def populateFormulaCaches(sheet: Sheet, wb: Workbook, refs: Vector[ARef]): Sheet =
    @tailrec
    def loop(currentSheet: Sheet, passesRemaining: Int): Sheet =
      val nextSheet = refs.foldLeft(currentSheet) { (s, ref) =>
        s.cells.get(ref).map(_.value) match
          case Some(CellValue.Formula(expr, _, _)) =>
            val cached =
              SheetEvaluator.evaluateCell(s)(ref, workbook = Some(wb.put(s))).toOption
            s.put(ref, CellValue.Formula(expr, cached))
          case _ => s
      }
      if nextSheet == currentSheet || passesRemaining <= 1 then nextSheet
      else loop(nextSheet, passesRemaining - 1)

    loop(sheet, refs.length.max(1))
