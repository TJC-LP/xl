package com.tjclp.xl.cli.helpers

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row}
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.formula.{FormulaParser, FormulaPrinter, FormulaShifter, SheetEvaluator}
import com.tjclp.xl.formula.eval.DependentRecalculation.*
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.sheets.styleSyntax.withCellStyle

import scala.annotation.tailrec

/**
 * Shared copy logic used by both the `copy` CLI command and the `copy` batch operation.
 *
 * Handles:
 *   - Overlapping copies (source and target ranges overlap in the same sheet). Source cells are
 *     snapshotted into an immutable map before mutation so each target read sees the pre-copy state
 *     regardless of iteration order.
 *   - Cross-sheet copies (source and target in different sheets).
 *   - Formula adjustment: relative references shifted by the source-to-target displacement.
 *   - `valuesOnly` mode: formulas materialized to their cached value, or evaluated against the
 *     source sheet if no cache is present.
 *   - Formula cache population against the final target-sheet state, so copied formulas see the
 *     destination context and any copied dependencies in the target range.
 *   - Transitive dependent recalculation on the target sheet after the copy completes.
 *   - Style preservation in all modes (source cell styles applied to target cells).
 */
object CopyOps:

  /** Formula cell written in phase 1; cache is populated after the target sheet is complete. */
  private final case class PendingFormulaCache(ref: ARef, expr: String)

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
   * Copy `sourceRange` cells from `sourceSheet` into `targetSheet` at `targetRange`.
   *
   * Requires `sourceRange.cells.size == targetRange.cells.size`; the caller must validate
   * dimensions first (see [[validateDimensions]]).
   *
   * When `sourceSheet.name == targetSheet.name` (same sheet), a single updated sheet is written
   * back. When they differ, only `targetSheet` is updated; the source sheet is untouched.
   *
   * Empty source cells (absent from `sourceSheet.cells`) are skipped — the target cell is left
   * unchanged rather than cleared.
   *
   * After phase-2 cache population, transitive dependents on the target sheet are recalculated so
   * that pre-existing formulas pointing into the target range see the new values.
   */
  def copyRange(
    wb: Workbook,
    sourceSheet: Sheet,
    sourceRange: CellRange,
    targetSheet: Sheet,
    targetRange: CellRange,
    valuesOnly: Boolean
  ): Workbook =
    // Snapshot source cells BEFORE any mutation so overlapping same-sheet copies are correct.
    val snapshot: Map[ARef, Cell] =
      sourceRange.cells.flatMap(ref => sourceSheet.cells.get(ref).map(ref -> _)).toMap

    val colDelta = Column.index0(targetRange.start.col) - Column.index0(sourceRange.start.col)
    val rowDelta = Row.index0(targetRange.start.row) - Row.index0(sourceRange.start.row)

    // Same-sheet vs cross-sheet: work on one sheet when both refer to the same one by name.
    val sameSheet = sourceSheet.name == targetSheet.name
    val workingSheet = if sameSheet then sourceSheet else targetSheet

    val (phase1Sheet, pendingFormulaCaches) =
      sourceRange.cells.foldLeft((workingSheet, Vector.empty[PendingFormulaCache])) {
        case ((s, pending), srcRef) =>
          val tgtRef = ARef.from0(
            Column.index0(srcRef.col) + colDelta,
            Row.index0(srcRef.row) + rowDelta
          )
          snapshot.get(srcRef) match
            case None => (s, pending) // Empty source cell: skip, leave target unchanged.
            case Some(srcCell) =>
              val (copiedValue, pendingFormula) =
                copiedCellValue(sourceSheet, wb, srcCell, colDelta, rowDelta, valuesOnly)
              val withValue = s.put(tgtRef, copiedValue)
              // Preserve style from source (looked up in source's registry, registered into target's).
              val withStyle = srcCell.styleId.flatMap(sourceSheet.styleRegistry.get) match
                case Some(style) => withValue.withCellStyle(tgtRef, style)
                case None => withValue
              val updatedPending = pendingFormula match
                case Some(expr) => pending :+ PendingFormulaCache(tgtRef, expr)
                case None => pending
              (withStyle, updatedPending)
      }

    val cachedSheet =
      if pendingFormulaCaches.isEmpty then phase1Sheet
      else populateFormulaCaches(phase1Sheet, wb, pendingFormulaCaches)

    // Recalculate transitive dependents (pre-existing formulas that reference the target range).
    val wbWithCopied = wb.put(cachedSheet)
    wbWithCopied.recalculateDependents(cachedSheet.name, targetRange.cells.toSet)

  /** Compute the copied value, shifting formulas unless `valuesOnly` is set. */
  private def copiedCellValue(
    sourceSheet: Sheet,
    wb: Workbook,
    srcCell: Cell,
    colDelta: Int,
    rowDelta: Int,
    valuesOnly: Boolean
  ): (CellValue, Option[String]) =
    srcCell.value match
      case CellValue.Formula(expr, cachedOpt) if valuesOnly =>
        // Materialize: prefer cached value, fall back to evaluating against source sheet.
        val materialized = cachedOpt.getOrElse {
          SheetEvaluator
            .evaluateFormula(sourceSheet)(
              s"=$expr",
              workbook = Some(wb),
              currentCell = Some(srcCell.ref)
            )
            .getOrElse(CellValue.Empty)
        }
        (materialized, None)

      case CellValue.Formula(expr, _) =>
        // Shift formula references by the displacement.
        FormulaParser.parse(s"=$expr") match
          case Right(parsed) =>
            val shifted = FormulaShifter.shift(parsed, colDelta, rowDelta)
            val shiftedExpr = FormulaPrinter.print(shifted, includeEquals = false)
            (CellValue.Formula(shiftedExpr, None), Some(shiftedExpr))
          case Left(_) =>
            // Unparseable formula: preserve as-is and try to cache in phase 2.
            (CellValue.Formula(expr, None), Some(expr))

      case other =>
        (other, None)

  /**
   * Populate caches for copied formulas against the final target-sheet state.
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
  private def populateFormulaCaches(
    sheet: Sheet,
    wb: Workbook,
    pendingFormulaCaches: Vector[PendingFormulaCache]
  ): Sheet =
    @tailrec
    def loop(currentSheet: Sheet, passesRemaining: Int): Sheet =
      val nextSheet = pendingFormulaCaches.foldLeft(currentSheet) {
        case (s, PendingFormulaCache(ref, expr)) =>
          val currentWorkbook = wb.put(s)
          val cached =
            SheetEvaluator
              .evaluateCell(s)(ref, workbook = Some(currentWorkbook))
              .toOption
          s.put(ref, CellValue.Formula(expr, cached))
      }

      if nextSheet == currentSheet || passesRemaining <= 1 then nextSheet
      else loop(nextSheet, passesRemaining - 1)

    loop(sheet, pendingFormulaCaches.length.max(1))
