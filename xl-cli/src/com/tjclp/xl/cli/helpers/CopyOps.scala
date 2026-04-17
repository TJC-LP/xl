package com.tjclp.xl.cli.helpers

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row}
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.formula.{FormulaParser, FormulaPrinter, FormulaShifter, SheetEvaluator}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.sheets.styleSyntax.withCellStyle

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
 *   - Style preservation in all modes (source cell styles applied to target cells).
 */
object CopyOps:

  /** Formula cell written in phase 1; cache is populated after the target sheet is complete. */
  private final case class PendingFormulaCache(ref: ARef, expr: String)

  /**
   * Copy `sourceRange` cells from `sourceSheet` into `targetSheet` at `targetRange`.
   *
   * Requires `sourceRange.cells.size == targetRange.cells.size`; the caller must validate
   * dimensions first.
   *
   * When `sourceSheet.name == targetSheet.name` (same sheet), a single updated sheet is written
   * back. When they differ, only `targetSheet` is updated; the source sheet is untouched.
   *
   * Empty source cells (absent from `sourceSheet.cells`) are skipped — the target cell is left
   * unchanged rather than cleared.
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

    val updated =
      if pendingFormulaCaches.isEmpty then phase1Sheet
      else populateFormulaCaches(phase1Sheet, wb, pendingFormulaCaches)

    wb.put(updated)

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

  /** Populate caches for copied formulas after the final target-sheet state has been assembled. */
  private def populateFormulaCaches(
    sheet: Sheet,
    wb: Workbook,
    pendingFormulaCaches: Vector[PendingFormulaCache]
  ): Sheet =
    pendingFormulaCaches.foldLeft(sheet) { case (s, PendingFormulaCache(ref, expr)) =>
      val workbookWithCurrentSheet = wb.put(s)
      val cached =
        SheetEvaluator
          .evaluateCell(s)(ref, workbook = Some(workbookWithCurrentSheet))
          .toOption
      s.put(ref, CellValue.Formula(expr, cached))
    }
