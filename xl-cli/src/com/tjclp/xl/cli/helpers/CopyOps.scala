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
 *   - Style preservation in all modes (source cell styles applied to target cells).
 */
object CopyOps:

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

    val updated = sourceRange.cells.foldLeft(workingSheet) { (s, srcRef) =>
      val tgtRef = ARef.from0(
        Column.index0(srcRef.col) + colDelta,
        Row.index0(srcRef.row) + rowDelta
      )
      snapshot.get(srcRef) match
        case None => s // Empty source cell: skip, leave target unchanged.
        case Some(srcCell) =>
          val withValue =
            writeCopiedValue(s, sourceSheet, wb, srcCell, tgtRef, colDelta, rowDelta, valuesOnly)
          // Preserve style from source (looked up in source's registry, registered into target's).
          srcCell.styleId.flatMap(sourceSheet.styleRegistry.get) match
            case Some(style) => withValue.withCellStyle(tgtRef, style)
            case None => withValue
    }

    wb.put(updated)

  /** Write the cell value at `tgtRef`, adjusting formulas unless `valuesOnly` is set. */
  private def writeCopiedValue(
    target: Sheet,
    sourceSheet: Sheet,
    wb: Workbook,
    srcCell: Cell,
    tgtRef: ARef,
    colDelta: Int,
    rowDelta: Int,
    valuesOnly: Boolean
  ): Sheet =
    srcCell.value match
      case CellValue.Formula(expr, cachedOpt) if valuesOnly =>
        // Materialize: prefer cached value, fall back to evaluating against source sheet.
        val materialized = cachedOpt.getOrElse {
          SheetEvaluator
            .evaluateFormula(sourceSheet)(s"=$expr", workbook = Some(wb))
            .getOrElse(CellValue.Empty)
        }
        target.put(tgtRef, materialized)

      case CellValue.Formula(expr, _) =>
        // Shift formula references by the displacement.
        FormulaParser.parse(s"=$expr") match
          case Right(parsed) =>
            val shifted = FormulaShifter.shift(parsed, colDelta, rowDelta)
            val shiftedExpr = FormulaPrinter.print(shifted, includeEquals = false)
            val cached = SheetEvaluator
              .evaluateFormula(sourceSheet)(s"=$shiftedExpr", workbook = Some(wb))
              .toOption
            target.put(tgtRef, CellValue.Formula(shiftedExpr, cached))
          case Left(_) =>
            // Unparseable formula: preserve as-is, try to cache.
            val cached = SheetEvaluator
              .evaluateFormula(sourceSheet)(s"=$expr", workbook = Some(wb))
              .toOption
            target.put(tgtRef, CellValue.Formula(expr, cached))

      case other =>
        target.put(tgtRef, other)
