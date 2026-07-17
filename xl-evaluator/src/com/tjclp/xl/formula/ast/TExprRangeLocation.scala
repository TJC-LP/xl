package com.tjclp.xl.formula.ast

import com.tjclp.xl.formula.functions.FunctionSpecs
import com.tjclp.xl.formula.eval.EvalError
import com.tjclp.xl.formula.functions.EvalContext

import com.tjclp.xl.{ARef, CellRange, Column, Row, SheetName}

trait TExprRangeLocation:
  /**
   * Where a range is located - same sheet or cross-sheet.
   *
   * This enum unifies local and cross-sheet range references, eliminating the need for paired TExpr
   * cases (e.g., Min + SheetMin). Used by TExpr.Aggregate for unified aggregation.
   */
  enum RangeLocation derives CanEqual:
    case Local(range: CellRange)
    case CrossSheet(sheet: SheetName, range: CellRange)

    /**
     * GH-353: external-workbook range in a range-typed argument slot — `SUMIF([2]Book1!A1:A9, …)`.
     *
     * The range analog of [[TExpr.ExternalRange]] for function arguments: `workbookIndex` is the
     * 1-based external-link index, `sheetName` is the raw sheet (or workbook-file) name in that
     * EXTERNAL workbook. The target cells live outside this workbook, so the location contributes
     * no dependency edges and can never resolve to a sheet — evaluation yields
     * `Evaluator.externalRefUnsupported`; cells CONTAINING such calls are pinned to their
     * Excel-written cache upstream (SheetEvaluator.pinnedExternalCache).
     */
    case External(workbookIndex: Int, sheetName: String, range: CellRange)

    /**
     * GH-394: a defined name in a range-typed argument slot — `=VLOOKUP(x, named_table, 2)`,
     * `=SUMIF(rev_range, ">1")`, optionally sheet-qualified (`=SUMIF(Model!rev_range, …)`).
     *
     * The target range lives behind the workbook's name table, so it is UNKNOWN until evaluation:
     * `Evaluator.resolveRangeLocation` is the single resolution boundary — it looks the name up
     * (sheet-scoped shadows workbook-scoped, case-insensitive, relative to `scope` when qualified),
     * parses the refersTo text, and yields the resolved (sheet, CellRange) pair for range-shaped
     * targets (single-cell targets act as 1×1 ranges). Non-range targets (constants, formulas) are
     * a clean per-cell error VALUE. This case deliberately has NO static range: the pre-resolution
     * accessors below (`staticRange`, `localCells*`, …) return their empty forms for it, and every
     * evaluation-time consumer goes through the resolution boundary.
     */
    case Name(name: String, scope: Option[SheetName])

  object RangeLocation:
    extension (loc: RangeLocation)
      /**
       * The range carried STATICALLY by the location — None for [[RangeLocation.Name]], whose
       * target range is only known after workbook resolution (GH-394). Evaluation-time consumers
       * must use `Evaluator.resolveRangeLocation`, which returns the resolved (sheet, CellRange)
       * pair for every case; this accessor exists for structural walks (printing, shifting,
       * dependency extraction) that inspect locations without a workbook.
       */
      def staticRange: Option[CellRange] = loc match
        case Local(r) => Some(r)
        case CrossSheet(_, r) => Some(r)
        case External(_, _, r) => Some(r)
        case Name(_, _) => None

      /** Get sheet name for cross-sheet, None for local or external-workbook locations */
      def sheetName: Option[SheetName] = loc match
        case CrossSheet(s, _) => Some(s)
        case _ => None

      /** Get cells for local ranges only (for intra-sheet dependency graphs) */
      def localCells: Set[ARef] = loc match
        case Local(r) => r.cells.toSet
        case _ => Set.empty

      /**
       * Get cells for local ranges, bounded by the sheet's used range.
       *
       * Preferred for dependency graph construction to avoid materializing 1M+ cells for full
       * column/row references like A:A or 1:1.
       *
       * @param bounds
       *   Optional bounding range (typically sheet.usedRange)
       * @return
       *   Set of cell references in the intersection of this range and bounds
       */
      def localCellsBounded(bounds: Option[CellRange]): Set[ARef] = loc match
        case Local(r) =>
          bounds match
            case Some(b) => r.intersect(b).map(_.cells.toSet).getOrElse(Set.empty)
            case None => r.cells.toSet
        case _ => Set.empty

      /** Check if this is a cross-sheet reference */
      def isCrossSheet: Boolean = loc match
        case CrossSheet(_, _) => true
        case _ => false

      /**
       * Get A1 string representation. Used in diagnostics (never re-parsed); GH-280: cell-ref-
       * shaped sheet names quote so a sheet literally named "A1" reads unambiguously.
       */
      def toA1: String = loc match
        case Local(r) => r.toA1
        case CrossSheet(s, r) => s"${SheetName.quoteForFormula(s.value)}!${r.toA1}"
        // GH-353: same quoting rule as FormulaPrinter.formatExternalSheet
        case External(i, n, r) =>
          s"${com.tjclp.xl.formula.printer.FormulaPrinter.formatExternalSheet(i, n)}!${r.toA1}"
        // GH-394: a name prints as its identifier (optionally sheet-qualified) — correct
        // diagnostics for the user's source text; the target range is not statically known
        case Name(n, scope) =>
          scope match
            case Some(s) => s"${SheetName.quoteForFormula(s.value)}!$n"
            case None => n
