package com.tjclp.xl.formula.functions

import com.tjclp.xl.formula.ast.{RangeForm, TExpr}

import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row}

/**
 * GH-631: Excel's rule for `SUMIF`'s `sum_range` and `AVERAGEIF`'s `average_range` — the third
 * argument "does not have to be the same size and shape as `range`": the cells actually summed
 * start at its upper-left cell and take `range`'s size and shape. `SUMIF(A1:A10, ">0", C1)` reads
 * C1:C10; `SUMIF(A1:A3, ">1", C1:D5)` reads C1:C3. The rule is applied in two places that must
 * agree — evaluation ([[FunctionSpecsAggregate]]) and dependency extraction
 * ([[com.tjclp.xl.formula.graph.DependencyGraph]]), so an edit to C5 dirties the formula that reads
 * it — and this object is the one statement of it. A pairing that would run past the grid edge
 * (`SUMIF(A1:A3, ">1", Z1048575)`) contributes nothing there, as LibreOffice computes it.
 *
 * The `*IFS` family is different: Excel requires every `criteria_range` to match `sum_range`'s
 * shape and returns `#VALUE!` otherwise; nothing is resized.
 */
object CriteriaRangeResize:

  /** The functions whose third argument is sized to their first. */
  val resizingFunctions: Set[String] = Set("SUMIF", "AVERAGEIF")

  /**
   * `target` resized to `shape`'s height and width from its upper-left cell, clipped to the grid.
   */
  def resize(target: CellRange, shape: CellRange): CellRange =
    val col0 = target.colStart.index0
    val row0 = target.rowStart.index0
    val colEnd = math.min(col0.toLong + shape.width - 1, Column.MaxIndex0.toLong).toInt
    val rowEnd = math.min(row0.toLong + shape.height - 1, Row.MaxIndex0.toLong).toInt
    CellRange(ARef.from0(col0, row0), ARef.from0(colEnd, rowEnd))

  /**
   * The cell of `target` paired with `testRef` of `shape`: the same offset from `target`'s
   * upper-left as `testRef` has from `shape`'s. None past the grid edge (that pairing reads
   * nothing).
   */
  def pairedCell(testRef: ARef, shape: CellRange, target: CellRange): Option[ARef] =
    val col = target.colStart.index0.toLong + (testRef.col.index0 - shape.colStart.index0)
    val row = target.rowStart.index0.toLong + (testRef.row.index0 - shape.rowStart.index0)
    Option.when(col <= Column.MaxIndex0 && row <= Row.MaxIndex0)(
      ARef.from0(col.toInt, row.toInt)
    )

  /**
   * A range-slot location resized to `shape` when it carries a range (local, cross-sheet or
   * external); a defined name or an error is returned as it is.
   */
  def resizeLocation(location: TExpr.RangeLocation, shape: CellRange): TExpr.RangeLocation =
    location match
      case TExpr.RangeLocation.Local(range, _) =>
        TExpr.RangeLocation.Local(resize(range, shape), RangeForm.Cells)
      case TExpr.RangeLocation.CrossSheet(sheet, range, _) =>
        TExpr.RangeLocation.CrossSheet(sheet, resize(range, shape), RangeForm.Cells)
      case TExpr.RangeLocation.External(index, name, range, _) =>
        TExpr.RangeLocation.External(index, name, resize(range, shape), RangeForm.Cells)
      case other @ (TExpr.RangeLocation.Name(_, _) | TExpr.RangeLocation.Error(_)) => other

  /**
   * The argument values of a call to `name` with the resize applied: for a resizing function whose
   * first argument's shape `shapeOf` can tell, the third argument's range is resized to it. Every
   * other call — and a call whose first argument's shape is unknown (a defined name without a
   * workbook to resolve it) — is returned as given. The dependency extractors read a call's
   * arguments through this so the graph records the cells evaluation actually reads.
   */
  def resizedArgs(
    name: String,
    values: List[ArgValue],
    shapeOf: TExpr.RangeLocation => Option[CellRange] = _.staticRange
  ): List[ArgValue] =
    if !resizingFunctions.contains(name.toUpperCase) then values
    else
      values match
        case (first @ ArgValue.Range(criteria)) :: criterion :: ArgValue.Range(target) :: rest =>
          shapeOf(criteria) match
            case Some(shape) =>
              first :: criterion :: ArgValue.Range(resizeLocation(target, shape)) :: rest
            case None => values
        case _ => values
