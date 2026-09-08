package com.tjclp.xl.formula.ast

import com.tjclp.xl.CellRange

/**
 * GH-612: the SURFACE FORM of a range reference in formula text.
 *
 * `A:A`, `$A:C` and `1:1` address the same cells as `A1:A1048576`, `$A1:C1048576` and `A1:XFD1`,
 * but Excel treats the two spellings differently: a whole-column reference drags and restructures
 * only along columns (its rows are "all of them", not coordinates), a whole-row reference only
 * along rows, and both print back as they were written. `CellRange` stays purely positional (its
 * corners are the cells addressed); the form rides on the formula AST's range nodes beside the
 * range, defaulting to [[RangeForm.Cells]] so a range built from two corners is a corner range.
 */
enum RangeForm derives CanEqual:
  /** `A1:B2` — two corner cells; both axes are coordinates. */
  case Cells

  /** `A:A`, `$A:C` — whole columns; only the column axis is a coordinate. */
  case Columns

  /** `1:1`, `$3:$10` — whole rows; only the row axis is a coordinate. */
  case Rows

object RangeForm:
  /**
   * The form Excel displays for range TEXT `start:end` (anchors allowed, no sheet qualifier) that
   * parsed to `range`. Letters-only parts are [[Columns]] and digits-only parts [[Rows]]; a CORNER
   * spelling that addresses every row (`A1:A1048576`) is [[Columns]] and one that addresses every
   * column (`A1:XFD1`) is [[Rows]], because Excel canonicalises such an entry to `A:A` / `1:1`
   * before it reaches the file — the explicit corner form only arrives from other producers
   * (openpyxl, a `putf`) and is treated as Excel would have treated it. A corner range spanning
   * every row AND every column is `1:1048576`. Anything else is [[Cells]].
   */
  def of(rangeText: String, range: CellRange): RangeForm =
    ofText(rangeText) match
      case Cells if range.isFullRow => Rows
      case Cells if range.isFullColumn => Columns
      case form => form

  /** The form SPELLED by the text alone: the parser's classification of the syntax it consumed. */
  private def ofText(rangeText: String): RangeForm =
    rangeText.split(':') match
      case Array(start, end) =>
        val s = start.stripPrefix("$")
        val e = end.stripPrefix("$")
        if s.nonEmpty && e.nonEmpty && s.forall(_.isLetter) && e.forall(_.isLetter) then Columns
        else if s.nonEmpty && e.nonEmpty && s.forall(_.isDigit) && e.forall(_.isDigit) then Rows
        else Cells
      case _ => Cells
