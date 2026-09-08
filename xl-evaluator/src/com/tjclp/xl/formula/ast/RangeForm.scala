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
 *
 * The pairing is a precondition the type does not enforce: `Columns` belongs on a range that spans
 * every row and `Rows` on one that spans every column (the parser only ever builds such pairs, and
 * `parse ∘ print = id` is stated for parser-built nodes). A node built by hand with an inconsistent
 * pairing is TREATED AS THE CORNER RANGE IT ADDRESSES — the printer and both shifters go through
 * [[RangeForm.actualFor]], so `RangeRef(CellRange(A1, B2), Columns)` prints `A1:B2`, never a
 * silently widened `A:B`.
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

  /**
   * The form SPELLED by the text alone — the parser's classification of the syntax it consumed,
   * made with the very predicates `CellRange.parse` uses to decide which corners to synthesize, so
   * the two can never disagree.
   */
  private def ofText(rangeText: String): RangeForm =
    rangeText.split(':') match
      case Array(start, end) =>
        if CellRange.spellsWholeColumn(start) && CellRange.spellsWholeColumn(end) then Columns
        else if CellRange.spellsWholeRow(start) && CellRange.spellsWholeRow(end) then Rows
        else Cells
      case _ => Cells

  extension (form: RangeForm)
    /**
     * The form the printer and the shifters honour for `range`: `Columns` only when the range
     * really spans every row, `Rows` only when it spans every column, `Cells` otherwise. This is
     * what keeps a hand-built inconsistent node truthful — it prints and moves as the corner range
     * it addresses instead of being widened to a whole column or row.
     */
    def actualFor(range: CellRange): RangeForm = form match
      case Columns if !range.isFullColumn => Cells
      case Rows if !range.isFullRow => Cells
      case consistent => consistent
