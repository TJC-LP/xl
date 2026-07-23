package com.tjclp.xl.sheets

import com.tjclp.xl.addressing.CellRange

/**
 * Sheet-level `<autoFilter>` overlay state (GH-429) — the lift-and-overlay tri-state.
 *
 * The full autoFilter element (filterColumn criteria, sortState) rides the GH-232 preservedKnown
 * passthrough untyped; only its `ref` range is lifted here so structural edits can move it. On
 * write the model state is overlaid onto the preserved element: `Ranged` replaces only the `ref`
 * attribute (children verbatim; a fresh minimal element when no source exists), `Remove` deletes
 * the element, and `Sheet.autoFilter = None` is the passive default (source XML rides through
 * verbatim — the FreezePane/tabColor precedent).
 *
 * The tri-state is required: a plain `Option[CellRange]` could not distinguish "never modeled" from
 * "collapsed by a structural delete", and would resurrect a stale filter.
 *
 * Documented limitations: interior column edits misalign `filterColumn@colId` (colId is
 * range-relative, so row edits and left-of-range column edits are safe — the dominant field cases);
 * `sortState` and `_xlnm._FilterDatabase` ride stale (Excel-tolerated).
 */
enum AutoFilterState derives CanEqual:
  /** The filter covers `ref`; structural edits shift it like every other sqref-shaped range. */
  case Ranged(ref: CellRange)

  /** Actively strip the sheet's autoFilter on write (the range collapsed into a deletion). */
  case Remove
