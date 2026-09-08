package com.tjclp.xl.ops

import com.tjclp.xl.addressing.ARef

/**
 * GH-628: one cell whose formula gained a `#REF!` because a fill, copy or drag carried a reference
 * off the grid (before column A / row 1, past XFD / 1048576). `references` are the voided
 * references as they were spelled in the source formula. Excel writes the `#REF!` silently; the
 * reporting edits ([[com.tjclp.xl.sheets.SheetEdits.fillReporting]], `copyRangeReporting`) return
 * these so a caller can warn — the CLI's `OFF_GRID_REF`.
 */
final case class OffGridRef(cell: ARef, references: Vector[String]) derives CanEqual
