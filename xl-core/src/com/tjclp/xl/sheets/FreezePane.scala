package com.tjclp.xl.sheets

import com.tjclp.xl.addressing.ARef

/**
 * Freeze pane configuration for a sheet.
 *
 * Freeze panes lock rows and/or columns so they remain visible while scrolling. The reference
 * specifies the top-left cell of the scrollable (unfrozen) area:
 *   - `At(ref"B2")` freezes row 1 and column A
 *   - `At(ref"A3")` freezes rows 1-2 (no column freeze)
 *   - `At(ref"C1")` freezes columns A-B (no row freeze)
 */
enum FreezePane derives CanEqual:
  /** Freeze rows above and columns to the left of the given cell. */
  case At(topLeftCell: ARef)

  /** Remove any existing freeze panes. */
  case Remove
