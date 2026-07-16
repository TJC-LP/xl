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
  /**
   * Freeze rows above and columns to the left of the given cell.
   *
   * NAMING WART: `topLeftCell` is the freeze ANCHOR — it maps to the pane's `xSplit`/`ySplit`
   * attributes, NOT to the OOXML `topLeftCell` attribute. The OOXML `<pane topLeftCell=".."/>`
   * attribute is the SCROLL target of the bottom-right pane, modeled here as [[scrolledTo]]
   * (GH-382): a sheet frozen at row 10 but scrolled down to show row 40 carries
   * `At(ref"A11", Some(ref"A40"))`. `None` means unscrolled — the pane attribute equals the anchor.
   */
  case At(topLeftCell: ARef, scrolledTo: Option[ARef] = None)

  /** Remove any existing freeze panes. */
  case Remove
