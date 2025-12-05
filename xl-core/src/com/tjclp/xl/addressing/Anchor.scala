package com.tjclp.xl.addressing

/**
 * Cell reference anchoring mode for formula dragging.
 *
 * Controls how references shift when a formula is copied to a new location:
 *   - `Relative`: Both column and row adjust (e.g., A1 → B2 when dragged right and down)
 *   - `AbsCol`: Column is fixed, row adjusts (e.g., $A1 → $A2 when dragged down)
 *   - `AbsRow`: Column adjusts, row is fixed (e.g., A$1 → B$1 when dragged right)
 *   - `Absolute`: Both column and row are fixed (e.g., $A$1 stays $A$1)
 */
enum Anchor derives CanEqual:
  case Relative // A1
  case AbsCol // $A1
  case AbsRow // A$1
  case Absolute // $A$1

object Anchor:
  /** Parse anchor from reference string, returning (cleanRef, anchor) */
  def parse(s: String): (String, Anchor) =
    val hasColAnchor = s.startsWith("$")
    val afterCol = if hasColAnchor then s.drop(1) else s
    val dollarInMiddle = afterCol.indexOf('$')

    if hasColAnchor && dollarInMiddle > 0 then
      // $A$1 → Absolute
      val cleanRef = afterCol.take(dollarInMiddle) + afterCol.drop(dollarInMiddle + 1)
      (cleanRef, Anchor.Absolute)
    else if hasColAnchor then
      // $A1 → AbsCol
      (afterCol, Anchor.AbsCol)
    else if dollarInMiddle > 0 then
      // A$1 → AbsRow
      val cleanRef = s.take(dollarInMiddle) + s.drop(dollarInMiddle + 1)
      (cleanRef, Anchor.AbsRow)
    else
      // A1 → Relative
      (s, Anchor.Relative)

  extension (anchor: Anchor)
    /** Is the column absolute (fixed)? */
    def isColAbsolute: Boolean = anchor match
      case AbsCol | Absolute => true
      case _ => false

    /** Is the row absolute (fixed)? */
    def isRowAbsolute: Boolean = anchor match
      case AbsRow | Absolute => true
      case _ => false
