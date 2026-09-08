package com.tjclp.xl.ops

import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row}
import com.tjclp.xl.patch.Patch
import com.tjclp.xl.sheets.{ColumnProperties, RowProperties, Sheet}

/**
 * `Patch.toEdits` (ADR-017 §2.12): the kernel's cases desugared into the edits that mean the same
 * thing on `sheet`, or `None` when no edit sequence reproduces the patch EXACTLY — a `ClearStyle`
 * on a cell that does not exist (the patch materializes an empty cell; no edit does), a `Remove`
 * inside a merged region (the edit also unmerges), column/row properties that change more than
 * width/height/visibility. Desugar-coherence law (EditLawsSpec): applying the edits equals
 * `Patch.applyPatch(sheet, p)`; a `Batch` threads the sheet so each leaf is read against the state
 * the kernel would have given it.
 */
private[xl] object PatchLowering:

  private def loc(ref: ARef): Loc = Loc(None, ref)
  private def cell(ref: ARef): Area = Area(None, CellRange(ref, ref))
  private def area(range: CellRange): Area = Area(None, range)

  def toEdits(patch: Patch, sheet: Sheet): Option[Vector[Edit]] = patch match
    case Patch.Put(ref, value) => Some(Vector(Edit.Put(loc(ref), value, None)))

    case Patch.SetStyle(ref, styleId) =>
      // Re-registering the registry's own style yields the same id; a stale numFmtId is the one
      // field an overlay cannot spell, so such a style has no exact desugaring.
      sheet.styleRegistry
        .get(styleId)
        .filter(_.numFmtId.isEmpty)
        .map(style => Vector(Edit.Style(cell(ref), StyleOverlay.of(style), StyleMode.Replace)))

    case Patch.SetCellStyle(ref, style) =>
      Option.when(style.numFmtId.isEmpty)(
        Vector(Edit.Style(cell(ref), StyleOverlay.of(style), StyleMode.Replace))
      )

    case Patch.SetRangeStyle(range, style) =>
      Option.when(style.numFmtId.isEmpty)(
        Vector(Edit.Style(area(range), StyleOverlay.of(style), StyleMode.Replace))
      )

    case Patch.MergeBorder(ref, border) =>
      val overlay = StyleOverlay.ofBorder(border)
      // An all-none border still stamps the cell's (default) style; an empty merge is refused.
      Option.when(!overlay.isEmpty)(Vector(Edit.Style(cell(ref), overlay, StyleMode.Merge)))

    case Patch.ClearStyle(ref) =>
      Option.when(sheet.contains(ref))(Vector(Edit.Clear(cell(ref), ClearWhat.styles)))

    case Patch.Merge(range) => Some(Vector(Edit.Merge(area(range))))
    case Patch.Unmerge(range) => Some(Vector(Edit.Unmerge(area(range))))

    case Patch.SetColumnProperties(col, props) => columnEdits(sheet, col, props)
    case Patch.SetRowProperties(row, props) => rowEdits(sheet, row, props)

    case Patch.Remove(ref) =>
      Option.when(!sheet.mergedRanges.exists(_.contains(ref)))(
        Vector(Edit.Clear(cell(ref), ClearWhat.contents))
      )

    case Patch.RemoveRange(range) =>
      Option.when(!sheet.mergedRanges.exists(_.intersects(range)))(
        Vector(Edit.Clear(area(range), ClearWhat.contents))
      )

    case Patch.PutArray(origin, values) =>
      // One row-major put per non-empty row; a row past the grid has no exact form (the kernel
      // clamps, an edit refuses).
      values.zipWithIndex.foldLeft[Option[Vector[Edit]]](Some(Vector.empty)) {
        case (Some(acc), (row, r)) if row.isEmpty => Some(acc)
        case (Some(acc), (row, r)) =>
          for
            start <- origin.tryShift(0, r)
            end <- start.tryShift(row.size - 1, 0)
          yield acc :+ Edit.PutValues(area(CellRange(start, end)), row, None)
        case (None, _) => None
      }

    case Patch.SetComment(ref, comment) => Some(Vector(Edit.SetComment(loc(ref), comment)))

    case Patch.SetConditionalFormat(ranges, rules) =>
      if ranges.isEmpty || rules.isEmpty then Some(Vector.empty)
      else Some(Vector(Edit.AddConditionalFormat(None, ranges, rules)))

    case Patch.Batch(patches) =>
      patches
        .foldLeft[Option[(Sheet, Vector[Edit])]](Some((sheet, Vector.empty))) {
          case (Some((s, acc)), p) => toEdits(p, s).map(es => (Patch.applyPatch(s, p), acc ++ es))
          case (None, _) => None
        }
        .map(_._2)

  /**
   * Width and visibility are the only column facts an edit sets one at a time; anything else in
   * `props` (a style, an outline level) has no exact edit. `ColWidth` re-stamps an explicit entry,
   * so an unchanged `props` desugars only when the entry already exists.
   */
  private def columnEdits(
    sheet: Sheet,
    col: Column,
    props: ColumnProperties
  ): Option[Vector[Edit]] =
    val current = sheet.getColumnProperties(col)
    val span = ColSpan.single(col)
    val visibility =
      if props.hidden == current.hidden then Vector.empty
      else Vector(if props.hidden then Edit.HideCols(None, span) else Edit.ShowCols(None, span))
    props.width match
      case Some(w) if props == current.copy(width = Some(w), hidden = props.hidden) =>
        Option.when(w >= 0 && w <= 255.0)(Edit.ColWidth(None, span, w) +: visibility)
      case None if props == current.copy(hidden = props.hidden) =>
        Option.when(visibility.nonEmpty || sheet.columnProperties.contains(col))(visibility)
      case _ => None

  private def rowEdits(sheet: Sheet, row: Row, props: RowProperties): Option[Vector[Edit]] =
    val current = sheet.getRowProperties(row)
    val span = RowSpan.single(row)
    val visibility =
      if props.hidden == current.hidden then Vector.empty
      else Vector(if props.hidden then Edit.HideRows(None, span) else Edit.ShowRows(None, span))
    props.height match
      case Some(h) if props == current.copy(height = Some(h), hidden = props.hidden) =>
        Option.when(h >= 0 && h <= 409.0)(Edit.RowHeight(None, span, h) +: visibility)
      case None if props == current.copy(hidden = props.hidden) =>
        Option.when(visibility.nonEmpty || sheet.rowProperties.contains(row))(visibility)
      case _ => None
