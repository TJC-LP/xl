package com.tjclp.xl.drawings

import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row}
import com.tjclp.xl.styles.units.{Emu, Px}

/**
 * A point in the drawing coordinate space: a cell plus an EMU offset inside it (DrawingML
 * CT_Marker). Drawing XML `col`/`row` are 0-based, matching [[ARef.from0]] — no ±1 on parse/emit.
 */
final case class AnchorPoint(cell: ARef, dx: Emu = Emu(0), dy: Emu = Emu(0)) derives CanEqual

/** A drawing size in EMUs (DrawingML CT_PositiveSize2D: `cx` width, `cy` height). */
final case class Extent(cx: Emu, cy: Emu) derives CanEqual

object Extent:
  /** Extent from pixel dimensions at 96 DPI (9525 EMU per pixel). */
  def fromPx(w: Int, h: Int): Extent =
    Extent(Px(w.toDouble).toEmu, Px(h.toDouble).toEmu)

/** How a two-cell-anchored drawing responds to row/column resizing (ST_EditAs). */
enum EditAs derives CanEqual:
  case TwoCell, OneCell, Absolute

/**
 * Placement of a drawing on the sheet — the three OOXML anchor forms (GH-221).
 *
 * Named `DrawingAnchor` because `Anchor` is taken by the addressing layer.
 */
enum DrawingAnchor derives CanEqual:
  /** Pinned to two cell corners; moves/sizes with the grid per `editAs`. */
  case TwoCell(from: AnchorPoint, to: AnchorPoint, editAs: EditAs = EditAs.TwoCell)

  /** Pinned to one cell corner with a fixed EMU extent; moves but does not resize. */
  case OneCell(from: AnchorPoint, extent: Extent)

  /** Fixed page position in EMUs, independent of the grid. */
  case Absolute(x: Emu, y: Emu, extent: Extent)

object DrawingAnchor:
  /** One-cell anchor at `ref` with the given extent and zero offsets. */
  def at(ref: ARef, extent: Extent): DrawingAnchor =
    DrawingAnchor.OneCell(AnchorPoint(ref), extent)

  /**
   * Two-cell anchor covering `range`: from = the range's start corner, to = the cell one past the
   * range's end, zero offsets. When the range touches the grid edge (column 16383 / row 1048575)
   * the to-marker clamps to the last valid index — a documented degenerate edge (the drawing covers
   * one cell less than requested on that axis).
   */
  def over(range: CellRange, editAs: EditAs = EditAs.TwoCell): DrawingAnchor =
    val toCol = math.min(range.end.col.index0 + 1, Column.MaxIndex0)
    val toRow = math.min(range.end.row.index0 + 1, Row.MaxIndex0)
    DrawingAnchor.TwoCell(
      AnchorPoint(range.start),
      AnchorPoint(ARef.from0(toCol, toRow)),
      editAs
    )

/**
 * A drawing object on a sheet (GH-221). Document order in `Sheet.drawings` is z-order is emission
 * order: appended drawings paint on top.
 *
 * Typed cases extend this enum under the accepted Patch.MergeBorder enum-extension policy
 * (ChartFrame added by GH-222; Shape still future).
 */
enum Drawing derives CanEqual:
  /**
   * An embedded picture.
   *
   * @param description
   *   alt text (DrawingML `cNvPr/@descr`)
   */
  case Picture(
    anchor: DrawingAnchor,
    image: ImageData,
    name: String = "",
    description: String = ""
  )

  /**
   * A typed chart hosted in a `graphicFrame` (GH-222). The chart part itself is regenerated from
   * the model whenever the enclosing drawings vector changes; an empty `name` emits as the writer
   * default "Chart {ordinal}".
   */
  case ChartFrame(
    anchor: DrawingAnchor,
    chart: com.tjclp.xl.charts.Chart,
    name: String = ""
  )

  /**
   * A drawing the typed model does not (yet) represent — charts, shapes, group shapes, connectors,
   * pictures with crops/effects/hyperlinks.
   *
   * CONTRACT: constructed only by xl-ooxml's drawing reader. The payload is the
   * scope-self-contained canonical XML of one whole anchor element, re-emitted verbatim in document
   * order on write. Users must not construct this case; a hand-built payload that is not canonical
   * anchor XML is silently dropped at emission.
   */
  case Preserved(xml: String)
