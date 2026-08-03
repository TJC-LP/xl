package com.tjclp.xl.sheets

import com.tjclp.xl.addressing.ARef

/**
 * Sheet view settings, serialized into the worksheet's `<sheetViews><sheetView .../>` element.
 *
 * These control how Excel displays the sheet (not how it prints — see [[PageSetup]] for print
 * settings). Professional templates are typically gridline-free: structure comes from cell borders,
 * and the default gridlines visibly cheapen the artifact in Excel and in HTML/SVG/PNG exports.
 *
 * On write, view settings share a single `<sheetView>` element with freeze panes
 * ([[Sheet.freezePane]]), so setting both never produces duplicate view elements. The reader
 * populates a SheetView whenever a modeled attribute is present on the first `<sheetView>`
 * (GH-258/GH-372), so read→modify→write keeps these settings through model regeneration.
 *
 * @param showGridLines
 *   Whether Excel draws the default cell gridlines (default: true, Excel's default)
 * @param zoomScale
 *   Zoom percentage (10-400) for the CURRENT view mode, or None for Excel's default (100)
 * @param tabSelected
 *   Whether this sheet's tab is selected in the UI (GH-372). Three-valued on write: `Some(_)`
 *   overlays the attribute, `None` leaves any preserved `tabSelected` attribute untouched. Note
 *   this is tab SELECTION (grouping), not the active sheet — that is `Workbook.activeSheetIndex`.
 * @param view
 *   View mode (GH-446): `"normal"`, `"pageBreakPreview"`, or `"pageLayout"`. Finance workbooks
 *   conventionally ship in page-break preview so the print extent is visible while editing.
 * @param zoomScaleNormal
 *   Zoom percentage (10-400) remembered for normal view, independent of [[zoomScale]] (which
 *   applies to whatever [[view]] is current)
 * @param zoomScaleSheetLayoutView
 *   Zoom percentage (10-400) remembered for page-break preview
 * @param topLeftCell
 *   Top-left visible cell of the view — the sheet's scroll position (GH-446). Distinct from the
 *   freeze pane's scroll target ([[Sheet.freezeAt]]'s `scrolledTo`), which lives on `<pane>`.
 */
final case class SheetView(
  showGridLines: Boolean = true,
  zoomScale: Option[Int] = None,
  tabSelected: Option[Boolean] = None,
  view: Option[String] = None,
  zoomScaleNormal: Option[Int] = None,
  zoomScaleSheetLayoutView: Option[Int] = None,
  topLeftCell: Option[ARef] = None
):
  require(
    zoomScale.forall(z => z >= 10 && z <= 400),
    s"Zoom scale must be 10-400, got: ${zoomScale.fold("None")(_.toString)}"
  )
  require(
    zoomScaleNormal.forall(z => z >= 10 && z <= 400),
    s"zoomScaleNormal must be 10-400, got: ${zoomScaleNormal.fold("None")(_.toString)}"
  )
  require(
    zoomScaleSheetLayoutView.forall(z => z >= 10 && z <= 400),
    s"zoomScaleSheetLayoutView must be 10-400, got: ${zoomScaleSheetLayoutView.fold("None")(_.toString)}"
  )
  require(
    view.forall(SheetView.validViews.contains),
    s"view must be one of ${SheetView.validViews.mkString(", ")}, got: ${view.getOrElse("None")}"
  )

object SheetView:
  val default: SheetView = SheetView()

  /** Legal values of the `sheetView/@view` attribute (ECMA-376 ST_SheetViewType). */
  val validViews: Set[String] = Set("normal", "pageBreakPreview", "pageLayout")
