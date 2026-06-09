package com.tjclp.xl.sheets

/**
 * Sheet view settings, serialized into the worksheet's `<sheetViews><sheetView .../>` element.
 *
 * These control how Excel displays the sheet (not how it prints — see [[PageSetup]] for print
 * settings). Professional templates are typically gridline-free: structure comes from cell borders,
 * and the default gridlines visibly cheapen the artifact in Excel and in HTML/SVG/PNG exports.
 *
 * On write, view settings share a single `<sheetView>` element with freeze panes
 * ([[Sheet.freezePane]]), so setting both never produces duplicate view elements.
 *
 * @param showGridLines
 *   Whether Excel draws the default cell gridlines (default: true, Excel's default)
 * @param zoomScale
 *   Zoom percentage (10-400), or None for Excel's default (100)
 */
final case class SheetView(
  showGridLines: Boolean = true,
  zoomScale: Option[Int] = None
):
  require(
    zoomScale.forall(z => z >= 10 && z <= 400),
    s"Zoom scale must be 10-400, got: ${zoomScale.fold("None")(_.toString)}"
  )

object SheetView:
  val default: SheetView = SheetView()
