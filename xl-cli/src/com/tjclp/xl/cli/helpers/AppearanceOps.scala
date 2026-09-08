package com.tjclp.xl.cli.helpers

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.CellRange
import com.tjclp.xl.cli.ColorParser

/**
 * Forwarders onto the sheet-appearance and print-setup merges of `Sheet.mergeSheetView` /
 * `mergePageSetup` / `mergeHeaderFooter` / `withAutoFilter` / `removeAutoFilter` (ADR-017 §2.12,
 * W2.1 — the GH-358 and GH-432 appliers moved into xl-core), shared by the CLI command handlers
 * (WriteCommands) and the batch ops (BatchParser). What stays here is the CLI's flag grammar — "at
 * least one option", "<range> and --clear are mutually exclusive" — and the colour parser; the
 * merge-into-current-settings rule and the value guards are the core's.
 */
object AppearanceOps:

  /**
   * Set or clear the sheet-level autoFilter (GH-432) through the GH-429 lift-and-overlay tri-state.
   * A range authors `<autoFilter ref="...">` (source filterColumn/sortState children ride
   * verbatim); `--clear` sets the active-removal state, stripping even an autoFilter preserved from
   * the source XML (unlike tab-color's passive clear, removal must not resurrect a stale filter).
   */
  def applyAutoFilter(
    sheet: Sheet,
    range: Option[CellRange],
    clear: Boolean
  ): Either[String, Sheet] =
    (range, clear) match
      case (Some(_), true) => Left("autofilter: <range> and --clear are mutually exclusive")
      case (None, false) => Left("autofilter requires a <range> argument or --clear")
      case (Some(r), false) => Right(sheet.withAutoFilter(r))
      case (None, true) => Right(sheet.removeAutoFilter)

  /** Merge view options (gridlines, zoom, tab selection) into the sheet's view settings. */
  def applySheetView(
    sheet: Sheet,
    gridlines: Option[Boolean],
    zoom: Option[Int],
    tabSelected: Option[Boolean]
  ): Either[String, Sheet] =
    if gridlines.isEmpty && zoom.isEmpty && tabSelected.isEmpty then
      Left("sheet-view requires at least one of: gridlines, zoom, tab-selected")
    else sheet.mergeSheetView(gridlines, zoom, tabSelected).left.map(_.message)

  /**
   * Set or clear the sheet tab color. Clearing removes the MODELED color only: on write, a
   * `tabColor` that arrived with the source file's XML is preserved when the model field is None
   * (preserve-if-None semantics) — the CLI cannot strip a preserved color.
   */
  def applyTabColor(
    sheet: Sheet,
    colorStr: Option[String],
    clear: Boolean
  ): Either[String, Sheet] =
    (colorStr, clear) match
      case (Some(_), true) => Left("tab-color: <color> and --clear are mutually exclusive")
      case (None, false) => Left("tab-color requires a <color> argument or --clear")
      case (Some(s), false) => ColorParser.parse(s).map(sheet.withTabColor)
      case (None, true) => Right(sheet.withoutTabColor)

  /** Merge print options (orientation, scale, fit-to) into the sheet's page setup. */
  def applyPageSetup(
    sheet: Sheet,
    orientation: Option[String],
    scale: Option[Int],
    fitToWidth: Option[Int],
    fitToHeight: Option[Int],
    fitToPage: Option[Boolean]
  ): Either[String, Sheet] =
    if orientation.isEmpty && scale.isEmpty && fitToWidth.isEmpty && fitToHeight.isEmpty &&
      fitToPage.isEmpty
    then
      Left(
        "page-setup requires at least one of: orientation, scale, fit-to-width, fit-to-height, fit-to-page"
      )
    else
      sheet
        .mergePageSetup(orientation, scale, fitToWidth, fitToHeight, fitToPage)
        .left
        .map(_.message)

  /**
   * Merge header/footer text into the sheet's page setup. Providing even-page text sets
   * `differentOddEven`, and first-page text sets `differentFirst` (Excel ignores the text while the
   * corresponding flag is off); the explicit flags force them on without text.
   */
  def applyHeaderFooter(
    sheet: Sheet,
    oddHeader: Option[String],
    oddFooter: Option[String],
    evenHeader: Option[String],
    evenFooter: Option[String],
    firstHeader: Option[String],
    firstFooter: Option[String],
    differentOddEven: Boolean,
    differentFirst: Boolean
  ): Either[String, Sheet] =
    val anyText =
      oddHeader.isDefined || oddFooter.isDefined || evenHeader.isDefined || evenFooter.isDefined ||
        firstHeader.isDefined || firstFooter.isDefined
    if !anyText && !differentOddEven && !differentFirst then
      Left(
        "header-footer requires at least one of: odd-header, odd-footer, even-header, " +
          "even-footer, first-header, first-footer, different-odd-even, different-first"
      )
    else
      Right(
        sheet.mergeHeaderFooter(
          oddHeader,
          oddFooter,
          evenHeader,
          evenFooter,
          firstHeader,
          firstFooter,
          differentOddEven,
          differentFirst
        )
      )

  /** Human-readable "key=value, ..." description of the provided (Some) options. */
  def describe(pairs: (String, Option[String])*): String =
    pairs.collect { case (k, Some(v)) => s"$k=$v" }.mkString(", ")
