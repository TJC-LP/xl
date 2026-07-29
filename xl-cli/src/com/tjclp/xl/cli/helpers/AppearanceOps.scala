package com.tjclp.xl.cli.helpers

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.CellRange
import com.tjclp.xl.cli.ColorParser
import com.tjclp.xl.sheets.{AutoFilterState, HeaderFooter, PageSetup, SheetView}

/**
 * Pure sheet-appearance and print-setup appliers (GH-358) plus the sheet-level autoFilter applier
 * (GH-432), shared by the CLI command handlers (WriteCommands) and the batch ops (BatchParser) so
 * the two paths cannot drift.
 *
 * Every applier merges into the sheet's CURRENT settings: unspecified fields are preserved
 * (read-modify-write on the whole case class — the model has no granular withers). Validation
 * happens here, returning Left with a clean message, so the model case classes' `require` guards
 * are never tripped from the CLI.
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
      case (Some(r), false) => Right(sheet.copy(autoFilter = Some(AutoFilterState.Ranged(r))))
      case (None, true) => Right(sheet.copy(autoFilter = Some(AutoFilterState.Remove)))

  /** Merge view options (gridlines, zoom, tab selection) into the sheet's view settings. */
  def applySheetView(
    sheet: Sheet,
    gridlines: Option[Boolean],
    zoom: Option[Int],
    tabSelected: Option[Boolean]
  ): Either[String, Sheet] =
    if gridlines.isEmpty && zoom.isEmpty && tabSelected.isEmpty then
      Left("sheet-view requires at least one of: gridlines, zoom, tab-selected")
    else
      zoom.filter(z => z < 10 || z > 400) match
        case Some(z) => Left(s"Zoom scale must be 10-400, got: $z")
        case None =>
          val current = sheet.viewSettings.getOrElse(SheetView.default)
          Right(
            sheet.withViewSettings(
              current.copy(
                showGridLines = gridlines.getOrElse(current.showGridLines),
                zoomScale = zoom.orElse(current.zoomScale),
                tabSelected = tabSelected.orElse(current.tabSelected)
              )
            )
          )

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
      for
        _ <- orientation
          .filterNot(o => o == "portrait" || o == "landscape")
          .map(o => s"Orientation must be 'portrait' or 'landscape', got: $o")
          .toLeft(())
        _ <- scale
          .filter(sc => sc < 10 || sc > 400)
          .map(sc => s"Scale must be 10-400, got: $sc")
          .toLeft(())
        _ <- fitToWidth
          .filter(_ < 1)
          .map(n => s"fit-to-width must be >= 1, got: $n")
          .toLeft(())
        _ <- fitToHeight
          .filter(_ < 1)
          .map(n => s"fit-to-height must be >= 1, got: $n")
          .toLeft(())
      yield
        val current = sheet.pageSetup.getOrElse(PageSetup.default)
        sheet.withPageSetup(
          current.copy(
            scale = scale.getOrElse(current.scale),
            orientation = orientation.orElse(current.orientation),
            fitToWidth = fitToWidth.orElse(current.fitToWidth),
            fitToHeight = fitToHeight.orElse(current.fitToHeight),
            // Tri-state (GH-284): None derives from fitToWidth/fitToHeight and preserves
            // any flag in the source file; Some(false) actively strips a preserved flag.
            fitToPage = fitToPage.orElse(current.fitToPage)
          )
        )

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
      val setup = sheet.pageSetup.getOrElse(PageSetup.default)
      val hf = setup.headerFooter.getOrElse(HeaderFooter())
      val updated = hf.copy(
        oddHeader = oddHeader.orElse(hf.oddHeader),
        oddFooter = oddFooter.orElse(hf.oddFooter),
        evenHeader = evenHeader.orElse(hf.evenHeader),
        evenFooter = evenFooter.orElse(hf.evenFooter),
        firstHeader = firstHeader.orElse(hf.firstHeader),
        firstFooter = firstFooter.orElse(hf.firstFooter),
        differentOddEven =
          hf.differentOddEven || differentOddEven || evenHeader.isDefined || evenFooter.isDefined,
        differentFirst =
          hf.differentFirst || differentFirst || firstHeader.isDefined || firstFooter.isDefined
      )
      Right(sheet.withPageSetup(setup.copy(headerFooter = Some(updated))))

  /** Human-readable "key=value, ..." description of the provided (Some) options. */
  def describe(pairs: (String, Option[String])*): String =
    pairs.collect { case (k, Some(v)) => s"$k=$v" }.mkString(", ")
