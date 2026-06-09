package com.tjclp.xl.sheets

import com.tjclp.xl.addressing.{CellRange, Row}

/**
 * Print header/footer text for a sheet, serialized to the worksheet's `<headerFooter>` element.
 *
 * Strings support Excel's formatting codes: `&P` (page number), `&N` (total pages), `&D` (date),
 * `&T` (time), `&F` (file name), `&A` (sheet name), plus the `&L`/`&C`/`&R` section markers for
 * left/center/right placement (e.g. `"&LTHE JORDAN COMPANY&RPage &P of &N"`).
 *
 * Only odd-page headers/footers are modeled for now; even/first-page variants present in a source
 * file are preserved verbatim on rewrite.
 *
 * @param oddHeader
 *   Header text for odd pages (all pages unless differentOddEven is set in the source file)
 * @param oddFooter
 *   Footer text for odd pages
 */
final case class HeaderFooter(
  oddHeader: Option[String] = None,
  oddFooter: Option[String] = None
)

/**
 * Page margins in inches, serialized to the worksheet's `<pageMargins>` element.
 *
 * Defaults match Excel's "Normal" margin preset.
 *
 * @param left
 *   Left margin (default 0.7")
 * @param right
 *   Right margin (default 0.7")
 * @param top
 *   Top margin (default 0.75")
 * @param bottom
 *   Bottom margin (default 0.75")
 * @param header
 *   Distance from page top to header (default 0.3")
 * @param footer
 *   Distance from page bottom to footer (default 0.3")
 */
final case class PageMargins(
  left: Double = 0.7,
  right: Double = 0.7,
  top: Double = 0.75,
  bottom: Double = 0.75,
  header: Double = 0.3,
  footer: Double = 0.3
):
  require(
    left >= 0 && right >= 0 && top >= 0 && bottom >= 0 && header >= 0 && footer >= 0,
    s"Margins must be non-negative, got: left=$left right=$right top=$top bottom=$bottom header=$header footer=$footer"
  )

object PageMargins:
  val default: PageMargins = PageMargins()

/**
 * Page setup settings for printing/PDF export.
 *
 * Scale/orientation/fit map to the worksheet's `<pageSetup>` element, margins to `<pageMargins>`,
 * header/footer to `<headerFooter>`. Print area and repeat rows are serialized as sheet-scoped
 * workbook defined names (`_xlnm.Print_Area` / `_xlnm.Print_Titles`).
 *
 * @param scale
 *   Print scale percentage (10-400, default 100)
 * @param orientation
 *   Page orientation: "portrait" or "landscape"
 * @param fitToWidth
 *   Fit to N pages wide (None = no fit)
 * @param fitToHeight
 *   Fit to N pages tall (None = no fit)
 * @param headerFooter
 *   Print header/footer text with Excel codes (None = no header/footer)
 * @param margins
 *   Page margins in inches (None = Excel defaults / preserve existing)
 * @param printArea
 *   Range to print (`_xlnm.Print_Area`); None prints the used range
 * @param repeatRows
 *   1-based inclusive row span repeated at the top of every printed page (`_xlnm.Print_Titles`),
 *   e.g. `Some((1, 2))` repeats rows 1-2
 */
final case class PageSetup(
  scale: Int = 100,
  orientation: Option[String] = None,
  fitToWidth: Option[Int] = None,
  fitToHeight: Option[Int] = None,
  headerFooter: Option[HeaderFooter] = None,
  margins: Option[PageMargins] = None,
  printArea: Option[CellRange] = None,
  repeatRows: Option[(Int, Int)] = None
):
  require(scale >= 10 && scale <= 400, s"Scale must be 10-400, got: $scale")
  require(
    orientation.forall(o => o == "portrait" || o == "landscape"),
    s"Orientation must be 'portrait' or 'landscape', got: ${orientation.getOrElse("None")}"
  )
  require(
    repeatRows.forall((start, end) => start >= 1 && start <= end && end <= Row.MaxIndex0 + 1),
    s"Repeat rows must be a 1-based span within 1-${Row.MaxIndex0 + 1}, got: ${repeatRows.fold("None")(_.toString)}"
  )

object PageSetup:
  val default: PageSetup = PageSetup()
