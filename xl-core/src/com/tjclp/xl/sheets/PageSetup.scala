package com.tjclp.xl.sheets

/**
 * Page setup settings for printing/PDF export.
 *
 * @param scale
 *   Print scale percentage (10-400, default 100)
 * @param orientation
 *   Page orientation: "portrait" or "landscape"
 * @param fitToWidth
 *   Fit to N pages wide (None = no fit)
 * @param fitToHeight
 *   Fit to N pages tall (None = no fit)
 */
final case class PageSetup(
  scale: Int = 100,
  orientation: Option[String] = None,
  fitToWidth: Option[Int] = None,
  fitToHeight: Option[Int] = None
):
  require(scale >= 10 && scale <= 400, s"Scale must be 10-400, got: $scale")
  require(
    orientation.forall(o => o == "portrait" || o == "landscape"),
    s"Orientation must be 'portrait' or 'landscape', got: ${orientation.getOrElse("None")}"
  )

object PageSetup:
  val default: PageSetup = PageSetup()
