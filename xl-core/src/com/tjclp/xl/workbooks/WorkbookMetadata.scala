package com.tjclp.xl.workbooks

import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.styles.color.ThemePalette
import com.tjclp.xl.styles.font.Font

/**
 * Workbook metadata including document properties, theme palette, defined names, and sheet states.
 *
 * @param sheetStates
 *   Sheet visibility overrides: SheetName -> state where state is None (visible), Some("hidden"),
 *   or Some("veryHidden"). Only sheets with non-default visibility are stored.
 * @param date1904
 *   True when the workbook uses the 1904 date system (`<workbookPr date1904="1"/>`, legacy Mac
 *   Excel): date serials count days since 1904-01-01 instead of the default 1900 system (GH-243).
 *   Read from workbookPr and preserved on write; DateTime cells are serialized with the matching
 *   epoch (see `CellValue.dateTimeToExcelSerial(dt, date1904)`).
 * @param calcPr
 *   Modeled `<calcPr>` settings (GH-373, GH-400): the iterate triple plus calcMode / fullCalcOnLoad
 *   / calcId. None reflects a file whose calcPr carries none of the modeled attributes (or has no
 *   calcPr at all); still-unmodeled attributes (refMode, concurrentCalc, ...) ride through on write
 *   either way. Author via `Workbook.withCalcPr` so surgical writes see the change; see [[CalcPr]]
 *   for the per-family write semantics.
 * @param defaultFont
 *   Workbook default (Normal cell style) font (GH-425): drives font slot 0 and the xfId-0
 *   cellStyleXf in styles.xml, so untouched grid cells — and cells authored without an explicit
 *   style — render in it. None means Excel's stock Calibri 11 (`Font.default`). Read back from the
 *   file's Normal font when it is non-default; author via `Workbook.withDefaultFont` so surgical
 *   writes see the change. Styles carrying the `Font.default` sentinel resolve to this font on
 *   write ("unspecified — inherit the workbook default"); explicitly styled fonts are untouched.
 */
final case class WorkbookMetadata(
  creator: Option[String] = None,
  created: Option[java.time.LocalDateTime] = None,
  modified: Option[java.time.LocalDateTime] = None,
  lastModifiedBy: Option[String] = None,
  application: Option[String] = Some("XL - Pure Scala 3.8 Excel Library"),
  appVersion: Option[String] = Some("0.18.0"),
  theme: ThemePalette = ThemePalette.office,
  definedNames: Vector[DefinedName] = Vector.empty,
  sheetStates: Map[SheetName, Option[String]] = Map.empty,
  date1904: Boolean = false,
  calcPr: Option[CalcPr] = None,
  defaultFont: Option[Font] = None
)
