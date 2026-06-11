package com.tjclp.xl.workbooks

import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.styles.color.ThemePalette

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
 */
final case class WorkbookMetadata(
  creator: Option[String] = None,
  created: Option[java.time.LocalDateTime] = None,
  modified: Option[java.time.LocalDateTime] = None,
  lastModifiedBy: Option[String] = None,
  application: Option[String] = Some("XL - Pure Scala 3.8 Excel Library"),
  appVersion: Option[String] = Some("0.12.2"),
  theme: ThemePalette = ThemePalette.office,
  definedNames: Vector[DefinedName] = Vector.empty,
  sheetStates: Map[SheetName, Option[String]] = Map.empty,
  date1904: Boolean = false
)
