package com.tjclp.xl.workbooks

import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.styles.color.ThemePalette

/**
 * Workbook metadata including document properties, theme palette, defined names, and sheet states.
 *
 * @param sheetStates
 *   Sheet visibility overrides: SheetName -> state where state is None (visible), Some("hidden"),
 *   or Some("veryHidden"). Only sheets with non-default visibility are stored.
 */
final case class WorkbookMetadata(
  creator: Option[String] = None,
  created: Option[java.time.LocalDateTime] = None,
  modified: Option[java.time.LocalDateTime] = None,
  lastModifiedBy: Option[String] = None,
  application: Option[String] = Some("XL - Pure Scala 3.7 Excel Library"),
  appVersion: Option[String] = Some("0.9.2"),
  theme: ThemePalette = ThemePalette.office,
  definedNames: Vector[DefinedName] = Vector.empty,
  sheetStates: Map[SheetName, Option[String]] = Map.empty
)
