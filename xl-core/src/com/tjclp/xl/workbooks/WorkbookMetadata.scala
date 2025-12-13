package com.tjclp.xl.workbooks

import com.tjclp.xl.styles.color.ThemePalette

/** Workbook metadata including document properties, theme palette, and defined names */
final case class WorkbookMetadata(
  creator: Option[String] = None,
  created: Option[java.time.LocalDateTime] = None,
  modified: Option[java.time.LocalDateTime] = None,
  lastModifiedBy: Option[String] = None,
  application: Option[String] = Some("XL - Pure Scala 3.7 Excel Library"),
  appVersion: Option[String] = Some("0.3.0"),
  theme: ThemePalette = ThemePalette.office,
  definedNames: Vector[DefinedName] = Vector.empty
)
