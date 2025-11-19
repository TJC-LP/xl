package com.tjclp.xl.workbooks

/** Workbook metadata */
case class WorkbookMetadata(
  creator: Option[String] = None,
  created: Option[java.time.LocalDateTime] = None,
  modified: Option[java.time.LocalDateTime] = None,
  lastModifiedBy: Option[String] = None,
  application: Option[String] = Some("XL - Pure Scala 3.7 Excel Library"),
  appVersion: Option[String] = Some("0.1.0-SNAPSHOT")
)
