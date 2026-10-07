package com.tjclp.xl.formula.ast

import com.tjclp.xl.SheetName

/**
 * GH-694: the sheet qualifier Excel keeps in front of an error literal — the `Sheet1!` of
 * `Sheet1!#REF!`, what Excel writes when a formula's or a name's target range is deleted (a
 * `_xlnm._FilterDatabase` left as `Support!#REF!` is the common shape). It changes nothing about
 * the value, which is the error, and carries no dependency edge; it exists so the text prints back
 * as Excel wrote it and follows a sheet rename.
 */
enum ErrorQualifier derives CanEqual:
  /** `Sheet1!#REF!` / `'My Sheet'!#REF!` */
  case Sheet(sheet: SheetName)

  /** `[1]Sheet1!#REF!`: the sheet of an external workbook, raw as in [[TExpr.ExternalRef]] */
  case External(workbookIndex: Int, sheetName: String)
