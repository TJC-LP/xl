package com.tjclp.xl.workbooks

/**
 * Represents an Excel named range (defined name).
 *
 * In Excel, defined names allow users to assign meaningful names to cell references, ranges, or
 * formulas. For example, "SalesTotal" might refer to "Sheet1!$A$1:$A$100".
 *
 * @param name
 *   The name identifier (e.g., "SalesTotal", "TaxRate")
 * @param formula
 *   The reference or formula the name points to (e.g., "Sheet1!$A$1:$A$10", "0.08")
 * @param localSheetId
 *   Optional sheet scope index. If None, the name is workbook-scoped (global). If Some(idx), the
 *   name is scoped to that sheet and only visible within it.
 * @param hidden
 *   Whether the name is hidden from the Name Manager UI
 * @param comment
 *   Optional comment describing the named range
 */
final case class DefinedName(
  name: String,
  formula: String,
  localSheetId: Option[Int] = None,
  hidden: Boolean = false,
  comment: Option[String] = None
)
