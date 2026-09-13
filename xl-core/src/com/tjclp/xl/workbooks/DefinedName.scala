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

object DefinedName:

  /**
   * Excel's identifier relation for defined names: `String.equalsIgnoreCase`, the relation
   * [[DefinedNameIndex]] hashes for resolution. The one rule every mutation path matches by
   * (GH-538), so that a write is the entry the evaluator then resolves. Deliberately narrower than
   * the lint's NFKC/kana fold: mutation is aligned with resolution, not with the linter.
   */
  def sameName(a: String, b: String): Boolean = a.equalsIgnoreCase(b)

  extension (dn: DefinedName)
    /** Whether `dn` is the entry `name` denotes in `scope` (a `localSheetId`; None = workbook). */
    def matches(name: String, scope: Option[Int]): Boolean =
      sameName(dn.name, name) && dn.localSheetId == scope
