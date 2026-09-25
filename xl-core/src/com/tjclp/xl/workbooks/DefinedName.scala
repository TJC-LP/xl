package com.tjclp.xl.workbooks

import com.tjclp.xl.addressing.{CellRange, SheetName}
import com.tjclp.xl.sheets.Sheet

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
   * Excel's name for a sheet's print area: sheet-scoped, and after a read its typed home is
   * `PageSetup.printArea` (the reader lifts the modelable form out of the table, GH-259).
   */
  val PrintArea: String = "_xlnm.Print_Area"

  /**
   * Excel's name for a sheet's repeated print titles: sheet-scoped; a pure row span is lifted into
   * `PageSetup.repeatRows` on read, column titles stay in the table verbatim.
   */
  val PrintTitles: String = "_xlnm.Print_Titles"

  /**
   * Excel's identifier relation for defined names: `String.equalsIgnoreCase`, the relation
   * [[DefinedNameIndex]] hashes for resolution. The one rule every mutation path matches by
   * (GH-538), so that a write is the entry the evaluator then resolves. Deliberately narrower than
   * the lint's NFKC/kana fold: mutation is aligned with resolution, not with the linter.
   */
  def sameName(a: String, b: String): Boolean = a.equalsIgnoreCase(b)

  /** A Print_Area formula, e.g. `Sheet1!$A$1:$D$20` or `'Q1 Report'!$A$1:$D$20` (GH-259). */
  private[xl] def printAreaFormula(sheet: SheetName, area: CellRange): String =
    val s = area.start
    val e = area.end
    s"${SheetName.quoteForFormula(sheet.value)}!$$${s.col.toLetter}$$${s.row.index1}:$$${e.col.toLetter}$$${e.row.index1}"

  /** A row-span Print_Titles formula, e.g. `Sheet1!$1:$3` (GH-259). */
  private[xl] def printTitlesFormula(sheet: SheetName, rows: (Int, Int)): String =
    s"${SheetName.quoteForFormula(sheet.value)}!$$${rows._1}:$$${rows._2}"

  /**
   * The sheet-scoped print names each sheet's PageSetup denotes (GH-259), in sheet order and, per
   * sheet, Print_Area before Print_Titles; `localSheetId` is the sheet's position.
   */
  private[xl] def fromPageSetups(sheets: Vector[Sheet]): Vector[DefinedName] =
    sheets.zipWithIndex.flatMap { (sheet, idx) =>
      val area = sheet.pageSetup.flatMap(_.printArea).map { range =>
        DefinedName(PrintArea, printAreaFormula(sheet.name, range), localSheetId = Some(idx))
      }
      val titles = sheet.pageSetup.flatMap(_.repeatRows).map { rows =>
        DefinedName(PrintTitles, printTitlesFormula(sheet.name, rows), localSheetId = Some(idx))
      }
      area.toList ++ titles.toList
    }

  extension (dn: DefinedName)
    /** Whether `dn` is the entry `name` denotes in `scope` (a `localSheetId`; None = workbook). */
    def matches(name: String, scope: Option[Int]): Boolean =
      sameName(dn.name, name) && dn.localSheetId == scope
