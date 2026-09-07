package com.tjclp.xl.formula.eval

import com.tjclp.xl.addressing.{CellRange, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.drawings.Drawing
import com.tjclp.xl.sheets.{AutoFilterState, FreezePane, Sheet}
import com.tjclp.xl.styles.color.Color
import com.tjclp.xl.workbooks.{CalcPr, DefinedName, Workbook}

/**
 * ADR-017 §2.10: one sheet's orientation card — what an agent needs to know before reading a cell.
 *
 * @param index
 *   1-based position in the workbook
 * @param state
 *   `None` visible, `Some("hidden")`, `Some("veryHidden")` — as `xl sheets` reports it
 * @param dimension
 *   the used range (bounding box of occupied cells), `None` for an empty sheet
 * @param autoFilter
 *   the autoFilter's range when one is set (a pending removal is `None`)
 * @param tables
 *   structured-table names, sorted
 * @param conditionalFormats
 *   conditional-format blocks (each may carry several rules)
 */
final case class SheetSummary(
  name: SheetName,
  index: Int,
  state: Option[String],
  dimension: Option[CellRange],
  cellCount: Int,
  formulaCount: Int,
  uncachedFormulas: Int,
  mergedRanges: Int,
  comments: Int,
  hyperlinks: Int,
  freeze: Option[FreezePane],
  tabColor: Option[Color],
  autoFilter: Option[CellRange],
  tables: Vector[String],
  charts: Int,
  pictures: Int,
  conditionalFormats: Int,
  dataValidations: Int,
  hiddenRows: Int,
  hiddenCols: Int
) derives CanEqual

/**
 * The whole workbook at a glance: one [[SheetSummary]] per sheet in workbook order, plus the
 * workbook facts that change how every number reads — defined names (hidden ones included), the
 * date system and the calculation properties. `WorkbookSummary.of` is total: any workbook the
 * reader accepts summarizes. This is what `xl describe --full` prints and `wb.describe` returns.
 */
final case class WorkbookSummary(
  sheets: Vector[SheetSummary],
  definedNames: Vector[DefinedName],
  date1904: Boolean,
  calcPr: Option[CalcPr]
) derives CanEqual

object WorkbookSummary:

  def of(wb: Workbook): WorkbookSummary =
    WorkbookSummary(
      sheets = wb.sheets.zipWithIndex.map((sheet, i) => summarize(wb, sheet, i + 1)),
      definedNames = wb.metadata.definedNames,
      date1904 = wb.metadata.date1904,
      calcPr = wb.metadata.calcPr
    )

  private def summarize(wb: Workbook, sheet: Sheet, index: Int): SheetSummary =
    val (formulas, uncached) = sheet.cells.valuesIterator.foldLeft((0, 0)) { case ((f, u), cell) =>
      cell.value match
        case CellValue.Formula(_, None, _) => (f + 1, u + 1)
        case CellValue.Formula(_, Some(_), _) => (f + 1, u)
        case _ => (f, u)
    }
    SheetSummary(
      name = sheet.name,
      index = index,
      state = wb.getSheetState(sheet.name),
      dimension = sheet.usedRange,
      cellCount = sheet.cells.size,
      formulaCount = formulas,
      uncachedFormulas = uncached,
      mergedRanges = sheet.mergedRanges.size,
      comments = sheet.comments.size,
      hyperlinks = sheet.cells.valuesIterator.count(_.hyperlink.isDefined),
      freeze = sheet.freezePane,
      tabColor = sheet.tabColor,
      autoFilter = sheet.autoFilter.collect { case AutoFilterState.Ranged(range) => range },
      tables = sheet.tables.keys.toVector.sorted,
      charts = sheet.drawings.count {
        case _: Drawing.ChartFrame => true
        case _ => false
      },
      pictures = sheet.drawings.count {
        case _: Drawing.Picture => true
        case _ => false
      },
      conditionalFormats = sheet.conditionalFormats.size,
      dataValidations = sheet.dataValidations.size,
      hiddenRows = sheet.rowProperties.valuesIterator.count(_.hidden),
      hiddenCols = sheet.columnProperties.valuesIterator.count(_.hidden)
    )

/**
 * `wb.describe` / `wb.audit` for scripts (exported as `WorkbookInspect.*` through the prelude). No
 * default arguments: extension methods reached through a wildcard export cannot carry them.
 */
object WorkbookInspect:
  extension (wb: Workbook)
    def describe: WorkbookSummary = WorkbookSummary.of(wb)
    def audit: WorkbookAudit = WorkbookAudit.of(wb)
