package com.tjclp.xl.sheets

import scala.annotation.targetName

import com.tjclp.xl.addressing.{CellRange, Column, Row}
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.ops.{ColSpan, RowSpan}

/**
 * Outline collapse and expand for row and column groups (GH-465, second half; GH-589).
 *
 * A collapsed Excel group is two things at once: the member rows/columns are hidden AND the summary
 * row/column immediately after the group carries `collapsed="1"`, which draws the "+" button (the
 * `summaryBelow`/`summaryRight` default). The CLI's `group-rows --collapsed` composes both through
 * `SheetEdits.groupRows`; these extensions are the same fold
 * (`SheetEdits.outlineRows`/`outlineCols`) so a collapse is never assembled by hand and the two
 * cannot drift: on an ungrouped sheet `collapseRows(span)` is
 * `groupRows(span, 1, collapsed = true)`. Members without an outline level become a level-1 group —
 * a hidden span with a dangling marker draws no outline symbol at all — while a member's existing
 * level is kept, so collapsing an outer group leaves its nested groups' levels intact. Expanding
 * unhides the members and clears the marker but keeps the level, as Excel's "+" button does for a
 * flat group (members of a still-collapsed inner group are unhidden too, where Excel would keep
 * them hidden). Total: spans are normalised, and a group ending on the last row/column has no
 * summary to mark.
 *
 * Spans are `RowSpan`/`ColSpan` (or a `(first, last)` pair): the axis is in the type, so
 * `ColSpan.parse("E:H")` can never reach `collapseRows`. The `CellRange` overloads accept only a
 * full-row range (`5:8`) or a full-column range (`E:H`) for their axis and refuse anything else
 * with `InvalidReference` — a column range projected onto the row axis would hide every row of the
 * sheet. Every row/column a collapse touches keeps an explicit properties entry so the writer emits
 * the attributes authoritatively (`SheetEdits.updateRow`); an expand updates only rows/columns that
 * have an entry, so expanding what was never hidden leaves the file byte-identical
 * (`SheetEdits.touchRow`).
 */
object outlineSyntax:
  extension (sheet: Sheet)

    /**
     * Collapse rows `first..last` (inclusive, either order): hide them, mark `last + 1` collapsed.
     */
    def collapseRows(first: Row, last: Row): Sheet = sheet.collapseRows(RowSpan(first, last))

    /** [[collapseRows]] over a row span — `RowSpan.parse("5:8")`. */
    @targetName("collapseRowsSpan")
    def collapseRows(rows: RowSpan): Sheet =
      SheetEdits.outlineRows(sheet, rows, Some(true), sparse = false) { p =>
        p.copy(hidden = true, outlineLevel = p.outlineLevel.orElse(Some(1)))
      }

    /**
     * [[collapseRows]] over a full-row `CellRange` (`CellRange.parse("5:8")`); a cell range or a
     * column range is refused rather than projected onto the row axis.
     */
    @targetName("collapseRowsRange")
    def collapseRows(rows: CellRange): XLResult[Sheet] =
      rowSpanOf(rows, "collapseRows").map(span => sheet.collapseRows(span))

    /** Expand rows `first..last`: unhide the members, clear the summary marker, keep the level. */
    def expandRows(first: Row, last: Row): Sheet = sheet.expandRows(RowSpan(first, last))

    /** [[expandRows]] over a row span. Rows without properties are left untouched. */
    @targetName("expandRowsSpan")
    def expandRows(rows: RowSpan): Sheet =
      SheetEdits.outlineRows(sheet, rows, Some(false), sparse = true)(_.copy(hidden = false))

    /** [[expandRows]] over a full-row `CellRange`; anything else is refused. */
    @targetName("expandRowsRange")
    def expandRows(rows: CellRange): XLResult[Sheet] =
      rowSpanOf(rows, "expandRows").map(span => sheet.expandRows(span))

    /**
     * Collapse columns `first..last` (inclusive, either order): hide them, mark the next collapsed.
     */
    def collapseCols(first: Column, last: Column): Sheet =
      sheet.collapseCols(ColSpan(first, last))

    /** [[collapseCols]] over a column span — `ColSpan.parse("E:H")`. */
    @targetName("collapseColsSpan")
    def collapseCols(cols: ColSpan): Sheet =
      SheetEdits.outlineCols(sheet, cols, Some(true), sparse = false) { p =>
        p.copy(hidden = true, outlineLevel = p.outlineLevel.orElse(Some(1)))
      }

    /**
     * [[collapseCols]] over a full-column `CellRange` (`CellRange.parse("E:H")`); a cell range or a
     * row range is refused rather than projected onto the column axis.
     */
    @targetName("collapseColsRange")
    def collapseCols(cols: CellRange): XLResult[Sheet] =
      colSpanOf(cols, "collapseCols").map(span => sheet.collapseCols(span))

    /** Expand columns `first..last`: unhide the members, clear the marker, keep the level. */
    def expandCols(first: Column, last: Column): Sheet = sheet.expandCols(ColSpan(first, last))

    /** [[expandCols]] over a column span. Columns without properties are left untouched. */
    @targetName("expandColsSpan")
    def expandCols(cols: ColSpan): Sheet =
      SheetEdits.outlineCols(sheet, cols, Some(false), sparse = true)(_.copy(hidden = false))

    /** [[expandCols]] over a full-column `CellRange`; anything else is refused. */
    @targetName("expandColsRange")
    def expandCols(cols: CellRange): XLResult[Sheet] =
      colSpanOf(cols, "expandCols").map(span => sheet.expandCols(span))

  /** The row span a full-row range denotes; a narrower range names the wrong axis. */
  private def rowSpanOf(range: CellRange, op: String): XLResult[RowSpan] =
    if range.isFullRow then Right(RowSpan(range.rowStart, range.rowEnd))
    else
      Left(
        XLError.InvalidReference(
          s"'${range.toA1}' is not a row span: $op takes full rows like 5:8 " +
            "(RowSpan.parse, or the (Row, Row) form); a column span or cell range is refused"
        )
      )

  /** The column span a full-column range denotes; a narrower range names the wrong axis. */
  private def colSpanOf(range: CellRange, op: String): XLResult[ColSpan] =
    if range.isFullColumn then Right(ColSpan(range.colStart, range.colEnd))
    else
      Left(
        XLError.InvalidReference(
          s"'${range.toA1}' is not a column span: $op takes full columns like E:H " +
            "(ColSpan.parse, or the (Column, Column) form); a row span or cell range is refused"
        )
      )
