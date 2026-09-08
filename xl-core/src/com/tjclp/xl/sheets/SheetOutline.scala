package com.tjclp.xl.sheets

import com.tjclp.xl.addressing.{CellRange, Column, Row}

/**
 * Outline collapse and expand for row and column groups (GH-465, second half; GH-589).
 *
 * A collapsed Excel group is two things at once: the member rows/columns are hidden AND the summary
 * row/column immediately after the group carries `collapsed="1"`, which draws the "+" button (the
 * `summaryBelow`/`summaryRight` default). The CLI's `group-rows --collapsed` composes both
 * (`GroupingOps`); these extensions give scripts the same composition so a collapse is never
 * assembled by hand. Members without an outline level become a level-1 group — a hidden span with a
 * dangling marker draws no outline symbol at all — while an existing level (nested groups) is kept.
 * Expanding unhides the members and clears the marker but keeps the level, exactly as Excel's "-"
 * button does. Total: spans are normalised, and a group ending on the last row/column has no
 * summary to mark. Every touched row/column keeps an explicit properties entry so the writer emits
 * the attributes authoritatively (see `GroupingOps.updateRow`).
 */
object outlineSyntax:
  extension (sheet: Sheet)

    /**
     * Collapse rows `first..last` (inclusive, either order): hide them, mark `last + 1` collapsed.
     */
    def collapseRows(first: Row, last: Row): Sheet =
      val (lo, hi) = ordered(first.index0, last.index0)
      val members = (lo to hi).foldLeft(sheet) { (s, r) =>
        updateRow(s, Row.from0(r))(p =>
          p.copy(hidden = true, outlineLevel = p.outlineLevel.orElse(Some(1)))
        )
      }
      if hi < Row.MaxIndex0 then updateRow(members, Row.from0(hi + 1))(_.copy(collapsed = true))
      else members

    /** [[collapseRows]] over the row span of `rows` — `CellRange.parse("5:8")`, or any range. */
    @annotation.targetName("collapseRowsRange")
    def collapseRows(rows: CellRange): Sheet = sheet.collapseRows(rows.rowStart, rows.rowEnd)

    /** Expand rows `first..last`: unhide the members, clear the summary marker, keep the level. */
    def expandRows(first: Row, last: Row): Sheet =
      val (lo, hi) = ordered(first.index0, last.index0)
      val members = (lo to hi).foldLeft(sheet) { (s, r) =>
        updateRow(s, Row.from0(r))(_.copy(hidden = false))
      }
      if hi < Row.MaxIndex0 then updateRow(members, Row.from0(hi + 1))(_.copy(collapsed = false))
      else members

    /** [[expandRows]] over the row span of `rows`. */
    @annotation.targetName("expandRowsRange")
    def expandRows(rows: CellRange): Sheet = sheet.expandRows(rows.rowStart, rows.rowEnd)

    /**
     * Collapse columns `first..last` (inclusive, either order): hide them, mark the next collapsed.
     */
    def collapseCols(first: Column, last: Column): Sheet =
      val (lo, hi) = ordered(first.index0, last.index0)
      val members = (lo to hi).foldLeft(sheet) { (s, c) =>
        updateCol(s, Column.from0(c))(p =>
          p.copy(hidden = true, outlineLevel = p.outlineLevel.orElse(Some(1)))
        )
      }
      if hi < Column.MaxIndex0 then
        updateCol(members, Column.from0(hi + 1))(_.copy(collapsed = true))
      else members

    /** [[collapseCols]] over the column span of `cols` — `CellRange.parse("E:H")`, or any range. */
    @annotation.targetName("collapseColsRange")
    def collapseCols(cols: CellRange): Sheet = sheet.collapseCols(cols.colStart, cols.colEnd)

    /** Expand columns `first..last`: unhide the members, clear the marker, keep the level. */
    def expandCols(first: Column, last: Column): Sheet =
      val (lo, hi) = ordered(first.index0, last.index0)
      val members = (lo to hi).foldLeft(sheet) { (s, c) =>
        updateCol(s, Column.from0(c))(_.copy(hidden = false))
      }
      if hi < Column.MaxIndex0 then
        updateCol(members, Column.from0(hi + 1))(_.copy(collapsed = false))
      else members

    /** [[expandCols]] over the column span of `cols`. */
    @annotation.targetName("expandColsRange")
    def expandCols(cols: CellRange): Sheet = sheet.expandCols(cols.colStart, cols.colEnd)

  private def ordered(a: Int, b: Int): (Int, Int) = if a <= b then (a, b) else (b, a)

  private def updateRow(sheet: Sheet, row: Row)(f: RowProperties => RowProperties): Sheet =
    sheet.setRowProperties(row, f(sheet.getRowProperties(row)))

  private def updateCol(sheet: Sheet, col: Column)(f: ColumnProperties => ColumnProperties): Sheet =
    sheet.setColumnProperties(col, f(sheet.getColumnProperties(col)))
