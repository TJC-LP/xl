package com.tjclp.xl.cli.helpers

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{Column, Row}
import com.tjclp.xl.sheets.{ColumnProperties, RowProperties}

/**
 * Pure row/column outline-grouping appliers (GH-421), shared by the CLI command handlers
 * (WriteCommands) and the batch ops (BatchParser) so the two paths cannot drift.
 *
 * Grouping writes through the existing `RowProperties`/`ColumnProperties` outlineLevel + collapsed
 * fields (both writer backends already persist them). `collapsed` follows Excel's convention:
 * members are hidden and the summary row/column AFTER the group (the summaryBelow/summaryRight
 * default) carries the collapsed marker that draws the "+" button. Ungrouping clears the outline
 * level and collapse markers but — like Excel — does not unhide members a collapse hid (use
 * `row`/`col --show`). Validation happens here, returning Left with a clean message, so the
 * properties' `require` guards are never tripped from the CLI.
 */
object GroupingOps:

  /** Group rows into a collapsible outline: every row in the spec gets `level`. */
  def groupRows(
    sheet: Sheet,
    spec: String,
    level: Int,
    collapsed: Boolean
  ): Either[String, Sheet] =
    for
      _ <- validateLevel(level)
      rows <- parseRowSpec(spec)
      (start, end) = rows
    yield
      val withMembers = (start to end).foldLeft(sheet) { (s, r) =>
        updateRow(s, Row.from1(r))(p =>
          p.copy(outlineLevel = Some(level), hidden = collapsed || p.hidden)
        )
      }
      if collapsed && end < Row.MaxIndex0 + 1 then
        updateRow(withMembers, Row.from1(end + 1))(_.copy(collapsed = true))
      else withMembers

  /** Group columns into a collapsible outline: every column in the spec gets `level`. */
  def groupCols(
    sheet: Sheet,
    spec: String,
    level: Int,
    collapsed: Boolean
  ): Either[String, Sheet] =
    for
      _ <- validateLevel(level)
      cols <- parseColSpec(spec)
      (start, end) = cols
    yield
      val withMembers = (start.index0 to end.index0).foldLeft(sheet) { (s, c) =>
        updateCol(s, Column.from0(c))(p =>
          p.copy(outlineLevel = Some(level), hidden = collapsed || p.hidden)
        )
      }
      if collapsed && end.index0 < Column.MaxIndex0 then
        updateCol(withMembers, Column.from0(end.index0 + 1))(_.copy(collapsed = true))
      else withMembers

  /** Clear outline level + collapse markers for the rows (and the group's summary row). */
  def ungroupRows(sheet: Sheet, spec: String): Either[String, Sheet] =
    parseRowSpec(spec).map { case (start, end) =>
      val cleared = (start to end).foldLeft(sheet) { (s, r) =>
        updateRow(s, Row.from1(r))(_.copy(outlineLevel = None, collapsed = false))
      }
      if end < Row.MaxIndex0 + 1 then
        updateRow(cleared, Row.from1(end + 1))(_.copy(collapsed = false))
      else cleared
    }

  /** Clear outline level + collapse markers for the columns (and the group's summary column). */
  def ungroupCols(sheet: Sheet, spec: String): Either[String, Sheet] =
    parseColSpec(spec).map { case (start, end) =>
      val cleared = (start.index0 to end.index0).foldLeft(sheet) { (s, c) =>
        updateCol(s, Column.from0(c))(_.copy(outlineLevel = None, collapsed = false))
      }
      if end.index0 < Column.MaxIndex0 then
        updateCol(cleared, Column.from0(end.index0 + 1))(_.copy(collapsed = false))
      else cleared
    }

  /** Parse a 1-based row spec: a single row ("10") or an inclusive range ("10:20"). */
  def parseRowSpec(spec: String): Either[String, (Int, Int)] =
    spec.split(':') match
      case Array(single) => parseRow1(single).map(r => (r, r))
      case Array(a, b) =>
        for
          ra <- parseRow1(a)
          rb <- parseRow1(b)
        yield (math.min(ra, rb), math.max(ra, rb))
      case _ => Left(s"Invalid row spec '$spec' (expected a row like 10 or a range like 10:20)")

  /** Parse a column spec: a single column ("E") or an inclusive range ("E:H"). */
  def parseColSpec(spec: String): Either[String, (Column, Column)] =
    spec.split(':') match
      case Array(single) => parseCol(single).map(c => (c, c))
      case Array(a, b) =>
        for
          ca <- parseCol(a)
          cb <- parseCol(b)
        yield
          if ca.index0 <= cb.index0 then (ca, cb)
          else (cb, ca)
      case _ =>
        Left(s"Invalid column spec '$spec' (expected a column like E or a range like E:H)")

  private def validateLevel(level: Int): Either[String, Unit] =
    if level >= 1 && level <= 7 then Right(())
    else Left(s"Outline level must be 1-7, got: $level")

  private def parseRow1(s: String): Either[String, Int] =
    s.trim.toIntOption
      .filter(r => r >= 1 && r <= Row.MaxIndex0 + 1)
      .toRight(s"Invalid row number '${s.trim}' (expected 1-${Row.MaxIndex0 + 1})")

  private def parseCol(s: String): Either[String, Column] =
    Column.fromLetter(s.trim)

  /**
   * Apply `f` to the row's properties, keeping the entry even when the result is all-default: an
   * explicit entry is authoritative on write (actively clearing preserved source attributes),
   * whereas a missing entry lets a preserved `<row>` ride through verbatim — pruning would
   * resurrect the very collapsed/outlineLevel attrs an ungroup just cleared. The reader drops
   * all-default rows, so the model converges on the next read.
   */
  private def updateRow(sheet: Sheet, row: Row)(f: RowProperties => RowProperties): Sheet =
    sheet.setRowProperties(row, f(sheet.getRowProperties(row)))

  /**
   * Apply `f` to the column's properties, keeping all-default entries for the same reason as
   * [[updateRow]]: an empty columnProperties map would resurrect the preserved source `<cols>`
   * element wholesale.
   */
  private def updateCol(sheet: Sheet, col: Column)(f: ColumnProperties => ColumnProperties): Sheet =
    sheet.setColumnProperties(col, f(sheet.getColumnProperties(col)))
