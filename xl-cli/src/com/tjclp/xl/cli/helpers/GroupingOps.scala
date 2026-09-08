package com.tjclp.xl.cli.helpers

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{Column, Row}
import com.tjclp.xl.ops.{ColSpan, RowSpan}

/**
 * Forwarders onto the outline-grouping semantics of `Sheet.groupRows` / `groupCols` / `ungroupRows`
 * / `ungroupCols` (ADR-017 §2.12, W2.1 — the GH-421 appliers moved into xl-core), shared by the CLI
 * command handlers (WriteCommands) and the batch ops (BatchParser). What stays here is the CLI's
 * spec syntax (`10:20`, `E:H`) and its messages; the level guard and Excel's collapsed-summary
 * convention are the core's, so the two paths cannot drift.
 */
object GroupingOps:

  /** Group rows into a collapsible outline: every row in the spec gets `level`. */
  def groupRows(
    sheet: Sheet,
    spec: String,
    level: Int,
    collapsed: Boolean
  ): Either[String, Sheet] =
    parseRowSpec(spec).flatMap((start, end) =>
      sheet.groupRows(rowSpan(start, end), level, collapsed).left.map(_.message)
    )

  /** Group columns into a collapsible outline: every column in the spec gets `level`. */
  def groupCols(
    sheet: Sheet,
    spec: String,
    level: Int,
    collapsed: Boolean
  ): Either[String, Sheet] =
    parseColSpec(spec).flatMap((start, end) =>
      sheet.groupCols(ColSpan(start, end), level, collapsed).left.map(_.message)
    )

  /** Clear outline level + collapse markers for the rows (and the group's summary row). */
  def ungroupRows(sheet: Sheet, spec: String): Either[String, Sheet] =
    parseRowSpec(spec).map((start, end) => sheet.ungroupRows(rowSpan(start, end)))

  /** Clear outline level + collapse markers for the columns (and the group's summary column). */
  def ungroupCols(sheet: Sheet, spec: String): Either[String, Sheet] =
    parseColSpec(spec).map((start, end) => sheet.ungroupCols(ColSpan(start, end)))

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

  private def rowSpan(start: Int, end: Int): RowSpan = RowSpan(Row.from1(start), Row.from1(end))

  private def parseRow1(s: String): Either[String, Int] =
    s.trim.toIntOption
      .filter(r => r >= 1 && r <= Row.MaxIndex0 + 1)
      .toRight(s"Invalid row number '${s.trim}' (expected 1-${Row.MaxIndex0 + 1})")

  private def parseCol(s: String): Either[String, Column] =
    Column.fromLetter(s.trim)
