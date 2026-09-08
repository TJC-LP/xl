package com.tjclp.xl.sheets

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{CellRange, Column, Row}
import com.tjclp.xl.error.XLError
import com.tjclp.xl.ops.{ColSpan, RowSpan}
import munit.FunSuite

/**
 * GH-465 (second half) / GH-589: `collapseRows`/`collapseCols` compose what the CLI's
 * `group-rows --collapsed` assembles — hide the members, mark the summary row/column after the
 * group `collapsed`, and make ungrouped members a level-1 group so Excel draws the outline button.
 * `expandRows`/`expandCols` are the inverse (members unhidden, marker cleared, level kept). Both
 * are the `SheetEdits.outlineRows`/`outlineCols` fold `groupRows`/`ungroupRows` use, and the
 * `CellRange` overloads refuse a range of the wrong axis instead of projecting it.
 */
class SheetOutlineSpec extends FunSuite:

  private def row(n: Int): Row = Row.from1(n)
  private def col(letter: String): Column =
    Column.fromLetter(letter).getOrElse(fail(s"bad column $letter"))
  private def range(s: String): CellRange =
    CellRange.parse(s).getOrElse(fail(s"bad range $s"))

  test("collapseRows hides the members and marks the summary row collapsed") {
    val sheet = Sheet("Outline").collapseRows(row(5), row(8))
    (5 to 8).foreach { r =>
      val props = sheet.getRowProperties(row(r))
      assert(props.hidden, s"row $r should be hidden")
      assertEquals(props.outlineLevel, Some(1), s"row $r should be a level-1 group")
      assert(!props.collapsed, s"member row $r must not carry the collapsed marker")
    }
    val summary = sheet.getRowProperties(row(9))
    assert(summary.collapsed, "summary row 9 carries the collapsed marker")
    assert(!summary.hidden, "summary row stays visible")
    assertEquals(summary.outlineLevel, None)
    assertEquals(sheet.rowProperties.size, 5)
  }

  test("collapseRows keeps an existing outline level and other properties") {
    val grouped = Sheet("Levels")
      .setRowProperties(row(5), RowProperties(height = Some(30.0), outlineLevel = Some(2)))
      .setRowProperties(row(6), RowProperties(outlineLevel = Some(2)))
    val sheet = grouped.collapseRows(row(5), row(6))
    assertEquals(sheet.getRowProperties(row(5)).outlineLevel, Some(2))
    assertEquals(sheet.getRowProperties(row(5)).height, Some(30.0))
    assert(sheet.getRowProperties(row(5)).hidden)
    assert(sheet.getRowProperties(row(7)).collapsed)
  }

  test("collapseRows normalizes a reversed span; RowSpan and full-row CellRange overloads agree") {
    val bySpan = Sheet("Norm").collapseRows(row(8), row(5))
    val byRowSpan = Sheet("Norm").collapseRows(RowSpan(row(5), row(8)))
    val byRange = Sheet("Norm").collapseRows(range("5:8"))
    assertEquals(byRowSpan, bySpan)
    assertEquals(byRange, Right(bySpan))
    assert(bySpan.getRowProperties(row(5)).hidden)
    assert(bySpan.getRowProperties(row(9)).collapsed)
  }

  test("collapseRows on an ungrouped sheet is groupRows(level 1, collapsed) — one composition") {
    val span = RowSpan(row(5), row(8))
    val sheet = Sheet("Same").put("A1" -> "x")
    assertEquals(sheet.groupRows(span, 1, collapsed = true), Right(sheet.collapseRows(span)))
    val cols = ColSpan(col("E"), col("H"))
    assertEquals(sheet.groupCols(cols, 1, collapsed = true), Right(sheet.collapseCols(cols)))
  }

  test("the CellRange overloads refuse a range of the wrong axis instead of projecting it") {
    val columns = range("E:H") // rows 1..1048576 — collapseRows would hide the whole sheet
    val rows = range("2:3") // columns A..XFD
    val cells = range("A5:D8")
    val sheet = Sheet("Axis")
    Seq(sheet.collapseRows(columns), sheet.expandRows(columns), sheet.collapseRows(cells)).foreach {
      case Left(XLError.InvalidReference(reason)) =>
        assert(reason.contains("is not a row span"), reason)
      case other => fail(s"expected InvalidReference, got $other")
    }
    Seq(sheet.collapseCols(rows), sheet.expandCols(rows), sheet.collapseCols(cells)).foreach {
      case Left(XLError.InvalidReference(reason)) =>
        assert(reason.contains("is not a column span"), reason)
      case other => fail(s"expected InvalidReference, got $other")
    }
    assertEquals(sheet.collapseCols(columns), Right(sheet.collapseCols(col("E"), col("H"))))
    assertEquals(sheet.collapseRows(rows), Right(sheet.collapseRows(row(2), row(3))))
  }

  test("collapseRows on the last row has no summary row to mark") {
    val last = Row.from0(Row.MaxIndex0)
    val sheet = Sheet("Edge").collapseRows(last, last)
    assert(sheet.getRowProperties(last).hidden)
    assertEquals(sheet.rowProperties.size, 1)
  }

  test("expandRows unhides members, clears the marker, keeps the outline level") {
    val collapsed = Sheet("Expand").collapseRows(row(5), row(8))
    val expanded = collapsed.expandRows(row(5), row(8))
    (5 to 8).foreach { r =>
      val props = expanded.getRowProperties(row(r))
      assert(!props.hidden, s"row $r should be visible again")
      assertEquals(props.outlineLevel, Some(1), s"row $r keeps its group level")
    }
    assert(!expanded.getRowProperties(row(9)).collapsed)
    assertEquals(expanded.expandRows(RowSpan(row(5), row(8))), expanded)
    assertEquals(expanded.expandRows(range("5:8")), Right(expanded))
  }

  test("expandRows/expandCols over rows and columns without properties are no-ops") {
    val sheet = Sheet("Untouched").put("A1" -> 1)
    assertEquals(sheet.expandRows(row(1), row(50)), sheet)
    assertEquals(sheet.expandRows(RowSpan(row(1), row(50))), sheet)
    assertEquals(sheet.expandCols(col("A"), col("Z")), sheet)
    assertEquals(sheet.rowProperties, Map.empty)
    assertEquals(sheet.columnProperties, Map.empty)
    // an entry that exists is updated in place (and kept when it becomes all-default)
    val hiddenOnly = sheet.setRowProperties(row(3), RowProperties(hidden = true))
    assertEquals(
      hiddenOnly.expandRows(row(1), row(5)).rowProperties,
      Map(row(3) -> RowProperties())
    )
  }

  test("collapseCols hides the members and marks the summary column collapsed") {
    val sheet = Sheet("Cols").collapseCols(col("E"), col("H"))
    Seq("E", "F", "G", "H").foreach { c =>
      val props = sheet.getColumnProperties(col(c))
      assert(props.hidden, s"column $c should be hidden")
      assertEquals(props.outlineLevel, Some(1))
      assert(!props.collapsed)
    }
    val summary = sheet.getColumnProperties(col("I"))
    assert(summary.collapsed)
    assert(!summary.hidden)
    assertEquals(sheet.columnProperties.size, 5)
  }

  test("collapseCols keeps width and each member's level; ColSpan, CellRange and reversed agree") {
    val grouped = Sheet("ColLevels")
      .setColumnProperties(col("E"), ColumnProperties(width = Some(18.0), outlineLevel = Some(3)))
    val sheet = grouped.collapseCols(col("E"), col("F"))
    assertEquals(sheet.getColumnProperties(col("E")).width, Some(18.0))
    assertEquals(sheet.getColumnProperties(col("E")).outlineLevel, Some(3))
    assertEquals(sheet.getColumnProperties(col("F")).outlineLevel, Some(1))
    val reversed = Sheet("ColNorm").collapseCols(col("F"), col("E"))
    val bySpan = Sheet("ColNorm").collapseCols(ColSpan(col("E"), col("F")))
    assertEquals(bySpan, reversed)
    assertEquals(Sheet("ColNorm").collapseCols(range("E:F")), Right(reversed))
  }

  test("collapseCols on the last column has no summary column to mark") {
    val last = Column.from0(Column.MaxIndex0)
    val sheet = Sheet("ColEdge").collapseCols(last, last)
    assert(sheet.getColumnProperties(last).hidden)
    assertEquals(sheet.columnProperties.size, 1)
  }

  test("expandCols is the inverse of collapseCols up to the retained outline level") {
    val expanded =
      Sheet("ColExpand").collapseCols(col("E"), col("H")).expandCols(col("E"), col("H"))
    Seq("E", "F", "G", "H").foreach { c =>
      assert(!expanded.getColumnProperties(col(c)).hidden)
      assertEquals(expanded.getColumnProperties(col(c)).outlineLevel, Some(1))
    }
    assert(!expanded.getColumnProperties(col("I")).collapsed)
    assertEquals(expanded.expandCols(ColSpan(col("E"), col("H"))), expanded)
    assertEquals(expanded.expandCols(range("E:H")), Right(expanded))
  }
