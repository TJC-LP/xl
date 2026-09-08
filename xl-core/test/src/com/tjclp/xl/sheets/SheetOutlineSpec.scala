package com.tjclp.xl.sheets

import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row}
import munit.FunSuite

/**
 * GH-465 (second half) / GH-589: `collapseRows`/`collapseCols` compose what the CLI's
 * `group-rows --collapsed` assembles by hand — hide the members, mark the summary row/column after
 * the group `collapsed`, and make ungrouped members a level-1 group so Excel draws the outline
 * button. `expandRows`/`expandCols` are the inverse (members unhidden, marker cleared, level kept).
 */
class SheetOutlineSpec extends FunSuite:

  private def row(n: Int): Row = Row.from1(n)
  private def col(letter: String): Column =
    Column.fromLetter(letter).getOrElse(fail(s"bad column $letter"))

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

  test("collapseRows normalizes a reversed span and accepts a CellRange") {
    val bySpan = Sheet("Norm").collapseRows(row(8), row(5))
    val byRange = Sheet("Norm").collapseRows(CellRange(ARef.from0(0, 4), ARef.from0(3, 7)))
    assertEquals(bySpan.rowProperties, byRange.rowProperties)
    assert(bySpan.getRowProperties(row(5)).hidden)
    assert(bySpan.getRowProperties(row(9)).collapsed)
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
    assertEquals(expanded.expandRows(CellRange(ARef.from0(0, 4), ARef.from0(0, 7))), expanded)
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

  test("collapseCols keeps width and level; CellRange overload and reversed span agree") {
    val grouped = Sheet("ColLevels")
      .setColumnProperties(col("E"), ColumnProperties(width = Some(18.0), outlineLevel = Some(3)))
    val sheet = grouped.collapseCols(col("E"), col("F"))
    assertEquals(sheet.getColumnProperties(col("E")).width, Some(18.0))
    assertEquals(sheet.getColumnProperties(col("E")).outlineLevel, Some(3))
    assertEquals(sheet.getColumnProperties(col("F")).outlineLevel, Some(1))
    val reversed = Sheet("ColNorm").collapseCols(col("F"), col("E"))
    val byRange = Sheet("ColNorm").collapseCols(CellRange(ARef.from0(4, 0), ARef.from0(5, 9)))
    assertEquals(reversed.columnProperties, byRange.columnProperties)
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
    assertEquals(expanded.expandCols(CellRange(ARef.from0(4, 0), ARef.from0(7, 0))), expanded)
  }
