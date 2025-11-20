package com.tjclp.xl

import com.tjclp.xl.api.*
import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.sheets.syntax.*
import munit.FunSuite
import com.tjclp.xl.macros.ref

class RangeCombinatorsSpec extends FunSuite:

  private def newSheet: Sheet =
    Sheet("Test").getOrElse(fail("Failed to create sheet"))

  test("putRange populates all cells when counts match") {
    val targetRange = ref"A1:B2"
    val values = Vector(
      CellValue.Text("A1"),
      CellValue.Text("B1"),
      CellValue.Text("A2"),
      CellValue.Text("B2")
    )

    val updated = newSheet.putRange(targetRange, values).getOrElse(fail("putRange returned error"))

    assertEquals(updated(ref"A1").value, CellValue.Text("A1"))
    assertEquals(updated(ref"B2").value, CellValue.Text("B2"))
    assertEquals(updated.cellCount, 4)
  }

  test("putRange detects mismatched value counts") {
    val result = newSheet.putRange(
      ref"A1:B2",
      Vector(
        CellValue.Text("A1"),
        CellValue.Text("B1"),
        CellValue.Text("A2")
      )
    )

    result match
      case Left(XLError.ValueCountMismatch(expected, actual, context)) =>
        assertEquals(expected, 4)
        assertEquals(actual, 3)
        assertEquals(context, "range A1:B2")
      case other =>
        fail(s"Expected ValueCountMismatch, got $other")
  }

  test("putRow populates contiguous columns and enforces bounds") {
    val values = Vector(CellValue.Text("Q1"), CellValue.Text("Q2"), CellValue.Text("Q3"))
    val updated =
      newSheet.putRow(Row.from1(1), Column.from1(1), values).getOrElse(fail("putRow failed"))

    assertEquals(updated(ref"A1").value, CellValue.Text("Q1"))
    assertEquals(updated(ref"C1").value, CellValue.Text("Q3"))
    assertEquals(updated.cellCount, 3)

    val overflow = newSheet.putRow(
      Row.from1(1),
      Column.from0(Column.MaxIndex0),
      Vector(CellValue.Text("last"), CellValue.Text("overflow"))
    )

    overflow match
      case Left(XLError.OutOfBounds(_, reason)) =>
        assert(reason.contains("Cannot write") && reason.contains("Excel limit"))
      case other =>
        fail(s"Expected OutOfBounds for column overflow, got $other")
  }

  test("putCol populates contiguous rows and enforces bounds") {
    val values = Vector(CellValue.Text("North"), CellValue.Text("South"))
    val updated =
      newSheet.putCol(Column.from1(1), Row.from1(1), values).getOrElse(fail("putCol failed"))

    assertEquals(updated(ref"A1").value, CellValue.Text("North"))
    assertEquals(updated(ref"A2").value, CellValue.Text("South"))
    assertEquals(updated.cellCount, 2)

    val overflow = newSheet.putCol(
      Column.from1(1),
      Row.from0(Row.MaxIndex0),
      Vector(CellValue.Text("last row"), CellValue.Text("overflow"))
    )

    overflow match
      case Left(XLError.OutOfBounds(_, reason)) =>
        assert(reason.contains("Cannot write") && reason.contains("Excel limit"))
      case other =>
        fail(s"Expected OutOfBounds for row overflow, got $other")
  }
