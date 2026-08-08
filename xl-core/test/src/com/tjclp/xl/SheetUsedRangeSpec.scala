package com.tjclp.xl

import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.sheets.Sheet
import munit.FunSuite

class SheetUsedRangeSpec extends FunSuite:

  test("usedRange is empty when the sheet has no non-empty cells"):
    val sheet = Sheet(SheetName.unsafe("S"))
      .put(Cell(ref"XFD1048576", CellValue.Empty))

    assertEquals(sheet.usedRange, None)

  test("usedRange computes the bounding box of non-empty cells"):
    val sheet = Sheet(SheetName.unsafe("S"))
      .put(ref"C7", CellValue.Number(BigDecimal(1)))
      .put(ref"A3", CellValue.Text("x"))
      .put(ref"B5", CellValue.Bool(true))

    assertEquals(sheet.usedRange, Some(CellRange(ref"A3", ref"C7")))
