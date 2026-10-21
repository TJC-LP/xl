package com.tjclp.xl.workbooks

import com.tjclp.xl.Workbook
import com.tjclp.xl.addressing.{ARef, RefType, SheetName}
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.sheets.Sheet
import munit.FunSuite

/**
 * #752: `Workbook.apply(SheetName)` resolves a sheet name the way Excel resolves a reference's
 * qualifier: the exact spelling first, then the one sheet whose name is equal ignoring case. Two
 * tabs that differ only by case (a book Excel refuses to save) make a case-folded lookup ambiguous,
 * and it is refused rather than guessed.
 */
class WorkbookSheetLookupSpec extends FunSuite:

  private def name(s: String): SheetName = SheetName.unsafe(s)

  private val a1 = ARef.from0(0, 0)
  private val data = Sheet("Data").put(a1, CellValue.Number(BigDecimal(1)))
  private val summary = Sheet("Summary")
  private val wb = Workbook(Vector(data, summary))

  private def sheetName(result: XLResult[Sheet]): String =
    result.fold(err => fail(s"expected a sheet, got $err"), _.name.value)

  test("an exact name resolves to its sheet") {
    assertEquals(sheetName(wb(name("Data"))), "Data")
  }

  test("a name in another case resolves to the sheet Excel would read") {
    assertEquals(sheetName(wb(name("data"))), "Data")
    assertEquals(sheetName(wb(name("DATA"))), "Data")
    assertEquals(sheetName(wb(name("sUMMARY"))), "Summary")
  }

  test("a qualified ref in another case reads the sheet's cell") {
    wb(RefType.QualifiedCell(name("data"), a1)) match
      case Right(cell: Cell) => assertEquals(cell.value, CellValue.Number(BigDecimal(1)))
      case other => fail(s"expected the cell, got $other")
  }

  test("a name no sheet carries in any case is SheetNotFound with the sheet names") {
    assertEquals(wb(name("Nope")), Left(XLError.SheetNotFound("Nope", Vector("Data", "Summary"))))
  }

  test("tabs differing only by case: the exact spelling wins, any other spelling is refused") {
    val twins = Workbook(Vector(data, Sheet("DATA"), summary))
    assertEquals(sheetName(twins(name("Data"))), "Data")
    assertEquals(sheetName(twins(name("DATA"))), "DATA")
    twins(name("data")) match
      case Left(XLError.InvalidReference(reason)) =>
        assert(reason.contains("'data'"), reason)
        assert(reason.contains("'Data', 'DATA'"), reason)
      case other => fail(s"expected an ambiguity refusal, got $other")
    // the rest of the book still folds
    assertEquals(sheetName(twins(name("summary"))), "Summary")
  }

  test("the shared fold is Excel's: lower case under the root locale") {
    assertEquals(name("Ärger").folded, name("äRGER").folded)
    assertEquals(name("SHEET1").folded, "sheet1")
  }
