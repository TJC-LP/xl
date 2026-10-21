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

  test("update, upsert, remove, delete, rename and setSheetState match a name the same way") {
    def names(result: XLResult[Workbook]): XLResult[Seq[String]] =
      result.map(_.sheetNames.map(_.value))
    val b2 = ARef.from0(1, 1)
    val updated = wb.update(name("data"), _.put(b2, CellValue.Number(BigDecimal(2))))
    assertEquals(names(updated), Right(Seq("Data", "Summary")))
    assertEquals(
      updated.flatMap(_(name("Data"))).map(_(b2).value),
      Right(CellValue.Number(BigDecimal(2)))
    )
    // upsert updates the tab rather than adding a second tab Excel treats as the same name
    val upserted = wb.upsert(name("DATA"), _.put(b2, CellValue.Number(BigDecimal(3))))
    assertEquals(upserted.sheetNames.map(_.value), Seq("Data", "Summary"))
    assertEquals(names(wb.remove(name("summary"))), Right(Seq("Data")))
    assertEquals(names(wb.delete(name("SUMMARY"))), Right(Seq("Data")))
    assertEquals(names(wb.rename(name("data"), name("Inputs"))), Right(Seq("Inputs", "Summary")))
    assertEquals(
      wb.setSheetState(name("summary"), Some("hidden")).map(_.getSheetState(name("Summary"))),
      Right(Some("hidden"))
    )
    assertEquals(
      wb.update(name("Nope"), identity),
      Left(XLError.SheetNotFound("Nope", Vector("Data", "Summary")))
    )
  }

  test("a defined name's sheet scope takes the exact spelling first on tabs differing by case") {
    val twins = Workbook(Vector(data, Sheet("DATA"), summary))
    def scopeOf(scope: String): XLResult[Option[Int]] =
      twins
        .withDefinedName("x", "DATA!$A$1", name(scope))
        .map(_.metadata.definedNames.find(_.name == "x").flatMap(_.localSheetId))
    assertEquals(scopeOf("Data"), Right(Some(0)))
    assertEquals(scopeOf("DATA"), Right(Some(1)))
    assertEquals(scopeOf("summary"), Right(Some(2)))
    assert(scopeOf("data").isLeft, "a third spelling of the twins is refused")
  }
