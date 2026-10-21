package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.{DefinedName, Workbook}
import munit.FunSuite

/**
 * #752: a reference's sheet qualifier resolves case-insensitively, as in Excel: on a book whose tab
 * is `Data`, `=data!B1`, `=SUM(DATA!A:A)`, `=INDIRECT("data!B1")` and `=data!rate` read `Data`, and
 * recalculation orders and evaluates them like the `Data` spelling.
 */
class SheetNameCaseSpec extends FunSuite:

  private def num(i: Int): CellValue = CellValue.Number(BigDecimal(i))
  private def formula(expr: String): CellValue = CellValue.Formula(expr, None)
  private def a1(s: String): ARef = ARef.parse(s).fold(err => fail(err), identity)

  private def sheetWith(name: String, cells: (String, CellValue)*): Sheet =
    cells.foldLeft(Sheet(SheetName.unsafe(name))) { case (s, (ref, value)) =>
      s.put(a1(ref), value)
    }

  private val data = sheetWith("Data", "A1" -> num(1), "A2" -> num(2), "B1" -> num(7))
  private val summary = Sheet(SheetName.unsafe("Summary"))
  private val book = Workbook(Vector(data, summary))

  private def cached(wb: Workbook, sheet: String, ref: String): Option[CellValue] =
    wb.sheets
      .find(_.name.value == sheet)
      .flatMap(_.cells.get(a1(ref)))
      .map(_.value)
      .collect { case CellValue.Formula(_, Some(v), _) => v }

  test("a qualified cell or range in another case evaluates as the tab's spelling") {
    assertEquals(book.evaluateFormula("=data!B1", "Summary"), Right(num(7)))
    assertEquals(book.evaluateFormula("=SUM(DATA!A:A)", "Summary"), Right(num(3)))
    assertEquals(book.evaluateFormula("=dAtA!A1+Data!A2", "Summary"), Right(num(3)))
  }

  test("INDIRECT text names the sheet in any case") {
    assertEquals(book.evaluateFormula("=INDIRECT(\"data!B1\")", "Summary"), Right(num(7)))
    assertEquals(book.evaluateFormula("=SUM(INDIRECT(\"DATA!A1:A2\"))", "Summary"), Right(num(3)))
  }

  test("a sheet-qualified defined name finds the name scoped to the tab in any case") {
    val scoped = book.copy(metadata =
      book.metadata.copy(definedNames =
        Vector(DefinedName("rate", "Data!$B$1", localSheetId = Some(0)))
      )
    )
    assertEquals(scoped.evaluateFormula("=data!rate", "Summary"), Right(num(7)))
    assertEquals(scoped.evaluateFormula("=Data!rate", "Summary"), Right(num(7)))
  }

  test("recalculation orders =data!B1 after Data!B1 and caches its value") {
    // Data!B1 is itself a formula, so a wrong order or a failed lookup leaves Summary!A1 wrong
    val wb = Workbook(
      Vector(
        sheetWith("Summary", "A1" -> formula("data!B1*2")),
        sheetWith("Data", "A1" -> num(5), "B1" -> formula("A1+1"))
      )
    )
    val result = wb.recalculate()
    assertEquals(result.errors, Vector.empty)
    assertEquals(cached(result.workbook, "Data", "B1"), Some(num(6)))
    assertEquals(cached(result.workbook, "Summary", "A1"), Some(num(12)))
  }

  test(
    "tabs differing only by case: an exact qualifier reads its own tab, a third spelling fails"
  ) {
    val twins = Workbook(
      Vector(sheetWith("Data", "A1" -> num(1)), sheetWith("DATA", "A1" -> num(2)), summary)
    )
    assertEquals(twins.evaluateFormula("=Data!A1", "Summary"), Right(num(1)))
    assertEquals(twins.evaluateFormula("=DATA!A1", "Summary"), Right(num(2)))
    assert(twins.evaluateFormula("=data!A1", "Summary").isLeft)
  }
