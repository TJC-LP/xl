package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * Tests for the workbook-level evaluateFormula (0.11.0): cross-sheet context is wired
 * automatically, so scripts never touch Sheet.evaluateFormula's Option[Workbook] parameter.
 */
class WorkbookEvaluateFormulaSpec extends FunSuite:

  private def num(i: Int): CellValue = CellValue.Number(BigDecimal(i))

  private val sales = Sheet(SheetName.unsafe("Sales"))
    .put(ARef.from0(0, 0), num(10))
    .put(ARef.from0(0, 1), num(20))

  private val summary = Sheet(SheetName.unsafe("Summary"))
    .put(ARef.from0(0, 0), num(5))

  private val wb = Workbook(sales, summary)

  test("evaluates intra-sheet formula on the named sheet"):
    assertEquals(wb.evaluateFormula("=SUM(A1:A2)", "Sales"), Right(num(30)))
    assertEquals(wb.evaluateFormula("=A1*2", "Summary"), Right(num(10)))

  test("cross-sheet references resolve without passing workbook explicitly"):
    assertEquals(wb.evaluateFormula("=Sales!A1 + Sales!A2", "Summary"), Right(num(30)))
    assertEquals(wb.evaluateFormula("=SUM(Sales!A1:A2) + A1", "Summary"), Right(num(35)))

  test("SheetName variant behaves identically"):
    assertEquals(
      wb.evaluateFormula("=SUM(Sales!A1:A2)", SheetName.unsafe("Summary")),
      Right(num(30))
    )

  test("missing sheet surfaces a structured error"):
    wb.evaluateFormula("=A1", "Nope") match
      case Left(err) => assert(err.message.contains("Nope"))
      case Right(v) => fail(s"expected SheetNotFound, got $v")
