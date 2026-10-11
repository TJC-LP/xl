package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.eval.DependentRecalculation.*
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

class ReviewLookupNamesIntegrationSpec extends FunSuite:
  private def n(value: Int): CellValue = CellValue.Number(BigDecimal(value))
  private val sheet = Sheet(SheetName.unsafe("S"))
    .put(ref"A1", n(1))
    .put(ref"A2", n(2))
    .put(ref"A3", n(3))
    .put(ref"C1", n(10))
    .put(ref"C2", n(20))
    .put(ref"C3", n(30))

  test("LET criteria references invalidate every cell in the resized SUMIF range") {
    val formula = "LET(r,IF(TRUE,A1:A3,A1:A3),SUMIF(r,\">0\",C1))"
    val original = sheet.put(ref"E1", CellValue.Formula(formula, Some(n(60))))
    assertEquals(original.evaluateCell(ref"E1"), Right(n(60)))
    val updated = original.put(ref"C2", n(200)).recalculateDependents(Set(ref"C2"))
    assertEquals(updated(ref"E1").value, CellValue.Formula(formula, Some(n(240))))
  }

  test("dynamic name criteria references invalidate every resized SUMIF target") {
    val formula = "SUMIF(criteria,\">0\",C1)"
    val original = sheet.put(ref"E1", CellValue.Formula(formula, Some(n(60))))
    val wb = Workbook(original).withDefinedName("criteria", "IF(TRUE,S!A1:A3,S!A1:A3)")
    assertEquals(original.evaluateCell(ref"E1", Clock.system, Some(wb)), Right(n(60)))
    val updated = wb
      .put(original.put(ref"C2", n(200)))
      .recalculateDependents(original.name, Set(ref"C2"))
    val target = updated(original.name).fold(e => fail(e.message), identity)
    assertEquals(target(ref"E1").value, CellValue.Formula(formula, Some(n(240))))
  }

  test("LET SUMIF target references invalidate every resized cell") {
    val formula = "LET(t,C1,SUMIF(A1:A3,\">0\",t))"
    val original = sheet.put(ref"E1", CellValue.Formula(formula, Some(n(60))))
    assertEquals(original.evaluateCell(ref"E1"), Right(n(60)))
    val updated = original.put(ref"C2", n(200)).recalculateDependents(Set(ref"C2"))
    assertEquals(updated(ref"E1").value, CellValue.Formula(formula, Some(n(240))))
  }

  test("text coercion and evaluating lookups work through dynamic LET ranges") {
    val data = sheet
      .put(ref"A2", CellValue.Formula("ABS(\"2\")", None))
      .put(ref"C2", CellValue.Formula("ABS(\"20\")", None))
    val formulas = List(
      "LET(r,OFFSET(A1,0,0,3,1),MATCH(2,r,0))" -> n(2),
      "LET(r,OFFSET(A1,0,0,3,1),t,OFFSET(C1,0,0,3,1),XLOOKUP(2,r,t))" -> n(20),
      "LET(r,OFFSET(A1,0,0,3,1),t,OFFSET(C1,0,0,3,1),IFNA(XLOOKUP(9,r,t),99))" -> n(99)
    )
    formulas.foreach { (formula, expected) =>
      assertEquals(data.evaluateFormula(formula), Right(expected), formula)
    }
  }

  test("lookup recalc diagnostics coexist with numeric text coercion") {
    val data = sheet
      .put(ref"E1", CellValue.Formula("XLOOKUP(99,A1:A3,C1:C3)", None))
      .put(ref"E2", CellValue.Formula("ABS(\"oops\")", None))
    val result = Workbook(data).recalculate()
    assert(result.errors.isEmpty)
    assertEquals(result.excelErrors.map(_._3).toSet, Set(CellError.NA, CellError.Value))
    assertEquals(result.errorCells.size, 2)
    assert(result.errorCells.forall(_.reason.nonEmpty))
  }
