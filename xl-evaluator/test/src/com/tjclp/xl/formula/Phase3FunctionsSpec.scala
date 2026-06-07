package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cells.{CellError, CellValue}
import munit.FunSuite

/** GH-76 (tier 1): conditional/selection functions IFS, SWITCH, CHOOSE. */
class Phase3FunctionsSpec extends FunSuite:
  private val s = new Sheet(name = SheetName.unsafe("S"))

  test("IFS returns the first TRUE branch") {
    assertEquals(s.evaluateFormula("=IFS(1>2,\"a\",2>1,\"b\")"), Right(CellValue.Text("b")))
  }

  test("IFS with no TRUE condition returns #N/A") {
    assertEquals(
      s.evaluateFormula("=IFS(1>2,\"a\",3>4,\"b\")"),
      Right(CellValue.Error(CellError.NA))
    )
  }

  test("SWITCH matches a case") {
    assertEquals(s.evaluateFormula("=SWITCH(2,1,\"a\",2,\"b\",\"def\")"), Right(CellValue.Text("b")))
  }

  test("SWITCH falls back to the trailing default") {
    assertEquals(s.evaluateFormula("=SWITCH(9,1,\"a\",\"def\")"), Right(CellValue.Text("def")))
  }

  test("SWITCH with no match and no default returns #N/A") {
    assertEquals(s.evaluateFormula("=SWITCH(9,1,\"a\")"), Right(CellValue.Error(CellError.NA)))
  }

  test("CHOOSE selects the 1-based index") {
    assertEquals(s.evaluateFormula("=CHOOSE(2,\"a\",\"b\",\"c\")"), Right(CellValue.Text("b")))
  }

  test("CHOOSE out of range returns #VALUE!") {
    assertEquals(s.evaluateFormula("=CHOOSE(9,\"a\")"), Right(CellValue.Error(CellError.Value)))
  }

  // GH-120: statistical functions over a range. Data A1:A5 = 3,1,4,1,5.
  private val nums =
    s.put("A1" -> 3).put("A2" -> 1).put("A3" -> 4).put("A4" -> 1).put("A5" -> 5)

  test("LARGE returns the k-th largest") {
    assertEquals(nums.evaluateFormula("=LARGE(A1:A5,1)"), Right(CellValue.Number(BigDecimal(5))))
    assertEquals(nums.evaluateFormula("=LARGE(A1:A5,2)"), Right(CellValue.Number(BigDecimal(4))))
  }

  test("SMALL returns the k-th smallest") {
    assertEquals(nums.evaluateFormula("=SMALL(A1:A5,1)"), Right(CellValue.Number(BigDecimal(1))))
    assertEquals(nums.evaluateFormula("=SMALL(A1:A5,3)"), Right(CellValue.Number(BigDecimal(3))))
  }

  test("RANK descending (default) and ascending") {
    assertEquals(nums.evaluateFormula("=RANK(4,A1:A5)"), Right(CellValue.Number(BigDecimal(2))))
    assertEquals(nums.evaluateFormula("=RANK(4,A1:A5,1)"), Right(CellValue.Number(BigDecimal(4))))
  }
