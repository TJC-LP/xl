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

  test("PERCENTILE inclusive, with interpolation") {
    // A1:A5 sorted = 1,1,3,4,5
    assertEquals(nums.evaluateFormula("=PERCENTILE(A1:A5,0.5)"), Right(CellValue.Number(BigDecimal(3))))
    assertEquals(nums.evaluateFormula("=PERCENTILE(A1:A5,0)"), Right(CellValue.Number(BigDecimal(1))))
    // A1:A4 sorted = 1,1,3,4; p=0.5 → 1 + 0.5*(3-1) = 2
    assertEquals(nums.evaluateFormula("=PERCENTILE(A1:A4,0.5)"), Right(CellValue.Number(BigDecimal(2))))
  }

  test("QUARTILE maps quart to percentiles") {
    assertEquals(nums.evaluateFormula("=QUARTILE(A1:A5,2)"), Right(CellValue.Number(BigDecimal(3))))
    assertEquals(nums.evaluateFormula("=QUARTILE(A1:A5,4)"), Right(CellValue.Number(BigDecimal(5))))
  }

  test("GH-55: XLOOKUP accepts binary search modes 2 and -2") {
    assertEquals(
      nums.evaluateFormula("=XLOOKUP(4,A1:A5,A1:A5,\"NA\",0,2)"),
      Right(CellValue.Number(BigDecimal(4)))
    )
    assertEquals(
      nums.evaluateFormula("=XLOOKUP(4,A1:A5,A1:A5,\"NA\",0,-2)"),
      Right(CellValue.Number(BigDecimal(4)))
    )
  }

  // GH-122: HLOOKUP (transpose of VLOOKUP).
  private val htable = s
    .put("A1" -> "a")
    .put("B1" -> "b")
    .put("C1" -> "c")
    .put("A2" -> 10)
    .put("B2" -> 20)
    .put("C2" -> 30)
    .put("A3" -> 100)
    .put("B3" -> 200)
    .put("C3" -> 300)

  test("HLOOKUP exact match returns from the given row index") {
    assertEquals(
      htable.evaluateFormula("=HLOOKUP(\"b\",A1:C3,2,FALSE)"),
      Right(CellValue.Number(BigDecimal(20)))
    )
    assertEquals(
      htable.evaluateFormula("=HLOOKUP(\"c\",A1:C3,3,FALSE)"),
      Right(CellValue.Number(BigDecimal(300)))
    )
  }

  test("HLOOKUP approximate (default) finds the largest key <= lookup") {
    val t = s.put("A1" -> 1).put("B1" -> 5).put("C1" -> 10).put("A2" -> "x").put("B2" -> "y").put(
      "C2" -> "z"
    )
    assertEquals(t.evaluateFormula("=HLOOKUP(7,A1:C2,2)"), Right(CellValue.Text("y")))
  }

  // GH-122: MAXIFS / MINIFS (criteria aggregates).
  private val crit = s
    .put("A1" -> 10)
    .put("A2" -> 20)
    .put("A3" -> 30)
    .put("A4" -> 40)
    .put("A5" -> 50)
    .put("B1" -> "x")
    .put("B2" -> "y")
    .put("B3" -> "x")
    .put("B4" -> "y")
    .put("B5" -> "x")

  test("MAXIFS / MINIFS reduce over matching cells") {
    assertEquals(
      crit.evaluateFormula("=MAXIFS(A1:A5,B1:B5,\"x\")"),
      Right(CellValue.Number(BigDecimal(50)))
    )
    assertEquals(
      crit.evaluateFormula("=MINIFS(A1:A5,B1:B5,\"x\")"),
      Right(CellValue.Number(BigDecimal(10)))
    )
  }

  test("MAXIFS returns 0 when nothing matches") {
    assertEquals(
      crit.evaluateFormula("=MAXIFS(A1:A5,B1:B5,\"z\")"),
      Right(CellValue.Number(BigDecimal(0)))
    )
  }
