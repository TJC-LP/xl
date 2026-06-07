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
