package com.tjclp.xl.formula

import munit.FunSuite

import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.eval.SheetEvaluator.*
import com.tjclp.xl.formula.eval.WorkbookEvaluator.*
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.syntax.*
import com.tjclp.xl.workbooks.Workbook
import com.tjclp.xl.SheetName

/**
 * GH-344: end-to-end pins for Excel error VALUES as first-class results, exact codes.
 *
 * Errors are constructed via error cells or /0 shapes (the parser has no error-literal tokens).
 * Each section pins one family from the #344 items 1-6 design: scalar channels (item 2), aggregates
 * (items 1/6), logicals (item 3), error-code fidelity and #N/A padding (item 4), TRUE/FALSE text
 * (item 5), and the recalculate() contract.
 */
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial"))
class ErrorValueSemanticsSpec extends FunSuite:

  private def num(n: Int): CellValue = CellValue.Number(BigDecimal(n))
  private val div0 = CellValue.Error(CellError.Div0)
  private val refErr = CellValue.Error(CellError.Ref)
  private val na = CellValue.Error(CellError.NA)

  private def err(e: CellError): Either[Nothing, CellValue] = Right(CellValue.Error(e))

  // ===== Item 2: scalar channels =====

  test("GH-344: =1/0 is #DIV/0! and =IFERROR(1/0,42) still catches it") {
    val sheet = Sheet("Test")
    assertEquals(sheet.evaluateFormula("=1/0"), err(CellError.Div0))
    assertEquals(sheet.evaluateFormula("=IFERROR(1/0,42)"), Right(num(42)))
  }

  test("GH-344: scalar comparison and equality against an error cell return the error") {
    val sheet = Sheet("Test").put(ref"A1", num(1)).put(ref"B1", refErr)
    assertEquals(sheet.evaluateFormula("=A1<B1"), err(CellError.Ref))
    assertEquals(sheet.evaluateFormula("=1=B1"), err(CellError.Ref))
    assertEquals(sheet.evaluateFormula("=B1=B1"), err(CellError.Ref))
  }

  test("GH-344: typed argument positions absorb the error with its code") {
    val sheet = Sheet("Test").put(ref"A1", na)
    assertEquals(sheet.evaluateFormula("=A1+1"), err(CellError.NA))
    assertEquals(sheet.evaluateFormula("=ABS(A1)"), err(CellError.NA))
    assertEquals(sheet.evaluateFormula("=YEAR(A1)"), err(CellError.NA))
    assertEquals(sheet.evaluateFormula("=LEFT(\"hey\",A1)"), err(CellError.NA))
  }

  test("GH-344: concatenation propagates an error operand instead of stringifying it") {
    val sheet = Sheet("Test").put(ref"A1", div0)
    assertEquals(sheet.evaluateFormula("=\"x\"&A1"), err(CellError.Div0))
    assertEquals(sheet.evaluateFormula("=A1&\"x\""), err(CellError.Div0))
    // An Any-typed call result carrying the error propagates too (the concatText guard)
    assertEquals(sheet.evaluateFormula("=\"x\"&IF(TRUE,A1,1)"), err(CellError.Div0))
  }

  test("GH-344: SWITCH propagates an error target/case with the RIGHT code (was #N/A)") {
    val sheet = Sheet("Test").put(ref"A1", div0)
    assertEquals(sheet.evaluateFormula("=SWITCH(A1,1,\"one\")"), err(CellError.Div0))
    assertEquals(sheet.evaluateFormula("=SWITCH(1,A1,\"x\",\"default\")"), err(CellError.Div0))
    // No-match without error stays #N/A
    assertEquals(
      Sheet("Test").put(ref"A1", num(9)).evaluateFormula("=SWITCH(A1,1,\"one\")"),
      err(CellError.NA)
    )
  }

  test("GH-344: LET binds error values as VALUES (Excel-exact laziness of consumption)") {
    val sheet = Sheet("Test")
    assertEquals(sheet.evaluateFormula("=LET(x,1/0,x)"), err(CellError.Div0))
    assertEquals(sheet.evaluateFormula("=LET(x,1/0,IFERROR(x,5))"), Right(num(5)))
    // An unused error binding does not poison the body (was a loud Left naming the binding)
    assertEquals(sheet.evaluateFormula("=LET(bad,1/0,42)"), Right(num(42)))
    // Host failures in a binding stay loud (missing workbook context for a cross-sheet ref)
    assert(sheet.evaluateFormula("=LET(x,Missing!A1,42)").isLeft)
  }

  test("GH-344: an uncached formula precedent that errors delivers its error value") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Formula("=1/0", None))
      .put(ref"B1", num(2))
    assertEquals(sheet.evaluateFormula("=A1+B1"), err(CellError.Div0))
    assertEquals(sheet.evaluateCell(ref"A1"), err(CellError.Div0))
  }

  test("GH-344: cross-sheet reads cascade error values cell-mediated") {
    val src = Sheet(SheetName.unsafe("Src")).put(ref"A1", CellValue.Formula("=1/0", None))
    val dst = Sheet(SheetName.unsafe("Dst"))
    val wb = Workbook(src, dst)
    assertEquals(
      wb.evaluateFormula("=Src!A1", SheetName.unsafe("Dst")),
      err(CellError.Div0)
    )
    assertEquals(
      wb.evaluateFormula("=Src!A1+1", SheetName.unsafe("Dst")),
      err(CellError.Div0)
    )
  }

  test("GH-344: scalar entry collapses a range to top-left — a carried error there propagates") {
    val errTopLeft = Sheet("Test").put(ref"A1", div0).put(ref"A2", num(5))
    assertEquals(errTopLeft.evaluateFormula("=A1:A2*2"), err(CellError.Div0))
    // An error NOT at top-left stays invisible to the collapsed scalar (implicit intersection)
    val errBelow = Sheet("Test").put(ref"A1", num(5)).put(ref"A2", div0)
    assertEquals(errBelow.evaluateFormula("=A1:A2*2"), Right(num(10)))
  }

  test("GH-344: evaluateArrayFormula spills a 1x1 error value at the origin") {
    val sheet = Sheet("Test")
    val result = sheet.evaluateArrayFormula("=1/0", ref"C1")
    assert(result.isRight, s"expected Right, got $result")
    val (updated, spill) = result.toOption.get
    assertEquals(spill.height, 1)
    assertEquals(spill.width, 1)
    assertEquals(updated(ref"C1").value, div0)
  }

  test("GH-344: dependency-ordered evaluation threads error cells instead of aborting") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Formula("=1/0", None))
      .put(ref"B1", CellValue.Formula("=A1+1", None))
      .put(ref"C1", CellValue.Formula("=IFERROR(B1,7)", None))
    val result = sheet.evaluateWithDependencyCheck()
    assert(result.isRight, s"expected Right, got $result")
    val values = result.toOption.get
    assertEquals(values.get(ref"A1"), Some(div0))
    assertEquals(values.get(ref"B1"), Some(div0))
    assertEquals(values.get(ref"C1"), Some(num(7)))
  }

  test("GH-344: ISERR vs ISERROR distinguish #N/A arriving through the Left channel") {
    val sheet = Sheet("Test").put(ref"A1", na)
    assertEquals(sheet.evaluateFormula("=ISERROR(1<A1)"), Right(CellValue.Bool(true)))
    assertEquals(sheet.evaluateFormula("=ISERR(1<A1)"), Right(CellValue.Bool(false)))
  }

  // ===== Items 1+6: aggregates propagate error VALUES; COUNT-family policies pinned =====

  test("GH-344: raw-range aggregates propagate an error cell as the error VALUE") {
    val sheet = Sheet("Test")
      .put(ref"A1", num(1))
      .put(ref"A2", div0)
      .put(ref"A3", num(2))
    assertEquals(sheet.evaluateFormula("=SUM(A1:A3)"), err(CellError.Div0))
    assertEquals(sheet.evaluateFormula("=MIN(A1:A3)"), err(CellError.Div0))
    assertEquals(sheet.evaluateFormula("=MAX(A1:A3)"), err(CellError.Div0))
    assertEquals(sheet.evaluateFormula("=AVERAGE(A1:A3)"), err(CellError.Div0))
    assertEquals(sheet.evaluateFormula("=MEDIAN(A1:A3)"), err(CellError.Div0))
    // IFERROR still catches the aggregate's error
    assertEquals(sheet.evaluateFormula("=IFERROR(SUM(A1:A3),0)"), Right(num(0)))
  }

  test("GH-344: COUNT skips error cells; COUNTA counts them; COUNTBLANK does not") {
    val sheet = Sheet("Test")
      .put(ref"A1", num(1))
      .put(ref"A2", div0)
      .put(ref"A3", num(2))
    assertEquals(sheet.evaluateFormula("=COUNT(A1:A3)"), Right(num(2)))
    assertEquals(sheet.evaluateFormula("=COUNTA(A1:A3)"), Right(num(3)))
    assertEquals(sheet.evaluateFormula("=COUNTBLANK(A1:A3)"), Right(num(0)))
  }

  test("GH-344: an UNCACHED formula cell that errors no longer vanishes from a raw range") {
    // The verified silent swallow: pre-GH-344 the uncached =1/0 recursed to an error and
    // mapped to None — SUM produced a wrong number instead of the error.
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Formula("=1/0", None))
      .put(ref"A2", num(2))
    assertEquals(sheet.evaluateFormula("=SUM(A1:A2)"), err(CellError.Div0))
    assertEquals(sheet.evaluateFormula("=COUNT(A1:A2)"), Right(num(1)))
  }

  test("GH-344: order statistics over an error-bearing range propagate") {
    val sheet = Sheet("Test")
      .put(ref"A1", num(3))
      .put(ref"A2", refErr)
      .put(ref"A3", num(1))
    assertEquals(sheet.evaluateFormula("=LARGE(A1:A3,1)"), err(CellError.Ref))
    assertEquals(sheet.evaluateFormula("=SMALL(A1:A3,1)"), err(CellError.Ref))
    assertEquals(sheet.evaluateFormula("=RANK(3,A1:A3)"), err(CellError.Ref))
    assertEquals(sheet.evaluateFormula("=PERCENTILE(A1:A3,0.5)"), err(CellError.Ref))
    assertEquals(sheet.evaluateFormula("=QUARTILE(A1:A3,2)"), err(CellError.Ref))
  }

  test("GH-344: SUMPRODUCT raw ranges propagate error cells (no longer coerce to 0)") {
    val sheet = Sheet("Test")
      .put(ref"A1", num(2))
      .put(ref"A2", div0)
      .put(ref"B1", num(3))
      .put(ref"B2", num(4))
    assertEquals(sheet.evaluateFormula("=SUMPRODUCT(A1:A2,B1:B2)"), err(CellError.Div0))
  }

  test("GH-344: expression-array aggregates surface the element's error VALUE") {
    val sheet = Sheet("Test")
      .put(ref"A1", num(3))
      .put(ref"A2", div0)
    assertEquals(sheet.evaluateFormula("=SUM((A1:A2)*1)"), err(CellError.Div0))
    assertEquals(sheet.evaluateFormula("=SUMPRODUCT((A1:A2)*1)"), err(CellError.Div0))
    assertEquals(sheet.evaluateFormula("=COUNT((A1:A2)*1)"), Right(num(1)))
  }
