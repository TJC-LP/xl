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

  // ===== Item 6 (SUMIF family): matched value cells propagate; error criteria never match =====

  test("GH-344: SUMIF/AVERAGEIF propagate an error in a MATCHED value cell") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Text("x"))
      .put(ref"A2", CellValue.Text("x"))
      .put(ref"B1", div0)
      .put(ref"B2", num(2))
    assertEquals(sheet.evaluateFormula("=SUMIF(A1:A2,\"x\",B1:B2)"), err(CellError.Div0))
    assertEquals(sheet.evaluateFormula("=AVERAGEIF(A1:A2,\"x\",B1:B2)"), err(CellError.Div0))
  }

  test("GH-344: an error value cell that is NOT matched never propagates") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Text("x"))
      .put(ref"A2", CellValue.Text("y"))
      .put(ref"B1", num(1))
      .put(ref"B2", div0)
    assertEquals(sheet.evaluateFormula("=SUMIF(A1:A2,\"x\",B1:B2)"), Right(num(1)))
  }

  test("GH-344: error cells in a CRITERIA range simply never match (Excel)") {
    val sheet = Sheet("Test")
      .put(ref"A1", refErr)
      .put(ref"A2", CellValue.Text("x"))
      .put(ref"B1", num(1))
      .put(ref"B2", num(2))
    assertEquals(sheet.evaluateFormula("=SUMIF(A1:A2,\"x\",B1:B2)"), Right(num(2)))
    assertEquals(sheet.evaluateFormula("=COUNTIF(A1:A2,\"x\")"), Right(num(1)))
  }

  test("GH-344: SUMIFS/MAXIFS/MINIFS propagate an error in a MATCHED value cell") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Text("x"))
      .put(ref"A2", CellValue.Text("x"))
      .put(ref"B1", na)
      .put(ref"B2", num(2))
    assertEquals(sheet.evaluateFormula("=SUMIFS(B1:B2,A1:A2,\"x\")"), err(CellError.NA))
    assertEquals(sheet.evaluateFormula("=MAXIFS(B1:B2,A1:A2,\"x\")"), err(CellError.NA))
    assertEquals(sheet.evaluateFormula("=MINIFS(B1:B2,A1:A2,\"x\")"), err(CellError.NA))
    assertEquals(sheet.evaluateFormula("=AVERAGEIFS(B1:B2,A1:A2,\"x\")"), err(CellError.NA))
  }

  // ===== Item 3: AND/OR/NOT propagate error values; AND/OR evaluate eagerly =====

  test("GH-344: AND/OR over an error operand return the error VALUE") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Bool(true))
      .put(ref"B1", refErr)
    assertEquals(sheet.evaluateFormula("=AND(A1,B1)"), err(CellError.Ref))
    assertEquals(sheet.evaluateFormula("=OR(A1,B1)"), err(CellError.Ref))
    assertEquals(sheet.evaluateFormula("=NOT(B1)"), err(CellError.Ref))
  }

  test("GH-344: AND/OR evaluate EVERY argument (Excel does not short-circuit logicals)") {
    val sheet = Sheet("Test").put(ref"B1", na)
    assertEquals(sheet.evaluateFormula("=AND(FALSE,1/0)"), err(CellError.Div0))
    assertEquals(sheet.evaluateFormula("=OR(TRUE,B1)"), err(CellError.NA))
    // Decisive values still fold when no argument fails
    assertEquals(sheet.evaluateFormula("=AND(FALSE,TRUE)"), Right(CellValue.Bool(false)))
    assertEquals(sheet.evaluateFormula("=OR(TRUE,FALSE)"), Right(CellValue.Bool(true)))
  }

  test("GH-344: first failing argument wins in argument order") {
    val sheet = Sheet("Test")
      .put(ref"A1", refErr)
      .put(ref"B1", div0)
    assertEquals(sheet.evaluateFormula("=AND(A1,B1)"), err(CellError.Ref))
    assertEquals(sheet.evaluateFormula("=AND(B1,A1)"), err(CellError.Div0))
  }

  test("GH-344: an error ELEMENT in an AND/OR array argument propagates as the error VALUE") {
    val sheet = Sheet("Test")
      .put(ref"A1", num(5))
      .put(ref"A2", div0)
    assertEquals(sheet.evaluateFormula("=AND(A1:A2>0)"), err(CellError.Div0))
    assertEquals(sheet.evaluateFormula("=OR(A1:A2>0)"), err(CellError.Div0))
    // IFERROR catches
    assertEquals(sheet.evaluateFormula("=IFERROR(AND(A1:A2>0),42)"), Right(num(42)))
  }

  test("GH-344: NOT over an array carries error elements positionally") {
    val sheet = Sheet("Test")
      .put(ref"A1", num(5))
      .put(ref"A2", div0)
    val result = sheet.evaluateArrayFormula("=NOT(A1:A2>0)", ref"C1")
    assert(result.isRight, s"expected Right, got $result")
    val (updated, _) = result.toOption.get
    assertEquals(updated(ref"C1").value, CellValue.Bool(false))
    assertEquals(updated(ref"C2").value, div0)
  }

  // ===== Item 4: #NUM!/code classification at source =====

  test("GH-344: pow overflow is #NUM! (scalar and element)") {
    // BigDecimal computes =2^1024 exactly (unlike Excel's doubles); the #NUM! arm fires on a
    // genuine ArithmeticException — a scale/magnitude overflow of the exact representation.
    val sheet = Sheet("Test").put(ref"A1", num(2)).put(ref"A2", num(3))
    assertEquals(sheet.evaluateFormula("=2^1.5E9"), err(CellError.Num))
    // Elementwise: the overflow element demotes to a #NUM! ELEMENT via the toCellError
    // identity arm; the scalar entry collapses to the top-left #NUM!.
    assertEquals(sheet.evaluateFormula("=A1:A2^1.5E9"), err(CellError.Num))
  }

  test("GH-344: math domain failures carry their Excel codes") {
    val sheet = Sheet("Test")
    assertEquals(sheet.evaluateFormula("=SQRT(-1)"), err(CellError.Num))
    assertEquals(sheet.evaluateFormula("=LOG(0)"), err(CellError.Num))
    assertEquals(sheet.evaluateFormula("=LOG(8,0)"), err(CellError.Num))
    assertEquals(sheet.evaluateFormula("=LOG(8,1)"), err(CellError.Div0))
    assertEquals(sheet.evaluateFormula("=LN(0)"), err(CellError.Num))
    assertEquals(sheet.evaluateFormula("=MOD(5,0)"), err(CellError.Div0))
  }

  test("GH-344: order-statistic domain failures carry their Excel codes") {
    val sheet = Sheet("Test").put(ref"A1", num(1)).put(ref"A2", num(2))
    assertEquals(sheet.evaluateFormula("=LARGE(A1:A2,5)"), err(CellError.Num))
    assertEquals(sheet.evaluateFormula("=SMALL(A1:A2,0)"), err(CellError.Num))
    assertEquals(sheet.evaluateFormula("=PERCENTILE(A1:A2,2)"), err(CellError.Num))
    assertEquals(sheet.evaluateFormula("=QUARTILE(A1:A2,9)"), err(CellError.Num))
    assertEquals(sheet.evaluateFormula("=RANK(99,A1:A2)"), err(CellError.NA))
  }

  test("GH-344: RANDBETWEEN with bottom > top is #NUM!") {
    val sheet = Sheet("Test")
    assertEquals(
      sheet
        .evaluateFormula("=RANDBETWEEN(10,1)", Clock.system, com.tjclp.xl.formula.Rng.seeded(1L)),
      err(CellError.Num)
    )
  }

  // ===== Item 4b: #N/A dimension padding in the shared broadcast machinery =====

  test("GH-344: mismatched broadcast dims pad with #N/A elements (Excel)") {
    val sheet = Sheet("Test")
      .put(ref"A1", num(1))
      .put(ref"A2", num(2))
      .put(ref"A3", num(3))
      .put(ref"B1", num(10))
      .put(ref"B2", num(20))
    val result = sheet.evaluateArrayFormula("=A1:A3*B1:B2", ref"D1")
    assert(result.isRight, s"expected Right, got $result")
    val (updated, spill) = result.toOption.get
    assertEquals(spill.height, 3)
    assertEquals(updated(ref"D1").value, num(10))
    assertEquals(updated(ref"D2").value, num(40))
    assertEquals(updated(ref"D3").value, na)
  }

  test("GH-344: dimension mismatch inside an IF pads with #N/A (was 1x1 #VALUE!)") {
    val sheet = Sheet("Test")
      .put(ref"A1", num(1))
      .put(ref"A2", num(2))
      .put(ref"A3", num(3))
      .put(ref"B1", num(10))
      .put(ref"B2", num(20))
    val result = sheet.evaluateArrayFormula("=IF(A1:A3>0,B1:B2,0)", ref"D1")
    assert(result.isRight, s"expected Right, got $result")
    val (updated, spill) = result.toOption.get
    assertEquals(spill.height, 3)
    assertEquals(updated(ref"D1").value, num(10))
    assertEquals(updated(ref"D2").value, num(20))
    assertEquals(updated(ref"D3").value, na)
  }

  // ===== Item 5: TRUE/FALSE text coerces in condition positions =====

  test("GH-344: literal \"TRUE\"/\"FALSE\" text coerces in condition positions") {
    val sheet = Sheet("Test")
    assertEquals(sheet.evaluateFormula("=IF(\"TRUE\",1,2)"), Right(num(1)))
    assertEquals(sheet.evaluateFormula("=IF(\"false\",1,2)"), Right(num(2)))
    assertEquals(sheet.evaluateFormula("=AND(\"TRUE\",\"true\")"), Right(CellValue.Bool(true)))
    assertEquals(sheet.evaluateFormula("=OR(\"FALSE\",\"false\")"), Right(CellValue.Bool(false)))
    assertEquals(sheet.evaluateFormula("=NOT(\"TRUE\")"), Right(CellValue.Bool(false)))
  }

  test("GH-344: recognition is exact — no trim, other text still refuses") {
    val sheet = Sheet("Test")
    assert(sheet.evaluateFormula("=IF(\" TRUE\",1,2)").isLeft, "\" TRUE\" must refuse (no trim)")
    assert(sheet.evaluateFormula("=IF(\"TRUEX\",1,2)").isLeft)
    assert(sheet.evaluateFormula("=IF(\"abc\",1,2)").isLeft)
  }

  test("GH-344: cell-sourced boolean text coerces too (documented micro-divergence)") {
    // Excel coerces only literal text; provenance is invisible at the decode layer, so the
    // direct/bound parity law wins and cell text coerces identically.
    val sheet = Sheet("Test").put(ref"A1", CellValue.Text("TRUE"))
    assertEquals(sheet.evaluateFormula("=IF(A1,1,2)"), Right(num(1)))
    assertEquals(sheet.evaluateFormula("=NOT(A1)"), Right(CellValue.Bool(false)))
  }

  test("GH-344: TRUE/FALSE text elements coerce in array condition positions") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Text("TRUE"))
      .put(ref"A2", CellValue.Text("false"))
    // AND folds the IF-selected text elements {"TRUE","TRUE"} via conditionTruthy
    assertEquals(
      sheet.evaluateFormula("=AND(IF(A1:A2<>0,\"TRUE\",\"FALSE\"))"),
      Right(CellValue.Bool(true))
    )
    // NOT broadcasts over the IF-selected cell text {TRUE, false} elementwise
    val result = sheet.evaluateArrayFormula("=NOT(IF(A1:A2<>\"\",A1:A2,\"FALSE\"))", ref"C1")
    assert(result.isRight, s"expected Right, got $result")
    val (updated, _) = result.toOption.get
    assertEquals(updated(ref"C1").value, CellValue.Bool(false))
    assertEquals(updated(ref"C2").value, CellValue.Bool(true))
  }

  test("GH-344: SUMPRODUCT keeps exact-dimension enforcement as #VALUE! (never padding)") {
    val sheet = Sheet("Test")
      .put(ref"A1", num(1))
      .put(ref"A2", num(2))
      .put(ref"B1", num(10))
      .put(ref"B2", num(20))
      .put(ref"B3", num(30))
    assertEquals(sheet.evaluateFormula("=SUMPRODUCT(A1:A2,B1:B3)"), err(CellError.Value))
  }

  // ===== recalculate() contract: error VALUES are results, not failures =====

  test("GH-344: recalculate() caches error values, isClean holds, excelErrors reports them") {
    val sheet = Sheet(SheetName.unsafe("S"))
      .put(ref"A1", CellValue.Formula("=1/0", None))
      .put(ref"B1", CellValue.Formula("=A1+1", None))
      .put(ref"C1", CellValue.Formula("=IFERROR(B1,7)", None))
    val result = Workbook(sheet).recalculate()
    // Every formula COMPUTED a value (two of them error values) — the run is clean
    assert(result.isClean, s"expected clean, got: ${result.errors.map(_.render)}")
    assertEquals(result.errors, Vector.empty)
    // The error values cache like any computed value (Excel writes them as t="e")
    def cachedValue(r: com.tjclp.xl.addressing.ARef): Option[CellValue] =
      result.workbook("S").toOption.flatMap(_.cells.get(r)).map(_.value).collect {
        case CellValue.Formula(_, Some(v), _) => v
      }
    assertEquals(cachedValue(ref"A1"), Some(div0))
    assertEquals(cachedValue(ref"B1"), Some(div0))
    assertEquals(cachedValue(ref"C1"), Some(num(7)))
    // excelErrors: deterministic (sheet, ref) order, code included
    assertEquals(
      result.excelErrors,
      Vector(
        (SheetName.unsafe("S"), ref"A1", CellError.Div0),
        (SheetName.unsafe("S"), ref"B1", CellError.Div0)
      )
    )
    assertEquals(result.toEither.map(_.sheets.size), Right(1))
  }

  test("GH-344: recalculate() keeps host failures in errors, disjoint from excelErrors") {
    val sheet = Sheet(SheetName.unsafe("S"))
      .put(ref"A1", CellValue.Formula("=NOSUCHFN(1)", None))
      .put(ref"B1", CellValue.Formula("=1/0", None))
    val result = Workbook(sheet).recalculate()
    assert(!result.isClean)
    assertEquals(result.errors.map(_.ref), Vector(ref"A1"))
    assertEquals(
      result.excelErrors,
      Vector((SheetName.unsafe("S"), ref"B1", CellError.Div0))
    )
  }
