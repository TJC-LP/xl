package com.tjclp.xl.formula

import munit.FunSuite

import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.eval.SheetEvaluator.*
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.syntax.*
import com.tjclp.xl.workbooks.Workbook

/**
 * GH-337/GH-344: aggregate error policy.
 *
 * Elementwise error carriage (GH-337) would silently CHANGE answers if aggregates kept skipping
 * non-numeric elements: an error element coercing to 0 produces a wrong number. GH-337 made those
 * consumptions loud Lefts; GH-344 completes the parity — an error element consumed by an aggregate
 * surfaces as the aggregate's Excel error VALUE (`Right(CellValue.Error(code))` at the boundary,
 * catchable by IFERROR), except COUNT, which genuinely counts numbers only. Raw-RANGE arguments
 * propagate identically (GH-344 item 6 flipped the pre-existing skip/coerce leniency): the
 * effective value of each cell is resolved first, THEN policed by the aggregator's policy.
 */
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial"))
class AggregateErrorPolicySpec extends FunSuite:

  private def num(n: Int): CellValue = CellValue.Number(BigDecimal(n))

  // A1=1, A2=#DIV/0!, A3=2 — the GH-337 issue-cited data shape.
  private val withError = Sheet("Test")
    .put(ref"A1", num(1))
    .put(ref"A2", CellValue.Error(CellError.Div0))
    .put(ref"A3", num(2))

  private val twoWithError = Sheet("Test")
    .put(ref"A1", num(3))
    .put(ref"A2", CellValue.Error(CellError.Div0))

  /** GH-344: the aggregate's failure IS the element's Excel error VALUE at the boundary. */
  private def assertErrorValue(sheet: Sheet, formula: String, code: CellError)(implicit
    loc: munit.Location
  ): Unit =
    assertEquals(sheet.evaluateFormula(formula), Right(CellValue.Error(code)), formula)

  // ===== The GH-337 headline: SUMPRODUCT must not treat error elements as 0 =====

  test("GH-344: =SUMPRODUCT((A1:A3<5)*1) with an error cell is #DIV/0!") {
    assertErrorValue(withError, "=SUMPRODUCT((A1:A3<5)*1)", CellError.Div0)
  }

  test("GH-344: SUMPRODUCT scalar (1x1) expression arm is guarded too") {
    assertErrorValue(twoWithError, "=SUMPRODUCT((A2:A2)*1)", CellError.Div0)
  }

  // ===== Variadic aggregates over expression arrays: the error value propagates =====

  test("GH-344: SUM/MIN/MAX/AVERAGE over an error-bearing expression array return the error") {
    assertErrorValue(twoWithError, "=SUM((A1:A2)*1)", CellError.Div0)
    assertErrorValue(twoWithError, "=MIN((A1:A2)*1)", CellError.Div0)
    assertErrorValue(twoWithError, "=MAX((A1:A2)*1)", CellError.Div0)
    assertErrorValue(twoWithError, "=AVERAGE((A1:A2)*1)", CellError.Div0)
  }

  test("GH-344: SUM over a mixed IF selection that SELECTS an error propagates it") {
    // GH-339 companion: the error is selected at the FALSE position, so the aggregate sees it.
    val mixed = Sheet("Test").put(ref"A1", num(5)).put(ref"A2", num(-3))
    assertErrorValue(mixed, "=SUM(IF(A1:A2>0,A1:A2,1/0))", CellError.Div0)
  }

  test("GH-344: =SUM(IFS(A1:A2>100,1)) with every element #N/A is #N/A") {
    val sheet = Sheet("Test").put(ref"A1", num(5)).put(ref"A2", num(-1))
    assertErrorValue(sheet, "=SUM(IFS(A1:A2>100,1))", CellError.NA)
  }

  // ===== COUNT-family parity (unchanged pins) =====

  test("GH-337: =COUNT((A1:A2)*1) with an error element counts numeric elements only") {
    assertEquals(twoWithError.evaluateFormula("=COUNT((A1:A2)*1)"), Right(num(1)))
  }

  test("GH-337: COUNTA counts error elements, COUNTBLANK does not") {
    assertEquals(twoWithError.evaluateFormula("=COUNTA((A1:A2)*1)"), Right(num(2)))
    assertEquals(twoWithError.evaluateFormula("=COUNTBLANK((A1:A2)*1)"), Right(num(0)))
  }

  // ===== The propagated errors stay IFERROR-catchable (unchanged pins) =====

  test("GH-337: =IFERROR(SUM((A1:A2)*1),0) catches the aggregate's error") {
    assertEquals(twoWithError.evaluateFormula("=IFERROR(SUM((A1:A2)*1),0)"), Right(num(0)))
    assertEquals(
      withError.evaluateFormula("=IFERROR(SUMPRODUCT((A1:A3<5)*1),-1)"),
      Right(num(-1))
    )
  }

  // ===== Raw-range boundary: GH-344 item 6 flips the leniency to propagation =====

  test("GH-344: raw-range =SUM(A1:A2) with an error cell propagates it (was a skip)") {
    assertErrorValue(twoWithError, "=SUM(A1:A2)", CellError.Div0)
  }

  test("GH-344: raw-range SUMPRODUCT propagates error cells (was coerce-to-0)") {
    val sheet = Sheet("Test")
      .put(ref"A1", num(2))
      .put(ref"A2", CellValue.Error(CellError.Div0))
      .put(ref"B1", num(3))
      .put(ref"B2", num(4))
    assertErrorValue(sheet, "=SUMPRODUCT(A1:A2,B1:B2)", CellError.Div0)
  }

  // ===== GH-630: an error VALUE as a direct argument — Excel's COUNT/COUNTA/COUNTBLANK rules =====
  // Confirmed against LibreOffice: COUNT(#REF!)=0, COUNT(1,#N/A,2)=2, COUNTA(#REF!)=1,
  // COUNTA(1,#N/A,2)=3, COUNT(1/0)=0, COUNTA(1/0)=1, COUNTBLANK(#REF!) is an error, SUM(1,#N/A)=#N/A.

  private val withNa = Sheet("Test")
    .put(ref"A1", CellValue.Formula("NA()", Some(CellValue.Error(CellError.NA))))
    .put(ref"A3", num(7))

  test("GH-630: COUNT ignores error arguments — literals, computed errors and error cells alike") {
    assertEquals(withError.evaluateFormula("=COUNT(#REF!)"), Right(num(0)))
    assertEquals(withError.evaluateFormula("=COUNT(1,#N/A,2)"), Right(num(2)))
    assertEquals(withError.evaluateFormula("=COUNT(1/0)"), Right(num(0)))
    assertEquals(withError.evaluateFormula("=COUNT(A2)"), Right(num(0)))
    assertEquals(withError.evaluateFormula("=COUNT(A1:A3)"), Right(num(2)))
    assertEquals(withNa.evaluateFormula("=COUNT(A1:A3)"), Right(num(1)))
    assertEquals(withNa.evaluateFormula("=COUNT(A1)"), Right(num(0)))
  }

  test("GH-630: COUNTA counts error arguments as non-empty") {
    assertEquals(withError.evaluateFormula("=COUNTA(#REF!)"), Right(num(1)))
    assertEquals(withError.evaluateFormula("=COUNTA(1,#N/A,2)"), Right(num(3)))
    assertEquals(withError.evaluateFormula("=COUNTA(1/0)"), Right(num(1)))
    assertEquals(withNa.evaluateFormula("=COUNTA(A1:A3)"), Right(num(2)))
    assertEquals(withNa.evaluateFormula("=COUNTA(A1)"), Right(num(1)))
  }

  test("GH-630: COUNTBLANK neither counts nor swallows an error; its range argument must exist") {
    assertEquals(withNa.evaluateFormula("=COUNTBLANK(A1:A3)"), Right(num(1)))
    assertEquals(
      withError.evaluateFormula("=COUNTBLANK(#REF!)"),
      Right(CellValue.Error(CellError.Ref))
    )
  }

  test("GH-630: every other aggregate still propagates an error argument") {
    assertErrorValue(withError, "=SUM(1,#N/A)", CellError.NA)
    assertErrorValue(withError, "=SUM(1,1/0)", CellError.Div0)
    assertErrorValue(withError, "=AVERAGE(#REF!,1)", CellError.Ref)
    assertErrorValue(withError, "=MAX(1,#NUM!)", CellError.Num)
    assertErrorValue(withNa, "=MAX(A1:A3)", CellError.NA)
    assertErrorValue(withError, "=MIN(#SPILL!)", CellError.Spill)
  }

  test("GH-630: a host failure inside COUNT stays loud — only Excel error VALUES are skipped") {
    val circular = Sheet("Test").put(ref"B1", CellValue.Formula("COUNT(B1)", None))
    assert(circular.evaluateCell(ref"B1").isLeft)
  }

  test("GH-630: a defined name bound to an error literal is triaged like the literal itself") {
    // LibreOffice: COUNT(bad)=0, COUNTA(bad)=1, SUM(bad)=#N/A, COUNT(bad,1)=1 with bad = #N/A
    val named = Workbook(withNa).withDefinedName("bad", "#N/A")
    val sheet = named.sheets.headOption.getOrElse(fail("sheet"))
    def evalNamed(formula: String): Either[?, CellValue] =
      sheet.evaluateFormula(formula, workbook = Some(named))
    assertEquals(evalNamed("=COUNT(bad)"), Right(num(0)))
    assertEquals(evalNamed("=COUNTA(bad)"), Right(num(1)))
    assertEquals(evalNamed("=COUNT(bad,1)"), Right(num(1)))
    assertEquals(evalNamed("=SUM(bad)"), Right(CellValue.Error(CellError.NA)))
    assertEquals(evalNamed("=MAX(bad,1)"), Right(CellValue.Error(CellError.NA)))
  }

  test("GH-630: the typed Aggregate node triages an error slot exactly like the Call form") {
    val evaluator = Evaluator.instance
    def node(name: String, location: TExpr.RangeLocation): Either[?, Any] =
      evaluator.eval(TExpr.Aggregate(name, location), withError)
    val error = TExpr.RangeLocation.Error(CellError.Ref)
    assertEquals(node("COUNT", error), Right(BigDecimal(0)))
    assertEquals(node("COUNTA", error), Right(BigDecimal(1)))
    assert(node("SUM", error).isLeft)
    assert(node("COUNTBLANK", error).isLeft)
    // and, through a workbook, the same for a name bound to #N/A
    val named = Workbook(withNa).withDefinedName("bad", "#N/A")
    val sheet = named.sheets.headOption.getOrElse(fail("sheet"))
    val bad = TExpr.RangeLocation.Name("bad", None)
    assertEquals(
      evaluator.eval(TExpr.Aggregate("COUNT", bad), sheet, workbook = Some(named)),
      Right(BigDecimal(0))
    )
    assertEquals(
      evaluator.eval(TExpr.Aggregate("COUNTA", bad), sheet, workbook = Some(named)),
      Right(BigDecimal(1))
    )
  }
