package com.tjclp.xl.formula

import munit.FunSuite

import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.eval.SheetEvaluator.*
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.syntax.*

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
