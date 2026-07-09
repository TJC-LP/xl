package com.tjclp.xl.formula

import munit.FunSuite

import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.eval.SheetEvaluator.*
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.syntax.*

/**
 * GH-337: aggregate error policy over expression arrays.
 *
 * Elementwise error carriage (GH-337) would silently CHANGE answers if aggregates kept skipping
 * non-numeric elements: pre-GH-337, =SUMPRODUCT((A1:A3<5)*1) with an error cell failed loudly; with
 * carriage but no guard, the error element would coerce to 0 and produce a wrong number. These
 * guards make the exact issue-cited case loud again, matching Excel: an error element consumed by
 * an aggregate surfaces as the aggregate's failure (catchable by IFERROR), except COUNT, which
 * genuinely counts numbers only.
 *
 * Raw-RANGE arguments (=SUM(A1:A2) over cells) keep their pre-existing skip/coerce semantics —
 * pinned below as documentation of that boundary.
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

  private def assertLoud(sheet: Sheet, formula: String, fnName: String, code: String)(implicit
    loc: munit.Location
  ): Unit =
    val result = sheet.evaluateFormula(formula)
    assert(result.isLeft, s"$formula: expected loud Left, got $result")
    val message = result.left.toOption.get.message
    assert(message.contains(code), s"$formula: expected '$code' in message, got: $message")
    assert(message.contains(fnName), s"$formula: expected '$fnName' in message, got: $message")

  // ===== The GH-337 headline: SUMPRODUCT must not treat error elements as 0 =====

  test("GH-337: =SUMPRODUCT((A1:A3<5)*1) with an error cell fails loudly naming #DIV/0!") {
    assertLoud(withError, "=SUMPRODUCT((A1:A3<5)*1)", "SUMPRODUCT", "#DIV/0!")
  }

  test("GH-337: SUMPRODUCT scalar (1x1) expression arm is guarded too") {
    assertLoud(twoWithError, "=SUMPRODUCT((A2:A2)*1)", "SUMPRODUCT", "#DIV/0!")
  }

  // ===== Variadic aggregates over expression arrays: loud by default =====

  test("GH-337: SUM/MIN/MAX/AVERAGE over an error-bearing expression array fail loudly") {
    assertLoud(twoWithError, "=SUM((A1:A2)*1)", "SUM", "#DIV/0!")
    assertLoud(twoWithError, "=MIN((A1:A2)*1)", "MIN", "#DIV/0!")
    assertLoud(twoWithError, "=MAX((A1:A2)*1)", "MAX", "#DIV/0!")
    assertLoud(twoWithError, "=AVERAGE((A1:A2)*1)", "AVERAGE", "#DIV/0!")
  }

  test("GH-337: SUM over a mixed IF selection that SELECTS an error is loud") {
    // GH-339 companion: the error is selected at the FALSE position, so the aggregate sees it.
    val mixed = Sheet("Test").put(ref"A1", num(5)).put(ref"A2", num(-3))
    assertLoud(mixed, "=SUM(IF(A1:A2>0,A1:A2,1/0))", "SUM", "#DIV/0!")
  }

  test("GH-337: =SUM(IFS(A1:A2>100,1)) with every element #N/A is loud (behavior flip)") {
    // Pre-GH-337 the all-#N/A IFS array summed to 0 by skipping; error elements now propagate.
    val sheet = Sheet("Test").put(ref"A1", num(5)).put(ref"A2", num(-1))
    assertLoud(sheet, "=SUM(IFS(A1:A2>100,1))", "SUM", "#N/A")
  }

  // ===== COUNT-family parity =====

  test("GH-337: =COUNT((A1:A2)*1) with an error element counts numeric elements only") {
    assertEquals(twoWithError.evaluateFormula("=COUNT((A1:A2)*1)"), Right(num(1)))
  }

  test("GH-337: COUNTA counts error elements, COUNTBLANK does not") {
    assertEquals(twoWithError.evaluateFormula("=COUNTA((A1:A2)*1)"), Right(num(2)))
    assertEquals(twoWithError.evaluateFormula("=COUNTBLANK((A1:A2)*1)"), Right(num(0)))
  }

  // ===== The loud Lefts stay IFERROR-catchable =====

  test("GH-337: =IFERROR(SUM((A1:A2)*1),0) catches the loud aggregate failure") {
    assertEquals(twoWithError.evaluateFormula("=IFERROR(SUM((A1:A2)*1),0)"), Right(num(0)))
    assertEquals(
      withError.evaluateFormula("=IFERROR(SUMPRODUCT((A1:A3<5)*1),-1)"),
      Right(num(-1))
    )
  }

  // ===== Raw-range boundary: unchanged, pinned as documentation =====

  test("GH-337: raw-range =SUM(A1:A2) with an error cell still skips it") {
    assertEquals(twoWithError.evaluateFormula("=SUM(A1:A2)"), Right(num(3)))
  }

  test("GH-337: raw-range SUMPRODUCT still coerces error cells to 0 (unchanged boundary)") {
    val sheet = Sheet("Test")
      .put(ref"A1", num(2))
      .put(ref"A2", CellValue.Error(CellError.Div0))
      .put(ref"B1", num(3))
      .put(ref"B2", num(4))
    assertEquals(sheet.evaluateFormula("=SUMPRODUCT(A1:A2,B1:B2)"), Right(num(6)))
  }
