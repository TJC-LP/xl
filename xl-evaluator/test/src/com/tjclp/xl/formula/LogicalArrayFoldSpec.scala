package com.tjclp.xl.formula

import munit.FunSuite

import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.eval.SheetEvaluator.*
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.syntax.*

/**
 * GH-338: AND/OR aggregate over array conditions; NOT broadcasts elementwise.
 *
 * Post-GH-333 the logical functions no longer crashed on array conditions, but the Coerced(Bool)
 * scalar boundary collapsed them by top-left implicit intersection — a plausible-but-wrong answer.
 * Excel's AND/OR natively accept array arguments and aggregate across EVERY element (no CSE entry
 * required); NOT broadcasts elementwise like the fixed IF. Condition elements follow the
 * broadcastIf conventions: booleans as-is, numbers zero/non-zero, empty is FALSE; text and error
 * elements refuse with a clean Left.
 */
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial"))
class LogicalArrayFoldSpec extends FunSuite:

  // The GH-338 repro data: mixed signs, truthy element FIRST (collapse used to answer TRUE).
  private val mixedTF = Sheet("Test")
    .put(ref"A1", CellValue.Number(5))
    .put(ref"A2", CellValue.Number(-1))

  // Mirror image: falsy element first (collapse used to answer FALSE for OR).
  private val mixedFT = Sheet("Test")
    .put(ref"A1", CellValue.Number(-1))
    .put(ref"A2", CellValue.Number(5))

  private val allNegative = Sheet("Test")
    .put(ref"A1", CellValue.Number(-5))
    .put(ref"A2", CellValue.Number(-1))

  private def assertScalar(sheet: Sheet, formula: String, expected: CellValue)(implicit
    loc: munit.Location
  ): Unit =
    assertEquals(sheet.evaluateFormula(formula), Right(expected), formula)

  private def assertArrayEntry(sheet: Sheet, formula: String, expected: CellValue)(implicit
    loc: munit.Location
  ): Unit =
    val result = sheet.evaluateArrayFormula(formula, ref"D1")
    assert(result.isRight, s"$formula: expected Right, got $result")
    val (updated, spill) = result.toOption.get
    assertEquals(spill.height, 1, s"$formula spill height")
    assertEquals(spill.width, 1, s"$formula spill width")
    assertEquals(updated(ref"D1").value, expected, formula)

  // ===== The exact GH-338 repro, via BOTH evaluation entries =====

  test("GH-338: AND over a mixed array comparison aggregates to FALSE (array entry)") {
    assertArrayEntry(mixedTF, "=AND(A1:A2>0)", CellValue.Bool(false))
  }

  test("GH-338: OR over a mixed array comparison aggregates to TRUE (array entry)") {
    assertArrayEntry(mixedFT, "=OR(A1:A2>0)", CellValue.Bool(true))
  }

  test("GH-338: AND aggregates via plain scalar evaluateFormula too") {
    assertScalar(mixedTF, "=AND(A1:A2>0)", CellValue.Bool(false))
  }

  test("GH-338: OR aggregates via plain scalar evaluateFormula too") {
    assertScalar(mixedFT, "=OR(A1:A2>0)", CellValue.Bool(true))
  }

  // ===== Element order must not matter =====

  test("GH-338: AND/OR aggregation is order-independent") {
    assertScalar(mixedFT, "=AND(A1:A2>0)", CellValue.Bool(false))
    assertScalar(mixedTF, "=OR(A1:A2>0)", CellValue.Bool(true))
  }

  // ===== Multi-argument mixes (scalars and arrays fold together) =====

  test("GH-338: mixed scalar and array arguments fold together") {
    assertScalar(mixedTF, "=AND(TRUE, A1:A2>0)", CellValue.Bool(false))
    assertScalar(mixedTF, "=AND(A1:A2>-10, TRUE)", CellValue.Bool(true))
    assertScalar(mixedTF, "=OR(FALSE, A1:A2>0)", CellValue.Bool(true))
    // A1=5, A2=-1: no element exceeds 5, and B1 is empty (decodes FALSE)
    assertScalar(mixedTF, "=OR(A1:A2>5, B1)", CellValue.Bool(false))
    val withB1 = mixedTF.put(ref"B1", CellValue.Bool(true))
    assertScalar(withB1, "=OR(A1:A2>5, B1)", CellValue.Bool(true))
  }

  test("GH-338: two array arguments both aggregate") {
    // {T,F} and {T,T}: AND folds both to FALSE && TRUE = FALSE; OR to TRUE
    assertScalar(mixedTF, "=AND(A1:A2>0, A1:A2>-10)", CellValue.Bool(false))
    assertScalar(mixedTF, "=OR(A1:A2>100, A1:A2>0)", CellValue.Bool(true))
  }

  // ===== All-TRUE / all-FALSE arrays =====

  test("GH-338: all-TRUE and all-FALSE arrays") {
    assertScalar(mixedTF, "=AND(A1:A2>-10)", CellValue.Bool(true))
    assertScalar(mixedTF, "=OR(A1:A2>100)", CellValue.Bool(false))
    assertScalar(allNegative, "=AND(A1:A2<0)", CellValue.Bool(true))
    assertScalar(allNegative, "=OR(A1:A2>0)", CellValue.Bool(false))
  }

  // ===== Empty cells =====

  test("GH-338: empty cells in the compared range coerce as 0") {
    val sparse = Sheet("Test").put(ref"A1", CellValue.Number(5)) // A2 empty
    assertScalar(sparse, "=AND(A1:A2>0)", CellValue.Bool(false)) // 0>0 is FALSE
    assertScalar(sparse, "=OR(A1:A2>0)", CellValue.Bool(true))
    assertScalar(sparse, "=AND(A1:A2>=0)", CellValue.Bool(true)) // 0>=0 is TRUE
  }

  test("GH-338: empty condition elements are FALSE (broadcastIf convention)") {
    // IF selects from an empty range, so the condition array AND/OR folds is all Empty
    assertScalar(mixedTF, "=OR(IF(A1:A2>0,C1:C2,C1:C2))", CellValue.Bool(false))
    assertScalar(mixedTF, "=OR(IF(A1:A2>0,C1:C2,C1:C2), TRUE)", CellValue.Bool(true))
    assertScalar(mixedTF, "=AND(IF(A1:A2>0,C1:C2,C1:C2))", CellValue.Bool(false))
  }

  // ===== NOT broadcasts elementwise =====

  test("GH-338: NOT broadcasts elementwise over an array condition (2x1 shape)") {
    val result = mixedTF.evaluateArrayFormula("=NOT(A1:A2>0)", ref"C1")
    assert(result.isRight, s"Expected Right, got $result")
    val (updated, spill) = result.toOption.get
    assertEquals(spill.height, 2)
    assertEquals(spill.width, 1)
    assertEquals(updated(ref"C1").value, CellValue.Bool(false))
    assertEquals(updated(ref"C2").value, CellValue.Bool(true))
  }

  test("GH-338: AND(NOT(...)) and OR(NOT(...)) compose end-to-end") {
    // mixedFT: NOT({F,T}) = {T,F} — collapse used to answer TRUE
    assertScalar(mixedFT, "=AND(NOT(A1:A2>0))", CellValue.Bool(false))
    assertScalar(mixedTF, "=AND(NOT(A1:A2>0))", CellValue.Bool(false))
    assertScalar(mixedTF, "=OR(NOT(A1:A2>0))", CellValue.Bool(true))
    assertScalar(allNegative, "=AND(NOT(A1:A2>0))", CellValue.Bool(true))
    assertScalar(allNegative, "=OR(NOT(A1:A2<0))", CellValue.Bool(false))
  }

  test("GH-338: scalar NOT keeps current behavior (including scalar-entry collapse)") {
    assertScalar(mixedTF, "=NOT(TRUE)", CellValue.Bool(false))
    assertScalar(mixedTF, "=NOT(0)", CellValue.Bool(true))
    assertScalar(mixedTF, "=NOT(A1)", CellValue.Bool(false)) // 5 is truthy
    // Scalar entry collapses the broadcast array to its top-left element (implicit intersection)
    assertScalar(mixedTF, "=NOT(A1:A2>0)", CellValue.Bool(false))
    assertScalar(mixedFT, "=NOT(A1:A2>0)", CellValue.Bool(true))
  }

  // ===== Text in a condition position refuses cleanly (total, no throw) =====

  test("GH-338: text condition elements refuse the AND/OR folds with a clean Left") {
    assert(mixedTF.evaluateFormula("=AND(IF(A1:A2>0,\"x\",\"y\"))").isLeft)
    assert(mixedTF.evaluateFormula("=OR(IF(A1:A2>0,\"x\",\"y\"))").isLeft)
    assert(mixedTF.evaluateArrayFormula("=AND(IF(A1:A2>0,\"x\",\"y\"))", ref"D1").isLeft)
    assert(mixedTF.evaluateFormula("=OR(\"abc\")").isLeft)
  }

  test("GH-344: NOT demotes text elements to #VALUE! elements (Excel broadcast semantics)") {
    // GH-344 supersedes the whole-formula refusal for NOT's array path: like broadcastIf, a
    // text element becomes a #VALUE! OUTPUT element; scalar entry collapses to the top-left.
    assertEquals(
      mixedTF.evaluateFormula("=NOT(IF(A1:A2>0,\"x\",\"y\"))"),
      Right(CellValue.Error(CellError.Value))
    )
  }

  // ===== GH-344: Excel does NOT short-circuit logical functions =====

  test("GH-344: every argument evaluates — a failing later argument surfaces its error") {
    // GH-344 supersedes the GH-338 cross-argument short-circuit pin: Excel evaluates every
    // AND/OR argument, so =AND(1>2, 1/0>0) is #DIV/0! even though the first is decisive.
    assertScalar(mixedTF, "=AND(1>2, 1/0>0)", CellValue.Error(CellError.Div0))
    assertScalar(mixedTF, "=OR(1<2, 1/0>0)", CellValue.Error(CellError.Div0))
    // Decisive folds still answer when every argument evaluates cleanly
    assertScalar(mixedTF, "=AND(1>2, 2>1)", CellValue.Bool(false))
    assertScalar(mixedTF, "=OR(1<2, 2<1)", CellValue.Bool(true))
  }

  test("GH-338: bare-range arguments keep their pre-existing error (out of scope)") {
    // Raw-range coercion parity (Excel skips text/blanks) is documented out of scope in GH-338
    val andResult = mixedTF.evaluateFormula("=AND(A1:A2)")
    assert(andResult.isLeft, s"Expected Left, got $andResult")
    assert(
      andResult.left.toOption.get.message.contains("must be used within a function"),
      s"unexpected error: $andResult"
    )
    assert(mixedTF.evaluateFormula("=OR(A1:A2)").isLeft)
    assert(mixedTF.evaluateFormula("=NOT(A1:A2)").isLeft)
  }

  // ===== IFS condition slots evaluate array-aware like IF =====

  test("GH-338: IFS over an array condition broadcasts elementwise (CSE semantics)") {
    val result = mixedTF.evaluateArrayFormula("=IFS(A1:A2>0, 1, TRUE, 0)", ref"D1")
    assert(result.isRight, s"Expected Right, got $result")
    val (updated, spill) = result.toOption.get
    assertEquals(spill.height, 2)
    assertEquals(updated(ref"D1").value, CellValue.Number(1))
    assertEquals(updated(ref"D2").value, CellValue.Number(0))
  }

  test("GH-338: SUM(IFS(...)) aggregates the broadcast array via scalar entry") {
    // {5>0, -1>0} selects {5, 0} elementwise
    assertScalar(mixedTF, "=SUM(IFS(A1:A2>0, A1:A2, TRUE, 0))", CellValue.Number(5))
  }

  test("GH-338: IFS with no matching pair yields elementwise #N/A") {
    val result = mixedTF.evaluateArrayFormula("=IFS(A1:A2>100, 1)", ref"D1")
    assert(result.isRight, s"Expected Right, got $result")
    val (updated, spill) = result.toOption.get
    assertEquals(spill.height, 2)
    assertEquals(updated(ref"D1").value, CellValue.Error(CellError.NA))
    assertEquals(updated(ref"D2").value, CellValue.Error(CellError.NA))
  }

  test("GH-338: scalar IFS keeps lazy pair selection") {
    // First TRUE condition wins without evaluating later pairs (which would divide by zero)
    assertScalar(mixedTF, "=IFS(TRUE, 42, 1/0>0, 99)", CellValue.Number(42))
    assertScalar(mixedTF, "=IFS(1>2, 1, 2>1, 7)", CellValue.Number(7))
    assertScalar(mixedTF, "=IFS(1>2, 1)", CellValue.Error(CellError.NA))
  }
