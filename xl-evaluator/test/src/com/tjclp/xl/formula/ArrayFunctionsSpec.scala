package com.tjclp.xl.formula

import munit.FunSuite

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.eval.ArrayResult
import com.tjclp.xl.formula.eval.SheetEvaluator.*
import com.tjclp.xl.patch.Patch
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.syntax.*

// Test code uses .get/.head for brevity in assertions
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial", "org.wartremover.warts.IterableOps"))
class ArrayFunctionsSpec extends FunSuite:

  // ========== ArrayResult Tests ==========

  test("ArrayResult.rows and cols for 2x3 array") {
    val arr = ArrayResult(
      Vector(
        Vector(CellValue.Number(1), CellValue.Number(2), CellValue.Number(3)),
        Vector(CellValue.Number(4), CellValue.Number(5), CellValue.Number(6))
      )
    )
    assertEquals(arr.rows, 2)
    assertEquals(arr.cols, 3)
  }

  test("ArrayResult.transpose swaps dimensions") {
    val arr = ArrayResult(
      Vector(
        Vector(CellValue.Number(1), CellValue.Number(2), CellValue.Number(3)),
        Vector(CellValue.Number(4), CellValue.Number(5), CellValue.Number(6))
      )
    )
    val transposed = arr.transpose
    assertEquals(transposed.rows, 3)
    assertEquals(transposed.cols, 2)
    assertEquals(transposed(0, 0), CellValue.Number(1))
    assertEquals(transposed(0, 1), CellValue.Number(4))
    assertEquals(transposed(1, 0), CellValue.Number(2))
    assertEquals(transposed(1, 1), CellValue.Number(5))
    assertEquals(transposed(2, 0), CellValue.Number(3))
    assertEquals(transposed(2, 1), CellValue.Number(6))
  }

  test("ArrayResult.empty is 0x0") {
    val arr = ArrayResult.empty
    assertEquals(arr.rows, 0)
    assertEquals(arr.cols, 0)
    assert(arr.isEmpty)
  }

  test("ArrayResult.single creates 1x1 array") {
    val arr = ArrayResult.single(CellValue.Number(42))
    assertEquals(arr.rows, 1)
    assertEquals(arr.cols, 1)
    assertEquals(arr(0, 0), CellValue.Number(42))
  }

  test("ArrayResult.apply returns Empty for out of bounds") {
    val arr = ArrayResult.single(CellValue.Number(42))
    assertEquals(arr(1, 0), CellValue.Empty)
    assertEquals(arr(0, 1), CellValue.Empty)
    assertEquals(arr(-1, 0), CellValue.Empty)
  }

  // ========== TRANSPOSE Function Tests ==========

  test("TRANSPOSE transposes 2x3 range to 3x2") {
    // Create sheet with 2x3 data:
    // A1=1, B1=2, C1=3
    // A2=4, B2=5, C2=6
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Number(1))
      .put(ref"B1", CellValue.Number(2))
      .put(ref"C1", CellValue.Number(3))
      .put(ref"A2", CellValue.Number(4))
      .put(ref"B2", CellValue.Number(5))
      .put(ref"C2", CellValue.Number(6))

    // Evaluate TRANSPOSE and spill to E1
    val result = sheet.evaluateArrayFormula("=TRANSPOSE(A1:C2)", ref"E1")

    assert(result.isRight, s"Expected Right, got $result")
    val (updatedSheet, affectedRange) = result.toOption.get

    // Should affect E1:F3 (3 rows x 2 cols)
    assertEquals(affectedRange.rowStart.index0, 0) // E1 row
    assertEquals(affectedRange.rowEnd.index0, 2) // F3 row
    assertEquals(affectedRange.colStart.index0, 4) // E column
    assertEquals(affectedRange.colEnd.index0, 5) // F column

    // Verify transposed values
    // Original: [[1,2,3],[4,5,6]]
    // Transposed: [[1,4],[2,5],[3,6]]
    assertEquals(updatedSheet(ref"E1").value, CellValue.Number(1))
    assertEquals(updatedSheet(ref"F1").value, CellValue.Number(4))
    assertEquals(updatedSheet(ref"E2").value, CellValue.Number(2))
    assertEquals(updatedSheet(ref"F2").value, CellValue.Number(5))
    assertEquals(updatedSheet(ref"E3").value, CellValue.Number(3))
    assertEquals(updatedSheet(ref"F3").value, CellValue.Number(6))
  }

  test("TRANSPOSE transposes 1x3 row to 3x1 column") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Number(10))
      .put(ref"B1", CellValue.Number(20))
      .put(ref"C1", CellValue.Number(30))

    val result = sheet.evaluateArrayFormula("=TRANSPOSE(A1:C1)", ref"E1")

    assert(result.isRight)
    val (updatedSheet, affectedRange) = result.toOption.get

    // 1x3 -> 3x1
    assertEquals(affectedRange.height, 3)
    assertEquals(affectedRange.width, 1)

    assertEquals(updatedSheet(ref"E1").value, CellValue.Number(10))
    assertEquals(updatedSheet(ref"E2").value, CellValue.Number(20))
    assertEquals(updatedSheet(ref"E3").value, CellValue.Number(30))
  }

  test("TRANSPOSE preserves text values") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Text("Hello"))
      .put(ref"B1", CellValue.Text("World"))

    val result = sheet.evaluateArrayFormula("=TRANSPOSE(A1:B1)", ref"D1")

    assert(result.isRight)
    val (updatedSheet, _) = result.toOption.get

    assertEquals(updatedSheet(ref"D1").value, CellValue.Text("Hello"))
    assertEquals(updatedSheet(ref"D2").value, CellValue.Text("World"))
  }

  test("TRANSPOSE preserves empty cells") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Number(1))
      // B1 is empty
      .put(ref"C1", CellValue.Number(3))

    val result = sheet.evaluateArrayFormula("=TRANSPOSE(A1:C1)", ref"E1")

    assert(result.isRight)
    val (updatedSheet, _) = result.toOption.get

    assertEquals(updatedSheet(ref"E1").value, CellValue.Number(1))
    assertEquals(updatedSheet(ref"E2").value, CellValue.Empty)
    assertEquals(updatedSheet(ref"E3").value, CellValue.Number(3))
  }

  test("TRANSPOSE in non-array context returns top-left value") {
    // When TRANSPOSE is evaluated as a regular formula (not array formula),
    // it should return the top-left value of the transposed result
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Number(1))
      .put(ref"B1", CellValue.Number(2))
      .put(ref"A2", CellValue.Number(3))
      .put(ref"B2", CellValue.Number(4))

    val result = sheet.evaluateFormula("=TRANSPOSE(A1:B2)")

    assert(result.isRight)
    // Top-left of transposed [[1,3],[2,4]] is 1
    assertEquals(result.toOption.get, CellValue.Number(1))
  }

  // ========== Patch.PutArray Tests ==========

  test("Patch.PutArray applies grid of values") {
    val sheet = Sheet("Test")
    val values = Vector(
      Vector(CellValue.Number(1), CellValue.Number(2)),
      Vector(CellValue.Number(3), CellValue.Number(4))
    )

    val patch = Patch.PutArray(ref"B2", values)
    val updated = Patch.applyPatch(sheet, patch)

    assertEquals(updated(ref"B2").value, CellValue.Number(1))
    assertEquals(updated(ref"C2").value, CellValue.Number(2))
    assertEquals(updated(ref"B3").value, CellValue.Number(3))
    assertEquals(updated(ref"C3").value, CellValue.Number(4))
  }

  // ========== GH-76 Dynamic Arrays ==========

  test("SEQUENCE(2,3) generates a 2x3 grid 1..6") {
    val result = Sheet("Test").evaluateArrayFormula("=SEQUENCE(2,3)", ref"E1")
    assert(result.isRight, s"$result")
    val (s, range) = result.toOption.get
    assertEquals(range.height, 2)
    assertEquals(range.width, 3)
    assertEquals(s(ref"E1").value, CellValue.Number(1))
    assertEquals(s(ref"G1").value, CellValue.Number(3))
    assertEquals(s(ref"E2").value, CellValue.Number(4))
    assertEquals(s(ref"G2").value, CellValue.Number(6))
  }

  test("SEQUENCE(3,1,10,5) honors start and step") {
    val (s, _) = Sheet("Test").evaluateArrayFormula("=SEQUENCE(3,1,10,5)", ref"E1").toOption.get
    assertEquals(s(ref"E1").value, CellValue.Number(10))
    assertEquals(s(ref"E2").value, CellValue.Number(15))
    assertEquals(s(ref"E3").value, CellValue.Number(20))
  }

  test("SORT sorts rows ascending by the first column") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Number(3))
      .put(ref"B1", CellValue.Text("c"))
      .put(ref"A2", CellValue.Number(1))
      .put(ref"B2", CellValue.Text("a"))
      .put(ref"A3", CellValue.Number(2))
      .put(ref"B3", CellValue.Text("b"))
    val (s, range) = sheet.evaluateArrayFormula("=SORT(A1:B3)", ref"E1").toOption.get
    assertEquals(range.height, 3)
    assertEquals(range.width, 2)
    assertEquals(s(ref"E1").value, CellValue.Number(1))
    assertEquals(s(ref"F1").value, CellValue.Text("a"))
    assertEquals(s(ref"E3").value, CellValue.Number(3))
  }

  test("SORT descending with sort_order -1") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Number(1))
      .put(ref"A2", CellValue.Number(3))
      .put(ref"A3", CellValue.Number(2))
    val (s, _) = sheet.evaluateArrayFormula("=SORT(A1:A3,1,-1)", ref"E1").toOption.get
    assertEquals(s(ref"E1").value, CellValue.Number(3))
    assertEquals(s(ref"E2").value, CellValue.Number(2))
    assertEquals(s(ref"E3").value, CellValue.Number(1))
  }

  test("GH-596/GH-654: SORT descending keeps blank keys last and ties in source order") {
    // F1 = 7, F2 blank, F3 = 5, F4 = 3 — the shape `xl sort --desc` orders 7,5,3,blank
    val sheet = Sheet("Test")
      .put(ref"F1", CellValue.Number(7))
      .put(ref"F3", CellValue.Number(5))
      .put(ref"F4", CellValue.Number(3))
    val (desc, _) = sheet.evaluateArrayFormula("=SORT(F1:F4,1,-1)", ref"H1").toOption.get
    assertEquals(desc(ref"H1").value, CellValue.Number(7))
    assertEquals(desc(ref"H2").value, CellValue.Number(5))
    assertEquals(desc(ref"H3").value, CellValue.Number(3))
    assertEquals(desc(ref"H4").value, CellValue.Empty, "the blank is last, not first")
    val (asc, _) = sheet.evaluateArrayFormula("=SORT(F1:F4)", ref"H1").toOption.get
    assertEquals(asc(ref"H1").value, CellValue.Number(3))
    assertEquals(asc(ref"H4").value, CellValue.Empty, "the blank is last ascending too")
    // ties: equal keys keep their source order in both directions (a reversed ascending sort
    // reversed them)
    val tied = Sheet("Test")
      .put(ref"A1", CellValue.Number(1))
      .put(ref"B1", CellValue.Text("first"))
      .put(ref"A2", CellValue.Number(1))
      .put(ref"B2", CellValue.Text("second"))
      .put(ref"A3", CellValue.Number(2))
      .put(ref"B3", CellValue.Text("third"))
    val (t, _) = tied.evaluateArrayFormula("=SORT(A1:B3,1,-1)", ref"E1").toOption.get
    assertEquals(t(ref"F1").value, CellValue.Text("third"))
    assertEquals(t(ref"F2").value, CellValue.Text("first"))
    assertEquals(t(ref"F3").value, CellValue.Text("second"))
  }

  test("UNIQUE returns distinct rows in first-seen order") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Number(1))
      .put(ref"A2", CellValue.Number(1))
      .put(ref"A3", CellValue.Number(2))
      .put(ref"A4", CellValue.Number(3))
      .put(ref"A5", CellValue.Number(3))
    val (s, range) = sheet.evaluateArrayFormula("=UNIQUE(A1:A5)", ref"E1").toOption.get
    assertEquals(range.height, 3)
    assertEquals(s(ref"E1").value, CellValue.Number(1))
    assertEquals(s(ref"E2").value, CellValue.Number(2))
    assertEquals(s(ref"E3").value, CellValue.Number(3))
  }

  test("UNIQUE exactly_once keeps only singletons") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Number(1))
      .put(ref"A2", CellValue.Number(1))
      .put(ref"A3", CellValue.Number(2))
      .put(ref"A4", CellValue.Number(3))
      .put(ref"A5", CellValue.Number(3))
    val (s, range) = sheet.evaluateArrayFormula("=UNIQUE(A1:A5,FALSE,TRUE)", ref"E1").toOption.get
    assertEquals(range.height, 1)
    assertEquals(s(ref"E1").value, CellValue.Number(2))
  }

  test("FILTER keeps rows where include is truthy") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Number(10))
      .put(ref"A2", CellValue.Number(20))
      .put(ref"A3", CellValue.Number(30))
      .put(ref"A4", CellValue.Number(40))
      .put(ref"B1", CellValue.Number(1))
      .put(ref"B2", CellValue.Number(0))
      .put(ref"B3", CellValue.Number(1))
      .put(ref"B4", CellValue.Number(0))
    val (s, range) = sheet.evaluateArrayFormula("=FILTER(A1:A4,B1:B4)", ref"E1").toOption.get
    assertEquals(range.height, 2)
    assertEquals(s(ref"E1").value, CellValue.Number(10))
    assertEquals(s(ref"E2").value, CellValue.Number(30))
  }

  test("FILTER returns if_empty when nothing matches") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Number(10))
      .put(ref"A2", CellValue.Number(20))
      .put(ref"B1", CellValue.Number(0))
      .put(ref"B2", CellValue.Number(0))
    val (s, _) = sheet.evaluateArrayFormula("=FILTER(A1:A2,B1:B2,\"none\")", ref"E1").toOption.get
    assertEquals(s(ref"E1").value, CellValue.Text("none"))
  }

  test("OFFSET spills a block (=OFFSET(A1,0,0,2,2))") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Number(1))
      .put(ref"B1", CellValue.Number(2))
      .put(ref"A2", CellValue.Number(3))
      .put(ref"B2", CellValue.Number(4))
    val (s, range) = sheet.evaluateArrayFormula("=OFFSET(A1,0,0,2,2)", ref"E1").toOption.get
    assertEquals(range.height, 2)
    assertEquals(range.width, 2)
    assertEquals(s(ref"E1").value, CellValue.Number(1))
    assertEquals(s(ref"F1").value, CellValue.Number(2))
    assertEquals(s(ref"E2").value, CellValue.Number(3))
    assertEquals(s(ref"F2").value, CellValue.Number(4))
  }

  // ===== GH-580: FILTER's include argument is any array-valued expression =====

  private val filterSheet = Sheet("Test")
    .put(ref"A1", CellValue.Text("a"))
    .put(ref"A2", CellValue.Text("b"))
    .put(ref"A3", CellValue.Text("a"))
    .put(ref"B1", CellValue.Number(1))
    .put(ref"B2", CellValue.Number(2))
    .put(ref"B3", CellValue.Number(3))

  test("GH-580: FILTER accepts a comparison over a range as include (=FILTER(B1:B3,B1:B3>1))") {
    val (s, range) =
      filterSheet.evaluateArrayFormula("=FILTER(B1:B3,B1:B3>1)", ref"E1").toOption.get
    assertEquals(range.height, 2)
    assertEquals(s(ref"E1").value, CellValue.Number(2))
    assertEquals(s(ref"E2").value, CellValue.Number(3))
  }

  test("GH-580: FILTER include may be a text comparison or a product of conditions") {
    val (s1, r1) =
      filterSheet.evaluateArrayFormula("=FILTER(B1:B3,A1:A3<>\"a\")", ref"E1").toOption.get
    assertEquals(r1.height, 1)
    assertEquals(s1(ref"E1").value, CellValue.Number(2))
    val (s2, r2) = filterSheet
      .evaluateArrayFormula("=FILTER(B1:B3,(A1:A3=\"a\")*(B1:B3>1))", ref"E1")
      .toOption
      .get
    assertEquals(r2.height, 1)
    assertEquals(s2(ref"E1").value, CellValue.Number(3))
  }

  test("GH-580: SUM over a FILTER with an array condition evaluates in a scalar formula") {
    assertEquals(
      filterSheet.evaluateFormula("=SUM(FILTER(B1:B3,A1:A3<>\"a\"))"),
      Right(CellValue.Number(2))
    )
    assertEquals(
      filterSheet.evaluateFormula("=SUM(FILTER(B1:B3,A1:A3=\"a\"))"),
      Right(CellValue.Number(4))
    )
  }

  test("GH-580: a one-row include as wide as the array filters columns") {
    val wide = Sheet("Test")
      .put(ref"A1", CellValue.Number(1))
      .put(ref"B1", CellValue.Number(5))
      .put(ref"C1", CellValue.Number(2))
      .put(ref"A2", CellValue.Number(10))
      .put(ref"B2", CellValue.Number(50))
      .put(ref"C2", CellValue.Number(20))
    val (s, range) = wide.evaluateArrayFormula("=FILTER(A1:C2,A1:C1>1)", ref"E1").toOption.get
    assertEquals((range.height, range.width), (2, 2))
    assertEquals(s(ref"E1").value, CellValue.Number(5))
    assertEquals(s(ref"F1").value, CellValue.Number(2))
    assertEquals(s(ref"E2").value, CellValue.Number(50))
    assertEquals(s(ref"F2").value, CellValue.Number(20))
  }

  test("GH-580: an include whose shape matches neither rows nor columns is #VALUE!") {
    val expr = FormulaParser.parse("=FILTER(B1:B3,B1:B2>1)").toOption.get
    Evaluator.arrayInstance.eval(expr, filterSheet) match
      case Left(EvalError.ErrorValue(CellError.Value, _)) => ()
      case other => fail(s"expected #VALUE!, got $other")
  }

  test("GH-654: FILTER's include may be a single cell — a bare reference, not just a range") {
    // D1 = TRUE keeps the one-row array whole; D1 = FALSE empties it
    val row = Sheet("Test")
      .put(ref"A1", CellValue.Number(1))
      .put(ref"B1", CellValue.Number(10))
      .put(ref"C1", CellValue.Number(7))
      .put(ref"D1", CellValue.Bool(true))
    val (s, range) = row.evaluateArrayFormula("=FILTER(A1:C1,D1)", ref"E3").toOption.get
    assertEquals(range.width, 3)
    assertEquals(s(ref"E3").value, CellValue.Number(1))
    assertEquals(s(ref"G3").value, CellValue.Number(7))
    assertEquals(row.evaluateFormula("=SUM(FILTER(A1:C1,D1))"), Right(CellValue.Number(18)))
    assertEquals(row.evaluateFormula("=SUM(FILTER(A1:A1,D1))"), Right(CellValue.Number(1)))
    assertEquals(
      row.evaluateFormula("=SUM(FILTER(A1:C1,D1))"),
      row.evaluateFormula("=SUM(FILTER(A1:C1,D1:D1))"),
      "the bare cell and the 1×1 range are the same include"
    )
    val off = row.put(ref"D1", CellValue.Bool(false))
    assertEquals(
      off.evaluateFormula("=FILTER(A1:C1,D1,\"none\")"),
      Right(CellValue.Text("none"))
    )
  }

  test("GH-580: FILTER with an array condition prints back and round-trips") {
    val parsed = FormulaParser.parse("=FILTER(B1:B3,B1:B3>1)")
    assertEquals(parsed.map(FormulaPrinter.print(_)), Right("=FILTER(B1:B3, B1:B3>1)"))
    assertEquals(parsed.map(FormulaPrinter.print(_)).flatMap(FormulaParser.parse), parsed)
  }
