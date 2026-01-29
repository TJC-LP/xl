package com.tjclp.xl.formula

import munit.FunSuite

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.eval.ArrayResult
import com.tjclp.xl.formula.eval.SheetEvaluator.*
import com.tjclp.xl.patch.Patch
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.syntax.*

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
