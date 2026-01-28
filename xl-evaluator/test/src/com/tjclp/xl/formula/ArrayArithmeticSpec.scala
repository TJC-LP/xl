package com.tjclp.xl.formula

import munit.FunSuite

import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.eval.{ArrayResult, ArrayArithmetic}
import com.tjclp.xl.formula.eval.SheetEvaluator.*
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.syntax.*

class ArrayArithmeticSpec extends FunSuite:

  // ========== Broadcasting Unit Tests ==========

  test("scalar * scalar = scalar (fast path)") {
    val result = ArrayArithmetic.broadcast(
      ArrayArithmetic.ArrayOperand.Scalar(BigDecimal(2)),
      ArrayArithmetic.ArrayOperand.Scalar(BigDecimal(3)),
      ArrayArithmetic.mul
    )
    assert(result.isRight, s"Expected Right, got $result")
    result match
      case Right(ArrayArithmetic.ArrayOperand.Scalar(v)) =>
        assertEquals(v, BigDecimal(6))
      case other => fail(s"Expected Scalar, got $other")
  }

  test("scalar + array = array (broadcast scalar)") {
    val arr = ArrayResult(
      Vector(
        Vector(CellValue.Number(1), CellValue.Number(2), CellValue.Number(3))
      )
    )
    val result = ArrayArithmetic.broadcast(
      ArrayArithmetic.ArrayOperand.Scalar(BigDecimal(10)),
      ArrayArithmetic.ArrayOperand.Array(arr),
      ArrayArithmetic.add
    )
    assert(result.isRight, s"Expected Right, got $result")
    result match
      case Right(ArrayArithmetic.ArrayOperand.Array(out)) =>
        assertEquals(out.rows, 1)
        assertEquals(out.cols, 3)
        assertEquals(out(0, 0), CellValue.Number(11))
        assertEquals(out(0, 1), CellValue.Number(12))
        assertEquals(out(0, 2), CellValue.Number(13))
      case other => fail(s"Expected Array, got $other")
  }

  test("MxN * 1xN = MxN (row broadcasts across rows)") {
    // 3x2 matrix * 1x2 row = 3x2 result
    val matrix = ArrayResult(
      Vector(
        Vector(CellValue.Number(1), CellValue.Number(2)),
        Vector(CellValue.Number(3), CellValue.Number(4)),
        Vector(CellValue.Number(5), CellValue.Number(6))
      )
    )
    val row = ArrayResult(
      Vector(
        Vector(CellValue.Number(10), CellValue.Number(100))
      )
    )

    val result = ArrayArithmetic.broadcast(
      ArrayArithmetic.ArrayOperand.Array(matrix),
      ArrayArithmetic.ArrayOperand.Array(row),
      ArrayArithmetic.mul
    )

    assert(result.isRight, s"Expected Right, got $result")
    result match
      case Right(ArrayArithmetic.ArrayOperand.Array(out)) =>
        assertEquals(out.rows, 3)
        assertEquals(out.cols, 2)
        assertEquals(out(0, 0), CellValue.Number(10)) // 1 * 10
        assertEquals(out(0, 1), CellValue.Number(200)) // 2 * 100
        assertEquals(out(1, 0), CellValue.Number(30)) // 3 * 10
        assertEquals(out(1, 1), CellValue.Number(400)) // 4 * 100
        assertEquals(out(2, 0), CellValue.Number(50)) // 5 * 10
        assertEquals(out(2, 1), CellValue.Number(600)) // 6 * 100
      case other => fail(s"Expected Array, got $other")
  }

  test("MxN * Mx1 = MxN (column broadcasts across columns)") {
    // 2x3 matrix * 2x1 column = 2x3 result
    val matrix = ArrayResult(
      Vector(
        Vector(CellValue.Number(1), CellValue.Number(2), CellValue.Number(3)),
        Vector(CellValue.Number(4), CellValue.Number(5), CellValue.Number(6))
      )
    )
    val col = ArrayResult(
      Vector(
        Vector(CellValue.Number(10)),
        Vector(CellValue.Number(100))
      )
    )

    val result = ArrayArithmetic.broadcast(
      ArrayArithmetic.ArrayOperand.Array(matrix),
      ArrayArithmetic.ArrayOperand.Array(col),
      ArrayArithmetic.mul
    )

    assert(result.isRight, s"Expected Right, got $result")
    result match
      case Right(ArrayArithmetic.ArrayOperand.Array(out)) =>
        assertEquals(out.rows, 2)
        assertEquals(out.cols, 3)
        assertEquals(out(0, 0), CellValue.Number(10)) // 1 * 10
        assertEquals(out(0, 1), CellValue.Number(20)) // 2 * 10
        assertEquals(out(0, 2), CellValue.Number(30)) // 3 * 10
        assertEquals(out(1, 0), CellValue.Number(400)) // 4 * 100
        assertEquals(out(1, 1), CellValue.Number(500)) // 5 * 100
        assertEquals(out(1, 2), CellValue.Number(600)) // 6 * 100
      case other => fail(s"Expected Array, got $other")
  }

  test("incompatible dimensions returns error") {
    // 2x3 * 2x2 -> error (cols don't match and neither is 1)
    val a = ArrayResult(
      Vector(
        Vector(CellValue.Number(1), CellValue.Number(2), CellValue.Number(3)),
        Vector(CellValue.Number(4), CellValue.Number(5), CellValue.Number(6))
      )
    )
    val b = ArrayResult(
      Vector(
        Vector(CellValue.Number(1), CellValue.Number(2)),
        Vector(CellValue.Number(3), CellValue.Number(4))
      )
    )

    val result = ArrayArithmetic.broadcast(
      ArrayArithmetic.ArrayOperand.Array(a),
      ArrayArithmetic.ArrayOperand.Array(b),
      ArrayArithmetic.mul
    )

    assert(result.isLeft, s"Expected Left (error), got $result")
    assert(result.swap.toOption.get.toString.contains("incompatible"))
  }

  test("division by zero in array returns error") {
    val arr = ArrayResult(
      Vector(
        Vector(CellValue.Number(1), CellValue.Number(0), CellValue.Number(3))
      )
    )
    val scalar = ArrayArithmetic.ArrayOperand.Scalar(BigDecimal(10))

    val result = ArrayArithmetic.broadcast(
      scalar,
      ArrayArithmetic.ArrayOperand.Array(arr),
      ArrayArithmetic.div
    )

    assert(result.isLeft, s"Expected Left (div by zero), got $result")
  }

  // ========== Integration Tests with evaluateArrayFormula ==========

  test("range * scalar via evaluateArrayFormula") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Number(1))
      .put(ref"B1", CellValue.Number(2))
      .put(ref"C1", CellValue.Number(3))

    val result = sheet.evaluateArrayFormula("=A1:C1*10", ref"E1")

    assert(result.isRight, s"Expected Right, got $result")
    val (updatedSheet, spillRange) = result.toOption.get

    assertEquals(spillRange.width, 3)
    assertEquals(spillRange.height, 1)
    assertEquals(updatedSheet(ref"E1").value, CellValue.Number(10))
    assertEquals(updatedSheet(ref"F1").value, CellValue.Number(20))
    assertEquals(updatedSheet(ref"G1").value, CellValue.Number(30))
  }

  test("range * TRANSPOSE(range) via evaluateArrayFormula") {
    // This is the key SpreadsheetBench 52292 pattern!
    // 1x3 row * TRANSPOSE(3x1 col) = 1x3 * 1x3 = 1x3 element-wise
    val sheet = Sheet("Test")
      // Row of rates
      .put(ref"A1", CellValue.Number(10))
      .put(ref"B1", CellValue.Number(20))
      .put(ref"C1", CellValue.Number(30))
      // Column of hours
      .put(ref"A3", CellValue.Number(2))
      .put(ref"A4", CellValue.Number(3))
      .put(ref"A5", CellValue.Number(4))

    // A1:C1 = [10, 20, 30] (1x3 row)
    // TRANSPOSE(A3:A5) = [2, 3, 4] (1x3 row after transpose from 3x1 col)
    // Result = [20, 60, 120] (element-wise)
    val result = sheet.evaluateArrayFormula("=A1:C1*TRANSPOSE(A3:A5)", ref"E1")

    assert(result.isRight, s"Expected Right, got $result")
    val (updatedSheet, spillRange) = result.toOption.get

    assertEquals(spillRange.width, 3)
    assertEquals(spillRange.height, 1)
    assertEquals(updatedSheet(ref"E1").value, CellValue.Number(20)) // 10 * 2
    assertEquals(updatedSheet(ref"F1").value, CellValue.Number(60)) // 20 * 3
    assertEquals(updatedSheet(ref"G1").value, CellValue.Number(120)) // 30 * 4
  }

  test("matrix * transposed row (broadcasting pattern from 52292)") {
    // The SpreadsheetBench 52292 pattern: 6x6 rates * transposed 6x1 hours
    // Simplified to 2x2 for testing
    val sheet = Sheet("Test")
      // 2x2 rates matrix
      .put(ref"A1", CellValue.Number(10))
      .put(ref"B1", CellValue.Number(20))
      .put(ref"A2", CellValue.Number(30))
      .put(ref"B2", CellValue.Number(40))
      // 2x1 hours column
      .put(ref"D1", CellValue.Number(2))
      .put(ref"D2", CellValue.Number(3))

    // A1:B2 = [[10, 20], [30, 40]] (2x2 matrix)
    // TRANSPOSE(D1:D2) = [2, 3] (1x2 row after transpose)
    // Result = 2x2 with row broadcast:
    //   [[10*2, 20*3], [30*2, 40*3]] = [[20, 60], [60, 120]]
    val result = sheet.evaluateArrayFormula("=A1:B2*TRANSPOSE(D1:D2)", ref"F1")

    assert(result.isRight, s"Expected Right, got $result")
    val (updatedSheet, spillRange) = result.toOption.get

    assertEquals(spillRange.width, 2)
    assertEquals(spillRange.height, 2)
    assertEquals(updatedSheet(ref"F1").value, CellValue.Number(20)) // 10 * 2
    assertEquals(updatedSheet(ref"G1").value, CellValue.Number(60)) // 20 * 3
    assertEquals(updatedSheet(ref"F2").value, CellValue.Number(60)) // 30 * 2
    assertEquals(updatedSheet(ref"G2").value, CellValue.Number(120)) // 40 * 3
  }

  test("scalar arithmetic still works (regression test)") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Number(10))
      .put(ref"B1", CellValue.Number(5))

    // Simple scalar formula should still work
    val result = sheet.evaluateFormula("=A1+B1*2")
    assert(result.isRight, s"Expected Right, got $result")
    assertEquals(result.toOption.get, CellValue.Number(20)) // 10 + 5*2
  }

  test("scalar division still works (regression test)") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Number(100))
      .put(ref"B1", CellValue.Number(4))

    val result = sheet.evaluateFormula("=A1/B1")
    assert(result.isRight, s"Expected Right, got $result")
    assertEquals(result.toOption.get, CellValue.Number(25))
  }

  test("division by zero still returns error (regression test)") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Number(100))
      .put(ref"B1", CellValue.Number(0))

    val result = sheet.evaluateFormula("=A1/B1")
    assert(result.isLeft, s"Expected Left (div by zero), got $result")
  }

  test("range subtraction") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Number(100))
      .put(ref"B1", CellValue.Number(200))
      .put(ref"A2", CellValue.Number(50))
      .put(ref"B2", CellValue.Number(25))

    val result = sheet.evaluateArrayFormula("=A1:B1-A2:B2", ref"D1")

    assert(result.isRight, s"Expected Right, got $result")
    val (updatedSheet, spillRange) = result.toOption.get

    assertEquals(spillRange.width, 2)
    assertEquals(spillRange.height, 1)
    assertEquals(updatedSheet(ref"D1").value, CellValue.Number(50)) // 100 - 50
    assertEquals(updatedSheet(ref"E1").value, CellValue.Number(175)) // 200 - 25
  }

  test("range addition with scalar") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Number(1))
      .put(ref"A2", CellValue.Number(2))
      .put(ref"A3", CellValue.Number(3))

    val result = sheet.evaluateArrayFormula("=A1:A3+100", ref"C1")

    assert(result.isRight, s"Expected Right, got $result")
    val (updatedSheet, spillRange) = result.toOption.get

    assertEquals(spillRange.width, 1)
    assertEquals(spillRange.height, 3)
    assertEquals(updatedSheet(ref"C1").value, CellValue.Number(101))
    assertEquals(updatedSheet(ref"C2").value, CellValue.Number(102))
    assertEquals(updatedSheet(ref"C3").value, CellValue.Number(103))
  }
