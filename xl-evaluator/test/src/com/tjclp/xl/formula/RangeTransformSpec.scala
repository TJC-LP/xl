package com.tjclp.xl.formula

import munit.FunSuite
import com.tjclp.xl.*
import com.tjclp.xl.formula.ast.TExpr

// Test code uses .get for brevity in assertions (guarded by isDefined check)
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial"))
class RangeTransformSpec extends FunSuite:

  test("collectRanges extracts ranges from SUMPRODUCT array expression") {
    FormulaParser.parse("=SUMPRODUCT((A:A>15)*B:B)") match
      case Right(expr) =>
        val ranges = TExpr.collectRanges(expr)
        println(s"Expression: $expr")
        println(s"Collected ranges: $ranges")
        // Should find two full-column ranges (A:A and B:B)
        // Note: CellRange.toA1 returns "A1:A1048576" for full-column ranges, not "A:A"
        assertEquals(ranges.length, 2, s"Should have 2 ranges, got: $ranges")
        val colARange = ranges.find(_._2.colStart.index0 == 0)
        val colBRange = ranges.find(_._2.colStart.index0 == 1)
        assert(colARange.isDefined, s"Should have range for column A")
        assert(colBRange.isDefined, s"Should have range for column B")
        assert(colARange.get._2.isFullColumn, s"Column A range should be full-column")
        assert(colBRange.get._2.isFullColumn, s"Column B range should be full-column")
      case Left(err) =>
        fail(s"Parse failed: $err")
  }

  test("transformRanges transforms ranges in comparison expression") {
    FormulaParser.parse("=(A:A>15)") match
      case Right(expr) =>
        println(s"Original: $expr")
        // Transform A:A to A1:A3
        val transformed = TExpr.transformRanges(
          expr,
          { (_, range) =>
            if range.isFullColumn then CellRange(ref"A1", ref"A3")
            else range
          }
        )
        println(s"Transformed: $transformed")
        // Check that ranges were transformed
        val newRanges = TExpr.collectRanges(transformed)
        println(s"New ranges: $newRanges")
        val rangeStrings = newRanges.map(_._2.toA1)
        assert(
          !rangeStrings.contains("A:A"),
          s"Should NOT contain A:A after transform, got: $rangeStrings"
        )
      case Left(err) =>
        fail(s"Parse failed: $err")
  }

  // ===== Array lifting: ranges inside lifted slots are element sources =====

  private def bounded(formula: String): List[String] =
    FormulaParser.parse(formula) match
      case Right(expr) =>
        val transformed = TExpr.transformRanges(
          expr,
          (_, range) => if range.isFullColumn then CellRange(range.start, range.start) else range
        )
        TExpr.collectRanges(transformed).map(_._2.toA1)
      case Left(err) => fail(s"Parse failed: $err")

  // the argument expressions SUMPRODUCT bounds before evaluating them
  test("transformRanges bounds ranges in lifted scalar slots, powers and concatenations") {
    assertEquals(bounded("=ABS(B:B)"), List("B1:B1"))
    assertEquals(bounded("=ROUND(B:B/A:A,2)"), List("B1:B1", "A1:A1"))
    assertEquals(bounded("=(A:A)^2"), List("A1:A1"))
    assertEquals(bounded("=LEN(A:A&\"\")"), List("A1:A1"))
    // IFERROR lifts its value only: the fallback keeps its shape
    assertEquals(bounded("=IFERROR(A:A,B:B)"), List("A1:A1", "B1:B1048576"))
  }

  test("transformRanges leaves shape-sensitive calls alone") {
    assertEquals(bounded("=ROWS(A:A)"), List("A1:A1048576"))
    // a lifted call keeps its range slots whole: only its scalar slots are element sources
    assertEquals(bounded("=INDEX(A:A,B:B)"), List("A1:A1048576", "B1:B1"))
  }

  test("SUMPRODUCT trims whole columns inside lifted slots to the used extent") {
    val sheet = Sheet("S")
      .put(ref"A1", CellValue.Number(BigDecimal(1)))
      .put(ref"A2", CellValue.Number(BigDecimal(-2)))
      .put(ref"A3", CellValue.Number(BigDecimal(3)))
      .put(ref"B1", CellValue.Number(BigDecimal(-4)))
      .put(ref"B2", CellValue.Number(BigDecimal(5)))
      .put(ref"B3", CellValue.Number(BigDecimal(-6)))
    def eval(formula: String) = sheet.evaluateFormula(formula)
    assertEquals(
      eval("=SUMPRODUCT(--(A:A>0),ABS(B:B))"),
      Right(CellValue.Number(BigDecimal(10))),
      "must not be #VALUE! for mismatched dimensions"
    )
    assertEquals(
      eval("=SUMPRODUCT(--(A:A>0),ABS(B:B))"),
      eval("=SUMPRODUCT(--(A1:A3>0),ABS(B1:B3))")
    )
    assertEquals(eval("=SUMPRODUCT((A:A)^2)"), Right(CellValue.Number(BigDecimal(14))))
    assertEquals(eval("=SUMPRODUCT(LEN(A:A&\"\"))"), Right(CellValue.Number(BigDecimal(4))))
  }
