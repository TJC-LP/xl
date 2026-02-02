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
        val transformed = TExpr.transformRanges(expr, { (_, range) =>
          if range.isFullColumn then
            CellRange(ref"A1", ref"A3")
          else
            range
        })
        println(s"Transformed: $transformed")
        // Check that ranges were transformed
        val newRanges = TExpr.collectRanges(transformed)
        println(s"New ranges: $newRanges")
        val rangeStrings = newRanges.map(_._2.toA1)
        assert(!rangeStrings.contains("A:A"), s"Should NOT contain A:A after transform, got: $rangeStrings")
      case Left(err) =>
        fail(s"Parse failed: $err")
  }
