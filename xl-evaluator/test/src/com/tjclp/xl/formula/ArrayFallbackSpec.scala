package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.eval.{ArrayResult, EvalError, Evaluator}
import com.tjclp.xl.formula.parser.FormulaParser
import munit.FunSuite

/** GH-754: Excel expectations recorded in the issue; LibreOffice 24.2 disagrees on IFNA. */
class ArrayFallbackSpec extends FunSuite:
  private def number(n: Int): CellValue = CellValue.Number(BigDecimal(n))
  private val sheet = (5 to 7).foldLeft(Sheet("Data")) { (s, r) =>
    s.put(ARef.from0(0, r - 1), number(r - 4))
      .put(ARef.from0(1, r - 1), number((r - 4) * 10))
  }

  for (fn, error) <- List("IFERROR" -> "1/0", "IFNA" -> "NA()") do
    for (suffix, expected) <- List(
        s"(A:A>0)*$fn($error,ROW(A:A))" -> 18,
        s"(A1:A10>0)*$fn($error,ROW(A1:A10))" -> 18,
        s"$fn($error,{10,20})" -> 30,
        s"(A:A>0)*$fn($error,B:B)" -> 60,
        s"$fn(5,{1,2})" -> 5,
        s"(A:A>0)*$fn(ROW(A:A),0)" -> 18
      )
    do
      test(s"GH-754: SUMPRODUCT($suffix) in placed and unplaced contexts") {
        val formula = s"=SUMPRODUCT($suffix)"
        assertEquals(sheet.evaluateFormula(formula), Right(number(expected)))
        val placed = sheet.put(ref"C6", CellValue.Formula(formula, None))
        assertEquals(placed.evaluateCell(ref"C6"), Right(number(expected)))
      }

  private def array(formula: String): Any =
    val expr = FormulaParser.parse(formula).fold(e => fail(e.toString), identity)
    Evaluator.arrayInstance
      .eval(expr, sheet.put(ref"D1", CellValue.Formula("=D1", None)))
      .fold(e => fail(e.toString), identity)

  for fn <- List("IFERROR", "IFNA") do
    test(s"GH-754: $fn replaces errors positionally and broadcasts row against column") {
      assertEquals(
        array(s"=$fn({1,#N/A;#N/A,4},{10,20})"),
        ArrayResult(Vector(Vector(number(1), number(20)), Vector(number(10), number(4))))
      )
      assertEquals(
        array(s"=$fn({#N/A;2},{10,20})"),
        ArrayResult(Vector(Vector(number(10), number(20)), Vector(number(2), number(2))))
      )
    }
    test(s"GH-754: $fn never evaluates an unused fallback") {
      assertEquals(array(s"=$fn(5,D1)"), number(5))
      assertEquals(array(s"=$fn({1,2},D1)"), ArrayResult(Vector(Vector(number(1), number(2)))))
    }

  test("GH-754: IFNA preserves other errors, including inside arrays") {
    assertEquals(
      array("=IFNA({#DIV/0!,#N/A},{10,20})"),
      ArrayResult(Vector(Vector(CellValue.Error(CellError.Div0), number(20))))
    )
  }

  for fn <- List("IFERROR", "IFNA") do
    test(s"GH-754: $fn carries a failing fallback only at selected elements") {
      assertEquals(
        array(s"=$fn({1,#N/A},1/0)"),
        ArrayResult(Vector(Vector(number(1), CellValue.Error(CellError.Div0))))
      )
    }
    test(s"GH-754: $fn broadcasts a scalar fallback over mixed array elements") {
      assertEquals(
        array(s"=$fn({1,#N/A;#N/A,4},10)"),
        ArrayResult(Vector(Vector(number(1), number(10)), Vector(number(10), number(4))))
      )
    }
    test(s"GH-754: $fn keeps plain-cell fallback reference intersection") {
      val formula = s"=$fn(NA(),B5:B7)"
      val placed = sheet.put(ref"C6", CellValue.Formula(formula, None))
      assertEquals(placed.evaluateCell(ref"C6"), Right(number(20)))
    }

  test("GH-754: static whole-axis trimming enters both error-guard slots") {
    def parse(f: String): TExpr[?] = FormulaParser.parse(f).fold(e => fail(e.toString), identity)
    val trim = (_: Option[SheetName], range: CellRange) =>
      if range.isFullColumn then CellRange(ref"B5", ref"B7") else range
    for fn <- List("IFERROR", "IFNA") do
      assertEquals(
        TExpr.transformRanges(parse(s"=$fn(A:A,B:B)"), trim),
        parse(s"=$fn(B5:B7,B5:B7)")
      )
  }

  test("GH-754: IFNA does not evaluate its fallback for a scalar division failure") {
    val cyclic = sheet.put(ref"D1", CellValue.Formula("=D1", None))
    assertEquals(
      cyclic.evaluateFormula("=IFNA(1/0,D1)"),
      Right(CellValue.Error(CellError.Div0))
    )
  }

  for (fn, error) <- List("IFERROR" -> "1/0", "IFNA" -> "NA()") do
    test(s"GH-754: $fn retains the selected fallback's exact array shape") {
      assertEquals(
        array(s"=$fn($error,{10,20})"),
        ArrayResult(Vector(Vector(number(10), number(20))))
      )
      array(s"=$fn($error,B:B)") match
        case result: ArrayResult =>
          assertEquals((result.rows, result.cols), (1048576, 1))
          assertEquals(result(4, 0), number(10))
          assertEquals(result(6, 0), number(30))
          assertEquals(result(1048575, 0), CellValue.Empty)
        case other => fail(s"expected a whole-column fallback array, got $other")
    }

  for fn <- List("IFERROR", "IFNA") do
    for fallback <- List("RAND()", "randomFallback") do
      test(s"GH-754: $fn retains per-selected-element draws from $fallback") {
        val expectedRng = Rng.seeded(42L)
        val first = CellValue.Number(BigDecimal(expectedRng.nextDouble()))
        val second = CellValue.Number(BigDecimal(expectedRng.nextDouble()))
        val formula = s"=$fn({#N/A,7,#N/A},$fallback)"
        val expr = FormulaParser.parse(formula).fold(e => fail(e.toString), identity)
        val wb = Workbook(sheet).withDefinedName("randomFallback", "RAND()")
        val actual: Either[EvalError, Any] = Evaluator
          .instance(Rng.seeded(42L))
          .withArrayResults
          .eval(
            expr,
            sheet,
            Clock.system,
            Some(wb)
          )
        assertEquals(actual, Right(ArrayResult(Vector(Vector(first, number(7), second)))))
      }
