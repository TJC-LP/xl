package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue, FormulaKind}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * The references a plain formula cell passes around: IF and CHOOSE return the reference they
 * select, OFFSET a range sized like its base, a dynamic named range the reference it computes. In a
 * value position a plain cell intersects them, in an aggregate's argument they stay whole; the
 * evaluation itself runs once (a selector draws RAND once). Array formulas keep array semantics
 * wherever they are evaluated, an iterative cycle included.
 */
class PlainCellReferenceSpec extends FunSuite:

  private def num(n: Double): CellValue = CellValue.Number(BigDecimal(n))

  /** A1:A10 = 1..10, B1:B10 = 2, Other!A1:A10 = 10..100. */
  private val book: Workbook =
    val sheet = (1 to 10).foldLeft(Sheet(SheetName.unsafe("S"))) { (s, i) =>
      s.put(ARef.from0(0, i - 1), num(i.toDouble)).put(ARef.from0(1, i - 1), num(2))
    }
    val other = (1 to 10).foldLeft(Sheet(SheetName.unsafe("Other"))) { (s, i) =>
      s.put(ARef.from0(0, i - 1), num(i * 10.0))
    }
    Workbook(Vector(sheet, other))
      .withDefinedName("dyn", "OFFSET(S!$A$1,0,0,10,1)")
      .withDefinedName("fixed", "S!$A$1:$A$10")

  private def sheetS: Sheet = book.sheets.headOption.getOrElse(fail("no sheet"))

  /** `formula` as a plain cell at `at` on S, evaluated at its position. */
  private def plain(formula: String, at: ARef): XLResult[CellValue] =
    val placed = sheetS.put(at, CellValue.Formula(formula, None))
    placed.evaluateCell(at, Clock.system, Some(book.put(placed)))

  test("OFFSET's height, width and sheet default to its base reference's, as in Excel") {
    assertEquals(plain("=SUM(OFFSET(A1:A10,0,1))", ref"D1"), Right(num(20)))
    assertEquals(plain("=ROWS(OFFSET(A1:A10,0,1))", ref"D1"), Right(num(10)))
    // in a value position the moved range is intersected: B4 from row 4
    assertEquals(plain("=OFFSET(A2:A5,0,1)*10", ref"D4"), Right(num(20)))
    assertEquals(plain("=OFFSET(Other!A1,1,0)", ref"D1"), Right(num(20)))
    assertEquals(plain("=SUM(OFFSET(Other!A1:A10,0,0))", ref"D1"), Right(num(550)))
  }

  test("a dynamic named range is a reference: intersected in a value position, whole in SUM") {
    assertEquals(plain("=dyn*2", ref"D5"), Right(num(10)))
    assertEquals(plain("=dyn*2", ref"D20"), Right(CellValue.Error(CellError.Value)))
    assertEquals(plain("=SUM(dyn)", ref"D20"), Right(num(55)))
    assertEquals(plain("=SUM(IF(TRUE,dyn,0))", ref"D20"), Right(num(55)))
    assertEquals(plain("=fixed*2", ref"D5"), Right(num(10)))
    assertEquals(plain("=SUM(IF(TRUE,fixed,0))", ref"D20"), Right(num(55)))
  }

  test("IF evaluates its condition once and keeps the reference it selects whole") {
    final class CountingRng(value: Double) extends Rng:
      @SuppressWarnings(Array("org.wartremover.warts.Var"))
      var draws: Int = 0
      def nextDouble(): Double =
        draws += 1
        value
    List(0.2 -> num(55), 0.8 -> num(0)).foreach { (draw, expected) =>
      val rng = CountingRng(draw)
      val placed = sheetS.put(ref"D5", CellValue.Formula("SUM(IF(RAND()<0.5,A1:A10,0))", None))
      val result = placed.evaluateCell(ref"D5", Clock.system, rng, Some(book.put(placed)))
      assertEquals(result, Right(expected), s"draw $draw")
      assertEquals(rng.draws, 1, "the condition evaluates once")
    }
  }

  test("an array formula in an iterative cycle evaluates as an array every round") {
    // C2 = {=SUM(A1:A3*1)+0*D2} (a CSE record) and D2 = C2 form a cycle: the array formula sums
    // A1:A3 (6) each round, never the plain cell's row-2 intersection (2)
    val cycle = sheetS
      .put(
        ref"C2",
        CellValue.Formula(
          "SUM(A1:A3*1)+0*D2",
          None,
          FormulaKind.ArrayFormula(CellRange(ref"C2", ref"C2"))
        )
      )
      .put(ref"D2", CellValue.Formula("C2", None))
    val result = book.put(cycle).recalculate(IterativeCalc(20, BigDecimal("0.001")))
    val c2 = result.workbook.sheets.headOption.flatMap(_.cells.get(ref"C2")).map(_.value)
    assertEquals(
      c2.collect { case CellValue.Formula(_, cached, _) => cached },
      Some(Some(num(6)))
    )
  }

  test("evaluateFormula at a cell is that plain cell's evaluateCell") {
    List(
      "=A1:A10*2",
      "=SUM(A1:A10*B1:B10)",
      "=IF(A1:A10>4,\"big\",\"small\")",
      "=SUM(IF(A1:A10>4,A1:A10,0))",
      "=dyn*2",
      "=COUNTIF(A1:A10,\">\"&A1:A10)"
    ).foreach { formula =>
      List(ref"D1", ref"D5", ref"D20").foreach { at =>
        assertEquals(
          sheetS.evaluateFormula(formula, Clock.system, Some(book), Some(at)),
          plain(formula, at),
          s"$formula at ${at.toA1}"
        )
      }
    }
  }

  test("the element law: a plain cell in row k reads element k of the array formula") {
    // for a column reference r and a plain cell in r's row k, plain(f) = arrayEval(f)(k); outside
    // r's rows it is #VALUE!
    List(
      "=A1:A10*3",
      "=ABS(A1:A10-5)",
      "=A1:A10&\"x\"",
      "=IF(A1:A10>4,A1:A10,-1)",
      "=ROUND(A1:A10/3,1)",
      "=COUNTIF(A1:A10,\"<\"&A1:A10)"
    ).foreach { formula =>
      val array = sheetS.evaluateArrayFormula(formula, ref"F1", Clock.system, Some(book)) match
        case Right((spilled, _)) => spilled
        case Left(err) => fail(s"$formula: ${err.message}")
      (1 to 10).foreach { row =>
        val at = ARef.from0(3, row - 1)
        val element = array.cells.get(ARef.from0(5, row - 1)).map(_.value)
        assertEquals(plain(formula, at).toOption, element, s"$formula at row $row")
      }
      assertEquals(plain(formula, ref"D11"), Right(CellValue.Error(CellError.Value)), formula)
    }
  }
