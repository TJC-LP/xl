package com.tjclp.xl.formula

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.eval.{SheetEvaluator, WorkbookEvaluator}
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * #962: a formula calling a worksheet function from an add-in (`_xll.` namespace) that xl cannot
 * resolve keeps its saved cached value, as Excel does when the add-in is not loaded, and the cells
 * that read it recalculate from that value instead of being blocked.
 */
class AddinCacheSpec extends FunSuite:

  import WorkbookEvaluator.*

  private def num(i: Int): CellValue = CellValue.Number(BigDecimal(i))
  private val s = SheetName.unsafe("S")
  private val a1 = ARef.from0(0, 0)
  private val a2 = ARef.from0(0, 1)
  private val a3 = ARef.from0(0, 2)
  private val b1 = ARef.from0(1, 0)
  private val b2 = ARef.from0(1, 1)

  private def cached(wb: Workbook, ref: ARef): Option[CellValue] =
    wb.sheets.headOption
      .flatMap(_.cells.get(ref))
      .map(_.value)
      .collect { case CellValue.Formula(_, Some(v), _) => v }

  private def addinBook: Workbook =
    Workbook(
      Sheet(s)
        .put(a1, CellValue.Formula("""=_xll.SomeAddin.FETCH("x",1)""", Some(num(40))))
        .put(a2, num(2))
        .put(a3, CellValue.Formula("=SUM(A1:A2)", Some(num(0))))
    )

  test("recalculate: a cached _xll. call keeps its value and a SUM reads it"):
    val result = addinBook.recalculate()
    assert(result.isClean, s"expected clean recalc, got ${result.errors.map(_.render)}")
    assertEquals(cached(result.workbook, a1), Some(num(40)), "add-in cache must survive verbatim")
    assertEquals(cached(result.workbook, a3), Some(num(42)))

  test("recalculate: the namespace matches in any case and inside a larger formula"):
    val wb = Workbook(
      Sheet(s)
        .put(a1, CellValue.Formula("""=1+_XLL.Vendor.Price("y")*2""", Some(num(7))))
        .put(b1, CellValue.Formula("=A1*10", None))
    )
    val result = wb.recalculate()
    assert(result.isClean, s"expected clean recalc, got ${result.errors.map(_.render)}")
    assertEquals(cached(result.workbook, a1), Some(num(7)))
    assertEquals(cached(result.workbook, b1), Some(num(70)))

  test("recalculate: an UNCACHED _xll. call still fails and blocks its dependents"):
    val wb = Workbook(
      Sheet(s)
        .put(a1, CellValue.Formula("""=_xll.SomeAddin.FETCH("x",1)""", None))
        .put(b1, CellValue.Formula("=A1+1", None))
    )
    val result = wb.recalculate()
    assertEquals(result.errors.map(_.ref).toSet, Set(a1, b1))

  test("recalculate: an unknown function outside the _xll. namespace is still an error"):
    val wb = Workbook(
      Sheet(s)
        .put(a1, CellValue.Formula("=SUMM(1,2)", Some(num(3))))
        .put(b1, CellValue.Formula("=A1+1", None))
    )
    val result = wb.recalculate()
    assertEquals(result.errors.map(_.ref).toSet, Set(a1, b1))

  test("recalculate: `_xll.` text inside a string of a formula that parses is evaluated normally"):
    val wb = Workbook(Sheet(s).put(b2, CellValue.Formula("""=LEN("_xll.x")""", Some(num(0)))))
    val result = wb.recalculate()
    assertEquals(cached(result.workbook, b2), Some(num(6)))

  test("addinCachedCells names each pinned add-in cell, sorted, and nothing else"):
    val wb = Workbook(
      Sheet(s)
        .put(b1, CellValue.Formula("""=_xll.A.F()""", Some(num(1))))
        .put(a1, CellValue.Formula("""=_xll.A.F()""", Some(num(1))))
        .put(a2, CellValue.Formula("""=_xll.A.F()""", None))
        .put(a3, CellValue.Formula("=SUM(A1:A2)", Some(num(0))))
    )
    assertEquals(
      SheetEvaluator.addinCachedCells(wb),
      Vector(QualifiedRef(s, a1), QualifiedRef(s, b1))
    )
