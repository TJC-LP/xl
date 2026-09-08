package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.formula.eval.DependentRecalculation
import com.tjclp.xl.formula.functions.CriteriaRangeResize
import com.tjclp.xl.formula.graph.DependencyGraph
import munit.FunSuite

/**
 * GH-631: Excel sizes SUMIF's `sum_range` and AVERAGEIF's `average_range` to the shape of `range`
 * from their upper-left cell — `SUMIF(A1:A10, ">0", C1)` sums C1:C10 — and accepts a single cell
 * wherever a range is expected. xl used to refuse the single cell at parse time
 * (`InvalidArguments(SUMIF,0,range,1 arguments)`) and, given a smaller range, to fail with a
 * dimension error. Values confirmed against LibreOffice (the same rule): with A1:A3 = 1,2,3 and
 * C1:C3 = 10,20,30, `SUMIF(A1:A3,">1",C1)` is 50, `…,C2)` is 30, `AVERAGEIF(A1:A3,">1",C1)` is 25,
 * `SUMIFS(C1,A1:A3,">1")` is #VALUE! (the *IFS family never resizes), `COUNTIF(A1,">0")` is 1.
 *
 * The graph half: dependency extraction records the RESIZED area, so an edit to C3 dirties the
 * formula and its cache is refreshed rather than left stale.
 */
class SumIfResizeSpec extends FunSuite:

  private def num(n: Int): CellValue = CellValue.Number(BigDecimal(n))

  private val data = SheetName.unsafe("Data")

  /** A1:A3 = 1,2,3; C1:C3 = 10,20,30; D1:D3 = 100,200,300. */
  private val sheet: Sheet =
    (1 to 3).foldLeft(Sheet("Data")) { (s, i) =>
      s.put(ARef.from0(0, i - 1), num(i))
        .put(ARef.from0(2, i - 1), num(i * 10))
        .put(ARef.from0(3, i - 1), num(i * 100))
    }

  private def eval(formula: String): CellValue =
    sheet
      .evaluateFormula(formula, workbook = Some(Workbook(sheet)))
      .fold(e => fail(e.message), identity)

  test("SUMIF sizes a single-cell sum_range to range from its upper-left cell") {
    assertEquals(eval("=SUMIF(A1:A3,\">1\",C1)"), num(50))
    assertEquals(eval("=SUMIF(A1:A3,\">1\",C2)"), num(30)) // C2:C4 = 20, 30, blank
    assertEquals(eval("=SUMIF(A1,\">0\",C1)"), num(10))
  }

  test("SUMIF sizes a smaller and a larger sum_range alike; AVERAGEIF follows the same rule") {
    assertEquals(eval("=SUMIF(A1:A3,\">1\",C1:C2)"), num(50))
    assertEquals(eval("=SUMIF(A1:A3,\">1\",C1:D5)"), num(50))
    assertEquals(eval("=AVERAGEIF(A1:A3,\">1\",C1)"), num(25))
    assertEquals(eval("=AVERAGEIF(A1:A3,\">1\",D1:D1)"), num(250))
  }

  // ===== the review repro: a used range that starts below row 1, whole columns, whole rows =====
  // LibreOffice: SUMIF(A:A,">0",C4)=4300, SUMIF(A:A,">0",C:C)=10, SUMIF(A3:A4,">0",C4)=30,
  // SUMIF(A:A,">0",C1)=10, AVERAGEIF(A:A,">0",C4)=2150, SUMIF(C:C,">0",A4)=0,
  // AVERAGEIF(A:A,">0",C:C)=10; on the row sheet SUMIF(1:1,">0",D3)=30, AVERAGEIF(1:1,">0",D3)=15,
  // SUMIF(1:1,">0",3:3)=0, SUMIF(1:1,">0",A3)=0.

  /** A3=1, A4=2; C4=10, C5=20, C6=300, C7=4000 — nothing in rows 1-2. */
  private val below: Sheet = Sheet("Off")
    .put(ref"A3", num(1))
    .put(ref"A4", num(2))
    .put(ref"C4", num(10))
    .put(ref"C5", num(20))
    .put(ref"C6", num(300))
    .put(ref"C7", num(4000))

  private def evalBelow(formula: String): CellValue =
    below
      .evaluateFormula(formula, workbook = Some(Workbook(below)))
      .fold(e => fail(e.message), identity)

  test("pairing uses the unconstrained origins when the used range starts below row 1") {
    // A:A is walked from row 3 (the used area) but A3 pairs with C6 — row 3 from C4's origin —
    // not with C4; the first cut of this fix offset from the clipped A3 and summed 30
    assertEquals(evalBelow("=SUMIF(A:A,\">0\",C4)"), num(4300))
    assertEquals(evalBelow("=AVERAGEIF(A:A,\">0\",C4)"), num(2150))
    assertEquals(evalBelow("=SUMIF(A3:A4,\">0\",C4)"), num(30))
    assertEquals(evalBelow("=SUMIF(A:A,\">0\",C1)"), num(10))
    assertEquals(evalBelow("=SUMIF(C:C,\">0\",A4)"), num(0))
  }

  test("two whole columns pair row by row") {
    assertEquals(evalBelow("=SUMIF(A:A,\">0\",C:C)"), num(10))
    assertEquals(evalBelow("=AVERAGEIF(A:A,\">0\",C:C)"), num(10))
  }

  test("the whole-row analogue: a used range that starts right of column A") {
    // C1=1, D1=2; F3=10, G3=20, H3=300 — 1:1 is walked from column C, C1 pairs with F3 (two
    // columns from D3's origin), D1 with G3
    val rows = Sheet("Row")
      .put(ref"C1", num(1))
      .put(ref"D1", num(2))
      .put(ref"F3", num(10))
      .put(ref"G3", num(20))
      .put(ref"H3", num(300))
    def evalRows(formula: String): CellValue =
      rows
        .evaluateFormula(formula, workbook = Some(Workbook(rows)))
        .fold(e => fail(e.message), identity)
    assertEquals(evalRows("=SUMIF(1:1,\">0\",D3)"), num(30))
    assertEquals(evalRows("=AVERAGEIF(1:1,\">0\",D3)"), num(15))
    assertEquals(evalRows("=SUMIF(1:1,\">0\",3:3)"), num(0))
    assertEquals(evalRows("=SUMIF(1:1,\">0\",A3)"), num(0))
  }

  test("a pairing past the grid edge contributes nothing, as LibreOffice computes it") {
    assertEquals(eval("=SUMIF(A1:A3,\">1\",Z1048575)"), num(0))
    assertEquals(
      CriteriaRangeResize
        .resize(CellRange(ref"Z1048575", ref"Z1048575"), CellRange(ref"A1", ref"A3"))
        .toA1,
      "Z1048575:Z1048576"
    )
  }

  test("the *IFS family never resizes: a shape mismatch is Excel's #VALUE! error value") {
    assertEquals(eval("=SUMIFS(C1,A1:A3,\">1\")"), CellValue.Error(CellError.Value))
    assertEquals(eval("=SUMIFS(C1:C2,A1:A3,\">1\")"), CellValue.Error(CellError.Value))
    assertEquals(eval("=AVERAGEIFS(C1,A1:A3,\">1\")"), CellValue.Error(CellError.Value))
    assertEquals(eval("=MAXIFS(C1,A1:A3,\">1\")"), CellValue.Error(CellError.Value))
    assertEquals(eval("=MINIFS(C1,A1:A3,\">1\")"), CellValue.Error(CellError.Value))
    assertEquals(eval("=COUNTIFS(A1:A3,\">1\",C1,\">0\")"), CellValue.Error(CellError.Value))
    assertEquals(eval("=IFERROR(SUMIFS(C1,A1:A3,\">1\"),-1)"), num(-1))
    // matching shapes keep working
    assertEquals(eval("=SUMIFS(C1:C3,A1:A3,\">1\")"), num(50))
  }

  test("a single cell is accepted in every range slot and prints back as written") {
    assertEquals(eval("=COUNTIF(A1,\">0\")"), num(1))
    assertEquals(eval("=COUNTIF(A2,\">5\")"), num(0))
    assertEquals(eval("=VLOOKUP(1,A1,1,FALSE)"), num(1))
    List(
      "=SUMIF(A1:A3, \">1\", C1)",
      "=COUNTIF(A1, \">0\")",
      "=SUMIF($A$1:$A$3, \">1\", $C$1)",
      "=SUMIF(Data!A1:A3, \">1\", Data!C1)",
      "=SUM(A1:A1)"
    ).foreach { source =>
      assertEquals(FormulaParser.parse(source).map(FormulaPrinter.print(_)), Right(source))
    }
    FormulaParser.parse("=COUNTIF(A1,\">0\")") match
      case Right(TExpr.Call(_, (TExpr.RangeLocation.Local(range, form), _))) =>
        assertEquals(range, CellRange(ref"A1", ref"A1"))
        assertEquals(form, RangeForm.Cell)
      case other => fail(s"expected COUNTIF over a 1x1 Local range, got $other")
  }

  test("a single-cell range slot drags like the cell it is") {
    val dragged =
      for expr <- FormulaParser.parse("=SUMIF($A$1:$A$3,\">1\",C1)")
      yield FormulaPrinter.print(FormulaShifter.shift(expr, 1, 2))
    assertEquals(dragged, Right("=SUMIF($A$1:$A$3, \">1\", D3)"))
    val voided =
      for expr <- FormulaParser.parse("=COUNTIF(A1,\">0\")")
      yield FormulaPrinter.print(FormulaShifter.shift(expr, 0, -1))
    assertEquals(voided, Right("=COUNTIF(#REF!, \">0\")"))
  }

  test("dependency extraction records the resized sum_range") {
    def deps(formula: String): Set[String] =
      FormulaParser
        .parse(formula)
        .fold(e => fail(e.toString), DependencyGraph.extractDependencies(_))
        .map(_.toA1)
    assertEquals(deps("=SUMIF(A1:A3,\">1\",C1)"), Set("A1", "A2", "A3", "C1", "C2", "C3"))
    assertEquals(deps("=AVERAGEIF(A1:A2,\">1\",D3)"), Set("A1", "A2", "D3", "D4"))
    assertEquals(deps("=SUMIF(A1:A3,\">1\",C1:D5)"), Set("A1", "A2", "A3", "C1", "C2", "C3"))
    // the *IFS family reads what it says
    assertEquals(deps("=SUMIFS(C1,A1:A3,\">1\")"), Set("A1", "A2", "A3", "C1"))
    val index = DependencyGraph.fromWorkbookDependencyIndex(
      Workbook(sheet.put(ref"F1", CellValue.Formula("SUMIF(A1:A3,\">1\",C1)", None)))
    )
    val ranges = index.rangeDependents.getOrElse(data, Vector.empty).map(_.range.toA1).toSet
    assert(ranges.contains("C1:C3"), ranges.toString)
  }

  test("an edit to a cell the resized sum_range reads refreshes the SUMIF cache") {
    val wb =
      Workbook(sheet.put(ref"F1", CellValue.Formula("SUMIF(A1:A3,\">1\",C1)", Some(num(50)))))
    val edited = wb.put(
      sheet
        .put(ref"F1", CellValue.Formula("SUMIF(A1:A3,\">1\",C1)", Some(num(50))))
        .put(ref"C3", num(300))
    )
    val result =
      DependentRecalculation.recalculateAfterEdit(edited, data, Set(ref"C3"), Clock.system)
    val refreshed = result.workbook(data).fold(e => fail(e.message), identity)(ref"F1").value
    assertEquals(refreshed, CellValue.Formula("SUMIF(A1:A3,\">1\",C1)", Some(num(320))))
  }
