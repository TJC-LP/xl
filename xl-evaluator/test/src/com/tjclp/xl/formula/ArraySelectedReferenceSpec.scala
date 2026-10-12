package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue, FormulaKind}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * GH-711: in an array context (a positionless formula, a CSE record, a conditional-format rule) the
 * reference IF, IFS, CHOOSE and SWITCH select is a reference, as in a plain cell (#683): an
 * aggregate, AND, OR or ROWS receives it whole and folds it by Excel's reference rule within the
 * used range; under SUMPRODUCT the selected whole column is trimmed to the shared used extent like
 * any written whole column (GH-192).
 */
class ArraySelectedReferenceSpec extends FunSuite:

  private def num(n: Double): CellValue = CellValue.Number(BigDecimal(n))

  /** A1:A3 = 4,5,6 and B1:B3 = 1,2,3; nothing else (the issue's book). */
  private val sheet: Sheet =
    Sheet(SheetName.unsafe("S"))
      .put(ref"A1", num(4))
      .put(ref"A2", num(5))
      .put(ref"A3", num(6))
      .put(ref"B1", num(1))
      .put(ref"B2", num(2))
      .put(ref"B3", num(3))

  /** `formula` with no cell position: an array formula, as `xl eval` evaluates it. */
  private def positionless(formula: String, on: Sheet = sheet): XLResult[CellValue] =
    on.evaluateFormula(formula, Clock.system, Some(Workbook(Vector(on))), None)

  /** `formula` as a one-cell CSE record at D1, recalculated; its cached value. */
  private def cse(formula: String, on: Sheet = sheet): Option[CellValue] =
    recalculatedAtD1(formula, FormulaKind.ArrayFormula(CellRange(ref"D1", ref"D1")), on)

  /** `formula` as a plain cell's (legacy) formula at D1, recalculated; its cached value. */
  private def plain(formula: String, on: Sheet): Option[CellValue] =
    recalculatedAtD1(formula, FormulaKind.Normal(), on)

  private def recalculatedAtD1(formula: String, kind: FormulaKind, on: Sheet): Option[CellValue] =
    val placed = on.put(ref"D1", CellValue.Formula(formula, None, kind))
    Workbook(Vector(on))
      .put(placed)
      .recalculate()
      .workbook
      .sheets
      .headOption
      .flatMap(_.cells.get(ref"D1"))
      .map(_.value)
      .collect { case CellValue.Formula(_, Some(cached), _) => cached }

  test("AND/OR over a selected whole column follow the reference rule: blanks are ignored") {
    assertEquals(positionless("=AND(B:B)"), Right(CellValue.Bool(true)))
    assertEquals(positionless("=AND(IF(TRUE,B:B,A:A))"), Right(CellValue.Bool(true)))
    assertEquals(positionless("=AND(CHOOSE(1,B:B,A:A))"), Right(CellValue.Bool(true)))
    assertEquals(positionless("=AND(IFS(FALSE,A:A,TRUE,B:B))"), Right(CellValue.Bool(true)))
    assertEquals(positionless("=AND(SWITCH(2,1,A:A,B:B))"), Right(CellValue.Bool(true)))
    assertEquals(positionless("=OR(IF(TRUE,B1:B5,A:A))"), Right(CellValue.Bool(true)))
    assertEquals(cse("AND(IF(TRUE,B:B,A:A))"), Some(CellValue.Bool(true)))
  }

  test("an aggregate over a selected whole column equals the aggregate over the column") {
    List("SUM", "COUNT", "COUNTA", "COUNTBLANK", "AVERAGE", "MIN", "MAX", "MEDIAN", "STDEV")
      .foreach { fn =>
        val direct = positionless(s"=$fn(B:B)")
        assert(direct.isRight, s"$fn(B:B): $direct")
        List("IF(TRUE,B:B,A:A)", "CHOOSE(2,A:A,B:B)", "IFS(TRUE,B:B)").foreach { selected =>
          assertEquals(positionless(s"=$fn($selected)"), direct, s"$fn($selected)")
        }
      }
    assertEquals(positionless("=COUNTA(IF(TRUE,B:B,A:A))"), Right(num(3)))
    assertEquals(positionless("=SUM(IF(TRUE,B:B,A:A))"), Right(num(6)))
    assertEquals(cse("SUM(IF(TRUE,B:B,A:A))"), Some(num(6)))
    assertEquals(cse("COUNTA(IF(TRUE,B:B,A:A))"), Some(num(3)))
    // every blank row of the column counts, as in Excel: 1048576 cells, three filled
    assertEquals(positionless("=COUNTBLANK(B:B)"), Right(num(1048573)))
    assertEquals(positionless("=COUNTBLANK(IF(TRUE,B:B,A:A))"), Right(num(1048573)))
    assertEquals(cse("COUNTBLANK(IF(TRUE,B:B,A:A))"), Some(num(1048573)))
  }

  test("AND, OR and the aggregates over a selected whole row follow the reference rule") {
    // row 2 holds A2=5 and B2=2, row 1 A1=4 and B1=1; the other 16382 cells of each are blank
    assertEquals(positionless("=AND(2:2)"), Right(CellValue.Bool(true)))
    assertEquals(positionless("=AND(IF(TRUE,2:2,1:1))"), Right(CellValue.Bool(true)))
    assertEquals(positionless("=OR(CHOOSE(1,2:2,1:1))"), Right(CellValue.Bool(true)))
    // the CSE record sits at D1, so it selects between rows 2 and 3 (row 1 would be circular)
    assertEquals(cse("AND(IF(TRUE,2:2,3:3))"), Some(CellValue.Bool(true)))
    assertEquals(positionless("=SUM(IF(TRUE,2:2,1:1))"), Right(num(7)))
    // every blank cell of the row counts, as in Excel and LibreOffice: 16384 cells, two filled
    assertEquals(positionless("=COUNTBLANK(1:1)"), Right(num(16382)))
    assertEquals(positionless("=COUNTBLANK(IF(TRUE,1:1,2:2))"), Right(num(16382)))
    assertEquals(cse("COUNTBLANK(IF(TRUE,2:2,3:3))"), Some(num(16382)))
  }

  test("COUNTBLANK counts the blanks above a used range that starts below row 1") {
    // only B5:B7 hold values: B1:B4 lie before the used range and B8 onwards after it
    val lower = Sheet(SheetName.unsafe("L"))
      .put(ref"B5", num(1))
      .put(ref"B6", num(2))
      .put(ref"B7", num(3))
    List("COUNTBLANK(B:B)", "COUNTBLANK(IF(TRUE,B:B,A:A))", "COUNTBLANK(CHOOSE(2,A:A,B:B))")
      .foreach { formula =>
        assertEquals(positionless("=" + formula, lower), Right(num(1048573)), formula)
        assertEquals(cse(formula, lower), Some(num(1048573)), formula)
        assertEquals(plain(formula, lower), Some(num(1048573)), formula)
      }
    // a column the used range does not cross is blank throughout; two columns, all but three
    assertEquals(positionless("=COUNTBLANK(A:A)", lower), Right(num(1048576)))
    assertEquals(positionless("=COUNTBLANK(B:C)", lower), Right(num(2097149)))
    // the other aggregates and AND skip those blanks
    assertEquals(positionless("=SUM(IF(TRUE,B:B,A:A))", lower), Right(num(6)))
    assertEquals(positionless("=COUNTA(IF(TRUE,B:B,A:A))", lower), Right(num(3)))
    assertEquals(positionless("=AND(IF(TRUE,B:B,A:A))", lower), Right(CellValue.Bool(true)))
  }

  test("COUNTBLANK over two whole columns, one union and a whole column OFFSET returns") {
    // A:B holds 2 * 1048576 cells, six filled (LibreOffice: 2097146)
    assertEquals(positionless("=COUNTBLANK(A:B)"), Right(num(2097146)))
    assertEquals(cse("COUNTBLANK(A:B)"), Some(num(2097146)))
    // COUNTBLANK takes one range: a union of two areas is #VALUE!, as in Excel (LibreOffice
    // counts both areas), now that each area folds through the trimmed-blank path too
    assertEquals(positionless("=COUNTBLANK((A:A,B:B))"), Right(CellValue.Error(CellError.Value)))
    // OFFSET's whole column reaches COUNTBLANK as a reference (LibreOffice: 1048573)
    assertEquals(positionless("=COUNTBLANK(OFFSET(A:A,0,1))"), Right(num(1048573)))
    assertEquals(positionless("=COUNTBLANK(IF(TRUE,OFFSET(A:A,0,1),A:A))"), Right(num(1048573)))
    assertEquals(cse("COUNTBLANK(OFFSET(A:A,0,1))"), Some(num(1048573)))
    assertEquals(positionless("=SUM(OFFSET(A:A,0,1))"), Right(num(6)))
  }

  test("ROWS of a selected whole column is the column's height") {
    assertEquals(positionless("=ROWS(IF(TRUE,B:B,A:A))"), Right(num(1048576)))
  }

  test("an array condition still lifts element-wise") {
    // A1:A3 > 4 selects A2, A3 element-wise, 0 for A1
    assertEquals(positionless("=SUM(IF(A1:A3>4,A1:A3,0))"), Right(num(11)))
    assertEquals(positionless("=AND(IF(A1:A3>4,TRUE,FALSE))"), Right(CellValue.Bool(false)))
    assertEquals(cse("SUM(IF(A1:A3>4,B1:B3,0))"), Some(num(5)))
  }

  test("a reference function lifted over an array stays the array of its values") {
    // INDEX over an array index selects B1 and B3, not one reference read at its top-left
    assertEquals(positionless("=SUM(INDEX(B1:B3,{1,3}))"), Right(num(4)))
    assertEquals(positionless("=SUM(IF(TRUE,INDEX(B1:B3,{1,3}),0))"), Right(num(4)))
    // over scalars INDEX returns the reference, folded whole within the used range
    assertEquals(positionless("=SUM(INDEX(A:B,0,2))"), Right(num(6)))
    assertEquals(positionless("=AND(INDEX(A:B,0,2))"), Right(CellValue.Bool(true)))
  }

  test("array arithmetic over a selected reference reads its values") {
    assertEquals(positionless("=SUM(IF(TRUE,B1:B3,A1:A3)*2)"), Right(num(12)))
    assertEquals(cse("SUM(IF(TRUE,B1:B3,A1:A3)*2)"), Some(num(12)))
  }

  test("SUMPRODUCT trims a selected whole column like a written one (GH-192)") {
    List(
      "IF(TRUE,B:B,A:A)" -> "B:B",
      "CHOOSE(2,A:A,B:B)" -> "B:B",
      "IFS(FALSE,A:A,TRUE,B:B)" -> "B:B",
      "SWITCH(1,1,B:B,A:A)" -> "B:B"
    ).foreach { (selected, column) =>
      List("SUMPRODUCT(%s)", "SUMPRODUCT(%s*2)", "SUMPRODUCT(--(%s=0))", "SUMPRODUCT(%s,A:A)")
        .foreach { shape =>
          val expected = positionless("=" + shape.format(column))
          assert(expected.isRight, s"${shape.format(column)}: $expected")
          assertEquals(positionless("=" + shape.format(selected)), expected, shape.format(selected))
        }
    }
    assertEquals(positionless("=SUMPRODUCT(IF(TRUE,B:B,A:A))"), Right(num(6)))
    // the selected column and a written one take the same bounds, so their sizes match
    assertEquals(positionless("=SUMPRODUCT(IF(TRUE,B:B,A:A),A:A)"), Right(num(32)))
    // an array condition broadcasts against branches trimmed to the same extent
    assertEquals(positionless("=SUMPRODUCT(IF(A:A>4,B:B,0))"), Right(num(5)))
    // a whole-column condition and two whole-column branches, all trimmed to one size: A1 (4 is
    // not > 4) and B2, B3, so 4+2+3 (LibreOffice: 9)
    assertEquals(positionless("=SUMPRODUCT(IF(A:A>4,B:B,A:A))"), Right(num(9)))
  }

  test("SUMPRODUCT trims a selector's condition and branches when the columns' data differ") {
    // A1:A5 = 1..5, B1:B2 = 10,20: B:B is trimmed to A's five rows, its blanks reading 0
    val ragged = (1 to 5)
      .foldLeft(Sheet(SheetName.unsafe("R")))((s, i) =>
        s.put(ARef.from0(0, i - 1), num(i.toDouble))
      )
      .put(ref"B1", num(10))
      .put(ref"B2", num(20))
    // LibreOffice: 21, 41 and 50
    assertEquals(positionless("=SUMPRODUCT(IF(A:A>1,B:B,A:A))", ragged), Right(num(21)))
    assertEquals(positionless("=SUMPRODUCT(IF(A:A>1,B:B,A:A),A:A)", ragged), Right(num(41)))
    assertEquals(positionless("=SUMPRODUCT(CHOOSE(1,B:B,A:A),A:A)", ragged), Right(num(50)))
  }

  test("drift: the selector flag marks IF, IFS, CHOOSE and SWITCH, each a reference function") {
    val flagged =
      com.tjclp.xl.formula.functions.FunctionRegistry.all.filter(_.flags.selector).map(_.name).toSet
    assertEquals(flagged, Set("IF", "IFS", "CHOOSE", "SWITCH"))
    // a reference position routes each to the reference it selects (Evaluator.referenceArgument)
    assert(flagged.subsetOf(com.tjclp.xl.formula.eval.Evaluator.referenceFunctions), flagged)
  }

  test("SUMPRODUCT bounds a selected column on another sheet with the shared extent") {
    val other = (1 to 5).foldLeft(Sheet(SheetName.unsafe("Other"))) { (s, i) =>
      s.put(ARef.from0(1, i - 1), num(i.toDouble))
    }
    val wb = Workbook(Vector(sheet, other))
    def eval(formula: String) = sheet.evaluateFormula(formula, Clock.system, Some(wb), None)
    assertEquals(eval("=SUMPRODUCT(IF(TRUE,Other!B:B,0),B:B)"), Right(num(14)))
    assertEquals(
      eval("=SUMPRODUCT(IF(TRUE,Other!B:B,0),B:B)"),
      eval("=SUMPRODUCT(Other!B:B,B:B)")
    )
    assertEquals(eval("=SUM(IF(TRUE,Other!B:B,0))"), Right(num(15)))
  }

  test("a conditional-format rule summing a selected whole column paints by its value") {
    val withRule = sheet
      .put(ref"C1", num(1))
      .copy(conditionalFormats =
        Vector(
          ConditionalFormat.Rules(
            Vector(CellRange(ref"A1", ref"A3")),
            Vector(
              CfRule.Expression(
                "SUM(IF($C$1>0,$B:$B,0))>5",
                Some(Dxf.fill(Color.Rgb(0xffff0000))),
                1,
                false
              )
            )
          )
        )
      )
    val ev = withRule.evaluateConditionalFormats(
      CellRange(ref"A1", ref"D5"),
      Some(Workbook(Vector(withRule))),
      Clock.system
    )
    assertEquals(
      ev.overlay.cells.keySet.map(_.toA1),
      Set("A1", "A2", "A3"),
      ev.toString
    )
    // a sum of 6 over the selected column, not an error and not a million-row array
    assertEquals(
      positionless("=SUM(IF($C$1>0,$B:$B,0))>5", withRule),
      Right(CellValue.Bool(true))
    )
  }

  test("a selected whole column on an empty sheet folds to nothing") {
    val empty = Sheet(SheetName.unsafe("E"))
    assertEquals(positionless("=SUM(IF(TRUE,B:B,A:A))", empty), Right(num(0)))
    assertEquals(positionless("=SUMPRODUCT(IF(TRUE,B:B,A:A))", empty), Right(num(0)))
    assertEquals(positionless("=COUNTBLANK(IF(TRUE,B:B,A:A))", empty), Right(num(1048576)))
    assertEquals(
      positionless("=AND(IF(TRUE,B:B,A:A))", empty),
      Right(CellValue.Error(CellError.Value))
    )
  }

  test("computed scalar reference slots are classified before reference dispatch") {
    assertEquals(positionless("=COUNTBLANK(INDEX(B:B,SUM(0)))"), Right(num(1048573)))
    assertEquals(positionless("=ROWS(INDEX(B:B,SUM(0)))"), Right(num(1048576)))
    assertEquals(positionless("=SUM(INDEX(B1:B3,SUM(1,1)))"), Right(num(2)))
    assertEquals(positionless("=ROWS(INDEX(B:B,{0}))"), Right(num(1048576)))
    assertEquals(positionless("=SUM(INDEX(B1:B3,{1,3}))"), Right(num(4)))
  }
