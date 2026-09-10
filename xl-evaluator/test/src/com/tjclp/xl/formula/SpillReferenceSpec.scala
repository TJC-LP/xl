package com.tjclp.xl.formula

import com.tjclp.xl.XLResult
import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.{CellError, CellValue, FormulaKind}
import com.tjclp.xl.formula.eval.SheetEvaluator.*
import com.tjclp.xl.formula.functions.FunctionSpecs
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.syntax.*
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * GH-655: `x#` — Excel 365's spill reference, stored as `_xlfn.ANCHORARRAY(x)`. The value is the
 * whole array the anchor cell's formula spills: the file's `<f t="array" ref>` span read through
 * its cached cells when Excel wrote one, else the anchor's formula evaluated as an array. A cell
 * that is not a spill anchor — a constant, a scalar formula, a spilled non-anchor value — is
 * `#REF!`. The parser accepts `x#` and `ANCHORARRAY(x)`; the printer emits `x#`.
 */
class SpillReferenceSpec extends FunSuite:

  private val n = (v: Int) => CellValue.Number(BigDecimal(v))

  /**
   * A1 spills SORT(B1:B3) without a cache (an xl-authored dynamic formula); C1 carries the file's
   * array record for SEQUENCE(3) with cached spill cells; D1 the same for a function xl does not
   * evaluate (SORTBY), so only the cached spill can answer; E1 is a constant, E2 a scalar formula.
   */
  private val sheet = Sheet("Data")
    .put(ref"B1", n(3))
    .put(ref"B2", n(1))
    .put(ref"B3", n(2))
    .put(ref"A1", CellValue.Formula("SORT(B1:B3)"))
    .put(
      ref"C1",
      CellValue.Formula(
        "SEQUENCE(3)",
        Some(n(1)),
        FormulaKind.ArrayFormula(CellRange(ref"C1", ref"C3"))
      )
    )
    .put(ref"C2", n(2))
    .put(ref"C3", n(3))
    .put(
      ref"D1",
      CellValue.Formula(
        "SORTBY(B1:B3,B1:B3)",
        Some(n(10)),
        FormulaKind.ArrayFormula(CellRange(ref"D1", ref"D3"))
      )
    )
    .put(ref"D2", n(20))
    .put(ref"D3", n(30))
    .put(ref"E1", n(5))
    .put(ref"E2", CellValue.Formula("B1*2"))

  private val wb = Workbook(sheet, Sheet("Other").put(ref"A1", n(1)))
    .withDefinedName("spill", "Data!$A$1")

  private def value(formula: String): XLResult[CellValue] =
    sheet.evaluateFormula(formula, Clock.system, Some(wb), Some(ref"H1"))

  private def number(formula: String): BigDecimal =
    value(formula) match
      case Right(CellValue.Number(v)) => v
      case other => fail(s"$formula: expected a number, got $other")

  private def isRef(result: XLResult[CellValue]): Boolean =
    result match
      case Left(err) => err.toString.contains("#REF!") || err.toString.contains("Ref")
      case Right(CellValue.Error(CellError.Ref)) => true
      case _ => false

  test("an uncached dynamic formula's spill is its array evaluation at the anchor") {
    assertEquals(number("=SUM(A1#)"), BigDecimal(6))
    assertEquals(number("=ROWS(A1#)"), BigDecimal(3))
    assertEquals(number("=SUM(A1#*10)"), BigDecimal(60), "the spill enters array arithmetic")
    assertEquals(number("=A1#"), BigDecimal(1), "a scalar position collapses to the top-left cell")
  }

  test("standalone, x# spills the anchor's array") {
    val (s, range) = sheet.evaluateArrayFormula("=A1#", ref"G1").toOption.get
    assertEquals(range, CellRange(ref"G1", ref"G3"))
    assertEquals(s(ref"G1").value, n(1))
    assertEquals(s(ref"G2").value, n(2))
    assertEquals(s(ref"G3").value, n(3))
  }

  test("the file's array record is Excel's extent, read through the cached spill cells") {
    assertEquals(number("=SUM(C1#)"), BigDecimal(6))
    // a spill from a function xl does not evaluate still reads — the caches carry it
    assertEquals(number("=SUM(D1#)"), BigDecimal(60))
    assertEquals(number("=ROWS(D1#)"), BigDecimal(3))
  }

  test("a constant, a scalar formula and a spilled non-anchor cell are #REF!") {
    assert(isRef(value("=SUM(E1#)")), value("=SUM(E1#)").toString)
    assert(isRef(value("=SUM(E2#)")), value("=SUM(E2#)").toString)
    assert(isRef(value("=SUM(C2#)")), value("=SUM(C2#)").toString)
    assert(isRef(value("=SUM(Z9#)")), value("=SUM(Z9#)").toString)
  }

  test("a sheet-qualified anchor and a defined name bound to the anchor spill the same array") {
    val other = wb.sheets.find(_.name.value == "Other").getOrElse(fail("no Other sheet"))
    assertEquals(
      other.evaluateFormula("=SUM(Data!A1#)", Clock.system, Some(wb), Some(ref"A2")),
      Right(n(6))
    )
    assertEquals(number("=SUM(spill#)"), BigDecimal(6))
    assertEquals(number("=SUM(Data!spill#)"), BigDecimal(6))
  }

  test("the stored spelling ANCHORARRAY(x) evaluates and prints as x#") {
    assertEquals(number("=SUM(ANCHORARRAY(A1))"), BigDecimal(6))
    assertEquals(number("=SUM(_xlfn.ANCHORARRAY(A1))"), BigDecimal(6))
    def printed(f: String): String =
      FormulaParser.parse(f).map(FormulaPrinter.print(_)).fold(e => fail(e.toString), identity)
    assertEquals(printed("=ANCHORARRAY(A1)"), "=A1#")
    assertEquals(printed("=SUM(A1#)"), "=SUM(A1#)")
    assertEquals(printed("=Data!A1#"), "=Data!A1#")
    assertEquals(printed("='My Sheet'!A1#"), "='My Sheet'!A1#")
    assertEquals(printed("=$A$1#"), "=$A$1#")
    assertEquals(printed("=spill#"), "=spill#")
    assertEquals(printed("=A1#*2"), "=A1#*2")
    assertEquals(printed("=@A1#"), "=@A1#")
    assertEquals(printed("=[1]Book!A1#"), "=[1]Book!A1#")
    // a call whose argument is not a shape `#` can follow keeps the call spelling
    assertEquals(printed("=ANCHORARRAY(A1:A3)"), "=ANCHORARRAY(A1:A3)")
    val fileForm = FormulaParser.parse("=SUM(A1#)").fold(e => fail(e.toString), identity)
    assertEquals(FormulaPrinter.printFileForm(fileForm), "SUM(A1#)")
  }

  test("parse ∘ print = id over the spill-reference spellings") {
    List(
      "=SUM(A1#)",
      "=Data!A1#*2",
      "='My Sheet'!$A$1#",
      "=spill#",
      "=@A1#",
      "=A1#%",
      "=INDEX(A1:A3,1)+A1#",
      "=ANCHORARRAY(A1:A3)"
    ).foreach { f =>
      val parsed = FormulaParser.parse(f)
      assert(parsed.isRight, s"$f: $parsed")
      val reprinted = parsed.map(FormulaPrinter.print(_))
      assertEquals(reprinted.flatMap(FormulaParser.parse), parsed, s"$f via $reprinted")
    }
  }

  test("# applies to one cell or name only: a range, a value or a call before it is an error") {
    assert(FormulaParser.parse("=A1:A3#").isLeft)
    assert(FormulaParser.parse("=5#").isLeft)
    assert(FormulaParser.parse("=SUM(A1)#").isLeft)
    assert(FormulaParser.parse("=\"x\"#").isLeft)
    assert(FormulaParser.parse("=#").isLeft)
  }

  test("the anchor's spill is not statically knowable: ANCHORARRAY is a dynamic-dependency call") {
    assert(FunctionSpecs.anchorArray.flags.dynamicDeps)
  }

  test("a spill reference that reaches its own anchor stops at the depth guard, never overflows") {
    val circular = Sheet("Loop").put(ref"A1", CellValue.Formula("ROWS(A1#)"))
    circular.evaluateFormula("=SUM(A1#)", Clock.system, None, Some(ref"B1")) match
      case Left(_) => ()
      case Right(v) => fail(s"a circular spill reference should not evaluate, got $v")
  }
