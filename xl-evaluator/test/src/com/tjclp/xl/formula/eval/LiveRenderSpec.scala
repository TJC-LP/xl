package com.tjclp.xl.formula.eval

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.Clock
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * `view --eval` pictures: the window's formulas evaluate with every formula its conditional
 * formatting reads, so a cell's paint never depends on the window asked for. Pins the reads
 * ([[CfEvaluator.reads]], never expanded) and the joint evaluation ([[LiveRender.evaluate]]).
 */
class LiveRenderSpec extends FunSuite:

  private val pink = Dxf.fill(Color.Rgb(0xffffc7ce))
  private val scale = CfRule.colorScale2(
    CfPoint(Cfvo.Min, Color.Rgb(0xffff0000)),
    CfPoint(Cfvo.Max, Color.Rgb(0xff00ff00))
  )

  private def range(a1: String): CellRange = CellRange.parse(a1).fold(fail(_), identity)
  private def aref(a1: String): ARef = ARef.parse(a1).fold(fail(_), identity)
  private def num(n: Int): CellValue = CellValue.Number(BigDecimal(n))
  private def reads(sheet: Sheet, window: String): Set[String] =
    CfEvaluator
      .reads(sheet, range(window))
      .map(r => if r.start == r.end then r.start.toA1 else r.toA1)
      .toSet

  /** B1:B10 = 1..10; A1:A10 = Bn, cached ten times too high (a file saved before B changed). */
  private val stale: Sheet =
    (1 to 10).foldLeft(Sheet("S")) { (s, i) =>
      s.put(aref(s"B$i"), num(i)).put(aref(s"A$i"), CellValue.Formula(s"B$i", Some(num(i * 10))))
    }

  test("a numeric rule reads its whole block, whatever part of it the window shows") {
    val sheet = stale.conditionalFormat(range("A1:A10"), scale)
    assertEquals(reads(sheet, "A1:B3"), Set("A1:A10"))
    assertEquals(reads(sheet, "Z1:Z3"), Set.empty[String], "a block off the window reads nothing")
  }

  test("a formula rule reads what it names at each window cell it covers, shifted") {
    val sheet = stale.conditionalFormat(range("A1:A100"), CfRule.expression("$C1>$D$1", pink))
    assertEquals(reads(sheet, "A2:B3"), Set("C2", "C3", "D1"))
  }

  test("a cell-value rule reads the cell it tests and its operands") {
    val sheet =
      stale.conditionalFormat(range("A1:A10"), CfRule.cellIs(CfOperator.GreaterThan, "$E$1", pink))
    assertEquals(reads(sheet, "A1:A2"), Set("A1", "A2", "E1"))
  }

  test("a whole-column reference stays one range; other sheets are left out") {
    val sheet = stale.conditionalFormat(
      range("A1:A10"),
      CfRule.expression("COUNTIF($C:$C,A1)+Other!Z1>1", pink)
    )
    assertEquals(reads(sheet, "A1"), Set("C1:C1048576", "A1"))
  }

  test("a value object's formula reads what it names at the block's anchor") {
    val bar = CfRule.dataBar(Color.Rgb(0xff638ec6), Cfvo.Formula("$F$1"), Cfvo.Max)
    val sheet = stale.conditionalFormat(range("A1:A10"), bar)
    assertEquals(reads(sheet, "A1:A2"), Set("A1:A10", "F1"))
  }

  test("a formula that does not parse reads nothing") {
    val sheet = stale.conditionalFormat(range("A1:A10"), CfRule.expression("=(", pink))
    assertEquals(reads(sheet, "A1:A2"), Set.empty[String])
  }

  test("the joint evaluation computes the block's formulas outside the window too") {
    val sheet = stale.conditionalFormat(range("A1:A10"), scale)
    val result = LiveRender.evaluate(sheet, range("A1:B3"), Clock.system, None)
    assertEquals(result.failures, Vector.empty)
    (1 to 10).foreach(i => assertEquals(result.values.get(aref(s"A$i")), Some(num(i)), s"A$i"))
  }

  test("without conditional formatting it is the window's own evaluation") {
    val result = LiveRender.evaluate(stale, range("A1:B3"), Clock.system, None)
    val window = SheetEvaluator.evaluateForRangePerCell(stale)(range("A1:B3"), Clock.system, None)
    assertEquals(result, window)
    assertEquals(result.values.keySet, Set(aref("A1"), aref("A2"), aref("A3")))
  }

  test("a formula the paint reads that fails is reported with the window's failures") {
    val broken = stale.put(aref("A9"), CellValue.Formula("NOSUCHFN(1)", Some(num(90))))
    val sheet = broken.conditionalFormat(range("A1:A10"), scale)
    val result =
      LiveRender.evaluate(sheet, range("A1:B3"), Clock.system, Some(Workbook(Vector(sheet))))
    assertEquals(result.failures.map(_.ref), Vector(aref("A9")))
    assertEquals(result.values.get(aref("A10")), Some(num(10)), "the rest of the block is live")
  }
