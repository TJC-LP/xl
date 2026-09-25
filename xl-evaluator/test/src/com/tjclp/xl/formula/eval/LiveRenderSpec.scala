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
  private def reads(
    sheet: Sheet,
    window: String,
    workbook: Option[Workbook] = None
  ): Set[String] =
    CfEvaluator
      .reads(sheet, range(window), workbook)
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

  test("named CF reads resolve aliases, local scope and case-variant sheet qualifiers") {
    val sheet = stale.conditionalFormat(range("A1:A3"), CfRule.expression("thresholdAlias>5", pink))
    val wb = Workbook(Vector(sheet))
      .withDefinedName("threshold", "S!$Z$1")
      .withDefinedName("threshold", "s!$C$1+SUM(s!$D$1:$D$2)", sheet.name)
      .fold(err => fail(err.message), identity)
      .withDefinedName("thresholdAlias", "S!threshold")
    assertEquals(reads(sheet, "A1:A3", Some(wb)), Set("C1", "D1:D2"))
  }

  test("named ranges and named value-object formulas remain symbolic reads") {
    val bar = CfRule.dataBar(Color.Rgb(0xff638ec6), Cfvo.Formula("minimum"), Cfvo.Max)
    val sheet = stale
      .conditionalFormat(range("A1:A3"), CfRule.expression("COUNTIF(drivers,A1)>0", pink))
      .conditionalFormat(range("A1:A3"), bar)
    val wb = Workbook(Vector(sheet))
      .withDefinedName("drivers", "S!$C:$C")
      .withDefinedName("minimum", "S!$F$1")
    assertEquals(reads(sheet, "A1", Some(wb)), Set("C1:C1048576", "A1", "A1:A3", "F1"))
  }

  test("a named formula outside the viewport is live before conditional formats paint") {
    val sheet = stale
      .put(aref("C1"), CellValue.Formula("B1+9", Some(num(0))))
      .conditionalFormat(range("A1:A3"), CfRule.expression("threshold>5", pink))
    val wb = Workbook(Vector(sheet)).withDefinedName("threshold", "S!$C$1")
    def paint(window: String): CfOverlay =
      val live = LiveRender.evaluate(sheet, range(window), Clock.system, Some(wb))
      assertEquals(live.failures, Vector.empty)
      assertEquals(live.values.get(aref("C1")), Some(num(10)))
      val rendered = live.values.foldLeft(sheet) { case (s, (at, value)) => s.put(at, value) }
      rendered
        .evaluateConditionalFormats(range("A1:A3"), Some(wb.put(rendered)), Clock.system)
        .overlay
    val narrow = paint("A1:A3")
    assertEquals(narrow.cells.keySet, Set(aref("A1"), aref("A2"), aref("A3")))
    assertEquals(narrow, paint("A1:C3"))
  }

  test("a failing named CF precedent outside the viewport is reported") {
    val sheet = stale
      .put(aref("C1"), CellValue.Formula("NOSUCHFN(1)", Some(num(0))))
      .conditionalFormat(range("A1:A3"), CfRule.expression("threshold>5", pink))
    val wb = Workbook(Vector(sheet)).withDefinedName("threshold", "S!$C$1")
    val result = LiveRender.evaluate(sheet, range("A1:A3"), Clock.system, Some(wb))
    assertEquals(result.failures.map(_.ref), Vector(aref("C1")))
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
