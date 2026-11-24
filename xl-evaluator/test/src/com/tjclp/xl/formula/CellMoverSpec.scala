package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cf.CfRule
import com.tjclp.xl.styles.Dxf
import com.tjclp.xl.formula.eval.CellMover
import com.tjclp.xl.formula.printer.FormulaOps
import munit.FunSuite

/**
 * GH-935: a merge over an empty upper-left cell moves the first non-empty cell into it as Excel's
 * cut and paste does, and every formula that read the moved cell is pointed at the upper-left cell
 * (checked against Excel for Mac: `E1 = B1+1`, merge `A1:B1` over an empty A1, gives `E1 = A1+1`,
 * value 8).
 */
class CellMoverSpec extends FunSuite:

  private val S1 = SheetName.unsafe("Sheet1")
  private val Other = SheetName.unsafe("Other Sheet")
  private def num(n: Int): CellValue = CellValue.Number(BigDecimal(n))
  private def f(text: String): CellValue = CellValue.Formula(text, None)
  private def text(wb: Workbook, sheet: SheetName, at: ARef): String =
    wb(sheet).toOption.flatMap(_.cells.get(at)).map(_.value) match
      case Some(CellValue.Formula(t, _, _)) => t
      case other => fail(s"${at.toA1} is not a formula: $other")
  private def cached(wb: Workbook, sheet: SheetName, at: ARef): Option[CellValue] =
    wb(sheet).toOption.flatMap(_.cells.get(at)).map(_.value) match
      case Some(CellValue.Formula(_, c, _)) => c
      case _ => None

  // ===== the text rewrite =====

  private def moved(t: String, host: Option[SheetName] = Some(S1)): String =
    FormulaOps.relocateCell(t, host, S1, ref"B1", ref"A1")

  test("relocateCell: single-cell references to the moved cell follow it, anchors kept") {
    assertEquals(moved("=B1+1"), "=A1+1")
    assertEquals(moved("=$B$1*2"), "=$A$1*2")
    assertEquals(moved("=B$1+$B1"), "=A$1+$A1")
    assertEquals(moved("=Sheet1!B1"), "=Sheet1!A1")
    assertEquals(moved("=sheet1!B1", Some(Other)), "=sheet1!A1")
  }

  test("relocateCell: a local reference on another sheet and wider ranges stay") {
    assertEquals(moved("=B1+1", Some(Other)), "=B1+1")
    assertEquals(moved("=SUM(B1:B3)"), "=SUM(B1:B3)")
    assertEquals(moved("=SUM(B:B)"), "=SUM(B:B)")
    assertEquals(moved("=B10+B2"), "=B10+B2")
    assertEquals(moved("=[1]Sheet1!B1"), "=[1]Sheet1!B1")
    assertEquals(moved("=INDIRECT(\"B1\")"), "=INDIRECT(\"B1\")")
    // a defined name (no host) reaches only qualified references
    assertEquals(moved("Sheet1!$B$1", None), "Sheet1!$A$1")
    assertEquals(moved("B1", None), "B1")
    // unparseable text is left as it is
    assertEquals(moved("=SUM(B1"), "=SUM(B1")
  }

  // ===== the workbook-level merge =====

  private val book: Workbook = Workbook(
    Vector(
      Sheet(S1)
        .put(ref"B1", num(7))
        .put(ref"E1", f("B1+1"))
        .conditionalFormat(
          CellRange(ref"E1", ref"E1"),
          CfRule.expression("$B$1>1", Dxf.fill(Color.Rgb(0xffffcc00)))
        ),
      Sheet(Other).put(ref"A1", f("Sheet1!B1*2")).put(ref"A2", f("'Sheet1'!B1:B1"))
    )
  ).withDefinedName("seed", "Sheet1!$B$1")

  test("GH-935 (Excel): merge points every reader of the moved cell at the upper-left cell") {
    val wb =
      CellMover.merge(book, S1, CellRange(ref"A1", ref"B1")).fold(e => fail(e.message), identity)
    val s = wb(S1).fold(e => fail(e.message), identity)
    assertEquals(s(ref"A1").value, num(7))
    assertEquals(s.cells.get(ref"B1"), None)
    assertEquals(text(wb, S1, ref"E1"), "A1+1")
    assertEquals(text(wb, Other, ref"A1"), "Sheet1!A1*2")
    assertEquals(text(wb, Other, ref"A2"), "Sheet1!A1:A1")
    assertEquals(
      wb.metadata.definedNames.find(_.name == "seed").map(_.formula),
      Some("Sheet1!$A$1")
    )
    s.conditionalFormats.headOption match
      case Some(com.tjclp.xl.cf.ConditionalFormat.Rules(_, Vector(r: CfRule.Expression), _)) =>
        assertEquals(r.formula, "$A$1>1")
      case other => fail(s"unexpected CF $other")
    val recalculated = wb.recalculate().workbook
    assertEquals(cached(recalculated, S1, ref"E1"), Some(num(8)))
    assertEquals(cached(recalculated, Other, ref"A1"), Some(num(14)))
  }

  test("GH-935: a merge whose upper-left cell has a value rewrites nothing") {
    val wb0 = book.update(S1, _.put(ref"A1", num(1))).fold(e => fail(e.message), identity)
    val wb =
      CellMover.merge(wb0, S1, CellRange(ref"A1", ref"B1")).fold(e => fail(e.message), identity)
    assertEquals(text(wb, S1, ref"E1"), "B1+1")
    assertEquals(text(wb, Other, ref"A1"), "Sheet1!B1*2")
  }

  test("GH-935: wb.edit(Edit.Merge) under EvalFormulaSupport does the same") {
    val wb = book
      .editIn(S1)(Edit.Merge(Area(None, CellRange(ref"A1", ref"B1"))))
      .fold(e => fail(e.message), identity)
    assertEquals(text(wb, S1, ref"E1"), "A1+1")
    assertEquals(text(wb, Other, ref"A1"), "Sheet1!A1*2")
    assertEquals(wb(S1).toOption.map(_(ref"A1").value), Some(num(7)))
  }

  test("GH-935: the sheet-level merge rewrites only this sheet's readers") {
    val s = book(S1).fold(e => fail(e.message), identity)
    val merged = CellMover.mergeSheet(s, CellRange(ref"A1", ref"B1"))
    assertEquals(merged(ref"A1").value, num(7))
    assertEquals(merged(ref"E1").value, f("A1+1"))
  }
