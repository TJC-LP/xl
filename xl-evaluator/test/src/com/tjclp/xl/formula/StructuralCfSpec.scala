package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cf.{CfOperator, CfPoint, CfRule, CfTextOp, Cfvo, ConditionalFormat}
import com.tjclp.xl.formula.eval.StructuralEditor
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.Dxf
import com.tjclp.xl.styles.color.Color
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * GH-136: structural edits rewrite TYPED conditional-format formulas (CellIs, Expression,
 * Cfvo.Formula) through the formula engine, mirroring the cell-formula behavior — fully-deleted
 * references degrade the formula text to "#REF!" (rule kept), Preserved payloads stay byte-stable,
 * and Text rules need no rewriting (their formulas are derived at emission).
 */
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial"))
class StructuralCfSpec extends FunSuite:

  private val S = SheetName.unsafe("S")
  private val dxf = Dxf.fill(Color.Rgb(0xffff0000))

  private def sheetNamed(wb: Workbook, n: String): Sheet =
    wb.sheets.find(_.name == SheetName.unsafe(n)).get

  private def rules(s: Sheet): Vector[CfRule] =
    s.typedConditionalFormats.flatMap(_.rules)

  test("insertRows shifts Expression formulas AND the block envelope together") {
    val s = new Sheet(name = S).conditionalFormat(ref"A5:A9", CfRule.expression("$B5>$C5", dxf))
    val r = StructuralEditor.insertRows(Workbook(Vector(s)), S, at = 0, count = 2)
    val s2 = sheetNamed(r, "S")
    assertEquals(s2.typedConditionalFormats.flatMap(_.ranges.map(_.toA1)), Vector("A7:A11"))
    rules(s2) match
      case Vector(CfRule.Expression(f, _, _, _)) => assertEquals(f, "$B7>$C7")
      case other => fail(s"expected Expression, got $other")
  }

  test("deleteColumns rewrites CellIs formula1/formula2") {
    val s = new Sheet(name = S).conditionalFormat(
      ref"D1:D9",
      CfRule.between("$B$1", "$C$1", dxf)
    )
    val r = StructuralEditor.deleteColumns(Workbook(Vector(s)), S, at = 0, count = 1) // drop col A
    rules(sheetNamed(r, "S")) match
      case Vector(CfRule.CellIs(CfOperator.Between, f1, f2, _, _, _)) =>
        assertEquals((f1, f2), ("$A$1", Some("$B$1")))
      case other => fail(s"expected Between CellIs, got $other")
  }

  test("a fully-deleted reference turns the formula text into #REF! and KEEPS the rule") {
    val s = new Sheet(name = S).conditionalFormat(
      ref"D1:D9",
      CfRule.cellIs(CfOperator.GreaterThan, "$B$1", dxf)
    )
    val r = StructuralEditor.deleteColumns(Workbook(Vector(s)), S, at = 1, count = 1) // drop col B
    rules(sheetNamed(r, "S")) match
      case Vector(CfRule.CellIs(_, f1, _, d, p, _)) =>
        assertEquals(f1, "#REF!")
        assertEquals(d, Some(dxf))
        assertEquals(p, 1)
      case other => fail(s"expected CellIs kept with #REF!, got $other")
  }

  test("Cfvo.Formula shifts inside ColorScale points and DataBar bounds") {
    val s = new Sheet(name = S).conditionalFormat(
      ref"A1:A9",
      CfRule.ColorScale(
        CfPoint(Cfvo.Formula("MIN($B$1:$B$9)"), Color.Rgb(0xffff0000)),
        None,
        CfPoint(Cfvo.Max, Color.Rgb(0xff00ff00)),
        1
      ),
      CfRule.DataBar(
        Cfvo.Num(BigDecimal(0)),
        Cfvo.Formula("MAX($B$1:$B$9)"),
        Color.Rgb(0xff638ec6),
        true,
        2
      )
    )
    val r = StructuralEditor.insertRows(Workbook(Vector(s)), S, at = 0, count = 3)
    rules(sheetNamed(r, "S")) match
      case Vector(cs: CfRule.ColorScale, db: CfRule.DataBar) =>
        assertEquals(cs.min.cfvo, Cfvo.Formula("MIN($B$4:$B$12)"))
        assertEquals(db.max, Cfvo.Formula("MAX($B$4:$B$12)"))
        assertEquals(db.min, Cfvo.Num(BigDecimal(0)))
      case other => fail(s"expected ColorScale + DataBar, got $other")
  }

  test("Text rules carry no stored formula: text unchanged, envelope shifts") {
    val s = new Sheet(name = S).conditionalFormat(ref"A5:A9", CfRule.containsText("x", dxf))
    val r = StructuralEditor.insertRows(Workbook(Vector(s)), S, at = 0, count = 1)
    val s2 = sheetNamed(r, "S")
    assertEquals(s2.typedConditionalFormats.flatMap(_.ranges.map(_.toA1)), Vector("A6:A10"))
    rules(s2) match
      case Vector(CfRule.Text(CfTextOp.Contains, text, _, _, _)) => assertEquals(text, "x")
      case other => fail(s"expected Text rule, got $other")
  }

  test("cross-sheet: cf on another sheet tracks the edited sheet's geometry") {
    val data = new Sheet(name = SheetName.unsafe("Data"))
    val view = new Sheet(name = SheetName.unsafe("View"))
      .conditionalFormat(ref"A1:A9", CfRule.expression("Data!$A$5>0", dxf))
    val r = StructuralEditor.insertRows(
      Workbook(Vector(data, view)),
      SheetName.unsafe("Data"),
      at = 0,
      count = 2
    )
    rules(sheetNamed(r, "View")) match
      case Vector(CfRule.Expression(f, _, _, _)) => assertEquals(f, "Data!$A$7>0")
      case other => fail(s"expected Expression, got $other")
    // the View sheet's own envelope is NOT shifted (its geometry did not change)
    assertEquals(
      sheetNamed(r, "View").typedConditionalFormats.flatMap(_.ranges.map(_.toA1)),
      Vector("A1:A9")
    )
  }

  test("Preserved blocks and rules are byte-stable through structural edits") {
    val blockXml =
      """<conditionalFormatting sqref="Z1"><cfRule type="iconSet" priority="9"/></conditionalFormatting>"""
    val ruleXml = """<cfRule type="timePeriod" timePeriod="today" priority="8"/>"""
    val s = new Sheet(name = S)
      .copy(conditionalFormats =
        Vector(
          ConditionalFormat.Preserved(blockXml),
          ConditionalFormat.Rules(
            Vector(ref"A5:A9": CellRange),
            Vector(CfRule.Preserved(ruleXml, Some(8)))
          )
        )
      )
    val r = StructuralEditor.insertRows(Workbook(Vector(s)), S, at = 0, count = 1)
    val s2 = sheetNamed(r, "S")
    assertEquals(s2.conditionalFormats(0), ConditionalFormat.Preserved(blockXml))
    s2.conditionalFormats(1) match
      case ConditionalFormat.Rules(ranges, rs, _) =>
        assertEquals(ranges.map(_.toA1), Vector("A6:A10")) // typed envelope shifts
        assertEquals(rs, Vector(CfRule.Preserved(ruleXml, Some(8)))) // payload untouched
      case other => fail(s"expected Rules, got $other")
  }

  test("unparseable typed formula text is left verbatim (the :93 precedent)") {
    val s = new Sheet(name = S).conditionalFormat(
      ref"A1:A9",
      CfRule.Expression("NOT A FORMULA ((", Some(dxf), 1)
    )
    val r = StructuralEditor.insertRows(Workbook(Vector(s)), S, at = 0, count = 1)
    rules(sheetNamed(r, "S")) match
      case Vector(CfRule.Expression(f, _, _, _)) => assertEquals(f, "NOT A FORMULA ((")
      case other => fail(s"expected Expression, got $other")
  }
