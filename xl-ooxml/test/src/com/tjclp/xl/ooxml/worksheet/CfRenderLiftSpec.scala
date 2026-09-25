package com.tjclp.xl.ooxml.worksheet

import scala.xml.Elem

import munit.FunSuite

import com.tjclp.xl.api.*
import com.tjclp.xl.cf.{CfOperator, CfRule, ConditionalFormat}
import com.tjclp.xl.ooxml.XmlSecurity
import com.tjclp.xl.styles.{Dxf, DxfFont}
import com.tjclp.xl.styles.color.Color
import com.tjclp.xl.styles.fill.Fill

/**
 * GH-497: the render-only lift of Excel's formula-backed rules that xl does not type — blanks,
 * errors and time periods — into expressions over the formula Excel stores for them, so the
 * renderers can paint them (and honour the "Blanks, Stop If True" guard).
 */
class CfRenderLiftSpec extends FunSuite:

  private def xml(s: String): Elem =
    XmlSecurity.parseSafe(s, "test").fold(e => fail(s"parse failed: ${e.message}"), identity)

  private val pinkDxf = xml(
    """<dxf><font><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>"""
  )
  private val pink = Dxf(
    font = Some(DxfFont(color = Some(Color.Rgb(0xff9c0006)))),
    fill = Some(Fill.Solid(Color.Rgb(0xffffc7ce)))
  )

  private def parse(
    rules: String,
    dxfs: Vector[Elem] = Vector(pinkDxf)
  ): Vector[ConditionalFormat] =
    CfCodec.parseAll(
      Seq(xml(s"""<conditionalFormatting sqref="A1:A9">$rules</conditionalFormatting>""")),
      dxfs
    )

  private def rulesOf(cfs: Vector[ConditionalFormat]): Vector[CfRule] =
    cfs.flatMap {
      case ConditionalFormat.Rules(_, rules, _) => rules
      case other => fail(s"expected a typed envelope, got $other")
    }

  test("containsBlanks lifts to an expression over Excel's stored formula, with its dxf") {
    val cfs = parse(
      """<cfRule type="containsBlanks" dxfId="0" priority="2" stopIfTrue="1"><formula>LEN(TRIM(A1))=0</formula></cfRule>"""
    )
    assert(
      rulesOf(cfs).forall {
        case _: CfRule.Preserved => true
        case _ => false
      },
      "precondition: the codec keeps it Preserved"
    )
    assertEquals(
      rulesOf(CfRenderLift.lift(cfs)),
      Vector(CfRule.Expression("LEN(TRIM(A1))=0", Some(pink), 2, stopIfTrue = true))
    )
  }

  test("the format-less guard lifts with no dxf, keeping stopIfTrue") {
    val cfs = parse(
      """<cfRule type="containsBlanks" priority="1" stopIfTrue="1"><formula>LEN(TRIM(A1))=0</formula></cfRule>"""
    )
    assertEquals(
      rulesOf(CfRenderLift.lift(cfs)),
      Vector(CfRule.Expression("LEN(TRIM(A1))=0", None, 1, stopIfTrue = true))
    )
  }

  test("notContainsBlanks, containsErrors, notContainsErrors and timePeriod lift too") {
    val cfs = parse(
      """<cfRule type="notContainsBlanks" dxfId="0" priority="1"><formula>LEN(TRIM(A1))&gt;0</formula></cfRule>""" +
        """<cfRule type="containsErrors" dxfId="0" priority="2"><formula>ISERROR(A1)</formula></cfRule>""" +
        """<cfRule type="notContainsErrors" dxfId="0" priority="3"><formula>NOT(ISERROR(A1))</formula></cfRule>""" +
        """<cfRule type="timePeriod" dxfId="0" priority="4" timePeriod="yesterday"><formula>FLOOR(A1,1)=TODAY()-1</formula></cfRule>"""
    )
    assertEquals(
      rulesOf(CfRenderLift.lift(cfs)),
      Vector(
        CfRule.Expression("LEN(TRIM(A1))>0", Some(pink), 1),
        CfRule.Expression("ISERROR(A1)", Some(pink), 2),
        CfRule.Expression("NOT(ISERROR(A1))", Some(pink), 3),
        CfRule.Expression("FLOOR(A1,1)=TODAY()-1", Some(pink), 4)
      )
    )
  }

  test("the stored formula crosses the storage boundary: _xlfn. prefixes are stripped") {
    val cfs = parse(
      """<cfRule type="containsErrors" dxfId="0" priority="1"><formula>_xlfn.ISFORMULA(A1)</formula></cfRule>"""
    )
    assertEquals(
      rulesOf(CfRenderLift.lift(cfs)),
      Vector(CfRule.Expression("ISFORMULA(A1)", Some(pink), 1))
    )
  }

  test("an untypeable dxf, an unknown attribute or an extra child keeps the rule Preserved") {
    val untypeable = parse(
      """<cfRule type="containsBlanks" dxfId="0" priority="1"><formula>LEN(TRIM(A1))=0</formula></cfRule>""",
      Vector(xml("""<dxf><alignment horizontal="center"/></dxf>"""))
    )
    val attr = parse(
      """<cfRule type="containsBlanks" priority="1" mystery="1"><formula>LEN(TRIM(A1))=0</formula></cfRule>"""
    )
    val child = parse(
      """<cfRule type="containsErrors" priority="1"><formula>ISERROR(A1)</formula><extLst/></cfRule>"""
    )
    List(untypeable, attr, child).foreach { cfs =>
      assertEquals(CfRenderLift.lift(cfs), cfs)
    }
  }

  test("kinds with no stored formula, and x14 data bars, stay Preserved") {
    val cfs = parse(
      """<cfRule type="iconSet" priority="1"><iconSet><cfvo type="percent" val="0"/><cfvo type="percent" val="50"/></iconSet></cfRule>""" +
        """<cfRule type="aboveAverage" dxfId="0" priority="2"/>""" +
        """<cfRule type="duplicateValues" dxfId="0" priority="3"/>""" +
        """<cfRule type="dataBar" priority="4"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar><extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"/></extLst></cfRule>"""
    )
    assertEquals(CfRenderLift.lift(cfs), cfs)
  }

  test("typed rules and unreadable blocks pass through; the lift is idempotent") {
    val typed = ConditionalFormat.Rules(
      Vector(CellRange.parse("B1:B3").fold(fail(_), identity)),
      Vector(CfRule.CellIs(CfOperator.LessThan, "0", None, Some(pink), 1))
    )
    val unreadable = ConditionalFormat.Preserved("""<conditionalFormatting sqref="??"/>""")
    val blanks = parse(
      """<cfRule type="containsBlanks" priority="3" stopIfTrue="1"><formula>LEN(TRIM(A1))=0</formula></cfRule>"""
    )
    val all = Vector(typed, unreadable) ++ blanks
    val once = CfRenderLift.lift(all)
    assertEquals(once.take(2), Vector(typed, unreadable))
    assertEquals(CfRenderLift.lift(once), once)
  }
