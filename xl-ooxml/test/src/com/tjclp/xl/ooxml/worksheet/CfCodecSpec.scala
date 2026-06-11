package com.tjclp.xl.ooxml.worksheet

import scala.xml.Elem

import munit.ScalaCheckSuite
import org.scalacheck.Prop.forAll

import com.tjclp.xl.Generators
import com.tjclp.xl.api.*
import com.tjclp.xl.cf.{CfOperator, CfPoint, CfRule, CfTextOp, Cfvo, ConditionalFormat}
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.XmlSecurity
import com.tjclp.xl.ooxml.style.DxfCodec
import com.tjclp.xl.styles.{Dxf, DxfFont}
import com.tjclp.xl.styles.color.Color

/**
 * GH-136: total conditional-formatting codec — typed parse of the six rule families with rule- and
 * block-level Preserved fallbacks, canonical emission, and the text-template derive/verify
 * contract. `parseAll(toElems(cfs)) == cfs` for typed content is the per-family round-trip law the
 * generative suite scales up.
 */
class CfCodecSpec extends ScalaCheckSuite:

  private val red = Dxf.fill(Color.Rgb(0xffff0000))
  private val bold = Dxf.font(DxfFont(bold = Some(true)))

  private def xml(s: String): Elem =
    XmlSecurity.parseSafe(s, "test").fold(e => fail(s"parse failed: ${e.message}"), identity)

  /** Round-trip helper: emit with a self-consistent dxf table, parse back. */
  private def roundTrip(cfs: Vector[ConditionalFormat]): Vector[ConditionalFormat] =
    val dxfs = CfCodec.collectDxfs(cfs).distinct
    val ids = dxfs.zipWithIndex.toMap
    val elems = CfCodec.toElems(cfs, ids).map(e => xml(e.toString))
    CfCodec.parseAll(elems, dxfs.map(DxfCodec.toXml))

  private def block(rules: CfRule*): Vector[ConditionalFormat] =
    Vector(ConditionalFormat.Rules(Vector(ref"A1:A5": CellRange), rules.toVector))

  // ===== per-family round-trips =====

  test("cellIs round-trips (incl. Between with two formulas)") {
    val cfs = block(
      CfRule.CellIs(CfOperator.GreaterThan, "100", None, Some(red), 1),
      CfRule.CellIs(CfOperator.Between, "1", Some("9"), Some(bold), 2, stopIfTrue = true),
      CfRule.CellIs(CfOperator.NotEqual, "\"x\"", None, None, 3)
    )
    assertEquals(roundTrip(cfs), cfs)
  }

  test("expression round-trips") {
    val cfs = block(CfRule.Expression("$B1>$C1", Some(red), 1, stopIfTrue = true))
    assertEquals(roundTrip(cfs), cfs)
  }

  test("colorScale (2- and 3-point) round-trips") {
    val cfs = block(
      CfRule.ColorScale(
        CfPoint(Cfvo.Min, Color.Rgb(0xffff0000)),
        Some(CfPoint(Cfvo.Percentile(BigDecimal(50)), Color.Rgb(0xffffff00))),
        CfPoint(Cfvo.Max, Color.Rgb(0xff00ff00)),
        1
      ),
      CfRule.ColorScale(
        CfPoint(Cfvo.Num(BigDecimal("1.5")), Color.Rgb(0xff111111)),
        None,
        CfPoint(
          Cfvo.Percent(BigDecimal(90)),
          Color.Theme(com.tjclp.xl.styles.color.ThemeSlot.Accent1, 0.25)
        ),
        2
      )
    )
    assertEquals(roundTrip(cfs), cfs)
  }

  test("dataBar round-trips (showValue=false emitted, =true omitted)") {
    val cfs = block(
      CfRule.DataBar(Cfvo.Min, Cfvo.Max, Color.Rgb(0xff638ec6), showValue = true, 1),
      CfRule.DataBar(
        Cfvo.Num(BigDecimal(0)),
        Cfvo.Formula("MAX($A$1:$A$5)"),
        Color.Rgb(0xffff0000),
        showValue = false,
        2
      )
    )
    assertEquals(roundTrip(cfs), cfs)
    val emitted = CfCodec.toElems(cfs, Map.empty).map(_.toString).mkString
    assert(emitted.contains("showValue=\"0\""), emitted)
    assert(!emitted.contains("showValue=\"1\""), emitted)
  }

  test("top10 round-trips (percent/bottom flags only when set)") {
    val cfs = block(
      CfRule.Top10(5, percent = false, bottom = false, Some(red), 1),
      CfRule.Top10(10, percent = true, bottom = true, Some(bold), 2, stopIfTrue = true)
    )
    assertEquals(roundTrip(cfs), cfs)
    val emitted = CfCodec.toElems(cfs, Map(red -> 0, bold -> 1)).map(_.toString).mkString
    assert(!emitted.contains("percent=\"0\""), emitted)
    assert(!emitted.contains("bottom=\"0\""), emitted)
  }

  test("text families round-trip via the derived-template contract") {
    val cfs = block(
      CfRule.Text(CfTextOp.Contains, "todo", Some(red), 1),
      CfRule.Text(CfTextOp.NotContains, "ok", Some(red), 2),
      CfRule.Text(CfTextOp.BeginsWith, "Q1", Some(bold), 3),
      CfRule.Text(CfTextOp.EndsWith, "z", None, 4, stopIfTrue = true)
    )
    assertEquals(roundTrip(cfs), cfs)
  }

  test("text with quotes: doubled in the derived formula, round-trips") {
    val cfs = block(CfRule.Text(CfTextOp.Contains, "say \"hi\"", Some(red), 1))
    val emitted = CfCodec.toElems(cfs, Map(red -> 0)).map(_.toString).mkString
    assert(emitted.contains("SEARCH(&quot;say &quot;&quot;hi&quot;&quot;&quot;,A1)"), emitted)
    assertEquals(roundTrip(cfs), cfs)
  }

  test("multi-range block: space-joined sqref, derived text formula uses FIRST range's top-left") {
    val cfs = Vector(
      ConditionalFormat.Rules(
        Vector(ref"B2:B9": CellRange, ref"D2:D9": CellRange),
        Vector(CfRule.Text(CfTextOp.Contains, "x", Some(red), 1))
      )
    )
    val emitted = CfCodec.toElems(cfs, Map(red -> 0)).map(_.toString).mkString
    assert(emitted.contains("sqref=\"B2:B9 D2:D9\""), emitted)
    assert(emitted.contains("SEARCH(&quot;x&quot;,B2)"), emitted)
    assertEquals(roundTrip(cfs), cfs)
  }

  test("pivot flag round-trips and is emitted only when set") {
    val cfs = Vector(
      ConditionalFormat.Rules(
        Vector(ref"A1:A5": CellRange),
        Vector(CfRule.Expression("A1>0", None, 1)),
        pivot = true
      )
    )
    assertEquals(roundTrip(cfs), cfs)
    val plain = CfCodec.toElems(block(CfRule.Expression("A1>0", None, 1)), Map.empty)
    assert(!plain.map(_.toString).mkString.contains("pivot"), plain.toString)
  }

  property("LAW: parseAll(toElems(cfs)) == cfs for generated typed blocks") {
    forAll(Generators.genConditionalFormats) { cfs =>
      assertEquals(roundTrip(cfs), cfs)
      true
    }
  }

  // ===== emission spec details =====

  test("cfRule attrs in Excel order: type, dxfId, priority, stopIfTrue, operator") {
    val cfs =
      block(CfRule.CellIs(CfOperator.GreaterThan, "1", None, Some(red), 3, stopIfTrue = true))
    val emitted = CfCodec.toElems(cfs, Map(red -> 7)).map(_.toString).mkString
    assert(
      emitted.contains(
        "<cfRule type=\"cellIs\" dxfId=\"7\" priority=\"3\" stopIfTrue=\"1\" operator=\"greaterThan\">"
      ),
      emitted
    )
  }

  test("colorScale emits all cfvos then all colors; cfvo val absent for min/max") {
    val cfs = block(
      CfRule.ColorScale(
        CfPoint(Cfvo.Min, Color.Rgb(0xffff0000)),
        Some(CfPoint(Cfvo.Percentile(BigDecimal(50)), Color.Rgb(0xffffff00))),
        CfPoint(Cfvo.Max, Color.Rgb(0xff00ff00)),
        1
      )
    )
    val emitted = CfCodec.toElems(cfs, Map.empty).map(_.toString).mkString
    assert(
      emitted.contains(
        "<colorScale><cfvo type=\"min\"/><cfvo type=\"percentile\" val=\"50\"/><cfvo type=\"max\"/>" +
          "<color rgb=\"FFFF0000\"/><color rgb=\"FFFFFF00\"/><color rgb=\"FF00FF00\"/></colorScale>"
      ),
      emitted
    )
  }

  test("a Rules block with empty ranges or rules is dropped at emission") {
    val empty = Vector(
      ConditionalFormat.Rules(Vector.empty, Vector(CfRule.Expression("A1>0", None, 1))),
      ConditionalFormat.Rules(Vector(ref"A1:A2": CellRange), Vector.empty)
    )
    assertEquals(CfCodec.toElems(empty, Map.empty), Seq.empty)
  }

  test(
    "emitter safety net: residual priority <= 0 assigned max+1.. in document order; explicit never renumbered"
  ) {
    val cfs = Vector(
      ConditionalFormat.Rules(
        Vector(ref"A1:A5": CellRange),
        Vector(
          CfRule.Expression("A1>0", None, priority = 0), // sentinel leaked via direct construction
          CfRule.Expression("A1>1", None, priority = 9),
          CfRule.Expression("A1>2", None, priority = 0)
        )
      )
    )
    val back = roundTrip(cfs)
    val priorities = back
      .collect { case ConditionalFormat.Rules(_, rules, _) => rules }
      .flatten
      .flatMap(CfRule.priorityOf)
    assertEquals(priorities, Vector(10, 9, 11))
  }

  // ===== Preserved fallbacks: rule level =====

  private def parseOne(blockXml: String, dxfs: Vector[Elem] = Vector.empty): ConditionalFormat =
    CfCodec.parseAll(Seq(xml(blockXml)), dxfs) match
      case Vector(one) => one
      case other => fail(s"expected one block, got $other")

  private def rulesOf(cf: ConditionalFormat): Vector[CfRule] = cf match
    case ConditionalFormat.Rules(_, rules, _) => rules
    case other => fail(s"expected typed envelope, got $other")

  test("iconSet rides Preserved with its parsed priority") {
    val cf = parseOne(
      """<conditionalFormatting sqref="A1:A5"><cfRule type="iconSet" priority="4"><iconSet iconSet="3Arrows"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule></conditionalFormatting>"""
    )
    rulesOf(cf) match
      case Vector(CfRule.Preserved(xmlStr, priority)) =>
        assertEquals(priority, Some(4))
        assert(xmlStr.contains("3Arrows"), xmlStr)
      case other => fail(s"expected Preserved iconSet, got $other")
  }

  test("unknown cfRule attr -> Preserved") {
    val cf = parseOne(
      """<conditionalFormatting sqref="A1"><cfRule type="expression" priority="1" mystery="x"><formula>A1&gt;0</formula></cfRule></conditionalFormatting>"""
    )
    assert(rulesOf(cf).forall(_.isInstanceOf[CfRule.Preserved]), cf.toString)
  }

  test("cfRule child extLst -> Preserved (protects x14 GUID pairings)") {
    val cf = parseOne(
      """<conditionalFormatting sqref="A1"><cfRule type="expression" priority="1"><formula>A1&gt;0</formula><extLst><ext uri="{x}"/></extLst></cfRule></conditionalFormatting>"""
    )
    assert(rulesOf(cf).forall(_.isInstanceOf[CfRule.Preserved]), cf.toString)
  }

  test("out-of-range dxfId and untypeable dxf -> Preserved") {
    val outOfRange = parseOne(
      """<conditionalFormatting sqref="A1"><cfRule type="cellIs" dxfId="5" priority="1" operator="equal"><formula>1</formula></cfRule></conditionalFormatting>""",
      Vector.empty
    )
    assert(rulesOf(outOfRange).forall(_.isInstanceOf[CfRule.Preserved]), outOfRange.toString)
    val untypeable = parseOne(
      """<conditionalFormatting sqref="A1"><cfRule type="cellIs" dxfId="0" priority="1" operator="equal"><formula>1</formula></cfRule></conditionalFormatting>""",
      Vector(xml("""<dxf><alignment horizontal="center"/></dxf>"""))
    )
    assert(rulesOf(untypeable).forall(_.isInstanceOf[CfRule.Preserved]), untypeable.toString)
  }

  test("text-template mismatch -> Preserved (typed = fully understood)") {
    // hand-edited formula no longer matches the canonical derivation for text="x" at A1
    val cf = parseOne(
      """<conditionalFormatting sqref="A1:A5"><cfRule type="containsText" dxfId="0" priority="1" operator="containsText" text="x"><formula>NOT(ISERROR(SEARCH("x",B7)))</formula></cfRule></conditionalFormatting>""",
      Vector(xml("""<dxf><font><b/></font></dxf>"""))
    )
    assert(rulesOf(cf).forall(_.isInstanceOf[CfRule.Preserved]), cf.toString)
  }

  test("text rule with the canonical formula parses typed") {
    val cf = parseOne(
      """<conditionalFormatting sqref="A1:A5"><cfRule type="containsText" dxfId="0" priority="1" operator="containsText" text="x"><formula>NOT(ISERROR(SEARCH("x",A1)))</formula></cfRule></conditionalFormatting>""",
      Vector(xml("""<dxf><font><b/></font></dxf>"""))
    )
    rulesOf(cf) match
      case Vector(CfRule.Text(CfTextOp.Contains, "x", Some(d), 1, false)) =>
        assertEquals(d, Dxf.font(DxfFont(bold = Some(true))))
      case other => fail(s"expected typed Text rule, got $other")
  }

  test("dataBar with minLength/maxLength (unmodeled attrs) -> Preserved") {
    val cf = parseOne(
      """<conditionalFormatting sqref="A1"><cfRule type="dataBar" priority="1"><dataBar minLength="10" maxLength="90"><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar></cfRule></conditionalFormatting>"""
    )
    assert(rulesOf(cf).forall(_.isInstanceOf[CfRule.Preserved]), cf.toString)
  }

  // ===== Preserved fallbacks: block level =====

  test("unknown envelope attr -> whole-block Preserved") {
    val src =
      """<conditionalFormatting sqref="A1" weird="1"><cfRule type="expression" priority="1"><formula>A1&gt;0</formula></cfRule></conditionalFormatting>"""
    parseOne(src) match
      case ConditionalFormat.Preserved(xmlStr) => assert(xmlStr.contains("weird"), xmlStr)
      case other => fail(s"expected block Preserved, got $other")
  }

  test("corrupt sqref -> whole-block Preserved") {
    val src =
      """<conditionalFormatting sqref="NOT A REF"><cfRule type="expression" priority="1"><formula>1</formula></cfRule></conditionalFormatting>"""
    assert(parseOne(src).isInstanceOf[ConditionalFormat.Preserved], parseOne(src).toString)
  }

  test("single-cell sqref token parses as a 1x1 range") {
    val cf = parseOne(
      """<conditionalFormatting sqref="B2"><cfRule type="expression" priority="1"><formula>B2&gt;0</formula></cfRule></conditionalFormatting>"""
    )
    cf match
      case ConditionalFormat.Rules(ranges, _, _) =>
        assertEquals(ranges, Vector(CellRange(ref"B2", ref"B2")))
      case other => fail(s"expected typed envelope, got $other")
  }

  test("Preserved (both levels) re-emit verbatim and survive a second parse byte-stably") {
    val ruleLevel =
      """<conditionalFormatting sqref="A1:A5"><cfRule type="iconSet" priority="2"><iconSet iconSet="5Rating"><cfvo type="percent" val="0"/><cfvo type="percent" val="20"/><cfvo type="percent" val="40"/><cfvo type="percent" val="60"/><cfvo type="percent" val="80"/></iconSet></cfRule></conditionalFormatting>"""
    val blockLevel =
      """<conditionalFormatting sqref="C1" unknown="1"><cfRule type="timePeriod" timePeriod="today" priority="3"/></conditionalFormatting>"""
    val first = CfCodec.parseAll(Seq(xml(ruleLevel), xml(blockLevel)), Vector.empty)
    val emitted = CfCodec.toElems(first, Map.empty).map(e => xml(e.toString))
    val second = CfCodec.parseAll(emitted, Vector.empty)
    assertEquals(second, first)
  }

  test("emitter safety net saturates at Int.MaxValue (never emits a negative priority)") {
    val cfs = block(
      CfRule.CellIs(CfOperator.GreaterThan, "1", None, None, priority = Int.MaxValue),
      CfRule.Expression("A1>0", None, CfRule.AutoPriority) // residual sentinel
    )
    val priorities = CfCodec
      .toElems(cfs, Map.empty)
      .flatMap(e => (e \ "cfRule").flatMap(_.attribute("priority")).map(_.text))
    assertEquals(priorities, Seq(Int.MaxValue.toString, Int.MaxValue.toString))
  }

  // ===== namespace-self-contained Preserved capture =====
  // WorksheetReader rebinds used prefixes onto the BLOCK root and severs every descendant scope
  // (cleanNamespaces), so a rule captured in isolation must re-resolve x14/xm against the block's
  // scope — or the stored payload is unbound XML that parsePreserved silently drops at emission.

  /** The dataBar extLst x14 pairing exactly as Excel/LibreOffice write it (inline binding). */
  private val x14ExtLstRule =
    """<cfRule type="dataBar" priority="5"><dataBar minLength="10" maxLength="90"><cfvo type="min" val="0"/><cfvo type="max" val="0"/><color rgb="FF638EC6"/></dataBar><extLst><ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"><x14:id>{26C9BFD7-9A20-468E-9236-D82474454A88}</x14:id></ext></extLst></cfRule>"""

  test("rule-level Preserved capture rebinds prefixes severed onto the block root (x14 extLst)") {
    val raw = s"""<conditionalFormatting sqref="D2:D9">$x14ExtLstRule</conditionalFormatting>"""
    // what WorksheetReader does before CfCodec ever sees the block
    val block = rebindUsedNamespaces(xml(raw))
    CfCodec.parseAll(Seq(block), Vector.empty) match
      case Vector(ConditionalFormat.Rules(_, Vector(CfRule.Preserved(payload, Some(5))), _)) =>
        assert(
          payload.contains("xmlns:x14="),
          s"x14 binding must travel with the fragment:\n$payload"
        )
        assert(
          XmlSecurity.parseSafe(payload, "rule payload").isRight,
          s"captured payload must be self-contained XML:\n$payload"
        )
      case other => fail(s"expected one typed envelope with one Preserved rule, got $other")
  }

  test("block-level Preserved capture keeps prefixed descendants bound (pin)") {
    val raw =
      s"""<conditionalFormatting sqref="D2:D9" weird="1">$x14ExtLstRule</conditionalFormatting>"""
    val block = rebindUsedNamespaces(xml(raw))
    CfCodec.parseAll(Seq(block), Vector.empty) match
      case Vector(ConditionalFormat.Preserved(payload)) =>
        assert(payload.contains("xmlns:x14="), payload)
        assert(XmlSecurity.parseSafe(payload, "block payload").isRight, payload)
      case other => fail(s"expected block-level Preserved, got $other")
  }
