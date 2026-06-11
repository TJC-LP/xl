package com.tjclp.xl.ooxml.style

import scala.xml.Elem

import munit.ScalaCheckSuite
import org.scalacheck.Prop.forAll

import com.tjclp.xl.Generators
import com.tjclp.xl.ooxml.XmlSecurity
import com.tjclp.xl.styles.{Dxf, DxfFont}
import com.tjclp.xl.styles.border.{Border, BorderSide, BorderStyle}
import com.tjclp.xl.styles.color.{Color, ThemeSlot}
import com.tjclp.xl.styles.fill.Fill
import com.tjclp.xl.styles.numfmt.NumFmt

/**
 * GH-136: strict-or-None differential-format codec. `parse(toXml(d)) == Some(d)` for every modeled
 * dxf; anything outside the modeled subset parses to None so the enclosing cfRule rides through
 * Preserved ("typed = fully understood").
 */
class DxfCodecSpec extends ScalaCheckSuite:

  private def xml(s: String): Elem =
    XmlSecurity.parseSafe(s, "test").fold(e => fail(s"parse failed: ${e.message}"), identity)

  // ===== round-trip law =====

  property("LAW: parse(toXml(d)) == Some(d)") {
    forAll(Generators.genDxf) { dxf =>
      assertEquals(DxfCodec.parse(DxfCodec.toXml(dxf)), Some(dxf))
      true
    }
  }

  test("round-trip: kitchen-sink dxf (font deltas + fill + border + numFmt)") {
    val dxf = Dxf(
      font = Some(
        DxfFont(
          bold = Some(true),
          italic = Some(false),
          strike = Some(true),
          underline = Some(false),
          color = Some(Color.Rgb(0xff9c0006))
        )
      ),
      fill = Some(Fill.Solid(Color.Rgb(0xffffc7ce))),
      border = Some(
        Border(
          left = BorderSide(BorderStyle.Thin, Some(Color.Rgb(0xff000000))),
          bottom = BorderSide(BorderStyle.MediumDashDot, Some(Color.Theme(ThemeSlot.Accent1, 0.25)))
        )
      ),
      numFmt = Some(NumFmt.Custom("0.000"))
    )
    assertEquals(DxfCodec.parse(DxfCodec.toXml(dxf)), Some(dxf))
  }

  test("toXml emits CT_Dxf child order font -> numFmt -> fill -> border") {
    val dxf = Dxf(
      font = Some(DxfFont(bold = Some(true))),
      fill = Some(Fill.Solid(Color.Rgb(0xff00ff00))),
      border = Some(Border(left = BorderSide(BorderStyle.Thin, None))),
      numFmt = Some(NumFmt.Custom("0.0"))
    )
    val labels = DxfCodec.toXml(dxf).child.collect { case e: Elem => e.label }
    assertEquals(labels, Seq("font", "numFmt", "fill", "border"))
  }

  test("toXml emits the Excel-native dxf fill dialect (bgColor, no patternType)") {
    val emitted = DxfCodec.toXml(Dxf.fill(Color.Rgb(0xff00b050))).toString
    assert(emitted.contains("<patternFill><bgColor rgb=\"FF00B050\"/></patternFill>"), emitted)
  }

  // ===== font deltas =====

  test("<b/>, <b val=\"1\"/>, <b val=\"true\"/> parse Some(true); <b val=\"0\"/> Some(false)") {
    def font(inner: String): Option[Dxf] = DxfCodec.parse(xml(s"<dxf><font>$inner</font></dxf>"))
    assertEquals(font("<b/>"), Some(Dxf.font(DxfFont(bold = Some(true)))))
    assertEquals(font("""<b val="1"/>"""), Some(Dxf.font(DxfFont(bold = Some(true)))))
    assertEquals(font("""<b val="true"/>"""), Some(Dxf.font(DxfFont(bold = Some(true)))))
    assertEquals(font("""<b val="0"/>"""), Some(Dxf.font(DxfFont(bold = Some(false)))))
    assertEquals(font("""<i val="false"/>"""), Some(Dxf.font(DxfFont(italic = Some(false)))))
    assertEquals(font("<strike/>"), Some(Dxf.font(DxfFont(strike = Some(true)))))
  }

  test("underline: <u/>/single -> Some(true), none -> Some(false), double -> degrade to None") {
    def font(inner: String): Option[Dxf] = DxfCodec.parse(xml(s"<dxf><font>$inner</font></dxf>"))
    assertEquals(font("<u/>"), Some(Dxf.font(DxfFont(underline = Some(true)))))
    assertEquals(font("""<u val="single"/>"""), Some(Dxf.font(DxfFont(underline = Some(true)))))
    assertEquals(font("""<u val="none"/>"""), Some(Dxf.font(DxfFont(underline = Some(false)))))
    assertEquals(font("""<u val="double"/>"""), None)
  }

  test("font name/sz children degrade the whole dxf to None") {
    assertEquals(DxfCodec.parse(xml("""<dxf><font><name val="Arial"/></font></dxf>""")), None)
    assertEquals(DxfCodec.parse(xml("""<dxf><font><sz val="14"/></font></dxf>""")), None)
    assertEquals(DxfCodec.parse(xml("""<dxf><font><scheme val="minor"/></font></dxf>""")), None)
  }

  // ===== fills: both wild dialects normalize to Fill.Solid =====

  test("Excel-native dxf fill (bgColor only, no patternType) parses to Fill.Solid") {
    val d = DxfCodec.parse(
      xml("""<dxf><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>""")
    )
    assertEquals(d, Some(Dxf.fill(Color.Rgb(0xffffc7ce))))
  }

  test("openpyxl dxf fill (patternType=solid, fgColor+bgColor) prefers bgColor") {
    val d = DxfCodec.parse(
      xml(
        """<dxf><fill><patternFill patternType="solid"><fgColor rgb="FF111111"/><bgColor rgb="FF222222"/></patternFill></fill></dxf>"""
      )
    )
    assertEquals(d, Some(Dxf.fill(Color.Rgb(0xff222222))))
  }

  test("solid fill with only fgColor falls back to fgColor") {
    val d = DxfCodec.parse(
      xml(
        """<dxf><fill><patternFill patternType="solid"><fgColor rgb="FF333333"/></patternFill></fill></dxf>"""
      )
    )
    assertEquals(d, Some(Dxf.fill(Color.Rgb(0xff333333))))
  }

  test("texture patternType / gradientFill / empty patternFill degrade to None") {
    assertEquals(
      DxfCodec.parse(
        xml("""<dxf><fill><patternFill patternType="darkGray"/></fill></dxf>""")
      ),
      None
    )
    assertEquals(
      DxfCodec.parse(xml("""<dxf><fill><gradientFill degree="90"/></fill></dxf>""")),
      None
    )
    assertEquals(DxfCodec.parse(xml("""<dxf><fill><patternFill/></fill></dxf>""")), None)
  }

  // ===== strictness fences =====

  test("alignment / protection / extLst anywhere in the dxf degrade to None") {
    assertEquals(DxfCodec.parse(xml("""<dxf><alignment horizontal="center"/></dxf>""")), None)
    assertEquals(DxfCodec.parse(xml("""<dxf><protection locked="1"/></dxf>""")), None)
    assertEquals(DxfCodec.parse(xml("""<dxf><extLst/></dxf>""")), None)
  }

  test("indexed color maps through the palette; unmapped index degrades") {
    val mapped = DxfCodec.parse(xml("""<dxf><font><color indexed="2"/></font></dxf>"""))
    assert(mapped.exists(_.font.exists(_.color.isDefined)), s"indexed=2 should map: $mapped")
    assertEquals(DxfCodec.parse(xml("""<dxf><font><color indexed="250"/></font></dxf>""")), None)
    assertEquals(DxfCodec.parse(xml("""<dxf><font><color auto="1"/></font></dxf>""")), None)
  }

  test("empty <dxf/> parses to the empty Dxf (all-inherit)") {
    assertEquals(DxfCodec.parse(xml("<dxf/>")), Some(Dxf()))
  }

  test("border outside the modeled subset (diagonal) degrades to None") {
    assertEquals(
      DxfCodec.parse(
        xml("""<dxf><border diagonalUp="1"><diagonal style="thin"/></border></dxf>""")
      ),
      None
    )
  }

  test("numFmt round-trips by format code") {
    val d = Dxf(numFmt = Some(NumFmt.Custom("#,##0.0")))
    assertEquals(DxfCodec.parse(DxfCodec.toXml(d)), Some(d))
    // built-in code strings parse back to the enum case
    val pct = DxfCodec.parse(
      xml("""<dxf><numFmt numFmtId="9" formatCode="0%"/></dxf>""")
    )
    assertEquals(pct, Some(Dxf(numFmt = Some(NumFmt.Percent))))
  }

/**
 * GH-136: preserved-prefix dxf table merging — append-only invariant, typed-equality index reuse,
 * no-rewrap fixpoint.
 */
class DxfTableSpec extends munit.FunSuite:

  private def xml(s: String): Elem =
    XmlSecurity.parseSafe(s, "test").fold(e => fail(s"parse failed: ${e.message}"), identity)

  private val sourceDxfs = xml(
    """<dxfs count="2"><dxf><font><b/></font></dxf><dxf><fill><patternFill><bgColor rgb="FF00FF00"/></patternFill></fill></dxf></dxfs>"""
  )

  test("existing typed-equal dxf reuses its source index (no growth)") {
    val needed = Vector(Dxf.fill(Color.Rgb(0xff00ff00)))
    val plan = DxfTable.plan(Some(sourceDxfs), needed)
    assertEquals(plan.dxfIds, Map(Dxf.fill(Color.Rgb(0xff00ff00)) -> 1))
    // nothing appended: the ORIGINAL elem rides through untouched (fixpoint)
    assert(plan.merged.exists(_ eq sourceDxfs), "expected the source <dxfs> object unchanged")
  }

  test("new dxf appends after the preserved prefix; source children byte-verbatim") {
    val fresh = Dxf.font(DxfFont(italic = Some(true)))
    val plan = DxfTable.plan(Some(sourceDxfs), Vector(fresh, Dxf.fill(Color.Rgb(0xff00ff00))))
    assertEquals(plan.dxfIds(fresh), 2)
    assertEquals(plan.dxfIds(Dxf.fill(Color.Rgb(0xff00ff00))), 1)
    val merged = plan.merged.getOrElse(fail("expected merged dxfs"))
    assertEquals(merged \@ "count", "3")
    val children = merged.child.collect { case e: Elem => e }
    assertEquals(children.size, 3)
    // preserved prefix is the same Elem objects (verbatim, never re-encoded)
    assert(children(0) eq sourceDxfs.child.collect { case e: Elem => e }.apply(0))
    assert(children(1) eq sourceDxfs.child.collect { case e: Elem => e }.apply(1))
  }

  test("unparseable source entries occupy indices but never match") {
    val withJunk = xml(
      """<dxfs count="2"><dxf><alignment horizontal="center"/></dxf><dxf><font><b/></font></dxf></dxfs>"""
    )
    val bold = Dxf.font(DxfFont(bold = Some(true)))
    val plan = DxfTable.plan(Some(withJunk), Vector(bold))
    assertEquals(plan.dxfIds(bold), 1)
  }

  test("duplicate source entries: first index wins") {
    val dup = xml(
      """<dxfs count="2"><dxf><font><b/></font></dxf><dxf><font><b/></font></dxf></dxfs>"""
    )
    val bold = Dxf.font(DxfFont(bold = Some(true)))
    assertEquals(DxfTable.plan(Some(dup), Vector(bold)).dxfIds(bold), 0)
  }

  test("no source: fresh table from scratch, deduped in encounter order") {
    val a = Dxf.fill(Color.Rgb(0xffff0000))
    val b = Dxf.font(DxfFont(bold = Some(true)))
    val plan = DxfTable.plan(None, Vector(a, b, a))
    assertEquals(plan.dxfIds, Map(a -> 0, b -> 1))
    val merged = plan.merged.getOrElse(fail("expected fresh dxfs"))
    assertEquals(merged \@ "count", "2")
  }

  test("no source and nothing needed: no dxfs element at all") {
    assertEquals(DxfTable.plan(None, Vector.empty), DxfTable.Plan(Map.empty, None))
  }

  test("fixpoint: planning the merged output again appends nothing") {
    val fresh = Dxf.font(DxfFont(italic = Some(true)))
    val first = DxfTable.plan(Some(sourceDxfs), Vector(fresh))
    val merged = first.merged.getOrElse(fail("expected merged"))
    val second = DxfTable.plan(Some(merged), Vector(fresh))
    assert(second.merged.exists(_ eq merged), "second plan must be the fixpoint (object-identical)")
    assertEquals(second.dxfIds(fresh), first.dxfIds(fresh))
  }
