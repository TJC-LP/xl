package com.tjclp.xl.styles

import com.tjclp.xl.Generators.{genCellStyle, genDxf}
import com.tjclp.xl.styles.border.{Border, BorderSide, BorderStyle}
import com.tjclp.xl.styles.color.Color
import com.tjclp.xl.styles.fill.Fill
import com.tjclp.xl.styles.font.{Font, Underline}
import com.tjclp.xl.styles.numfmt.NumFmt
import munit.ScalaCheckSuite
import org.scalacheck.Prop.forAll

/**
 * GH-497: the differential format as conditional formatting composes it — `orElse` is the monoid
 * the rule fold accumulates (the higher-precedence dxf wins each property it sets), and `applyTo`
 * lays the composed delta over a cell's base style, a homomorphism from that monoid to style
 * endomorphisms.
 */
class DxfSpec extends ScalaCheckSuite:

  private val red = Color.Rgb(0xffff0000)
  private val pink = Color.Rgb(0xffffc7ce)
  private val thin = BorderSide(BorderStyle.Thin, Some(red))

  property("orElse is associative") {
    forAll(genDxf, genDxf, genDxf) { (a, b, c) =>
      assertEquals(a.orElse(b).orElse(c), a.orElse(b.orElse(c)))
    }
  }

  property("Dxf() is the left and right identity of orElse") {
    forAll(genDxf) { d =>
      assertEquals(Dxf().orElse(d), d)
      assertEquals(d.orElse(Dxf()), d)
    }
  }

  property("applyTo is a homomorphism: (hi orElse lo).applyTo(s) == hi.applyTo(lo.applyTo(s))") {
    forAll(genDxf, genDxf, genCellStyle) { (hi, lo, s) =>
      assertEquals(hi.orElse(lo).applyTo(s), hi.applyTo(lo.applyTo(s)))
    }
  }

  property("Dxf().applyTo is the identity") {
    forAll(genCellStyle)(s => assertEquals(Dxf().applyTo(s), s))
  }

  test("orElse: the higher dxf wins a conflict, non-conflicting properties combine") {
    val hi = Dxf(font = Some(DxfFont(bold = Some(true))), fill = Some(Fill.Solid(red)))
    val lo = Dxf(
      font = Some(DxfFont(bold = Some(false), italic = Some(true))),
      fill = Some(Fill.Solid(pink)),
      numFmt = Some(NumFmt.Percent)
    )
    assertEquals(
      hi.orElse(lo),
      Dxf(
        font = Some(DxfFont(bold = Some(true), italic = Some(true))),
        fill = Some(Fill.Solid(red)),
        numFmt = Some(NumFmt.Percent)
      )
    )
  }

  test("orElse: a border side whose style is None is absent, so the lower side shows through") {
    val hi = Dxf(border = Some(Border(bottom = thin)))
    val lo = Dxf(border = Some(Border(top = thin, bottom = BorderSide(BorderStyle.Double))))
    assertEquals(hi.orElse(lo), Dxf(border = Some(Border(top = thin, bottom = thin))))
  }

  test("applyTo: a solid fill replaces the base fill") {
    val base = CellStyle.default.withFill(Fill.Solid(pink))
    assertEquals(Dxf.fill(red).applyTo(base).fill, Fill.Solid(red))
  }

  test("applyTo: Some(false) turns a bold base off, None inherits it") {
    val bold = CellStyle.default.withFont(Font.default.withBold(true))
    assertEquals(Dxf.font(DxfFont(bold = Some(false))).applyTo(bold).font.bold, false)
    assertEquals(Dxf.font(DxfFont(italic = Some(true))).applyTo(bold).font.bold, true)
  }

  test("applyTo: Some(Underline.None) clears the underline; a colour overrides the base's") {
    val base = CellStyle.default.withFont(Font.default.withUnderline(Underline.Single))
    val out = Dxf.font(DxfFont(underline = Some(Underline.None), color = Some(red))).applyTo(base)
    assertEquals(out.font.underline, Underline.None)
    assertEquals(out.font.color, Some(red))
    // name and size always come from the base: a dxf cannot express them
    assertEquals((out.font.name, out.font.sizePt), (base.font.name, base.font.sizePt))
  }

  test("applyTo: a border with only its bottom side set keeps the base's other sides") {
    val base = CellStyle.default.withBorder(Border.all(BorderStyle.Medium))
    val out = Dxf(border = Some(Border(bottom = thin))).applyTo(base)
    assertEquals(out.border, Border.all(BorderStyle.Medium).withBottom(thin))
  }

  test("applyTo: a numFmt overrides the base's and clears its numFmtId") {
    val base = CellStyle.default.withNumFmtId(3, NumFmt.Integer)
    val out = Dxf(numFmt = Some(NumFmt.Percent)).applyTo(base)
    assertEquals((out.numFmt, out.numFmtId), (NumFmt.Percent, None))
  }
