package com.tjclp.xl.render

import com.tjclp.xl.Generators
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.macros.ref
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.sheets.styleSyntax.*
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.alignment.{Align, HAlign}
import com.tjclp.xl.styles.numfmt.NumFmt

import munit.ScalaCheckSuite
import org.scalacheck.Gen
import org.scalacheck.Prop.forAll

import java.time.LocalDateTime

/**
 * The effective rendered content of a cell — its kind and formatted text — is the single input to
 * the three overflow decisions (hash gate, General alignment, colspan measurement), so all three
 * agree with what the renderers draw (GH-500, GH-501, GH-502).
 */
class RenderUtilsSpec extends ScalaCheckSuite:

  import RenderUtils.*

  /** genCellValue has no DateTime or Formula arm; both matter for the resolver. */
  private val genValue: Gen[CellValue] = Gen.frequency(
    5 -> Generators.genCellValue,
    1 -> Generators.genExcelDateTime.map(CellValue.DateTime.apply),
    1 -> Generators.genCellValue.map(v => CellValue.Formula("A1", Some(v))),
    1 -> Gen.const(CellValue.Formula("A1+1", None))
  )

  private val genStyle: Gen[Option[CellStyle]] =
    Gen.option(Generators.genCellStyle)

  private def numFmtOf(style: Option[CellStyle]): NumFmt =
    style.map(_.numFmt).getOrElse(NumFmt.General)

  private def hashable(kind: RenderedKind): Boolean = kind match
    case RenderedKind.Numeric | RenderedKind.Bool | RenderedKind.Error => true
    case RenderedKind.Text | RenderedKind.Empty => false

  // ========== Resolver ==========

  test("renderedContent: kinds follow the value, text follows the format") {
    val dt = LocalDateTime.of(2025, 11, 21, 0, 0, 0)
    assertEquals(
      renderedContent(CellValue.Number(BigDecimal("1234567.9")), NumFmt.General),
      RenderedContent(RenderedKind.Numeric, "1234567.9")
    )
    assertEquals(
      renderedContent(CellValue.Number(BigDecimal("1234.5")), NumFmt.Currency),
      RenderedContent(RenderedKind.Numeric, "$1,234.50")
    )
    assertEquals(
      renderedContent(CellValue.DateTime(dt), NumFmt.General),
      RenderedContent(RenderedKind.Numeric, "45982")
    )
    assertEquals(
      renderedContent(CellValue.Bool(true), NumFmt.General),
      RenderedContent(RenderedKind.Bool, "TRUE")
    )
    assertEquals(
      renderedContent(CellValue.Error(CellError.Div0), NumFmt.Decimal),
      RenderedContent(RenderedKind.Error, "#DIV/0!")
    )
    assertEquals(
      renderedContent(CellValue.Text("abc"), NumFmt.Decimal),
      RenderedContent(RenderedKind.Text, "abc")
    )
    assertEquals(
      renderedContent(CellValue.Empty, NumFmt.General),
      RenderedContent(RenderedKind.Empty, "")
    )
  }

  test("renderedContent: a formula resolves through its cached value, or to its text (GH-500)") {
    assertEquals(
      renderedContent(
        CellValue.Formula("1/0", Some(CellValue.Error(CellError.Div0))),
        NumFmt.General
      ),
      RenderedContent(RenderedKind.Error, "#DIV/0!")
    )
    assertEquals(
      renderedContent(CellValue.Formula("A1>1", Some(CellValue.Bool(false))), NumFmt.General),
      RenderedContent(RenderedKind.Bool, "FALSE")
    )
    assertEquals(
      renderedContent(CellValue.Formula("A1+1", None), NumFmt.General),
      RenderedContent(RenderedKind.Text, "=A1+1")
    )
  }

  test("renderedContent: a number under a text-only format is Text (GH-501)") {
    val n = CellValue.Number(BigDecimal("1234567.9"))
    assertEquals(renderedContent(n, NumFmt.Text), RenderedContent(RenderedKind.Text, "1234567.9"))
    assertEquals(
      renderedContent(n, NumFmt.Custom("@")),
      RenderedContent(RenderedKind.Text, "1234567.9")
    )
    // Dates are numbers: under @ Excel shows the serial as General text
    val dt = LocalDateTime.of(2025, 11, 21, 0, 0, 0)
    assertEquals(
      renderedContent(CellValue.DateTime(dt), NumFmt.Text),
      RenderedContent(RenderedKind.Text, "45982")
    )
    // A 4-section accounting code keeps its @ arm as the text section of a NUMERIC format
    val acct = NumFmt.Custom("_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)")
    assertEquals(renderedContent(n, acct).kind, RenderedKind.Numeric)
    // "General" spelled as a custom code is General, not text
    assertEquals(renderedContent(n, NumFmt.Custom("General")).kind, RenderedKind.Numeric)
  }

  property("renderedContent(v, f).text == cellValueToText(v, f)") {
    forAll(genValue, Generators.genNumFmt) { (v, f) =>
      assertEquals(renderedContent(v, f).text, cellValueToText(v, f))
    }
  }

  // ========== Alignment ==========

  test("alignmentFor: Excel's General alignment by kind — errors centre like logicals (GH-500)") {
    assertEquals(alignmentFor(RenderedKind.Numeric), HAlign.Right)
    assertEquals(alignmentFor(RenderedKind.Bool), HAlign.Center)
    assertEquals(alignmentFor(RenderedKind.Error), HAlign.Center)
    assertEquals(alignmentFor(RenderedKind.Text), HAlign.Left)
    assertEquals(alignmentFor(RenderedKind.Empty), HAlign.Left)
    assertEquals(contentBasedAlignment(CellValue.Error(CellError.NA)), HAlign.Center)
  }

  test("resolveHAlign: an explicit alignment wins over the kind") {
    val right = Some(CellStyle.default.withAlign(Align(horizontal = HAlign.Right)))
    val content = renderedContent(CellValue.Text("abc"), NumFmt.General)
    assertEquals(resolveHAlign(right, content), HAlign.Right)
    assertEquals(resolveHAlign(None, content), HAlign.Left)
    val textFmt = Some(CellStyle.default.withNumFmt(NumFmt.Text))
    val numUnderText = renderedContent(CellValue.Number(BigDecimal(1)), NumFmt.Text)
    assertEquals(resolveHAlign(textFmt, numUnderText), HAlign.Left)
  }

  // ========== Measurement ==========

  test("measureCellValueWidth: measures the formatted text the renderers draw (GH-502)") {
    val err = CellValue.Error(CellError.Div0)
    assertEquals(measureCellValueWidth(err, None), measureTextWidth("#DIV/0!", None))
    val n = CellValue.Number(BigDecimal("1234.5"))
    assertEquals(
      measureCellValueWidth(n, NumFmt.Currency, None),
      measureTextWidth("$1,234.50", None)
    )
    assertEquals(
      measureCellValueWidth(CellValue.Number(BigDecimal("0.123456789")), NumFmt.Decimal, None),
      measureTextWidth("0.12", None)
    )
  }

  // ========== Hash gate ==========

  property("hashOverflowText hashes only Numeric, Bool and Error content") {
    forAll(genValue, genStyle, Gen.choose(1, 200)) { (v, style, width) =>
      val kind = renderedContent(v, numFmtOf(style)).kind
      hashOverflowText(v, style, width) match
        case Some(marker) => hashable(kind) && marker.nonEmpty && marker.forall(_ == '#')
        case None => true
    }
  }

  property("calculateOverflowColspan is 1 whenever the resolved alignment is Right") {
    forAll(genValue, genStyle, Gen.choose(1, 120)) { (v, style, width) =>
      val base = Sheet("Test").put(ref"A1", v)
      val sheet = style.fold(base)(s => base.withCellStyle(ref"A1", s))
      val colWidths = IndexedSeq(width, 72, 72, 72)
      sheet.cells.get(ref"A1").forall { cell =>
        val content = renderedContent(v, numFmtOf(style))
        val colspan = calculateOverflowColspan(cell, ref"A1", width, colWidths, sheet, 0, 3)
        resolveHAlign(style, content) != HAlign.Right || colspan == 1
      }
    }
  }
