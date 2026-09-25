package com.tjclp.xl.render

import com.tjclp.xl.Generators
import com.tjclp.xl.addressing.Column
import com.tjclp.xl.cells.{Cell, CellError, CellValue}
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.display.NumFmtFormatter
import com.tjclp.xl.macros.ref
import com.tjclp.xl.render.syntax.*
import com.tjclp.xl.sheets.{ColumnProperties, RowProperties, Sheet}
import com.tjclp.xl.sheets.styleSyntax.*
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.alignment.{Align, HAlign}
import com.tjclp.xl.styles.color.ThemePalette
import com.tjclp.xl.styles.numfmt.NumFmt
import com.tjclp.xl.unsafe.*

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

  property("renderedContent(v, f).text is what the display formatter shows for the value") {
    // The oracle is the display layer, not the resolver: a formula shows its cached value (or
    // its source when uncached), everything else its NumFmtFormatter text under the format.
    forAll(genValue, Generators.genNumFmt) { (v, f) =>
      val shown = v match
        case CellValue.Formula(_, Some(cached), _) => NumFmtFormatter.formatValue(cached, f)
        case CellValue.Formula(expr, None, _) => s"=$expr"
        case other => NumFmtFormatter.formatValue(other, f)
      assertEquals(renderedContent(v, f).text, shown)
      assertEquals(cellValueToText(v, f), shown)
    }
  }

  property("renderedContent: a cached formula resolves exactly as its cached value does") {
    forAll(genValue, Generators.genNumFmt) { (v, f) =>
      assertEquals(renderedContent(CellValue.Formula("A1", Some(v)), f), renderedContent(v, f))
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

  property("hashOverflowText decides from the measured formatted text and the box alone") {
    // Two-sided, without restating the box arithmetic: text that fits the SMALLEST box (left,
    // padded, indented) is never hashed; text wider than the whole cell always is, for a
    // hashable kind that does not wrap. Changing what the resolver formats, or hashing on
    // anything but the measured text, fails this.
    forAll(genValue, genStyle, Gen.choose(1, 200)) { (v, style, width) =>
      val content = renderedContent(v, numFmtOf(style))
      val font = style.map(_.font)
      val indentPx = style.map(_.align.indent).getOrElse(0) * IndentPxPerLevel
      val measured = measureTextWidth(content.text, font)
      val wraps = style.exists(_.align.wrapText)
      val decision = hashOverflowText(v, style, width)
      val fitsSmallestBox = measured <= width - CellPaddingX - indentPx
      val widerThanCell = measured > width
      if fitsSmallestBox then decision.isEmpty
      else if widerThanCell && hashable(content.kind) && !wraps && content.text.nonEmpty then
        decision.exists(m => m.nonEmpty && m.forall(_ == '#'))
      else true
    }
  }

  property("hashOverflowText: a cached formula hashes exactly as its cached value does") {
    forAll(genValue, genStyle, Gen.choose(1, 200)) { (v, style, width) =>
      assertEquals(
        hashOverflowText(CellValue.Formula("A1", Some(v)), style, width),
        hashOverflowText(v, style, width)
      )
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

  property("calculateOverflowColspan is 1 for every hashable kind: only text spills (GH-500)") {
    // Excel never lets a number, date, logical or error borrow a neighbour, whatever its
    // alignment and however empty the cells to the right are.
    forAll(genValue, genStyle, Gen.choose(1, 120)) { (v, style, width) =>
      val base = Sheet("Test").put(ref"A1", v)
      val sheet = style.fold(base)(s => base.withCellStyle(ref"A1", s))
      val colWidths = IndexedSeq(width, 72, 72, 72)
      sheet.cells.get(ref"A1").forall { cell =>
        val kind = renderedContent(v, numFmtOf(style)).kind
        val colspan = calculateOverflowColspan(cell, ref"A1", width, colWidths, sheet, 0, 3)
        !hashable(kind) || colspan == 1
      }
    }
  }

  // ========== Spill spans (Excel text overflow) ==========

  private val longLabel = "A label far wider than three default columns put together, and then some"

  private def spanOf(sheet: Sheet, colWidths: IndexedSeq[Int]): Option[OverflowSpan] =
    sheet.cells.get(ref"B1").map { cell =>
      overflowSpan(ResolvedCell(cell, sheet), ref"B1", colWidths(1), colWidths, sheet, 0, 2)
    }

  property("overflowSpan is none for every hashable kind, on both sides (GH-500)") {
    forAll(genValue, genStyle, Gen.choose(1, 120)) { (v, style, width) =>
      val base = Sheet("Test").put(ref"B1", v)
      val sheet = style.fold(base)(s => base.withCellStyle(ref"B1", s))
      val kind = renderedContent(v, numFmtOf(style)).kind
      !hashable(kind) || spanOf(sheet, IndexedSeq(72, width, 72)).forall(_ == OverflowSpan.none)
    }
  }

  property("overflowSpan: the side a text spills to follows its resolved alignment") {
    forAll(genStyle, Gen.choose(1, 120)) { (style, width) =>
      val base = Sheet("Test").put(ref"B1", CellValue.Text(longLabel))
      val sheet = style.fold(base)(s => base.withCellStyle(ref"B1", s))
      val content = renderedContent(CellValue.Text(longLabel), numFmtOf(style))
      spanOf(sheet, IndexedSeq(72, width, 72)).forall { span =>
        resolveHAlign(style, content) match
          case HAlign.Left | HAlign.General => span.left == 0
          case HAlign.Right => span.right == 0
          case HAlign.Center | HAlign.CenterContinuous => true
          case _ => span == OverflowSpan.none
      }
    }
  }

  property("overflowSpan: a neighbour's formatting never blocks a spill; any value does") {
    forAll(Generators.genCellStyle, Gen.oneOf(HAlign.General, HAlign.Right, HAlign.Center)) {
      (neighbour, h) =>
        val source = CellStyle.default.withAlign(Align(horizontal = h))
        val base =
          Sheet("Test").put(ref"B1", CellValue.Text(longLabel)).withCellStyle(ref"B1", source)
        val formatted = base.withCellStyle(ref"A1", neighbour).withCellStyle(ref"C1", neighbour)
        val filled = base.put(ref"A1", CellValue.Text("")).put(ref"C1", CellValue.Text(""))
        val widths = IndexedSeq(72, 30, 72)
        val free = spanOf(base, widths)
        assert(free.exists(_.spills), s"$h: the label spills into empty neighbours")
        assertEquals(spanOf(formatted, widths), free, s"$h: formatting is not a value")
        assertEquals(spanOf(filled, widths), Some(OverflowSpan.none), s"$h: an empty string is")
    }
  }

  test("calculateOverflowColspan is the rightward extent of overflowSpan") {
    val centred = CellStyle.default.withAlign(Align(horizontal = HAlign.Center))
    val sheet =
      Sheet("Test").put(ref"B1", CellValue.Text(longLabel)).withCellStyle(ref"B1", centred)
    val widths = IndexedSeq(72, 30, 72)
    val cell = sheet.cells.get(ref"B1").getOrElse(fail("B1"))
    assertEquals(spanOf(sheet, widths), Some(OverflowSpan(1, 1)))
    assertEquals(calculateOverflowColspan(cell, ref"B1", 30, widths, sheet, 0, 2), 2)
  }

  // ========== Row autofit ==========

  test("autofitLineHeightPx: Excel's Calibri autofit rows, snapped to whole pixels") {
    assertEquals(
      List(11.0, 14.0, 18.0, 24.0).map(autofitLineHeightPx),
      List(20, 25, 31, 42),
      "15pt, 18.75pt, 23.25pt and 31.5pt at 96 DPI"
    )
  }

  property("autofitLineHeightPx grows with the font and always fits the font's em") {
    forAll(Gen.choose(1.0, 200.0), Gen.choose(0.0, 50.0)) { (pt, more) =>
      val h = autofitLineHeightPx(pt)
      h >= pt * 4.0 / 3.0 && autofitLineHeightPx(pt + more) >= h
    }
  }

  property("calculateRowHeights: never below the default, and an explicit height is kept exactly") {
    forAll(Gen.choose(1.0, 60.0), Gen.option(Gen.choose(1.0, 100.0))) { (pt, explicit) =>
      val big = CellStyle.default.withFont(com.tjclp.xl.styles.font.Font.default.withSize(pt))
      val base = Sheet("Test").put(ref"A1", CellValue.Text("x")).withCellStyle(ref"A1", big)
      val sheet = explicit.fold(base) { h =>
        base.setRowProperties(com.tjclp.xl.addressing.Row.from0(0), RowProperties(height = Some(h)))
      }
      val height = calculateRowHeights(sheet, ref"A1:A1").headOption.getOrElse(-1)
      explicit match
        case Some(h) => height == excelRowHeightToPixels(h)
        case None => height == math.max(DefaultCellHeightPx, autofitLineHeightPx(pt))
    }
  }

  // ========== One resolution per cell ==========

  /**
   * A sheet whose cells exercise every per-cell path: a Custom numeric code (parsed on each
   * resolution), a date code, a logical, an error, text that fits, text that spills, a left-aligned
   * number that hashes, a wrapped cell and an uncached formula.
   */
  private val perCellPathsSheet: Sheet =
    val acct = CellStyle.default.withNumFmt(
      NumFmt.Custom("_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)")
    )
    val date = CellStyle.default.withNumFmt(NumFmt.Custom("m/d/yyyy h:mm"))
    val left = CellStyle.default.withAlign(Align(horizontal = HAlign.Left))
    val wrap = CellStyle.default.withAlign(Align(wrapText = true))
    Sheet("Test")
      .put(ref"A1" -> 1234.87)
      .put(ref"B1", CellValue.DateTime(LocalDateTime.of(2025, 11, 21, 9, 30)))
      .put(ref"C1" -> true)
      .put(ref"A2", CellValue.Error(CellError.Div0))
      .put(ref"B2" -> "fits")
      .put(ref"C2" -> "a long text that spills into the empty cells to its right")
      .put(ref"A3" -> 1234567.9)
      .put(ref"B3" -> "wrapped text over several lines")
      .put(ref"C3", CellValue.Formula("A1+1", None))
      .unsafe
      .withCellStyle(ref"A1", acct)
      .withCellStyle(ref"B1", date)
      .withCellStyle(ref"A3", left)
      .withCellStyle(ref"B3", wrap)
      .setColumnProperties(Column.from0(0), ColumnProperties(width = Some(6.0)))

  test("SvgRenderer resolves each cell's rendered content once, and the seam changes no byte") {
    val range = ref"A1:E4"
    val cells = perCellPathsSheet.cells.size
    var resolutions = 0
    val counting: (Cell, Sheet) => ResolvedCell = (cell, sheet) =>
      resolutions += 1
      ResolvedCell(cell, sheet)
    val svg = SvgRenderer.toSvgResolving(counting)(
      perCellPathsSheet,
      range,
      includeStyles = true,
      theme = ThemePalette.office,
      showLabels = false,
      showGridlines = false
    )
    assertEquals(resolutions, cells, "one resolution per rendered cell (a Custom code parsed once)")
    assertEquals(svg, perCellPathsSheet.toSvg(range))
  }

  test("HtmlRenderer resolves each cell's rendered content once, and the seam changes no byte") {
    val range = ref"A1:E4"
    val cells = perCellPathsSheet.cells.size
    var resolutions = 0
    val counting: (Cell, Sheet) => ResolvedCell = (cell, sheet) =>
      resolutions += 1
      ResolvedCell(cell, sheet)
    val html = HtmlRenderer.toHtmlResolving(counting)(
      perCellPathsSheet,
      range,
      includeStyles = true,
      includeComments = true,
      theme = ThemePalette.office,
      applyPrintScale = false,
      showLabels = false
    )
    assertEquals(resolutions, cells, "one resolution per rendered cell (a Custom code parsed once)")
    assertEquals(html, perCellPathsSheet.toHtml(range))
  }

  test("both renderers take text, alignment and hash from the injected content alone") {
    // A resolver that calls every cell the text "X": a decision that re-resolved from the raw
    // value would show that cell's digits, hash it or right-align it instead.
    val range = ref"A1:E4"
    val cells = perCellPathsSheet.cells.size
    val asText: (Cell, Sheet) => ResolvedCell = (cell, sheet) =>
      ResolvedCell(cell, sheet).copy(content = RenderedContent(RenderedKind.Text, "X"))

    val svg = SvgRenderer.toSvgResolving(asText)(
      perCellPathsSheet,
      range,
      includeStyles = true,
      theme = ThemePalette.office,
      showLabels = false,
      showGridlines = false
    )
    val plainTexts = """<text[^>]*>([^<]*)</text>""".r.findAllMatchIn(svg).map(_.group(1)).toList
    val wrappedLines =
      """<tspan[^>]*>([^<]*)</tspan>""".r.findAllMatchIn(svg).map(_.group(1)).toList
    assertEquals(plainTexts, List.fill(cells - 1)("X"), s"every unwrapped cell draws X: $svg")
    assertEquals(wrappedLines, List("X"), s"the wrapped cell draws X: $svg")
    val anchors = """text-anchor="([a-z]+)"""".r.findAllMatchIn(svg).map(_.group(1)).toList
    assertEquals(anchors.distinct, List("start"), s"text is left-anchored: $svg")

    val html = HtmlRenderer.toHtmlResolving(asText)(
      perCellPathsSheet,
      range,
      includeStyles = true,
      includeComments = true,
      theme = ThemePalette.office,
      applyPrintScale = false,
      showLabels = false
    )
    val bodies = """<td[^>]*>([^<]*)</td>""".r.findAllMatchIn(html).map(_.group(1)).toList
    assertEquals(bodies.filter(_.nonEmpty), List.fill(cells)("X"), s"every cell shows X: $html")
    assert(!html.contains("text-align: right") && !html.contains("text-align: center"), html)
  }
