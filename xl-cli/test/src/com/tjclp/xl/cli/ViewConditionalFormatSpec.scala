package com.tjclp.xl.cli

import java.nio.charset.StandardCharsets.UTF_8
import java.nio.file.Files

import cats.effect.{IO, Ref}
import cats.syntax.all.*
import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.contract.{
  CliError,
  CliException,
  ErrorCode,
  Outcome,
  OutputMode,
  Render,
  Warning,
  WarningCode
}
import com.tjclp.xl.cli.raster.BatikRasterizer
import com.tjclp.xl.cli.read.{InMemorySource, ReadTestKit}
import com.tjclp.xl.formula.eval.{CfEvaluation, CfUnevaluated}
import com.tjclp.xl.sheets.Sheet

/**
 * GH-497: `view` paints conditional formatting in its pictures (svg, html and every raster built
 * from the SVG) as Excel does, with no flag; a rule it cannot paint is named by an informational
 * CF_NOT_RENDERED warning that never gates, and a failing evaluation degrades to an unpainted view,
 * never INTERNAL.
 */
class ViewConditionalFormatSpec extends CatsEffectSuite:

  private val redFont = Dxf.font(DxfFont(color = Some(Color.Rgb(0xffff0000))))
  private val pink = Dxf.fill(Color.Rgb(0xffffc7ce))

  private def range(a1: String): CellRange = CellRange.parse(a1).fold(fail(_), identity)
  private def aref(a1: String): ARef = ARef.parse(a1).fold(fail(_), identity)
  private def num(n: Double): CellValue = CellValue.Number(BigDecimal(n))

  /**
   * Weaver's shape: a header row, then numbers under a "< 0" red-font rule (priority 1) and a
   * 3-colour scale (priority 2).
   */
  private val weaver: Workbook =
    val sheet = Sheet("Data")
      .put(aref("A1"), CellValue.Text("Item"))
      .put(aref("B1"), CellValue.Text("Delta"))
      .put(aref("A2"), CellValue.Text("north"))
      .put(aref("B2"), num(12))
      .put(aref("A3"), CellValue.Text("south"))
      .put(aref("B3"), num(-3))
      .put(aref("A4"), CellValue.Text("east"))
      .put(aref("B4"), num(0))
      .conditionalFormat(
        range("B2:B4"),
        CfRule.cellIs(CfOperator.LessThan, "0", redFont),
        CfRule.colorScale3(
          CfPoint(Cfvo.Min, Color.Rgb(0xfff8696b)),
          CfPoint(Cfvo.Percentile(BigDecimal(50)), Color.Rgb(0xffffeb84)),
          CfPoint(Cfvo.Max, Color.Rgb(0xff63be7b))
        )
      )
    Workbook(Vector(sheet))

  private def view(
    wb: Workbook,
    window: String,
    format: ViewFormat,
    evalFormulas: Boolean = false,
    strict: Boolean = false,
    mode: OutputMode = OutputMode.Text
  ): IO[Outcome] =
    ReadTestKit.inMemory(
      wb,
      Some("Data"),
      ReadTestKit.view(Some(window), format, evalFormulas = evalFormulas, strict = strict),
      mode
    )

  private def cfWarnings(outcome: Outcome): Vector[String] =
    outcome.warnings.filter(_.code == WarningCode.CF_NOT_RENDERED).map(_.message)

  /** The fill of the SVG background rect drawn at (x, y). */
  private def rectFill(svg: String, x: Int, y: Int): Option[String] =
    s"""<rect x="$x" y="$y" width="\\d+" height="\\d+" fill="(#[0-9A-F]{6})"[^>]*class="cell"/>""".r
      .findFirstMatchIn(svg)
      .map(_.group(1))

  /** The fill attribute of the text drawn under a cell's clip. */
  private def textFill(svg: String, a1: String): Option[String] =
    s"""<text [^>]*fill="(#[0-9A-F]{6})"[^>]*clip-path="url\\(#clip-$a1\\)"""".r
      .findFirstMatchIn(svg)
      .map(_.group(1))

  test("svg paints Weaver's book: the red-font rule on the negative, the scale on every number") {
    view(weaver, "A1:B4", ViewFormat.Svg).map { outcome =>
      val svg = ReadTestKit.text(outcome)
      assertEquals(rectFill(svg, 72, 0), Some("#FFFFFF"), s"the header is unpainted: $svg")
      assertEquals(rectFill(svg, 72, 20), Some("#63BE7B"), "12 is the maximum")
      assertEquals(rectFill(svg, 72, 40), Some("#F8696B"), "-3 is the minimum")
      assertEquals(rectFill(svg, 72, 60), Some("#FFEB84"), "0 is the median")
      assertEquals(textFill(svg, "B3"), Some("#FF0000"), "the negative is drawn red")
      assertEquals(textFill(svg, "B2"), Some("#000000"))
      assertEquals(outcome.warnings, Vector.empty)
    }
  }

  test("html paints the same cells") {
    view(weaver, "A1:B4", ViewFormat.Html).map { outcome =>
      val html = ReadTestKit.text(outcome)
      val negative = """<td style="([^"]*)">-3</td>""".r
        .findFirstMatchIn(html)
        .map(_.group(1))
        .getOrElse(fail(html))
      assert(negative.contains("color: #FF0000"), negative)
      assert(negative.contains("background-color: #F8696B"), negative)
      assert(html.contains("background-color: #63BE7B"), html)
    }
  }

  test("a sheet without conditional formatting renders exactly as the CF-blind library call") {
    val sheet = Sheet("Data").put(aref("A1"), num(1)).put(aref("B2"), num(-1))
    val plain = Workbook(Vector(sheet))
    view(plain, "A1:B2", ViewFormat.Svg).map { outcome =>
      val expected = sheet.toSvg(range("A1:B2"), theme = plain.metadata.theme)
      assertEquals(ReadTestKit.text(outcome), expected)
    }
  }

  test("rules see cached values, and live ones under --eval") {
    val sheet = Sheet("Data")
      .put(aref("A1"), CellValue.Formula("B1*2", Some(num(0))))
      .put(aref("B1"), num(75))
      .conditionalFormat(range("A1"), CfRule.cellIs(CfOperator.GreaterThan, "100", pink))
    val wb = Workbook(Vector(sheet))
    for
      cached <- view(wb, "A1:B1", ViewFormat.Svg)
      live <- view(wb, "A1:B1", ViewFormat.Svg, evalFormulas = true)
    yield
      assertEquals(rectFill(ReadTestKit.text(cached), 0, 0), Some("#FFFFFF"), "cached 0")
      assertEquals(rectFill(ReadTestKit.text(live), 0, 0), Some("#FFC7CE"), "live 150")
  }

  /**
   * B1:B10 = 1..10; A1:A10 = Bn, each cached ten times too high (saved before B changed), under a
   * red-to-green scale on A1:A10.
   */
  private val staleScale: Workbook =
    val cells = (1 to 10).foldLeft(Sheet("Data")) { (s, i) =>
      s.put(aref(s"B$i"), num(i.toDouble))
        .put(aref(s"A$i"), CellValue.Formula(s"B$i", Some(num(i * 10.0))))
    }
    val sheet = cells.conditionalFormat(
      range("A1:A10"),
      CfRule.colorScale2(
        CfPoint(Cfvo.Min, Color.Rgb(0xffff0000)),
        CfPoint(Cfvo.Max, Color.Rgb(0xff00ff00))
      )
    )
    Workbook(Vector(sheet))

  test("under --eval a scale paints from the live block, whatever window is asked for") {
    // live 1..10: 2 and 3 sit at 1/9 and 2/9 of the span; the stale caches (max 100) never count
    for
      short <- view(staleScale, "A1:B3", ViewFormat.Svg, evalFormulas = true)
      full <- view(staleScale, "A1:B10", ViewFormat.Svg, evalFormulas = true)
    yield List(short, full).foreach { outcome =>
      val svg = ReadTestKit.text(outcome)
      assertEquals(rectFill(svg, 0, 20), Some("#E31C00"), s"A2 = 2: $svg")
      assertEquals(rectFill(svg, 0, 40), Some("#C63900"), "A3 = 3")
      assertEquals(outcome.warnings, Vector.empty)
    }
  }

  test("under --eval a formula rule reads live values outside the window") {
    val sheet = Sheet("Data")
      .put(aref("A1"), num(1))
      .put(aref("C1"), CellValue.Formula("D1*2", Some(num(0))))
      .put(aref("D1"), num(10))
      .conditionalFormat(range("A1:A3"), CfRule.expression("$C1>5", pink))
    for
      cached <- view(Workbook(Vector(sheet)), "A1:A1", ViewFormat.Svg)
      live <- view(Workbook(Vector(sheet)), "A1:A1", ViewFormat.Svg, evalFormulas = true)
    yield
      assertEquals(rectFill(ReadTestKit.text(cached), 0, 0), Some("#FFFFFF"), "cached C1 = 0")
      assertEquals(rectFill(ReadTestKit.text(live), 0, 0), Some("#FFC7CE"), "live C1 = 20")
  }

  test("under --eval a block cell that fails is EVAL_FAILED, and --strict gates on it") {
    val broken = staleScale.sheets.headOption.getOrElse(fail("no sheet"))
    val wb =
      Workbook(Vector(broken.put(aref("A9"), CellValue.Formula("NOSUCHFN(1)", Some(num(90))))))
    for
      warned <- view(wb, "A1:B3", ViewFormat.Svg, evalFormulas = true)
      gated <- view(wb, "A1:B3", ViewFormat.Svg, evalFormulas = true, strict = true).attempt
    yield
      assert(warned.ok, warned.error.toString)
      warned.warnings.filter(_.code == WarningCode.EVAL_FAILED) match
        case Vector(warning) => assertEquals(warning.location.flatMap(_.ref), Some("A9"))
        case other => fail(s"expected one EVAL_FAILED, got $other")
      val code = gated match
        case Right(outcome) => outcome.error.map(_.code)
        case Left(e: CliException) => Some(e.error.code)
        case Left(other) => fail(s"unexpected $other")
      assertEquals(code, Some(ErrorCode.RECALC_GATE))
  }

  test("a rule that cannot be evaluated is one CF_NOT_RENDERED warning; the render proceeds") {
    val sheet = Sheet("Data")
      .put(aref("A1"), num(1))
      .put(aref("A2"), num(2))
      .conditionalFormat(range("A1:A3"), CfRule.expression("NOSUCHFN(A1)>0", pink))
    view(Workbook(Vector(sheet)), "A1:A3", ViewFormat.Svg).map { outcome =>
      assert(outcome.ok, outcome.error.toString)
      assert(ReadTestKit.text(outcome).startsWith("<svg"), "the picture is still drawn")
      cfWarnings(outcome) match
        case Vector(message) =>
          assert(message.contains("expression"), message)
          assert(message.contains("(priority 1)"), message)
          assert(message.contains("on A1:A3"), message)
        case other => fail(s"expected one CF_NOT_RENDERED, got $other")
    }
  }

  test("--json: ok, with the CF_NOT_RENDERED warning in the envelope") {
    val sheet = Sheet("Data")
      .put(aref("A1"), num(1))
      .conditionalFormat(range("A1"), CfRule.expression("NOSUCHFN(A1)>0", pink))
    view(Workbook(Vector(sheet)), "A1", ViewFormat.Svg, mode = OutputMode.Json).map { outcome =>
      val envelope = ujson.read(Render.json(outcome, "test").stdout)
      assertEquals(envelope("ok").bool, true)
      assertEquals(envelope("exitCode").num, 0.0)
      assertEquals(
        envelope("warnings").arr.map(_("code").str).toVector,
        Vector(WarningCode.CF_NOT_RENDERED)
      )
    }
  }

  test("CF_NOT_RENDERED never gates: view --eval --strict still succeeds") {
    val sheet = Sheet("Data")
      .put(aref("A1"), CellValue.Formula("1+1", None))
      .conditionalFormat(range("A1"), CfRule.expression("NOSUCHFN(A1)>0", pink))
    view(Workbook(Vector(sheet)), "A1", ViewFormat.Svg, evalFormulas = true, strict = true).map {
      outcome =>
        assert(outcome.ok, outcome.error.toString)
        assertEquals(outcome.exitCode.code, 0)
        assertEquals(cfWarnings(outcome).size, 1)
    }
  }

  /**
   * A1 = 0, then An = A(n-1)+1 down to A300 (uncached, as openpyxl writes it, unless `cached`),
   * under a red-to-green scale on A1:A300; `broken` replaces A3 with a function xl lacks.
   */
  private def chain(cached: Boolean, broken: Boolean = false): Workbook =
    val filled = (2 to 300).foldLeft(Sheet("Data").put(aref("A1"), num(0))) { (s, i) =>
      val value = Option.when(cached)(num((i - 1).toDouble))
      s.put(ARef.from0(0, i - 1), CellValue.Formula(s"A${i - 1}+1", value))
    }
    val cells =
      if broken then filled.put(aref("A3"), CellValue.Formula("NOSUCHFN(1)", None)) else filled
    val sheet = cells
      .conditionalFormat(
        range("A1:A300"),
        CfRule.colorScale2(
          CfPoint(Cfvo.Min, Color.Rgb(0xffff0000)),
          CfPoint(Cfvo.Max, Color.Rgb(0xff00ff00))
        )
      )
    Workbook(Vector(sheet))

  test("view --eval paints by the whole uncached chain, however deep, as if it were cached") {
    for
      live <- view(chain(cached = false), "A1:A5", ViewFormat.Svg, evalFormulas = true)
      plain <- view(chain(cached = false), "A1:A5", ViewFormat.Svg)
      cached <- view(chain(cached = true), "A1:A5", ViewFormat.Svg)
    yield List(live, plain).foreach { outcome =>
      assertEquals(outcome.warnings, Vector.empty)
      val svg = ReadTestKit.text(outcome)
      assertEquals(rectFill(svg, 0, 20), Some("#FE0100"), s"1 of 0..299: $svg")
      (0 to 4).foreach { row =>
        val y = row * 20
        assertEquals(rectFill(svg, 0, y), rectFill(ReadTestKit.text(cached), 0, y), s"row $row")
      }
    }
  }

  test("a cell the scale's statistics cannot compute is one CF_NOT_RENDERED; nothing is guessed") {
    view(chain(cached = false, broken = true), "A1:A5", ViewFormat.Svg, evalFormulas = true).map {
      outcome =>
        assert(outcome.ok, outcome.error.toString)
        val svg = ReadTestKit.text(outcome)
        (0 to 4).foreach(row => assertEquals(rectFill(svg, 0, row * 20), Some("#FFFFFF"), svg))
        cfWarnings(outcome) match
          case Vector(message) =>
            assert(message.contains("colorScale (priority 1) on A1:A300"), message)
            assert(message.contains("NOSUCHFN(1)") && message.contains("A3"), message)
          case other => fail(s"expected one CF_NOT_RENDERED, got $other")
    }
  }

  test("a scale that cannot compute its cells never stops another block's scale, in any order") {
    // A3000 = 2999 by a 3000-cell uncached chain: A:A reads it row-major, B1 in one recursion
    val cells = (2 to 3000)
      .foldLeft(Sheet("Data").put(aref("A1"), num(0))) { (s, i) =>
        s.put(ARef.from0(0, i - 1), CellValue.Formula(s"A${i - 1}+1", None))
      }
      .put(aref("B1"), CellValue.Formula("A3000", None))
      .put(aref("B2"), num(5))
    val scale = CfRule.colorScale2(
      CfPoint(Cfvo.Min, Color.Rgb(0xffff0000)),
      CfPoint(Cfvo.Max, Color.Rgb(0xff00ff00))
    )
    val x = ConditionalFormat.Rules(Vector(range("B1:B2")), Vector(scale))
    val y = ConditionalFormat.Rules(Vector(range("A:A")), Vector(scale))
    List(Vector(x, y), Vector(y, x)).traverse_ { order =>
      val wb = Workbook(Vector(cells.copy(conditionalFormats = order)))
      view(wb, "A1:B5", ViewFormat.Svg).map { outcome =>
        val svg = ReadTestKit.text(outcome)
        assertEquals(rectFill(svg, 0, 0), Some("#FF0000"), s"A:A paints its minimum: $svg")
        assertEquals(rectFill(svg, 72, 0), Some("#FFFFFF"), "B1:B2 is not guessed")
        cfWarnings(outcome) match
          case Vector(message) =>
            assert(message.contains("on B1:B2") && message.contains("(cell B1)"), message)
          case other => fail(s"expected one CF_NOT_RENDERED for B1:B2, got $other")
      }
    }
  }

  // ========== Files another producer wrote ==========

  /** The book written to a temp file, `blocks` injected before its first CF block, read back. */
  private def withInjectedCf(wb: Workbook, blocks: String)(test: Workbook => IO[Unit]): IO[Unit] =
    ReadTestKit.withTempWorkbook(wb) { path =>
      ReadTestKit.rewriteZip(path) { (name, bytes) =>
        if name != "xl/worksheets/sheet1.xml" then bytes
        else
          val xml = new String(bytes, UTF_8)
          val at = xml.indexOf("<conditionalFormatting")
          assert(at >= 0, s"the writer emitted no CF block: $xml")
          (xml.substring(0, at) + blocks + xml.substring(at)).getBytes(UTF_8)
      } *> ReadTestKit.excel.read(path).flatMap(test)
    }

  private val lessThanFive: Workbook =
    val sheet = Sheet("Data")
      .put(aref("A1"), num(3))
      .put(aref("A3"), num(9))
      .conditionalFormat(
        range("A1:A3"),
        CfRule.CellIs(CfOperator.LessThan, "5", None, Some(pink), 2)
      )
    Workbook(Vector(sheet))

  test("Excel's 'Blanks, Stop If True' guard keeps the blank cell unpainted") {
    val guard =
      """<conditionalFormatting sqref="A1:A3"><cfRule type="containsBlanks" priority="1" stopIfTrue="1"><formula>LEN(TRIM(A1))=0</formula></cfRule></conditionalFormatting>"""
    withInjectedCf(lessThanFive, guard) { wb =>
      view(wb, "A1:A3", ViewFormat.Svg).map { outcome =>
        val svg = ReadTestKit.text(outcome)
        assertEquals(rectFill(svg, 0, 0), Some("#FFC7CE"), s"3 < 5 is painted: $svg")
        assertEquals(rectFill(svg, 0, 20), Some("#FFFFFF"), "the blank A2 is guarded")
        assertEquals(rectFill(svg, 0, 40), Some("#FFFFFF"), "9 is not < 5")
        assertEquals(outcome.warnings, Vector.empty)
      }
    }
  }

  test("an icon set is named by CF_NOT_RENDERED and the rest still paints") {
    val icons =
      """<conditionalFormatting sqref="A1:A3"><cfRule type="iconSet" priority="5"><iconSet><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule></conditionalFormatting>"""
    withInjectedCf(lessThanFive, icons) { wb =>
      view(wb, "A1:A3", ViewFormat.Svg).map { outcome =>
        assert(outcome.ok, outcome.error.toString)
        assertEquals(rectFill(ReadTestKit.text(outcome), 0, 0), Some("#FFC7CE"))
        cfWarnings(outcome) match
          case Vector(message) => assert(message.contains("iconSet (priority 5) on A1:A3"), message)
          case other => fail(s"expected one CF_NOT_RENDERED, got $other")
      }
    }
  }

  test("png export paints conditional formatting through the SVG (Batik)") {
    BatikRasterizer.isAvailable.flatMap { batikAvailable =>
      assume(batikAvailable, "AWT not available - skipping the raster smoke test")
      val out = Files.createTempFile("xl-cli-cf-", ".png")
      out.toFile.deleteOnExit()
      ReadTestKit
        .inMemory(
          weaver,
          Some("Data"),
          ReadTestKit.view(Some("A1:B4"), ViewFormat.Png, rasterOutput = Some(out))
        )
        .map { outcome =>
          assert(outcome.ok, outcome.error.toString)
          assert(ReadTestKit.text(outcome).contains("Exported:"), ReadTestKit.text(outcome))
          assert(Files.size(out) > 0, "the PNG has content")
          assertEquals(outcome.warnings, Vector.empty)
        }
        .guarantee(IO(Files.deleteIfExists(out)).void)
    }
  }

  // ========== The effect boundary ==========

  private def paintWith(evaluate: => CfEvaluation): IO[(CfOverlay, Vector[Warning])] =
    for
      seen <- Ref.of[IO, Vector[Warning]](Vector.empty)
      overlay <- InMemorySource.conditionalPaint(evaluate, w => seen.update(_ :+ w))
      warnings <- seen.get
    yield (overlay, warnings)

  test("the evaluation's unpainted rules become CF_NOT_RENDERED warnings, in order") {
    val paint = CfOverlay(Map(aref("A1") -> CfPaint(pink, None)))
    val u = (p: Int) =>
      CfUnevaluated(Some(p), "iconSet", Vector(range("A1:A3")), CfUnevaluated.Reason.NotModeled)
    paintWith(CfEvaluation(paint, Vector(u(2), u(4)))).map { (overlay, warnings) =>
      assertEquals(overlay, paint)
      assertEquals(warnings.map(_.code), Vector.fill(2)(WarningCode.CF_NOT_RENDERED))
      assertEquals(warnings.map(_.message), Vector(u(2).message, u(4).message))
    }
  }

  test("an evaluator defect degrades to an unpainted render and one warning, never INTERNAL") {
    List[Throwable](
      new ClassCastException("CellValue$Number cannot be cast to BigDecimal"),
      new StackOverflowError()
    ).traverse { thrown =>
      paintWith(throw thrown).map { (overlay, warnings) =>
        assertEquals(overlay, CfOverlay.empty)
        assertEquals(warnings.map(_.code), Vector(WarningCode.CF_NOT_RENDERED))
        assert(warnings.forall(_.message.contains(thrown.getClass.getSimpleName)), warnings)
      }
    }.void
  }

  test("a typed CLI failure (RESOURCE_LIMIT) is re-raised, not degraded") {
    val limit = CliException(CliError(ErrorCode.RESOURCE_LIMIT, "heap exhausted"))
    paintWith(throw limit).attempt.map {
      case Left(e: CliException) => assertEquals(e.error.code, ErrorCode.RESOURCE_LIMIT)
      case other => fail(s"expected RESOURCE_LIMIT, got $other")
    }
  }
