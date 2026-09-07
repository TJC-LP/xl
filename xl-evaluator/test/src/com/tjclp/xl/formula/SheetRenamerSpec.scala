package com.tjclp.xl.formula

import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.{CellError, CellValue, FormulaKind}
import com.tjclp.xl.cf.{CfOperator, CfPoint, CfRule, Cfvo, ConditionalFormat}
import com.tjclp.xl.charts.{Chart, ChartType, DataRef, Series, SeriesName}
import com.tjclp.xl.formula.eval.SheetRenamer
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.formula.printer.FormulaShifter
import com.tjclp.xl.ooxml.{XlsxReader, XlsxWriter}
import com.tjclp.xl.sheets.{DataValidation, DvBoundedType, DvKind, DvOperator}
import com.tjclp.xl.styles.Dxf
import com.tjclp.xl.styles.color.Color
import com.tjclp.xl.workbooks.{DefinedName, WorkbookMetadata}
import munit.ScalaCheckSuite
import org.scalacheck.Gen
import org.scalacheck.Prop.*

/**
 * GH-559: renaming a sheet must rewrite every reference to it — cell formulas on every sheet
 * (caches and record kinds preserved: a rename changes no value), defined names, conditional-format
 * and data-validation formulas — refusing before mutating when a text that mentions the sheet
 * cannot be parsed. `Workbook.rename` itself stays tab-only.
 *
 * ADR-017 invariant 1 is pinned through a REAL FILE: the surgical writer copies unmodified sheets
 * byte-for-byte from the source zip, so a renamer that edited other sheets via `copy(sheets = …)`
 * would succeed in memory and still write `Sheet1!A1` to disk.
 */
@SuppressWarnings(Array("org.wartremover.warts.IterableOps"))
class SheetRenamerSpec extends ScalaCheckSuite:

  private val Sheet1 = SheetName.unsafe("Sheet1")
  private val Data = SheetName.unsafe("Data")
  private val Q1Data = SheetName.unsafe("Q1 Data")
  private val red = Dxf.fill(Color.Rgb(0xffff0000))
  private val color = Color.Rgb(0xff0000ff)

  private def num(i: Int): CellValue = CellValue.Number(BigDecimal(i))
  private def f(text: String, cached: Option[CellValue] = None): CellValue =
    CellValue.Formula(text, cached)

  private def sheetNamed(wb: Workbook, name: String): Sheet =
    wb.sheets.find(_.name.value == name).getOrElse(fail(s"missing sheet $name in ${wb.sheetNames}"))

  private def formulaText(wb: Workbook, sheet: String, ref: ARef): String =
    sheetNamed(wb, sheet)(ref).value match
      case CellValue.Formula(text, _, _) => text
      case other => fail(s"expected a formula at $sheet!${ref.toA1}, got $other")

  private def cachedOf(wb: Workbook, sheet: String, ref: ARef): Option[CellValue] =
    sheetNamed(wb, sheet)(ref).value match
      case CellValue.Formula(_, cached, _) => cached
      case other => fail(s"expected a formula at $sheet!${ref.toA1}, got $other")

  private def rename(wb: Workbook, from: SheetName, to: SheetName): Workbook =
    SheetRenamer.rename(wb, from, to).fold(e => fail(s"rename refused: ${e.message}"), identity)

  /** The #559 repro: Sheet2 reads Sheet1 three ways and every formula carries a cache. */
  private def repro: Workbook =
    Workbook(
      Sheet("Sheet1").put(ref"A1", num(5)).put(ref"B1", f("Sheet1!A1+1", Some(num(6)))),
      Sheet("Sheet2")
        .put(ref"A1", f("Sheet1!A1*2", Some(num(10))))
        .put(ref"B1", f("SUM(Sheet1!A1:A1)", Some(num(5))))
        .put(ref"C1", f("A1*2", Some(num(20))))
        .put(ref"D1", f("\"Sheet1\"", Some(CellValue.Text("Sheet1"))))
        .put(ref"E1", f("Sheet1!A1/0", Some(CellValue.Error(CellError.Div0))))
    )

  // ===== cell formulas =====

  test("GH-559 repro: every dependent formula follows the rename, caches preserved") {
    val renamed = rename(repro, Sheet1, Data)
    assertEquals(renamed.sheets.map(_.name.value), Vector("Data", "Sheet2"))
    assertEquals(formulaText(renamed, "Sheet2", ref"A1"), "Data!A1*2")
    assertEquals(cachedOf(renamed, "Sheet2", ref"A1"), Some(num(10)))
    assertEquals(formulaText(renamed, "Sheet2", ref"B1"), "SUM(Data!A1:A1)")
    assertEquals(cachedOf(renamed, "Sheet2", ref"B1"), Some(num(5)))
    // a cached ERROR value is a value too — preserved verbatim
    assertEquals(formulaText(renamed, "Sheet2", ref"E1"), "Data!A1/0")
    assertEquals(cachedOf(renamed, "Sheet2", ref"E1"), Some(CellValue.Error(CellError.Div0)))
    // the renamed sheet's own self-qualified reference follows too
    assertEquals(formulaText(renamed, "Data", ref"B1"), "Data!A1+1")
    assertEquals(cachedOf(renamed, "Data", ref"B1"), Some(num(6)))
    // values on the renamed sheet are untouched
    assertEquals(sheetNamed(renamed, "Data")(ref"A1").value, num(5))
  }

  test("formulas that never mention the sheet stay byte-identical (no canonicalising reprint)") {
    val wb = repro.put(
      sheetNamed(repro, "Sheet2").put(ref"F1", f("SUM(A1, B1)", Some(num(15))))
    )
    val renamed = rename(wb, Sheet1, Data)
    assertEquals(sheetNamed(renamed, "Sheet2")(ref"C1").value, f("A1*2", Some(num(20))))
    // spaced separators survive: the text was never parsed and reprinted
    assertEquals(sheetNamed(renamed, "Sheet2")(ref"F1").value, f("SUM(A1, B1)", Some(num(15))))
  }

  test("a string literal spelling the sheet name is not a reference") {
    val renamed = rename(repro, Sheet1, Data)
    assertEquals(
      sheetNamed(renamed, "Sheet2")(ref"D1").value,
      f("\"Sheet1\"", Some(CellValue.Text("Sheet1")))
    )
    val literalRef = Workbook(
      Sheet("Sheet1").put(ref"A1", num(1)),
      Sheet("Sheet2").put(ref"A1", f("IF(A2=\"Sheet1!A1\",1,0)", Some(num(0))))
    )
    assertEquals(
      formulaText(rename(literalRef, Sheet1, Data), "Sheet2", ref"A1"),
      "IF(A2=\"Sheet1!A1\",1,0)"
    )
  }

  test("a new name that needs quoting is quoted in every rewritten formula") {
    val renamed = rename(repro, Sheet1, Q1Data)
    assertEquals(renamed.sheets.headOption.map(_.name.value), Some("Q1 Data"))
    assertEquals(formulaText(renamed, "Sheet2", ref"A1"), "'Q1 Data'!A1*2")
    assertEquals(formulaText(renamed, "Sheet2", ref"B1"), "SUM('Q1 Data'!A1:A1)")
    // and a quoted OLD name is recognised on the way back
    val back = rename(renamed, Q1Data, Sheet1)
    assertEquals(formulaText(back, "Sheet2", ref"A1"), "Sheet1!A1*2")
    assertEquals(formulaText(back, "Sheet2", ref"B1"), "SUM(Sheet1!A1:A1)")
  }

  test("the old name matches case-insensitively, the way Excel resolves sheet names") {
    val wb = Workbook(
      Sheet("Sheet1").put(ref"A1", num(1)),
      Sheet("Sheet2").put(ref"A1", f("sheet1!A1+SHEET1!A1", Some(num(2))))
    )
    assertEquals(formulaText(rename(wb, Sheet1, Data), "Sheet2", ref"A1"), "Data!A1+Data!A1")
  }

  test("a sheet-qualified defined name (SheetNameRef) is rewritten") {
    val wb = Workbook(
      Vector(
        Sheet("Sheet1").put(ref"A1", num(3)),
        Sheet("Sheet2").put(ref"A1", f("Sheet1!rate*2"))
      ),
      metadata = WorkbookMetadata(definedNames =
        Vector(DefinedName("rate", "Sheet1!$A$1", localSheetId = Some(0)))
      )
    )
    val renamed = rename(wb, Sheet1, Data)
    assertEquals(formulaText(renamed, "Sheet2", ref"A1"), "Data!rate*2")
    assertEquals(renamed.metadata.definedNames.map(_.formula), Vector("Data!$A$1"))
  }

  test("external-workbook references to a same-named sheet are identity") {
    val wb = Workbook(
      Sheet("Sheet1").put(ref"A1", num(1)),
      Sheet("Sheet2")
        .put(ref"A1", f("[2]Sheet1!A1*2", Some(num(4))))
        .put(ref"B1", f("SUM([2]Sheet1!A1:B2)", Some(num(9))))
        .put(ref"C1", f("Sheet1!A1+[2]Sheet1!A1", Some(num(3))))
    )
    val renamed = rename(wb, Sheet1, Data)
    assertEquals(sheetNamed(renamed, "Sheet2")(ref"A1").value, f("[2]Sheet1!A1*2", Some(num(4))))
    assertEquals(
      sheetNamed(renamed, "Sheet2")(ref"B1").value,
      f("SUM([2]Sheet1!A1:B2)", Some(num(9)))
    )
    assertEquals(formulaText(renamed, "Sheet2", ref"C1"), "Data!A1+[2]Sheet1!A1")
  }

  test("a sibling whose name merely CONTAINS the old name is untouched") {
    val wb = Workbook(
      Sheet("Sheet1").put(ref"A1", num(1)),
      Sheet("Sheet10").put(ref"A1", num(2)),
      Sheet("Sheet2").put(ref"A1", f("Sheet10!A1+Sheet1!A1", Some(num(3))))
    )
    assertEquals(formulaText(rename(wb, Sheet1, Data), "Sheet2", ref"A1"), "Sheet10!A1+Data!A1")
  }

  test("array-formula records keep their kind; data-table records are never parsed") {
    val kind = FormulaKind.ArrayFormula(ref"C1:C2", ca = true)
    val wb = Workbook(
      Sheet("Sheet1").put(ref"A1", num(1)).put(ref"A2", num(2)),
      Sheet("Sheet2").put(ref"C1", CellValue.Formula("Sheet1!A1:A2*2", Some(num(2)), kind))
    )
    val renamed = rename(wb, Sheet1, Data)
    assertEquals(
      sheetNamed(renamed, "Sheet2")(ref"C1").value,
      CellValue.Formula("Data!A1:A2*2", Some(num(2)), kind)
    )
  }

  // ===== defined names, CF, DV =====

  test(
    "defined names follow the rename, including top-level comma unions; constants ride verbatim"
  ) {
    val wb = Workbook(
      Vector(Sheet("Sheet1"), Sheet("Other")),
      metadata = WorkbookMetadata(definedNames =
        Vector(
          DefinedName("Total", "Sheet1!$A$1"),
          DefinedName("Union", "Sheet1!$A$1:$B$2,Sheet1!$D$10,Other!$A$1"),
          DefinedName("Rate", "0.08"),
          DefinedName("Label", "\"Sheet1!A1, please\""),
          DefinedName("Elsewhere", "Other!$B$2", localSheetId = Some(1))
        )
      )
    )
    val renamed = rename(wb, Sheet1, Data)
    assertEquals(
      renamed.metadata.definedNames.map(_.formula),
      Vector(
        "Data!$A$1",
        "Data!$A$1:$B$2,Data!$D$10,Other!$A$1",
        "0.08",
        "\"Sheet1!A1, please\"",
        "Other!$B$2"
      )
    )
    assertEquals(
      renamed.metadata.definedNames.map(_.localSheetId),
      Vector(None, None, None, None, Some(1))
    )
  }

  test(
    "conditional-format formulas follow the rename (CellIs, Expression, Cfvo.Formula); the rest is identity"
  ) {
    val preservedRule = CfRule.Preserved("<cfRule type=\"iconSet\" priority=\"9\"/>", Some(9))
    val target = Sheet("Sheet2")
      .put(ref"A1", num(1))
      .conditionalFormat(ref"A1:A9", CfRule.expression("Sheet1!A1>0", red))
      .conditionalFormat(ref"B1:B9", CfRule.between("Sheet1!$B$1", "Sheet1!$B$2", red))
      .conditionalFormat(
        ref"C1:C9",
        CfRule.colorScale3(
          CfPoint(Cfvo.Formula("Sheet1!$C$1"), color),
          CfPoint(Cfvo.Percentile(BigDecimal(50)), color),
          CfPoint(Cfvo.Formula("MAX(Sheet1!$C$1:$C$9)"), color)
        )
      )
      .conditionalFormat(
        ref"D1:D9",
        CfRule.dataBar(color, min = Cfvo.Min, max = Cfvo.Formula("Sheet1!$D$1"))
      )
      .conditionalFormat(ref"E1:E9", CfRule.containsText("Sheet1", red))
      .conditionalFormat(ref"F1:F9", preservedRule)
    val wb = Workbook(Sheet("Sheet1").put(ref"A1", num(1)), target)
    val rules =
      sheetNamed(rename(wb, Sheet1, Data), "Sheet2").typedConditionalFormats.flatMap(_.rules)
    rules match
      case Vector(
            CfRule.Expression(expr, _, _, _),
            CfRule.CellIs(CfOperator.Between, f1, f2, _, _, _),
            CfRule.ColorScale(min, Some(mid), max, _),
            CfRule.DataBar(dbMin, dbMax, _, _, _),
            CfRule.Text(_, text, _, _, _),
            preserved: CfRule.Preserved
          ) =>
        assertEquals(expr, "Data!A1>0")
        assertEquals((f1, f2), ("Data!$B$1", Some("Data!$B$2")))
        assertEquals(min.cfvo, Cfvo.Formula("Data!$C$1"): Cfvo)
        assertEquals(mid.cfvo, Cfvo.Percentile(BigDecimal(50)): Cfvo)
        assertEquals(max.cfvo, Cfvo.Formula("MAX(Data!$C$1:$C$9)"): Cfvo)
        assertEquals(dbMin, Cfvo.Min: Cfvo)
        assertEquals(dbMax, Cfvo.Formula("Data!$D$1"): Cfvo)
        assertEquals(text, "Sheet1") // a text rule's payload is text, not a reference
        assertEquals(preserved, preservedRule)
      case other => fail(s"unexpected rules: $other")
  }

  test(
    "data-validation formulas follow the rename (List, Custom, Bounded); inline lists and Preserved are identity"
  ) {
    val preserved = DataValidation.Preserved("<dataValidation type=\"date\" sqref=\"Z1\"/>")
    val target = Sheet("Sheet2")
      .withDataValidation(ref"A1:A9", DataValidation.list("Sheet1!$A$1:$A$3"))
      .withDataValidation(ref"B1:B9", DataValidation.list("\"Sheet1,Sheet2\""))
      .withDataValidation(ref"C1:C9", DataValidation.custom("Sheet1!$A$1>0"))
      .withDataValidation(
        ref"D1:D9",
        DataValidation.whole(DvOperator.Between, "Sheet1!$A$1", Some("Sheet1!$A$2"))
      )
    val withPreserved = target.copy(dataValidations = target.dataValidations :+ preserved)
    val wb = Workbook(Sheet("Sheet1").put(ref"A1", num(1)), withPreserved)
    val dvs = sheetNamed(rename(wb, Sheet1, Data), "Sheet2").dataValidations
    val kinds = dvs.collect { case DataValidation.Rules(_, kind, _, _, _) => kind }
    assertEquals(
      kinds,
      Vector[DvKind](
        DvKind.List("Data!$A$1:$A$3"),
        DvKind.List("\"Sheet1,Sheet2\""),
        DvKind.Custom("Data!$A$1>0"),
        DvKind.Bounded(DvBoundedType.Whole, DvOperator.Between, "Data!$A$1", Some("Data!$A$2"))
      )
    )
    assertEquals(dvs.lastOption, Some(preserved: DataValidation))
  }

  // ===== refusal, Workbook.rename, references =====

  test("refuses before mutating when a text that mentions the sheet cannot be parsed") {
    val wb = Workbook(
      Sheet("Sheet1").put(ref"A1", num(1)),
      Sheet("Sheet2")
        .put(ref"A1", f("Sheet1!A1*2", Some(num(2))))
        .put(ref"B1", f("Sheet1!A1+", None))
    )
    SheetRenamer.rename(wb, Sheet1, Data) match
      case Left(XLError.FormulaError(formula, reason)) =>
        assertEquals(formula, "Sheet1!A1+")
        assert(reason.contains("Sheet2!B1"), reason)
        assert(reason.contains("Sheet1"), reason)
      case other => fail(s"expected a FormulaError refusal, got $other")
    // an unparsable text that does NOT mention the sheet is no obstacle
    val unrelated = Workbook(
      Sheet("Sheet1").put(ref"A1", num(1)),
      Sheet("Sheet2").put(ref"A1", f("Sheet1!A1*2", Some(num(2)))).put(ref"B1", f("A1+", None))
    )
    assertEquals(formulaText(rename(unrelated, Sheet1, Data), "Sheet2", ref"A1"), "Data!A1*2")
    assertEquals(
      sheetNamed(rename(unrelated, Sheet1, Data), "Sheet2")(ref"B1").value,
      f("A1+", None)
    )
  }

  test("refuses on an unparsable defined name that mentions the sheet") {
    val wb = Workbook(
      Vector(Sheet("Sheet1"), Sheet("Sheet2")),
      metadata = WorkbookMetadata(definedNames =
        Vector(DefinedName("Bad", "Sheet1!$A$1 Sheet1!$B$1:$B$9 +"))
      )
    )
    SheetRenamer.rename(wb, Sheet1, Data) match
      case Left(XLError.FormulaError(_, reason)) => assert(reason.contains("'Bad'"), reason)
      case other => fail(s"expected a FormulaError refusal, got $other")
  }

  test("the usual Workbook.rename refusals still apply") {
    assertEquals(
      SheetRenamer.rename(repro, SheetName.unsafe("Nope"), Data),
      Left(XLError.SheetNotFound("Nope")): XLResult[Workbook]
    )
    assertEquals(
      SheetRenamer.rename(repro, Sheet1, SheetName.unsafe("Sheet2")),
      Left(XLError.DuplicateSheet("Sheet2")): XLResult[Workbook]
    )
  }

  test("Workbook.rename's refusals win over a formula refusal (rename runs first, purely)") {
    val broken = repro.put(sheetNamed(repro, "Sheet2").put(ref"F1", f("Sheet1!A1+", None)))
    assertEquals(
      SheetRenamer.rename(broken, SheetName.unsafe("Nope"), Data),
      Left(XLError.SheetNotFound("Nope")): XLResult[Workbook]
    )
    assertEquals(
      SheetRenamer.rename(broken, Sheet1, SheetName.unsafe("Sheet2")),
      Left(XLError.DuplicateSheet("Sheet2")): XLResult[Workbook]
    )
    // and the formula refusal still fires when the tab rename itself is fine
    assert(SheetRenamer.rename(broken, Sheet1, Data).isLeft)
  }

  // ===== GH-222: the typed-chart remap Workbook.rename performs must survive =====

  private def chartSourcing(sheet: SheetName, name: SeriesName): Chart =
    Chart(
      ChartType.Bar(),
      Vector(
        Series(
          values = DataRef(sheet, ref"A1:A3"),
          categories = Some(DataRef(sheet, ref"B1:B3")),
          name = Some(name)
        )
      )
    )

  private def chartData: Sheet =
    Sheet("Sheet1")
      .put(ref"A1", num(1))
      .put(ref"A2", num(2))
      .put(ref"A3", num(3))
      .put(ref"B1", CellValue.Text("x"))
      .put(ref"B2", CellValue.Text("y"))
      .put(ref"B3", CellValue.Text("z"))
      .put(ref"C1", CellValue.Text("Units"))

  test(
    "a sheet holding both a formula and a typed chart sourcing the renamed sheet keeps both rewrites"
  ) {
    val dashboard = Sheet("Dashboard")
      .put(ref"A1", f("Sheet1!A1*2", Some(num(2))))
      .addChart(chartSourcing(Sheet1, SeriesName.FromCell(Sheet1, ref"C1")), ref"D2:K15")
    val renamed = rename(Workbook(chartData, dashboard), Sheet1, Data)
    assertEquals(formulaText(renamed, "Dashboard", ref"A1"), "Data!A1*2")
    val chart = sheetNamed(renamed, "Dashboard").charts.headOption
      .map(_.chart)
      .getOrElse(fail("the dashboard lost its chart"))
    chart.series match
      case Vector(series) =>
        assertEquals(series.values, DataRef(Data, ref"A1:A3"))
        assertEquals(series.categories, Some(DataRef(Data, ref"B1:B3")))
        assertEquals(series.name, Some(SeriesName.FromCell(Data, ref"C1")): Option[SeriesName])
      case other => fail(s"expected one series, got $other")
    // the same shape as Workbook.rename alone produces for the chart
    val tabOnly =
      Workbook(chartData, dashboard).rename(Sheet1, Data).fold(e => fail(e.message), identity)
    assertEquals(
      sheetNamed(renamed, "Dashboard").charts.map(_.chart),
      sheetNamed(tabOnly, "Dashboard").charts.map(_.chart)
    )
  }

  test("through a file: the chart part names the new sheet and the formula is rewritten") {
    val dir = Files.createTempDirectory("xl-renamer-chart")
    val in = dir.resolve("in-chart.xlsx")
    val out = dir.resolve("out-chart.xlsx")
    val chart = Chart
      .bar(
        Vector(
          Series(
            values = DataRef(Sheet1, ref"A1:A3"),
            categories = Some(DataRef(Sheet1, ref"B1:B3")),
            name = Some(SeriesName.Literal("Units"))
          )
        ),
        title = Some("Rename")
      )
      .fold(e => fail(e.message), identity)
    val dashboard = Sheet("Dashboard")
      .put(ref"A1", f("Sheet1!A1*2", Some(num(2))))
      .addChart(chart, ref"D2:K15")
    XlsxWriter.write(Workbook(chartData, dashboard), in).fold(e => fail(e.message), identity)
    val read = XlsxReader.read(in).fold(e => fail(e.message), identity)
    val renamed = rename(read, Sheet1, Data)
    XlsxWriter.write(renamed, out).fold(e => fail(e.message), identity)
    val chartParts = zipEntriesUnder(out, "xl/charts/chart")
    assert(chartParts.nonEmpty, s"no chart part written to $out")
    chartParts.foreach { (name, xml) =>
      assert(xml.contains("Data!$A$1:$A$3"), s"$name: $xml")
      assert(!xml.contains("Sheet1!"), s"$name still names Sheet1: $xml")
    }
    val reread = XlsxReader.read(out).fold(e => fail(e.message), identity)
    assertEquals(formulaText(reread, "Dashboard", ref"A1"), "Data!A1*2")
    assertEquals(reread.sheets.map(_.name.value), Vector("Data", "Dashboard"))
  }

  test("renaming a sheet to itself changes nothing but the tracker") {
    val same = rename(repro, Sheet1, Sheet1)
    assertEquals(same.sheets, repro.sheets)
  }

  test("Workbook.rename itself stays tab-only (formula-blind)") {
    val tabOnly = repro.rename(Sheet1, Data).fold(e => fail(e.message), identity)
    assertEquals(tabOnly.sheets.map(_.name.value), Vector("Data", "Sheet2"))
    assertEquals(formulaText(tabOnly, "Sheet2", ref"A1"), "Sheet1!A1*2")
  }

  test("references lists the cells whose text mentions the sheet, sheet order then row-major") {
    val wb = repro.put(
      Sheet("Zed")
        .put(ref"B2", f("Sheet1!A1"))
        .put(ref"A3", f("[2]Sheet1!A1"))
        .put(ref"A2", f("Sheet1!A1+1"))
    )
    assertEquals(
      SheetRenamer.references(wb, Sheet1),
      Vector(
        QualifiedRef(Sheet1, ref"B1"),
        QualifiedRef(SheetName.unsafe("Sheet2"), ref"A1"),
        QualifiedRef(SheetName.unsafe("Sheet2"), ref"B1"),
        QualifiedRef(SheetName.unsafe("Sheet2"), ref"E1"),
        QualifiedRef(SheetName.unsafe("Zed"), ref"A2"),
        QualifiedRef(SheetName.unsafe("Zed"), ref"B2")
      )
    )
    assertEquals(SheetRenamer.references(wb, SheetName.unsafe("Nope")), Vector.empty[QualifiedRef])
  }

  // ===== the AST law =====

  private val genA1: Gen[String] =
    for
      col <- Gen.choose(0, 30)
      row <- Gen.choose(0, 30)
    yield ARef.from0(col, row).toA1

  private val genRangeText: Gen[String] =
    for
      col <- Gen.choose(0, 20)
      row <- Gen.choose(0, 20)
      w <- Gen.choose(0, 4)
      h <- Gen.choose(0, 4)
    yield s"${ARef.from0(col, row).toA1}:${ARef.from0(col + w, row + h).toA1}"

  /** Formula texts over sheet `Alpha` (exact case, never `Beta`), locals, externals and names. */
  private def genFormulaText(depth: Int): Gen[String] =
    val leaf: Gen[String] = Gen.oneOf(
      genA1.map(a => s"Alpha!$a"),
      genA1.map(a => "'Alpha'!$" + a),
      genA1.map(a => s"Other!$a"),
      genA1,
      Gen.choose(1, 99).map(_.toString),
      genA1.map(a => s"[2]Alpha!$a"),
      Gen.const("Alpha!rate"),
      Gen.const("rate"),
      Gen.const("LEN(\"Alpha!A1\")"),
      genRangeText.map(r => s"SUM(Alpha!$r)"),
      genRangeText.map(r => s"SUM($r)"),
      genRangeText.map(r => s"SUM([3]Alpha!$r)"),
      genRangeText.map(r => s"SUMIF(Alpha!$r,\">0\")")
    )
    if depth >= 2 then leaf
    else
      Gen.frequency(
        2 -> leaf,
        1 -> Gen.lzy(for
          x <- genFormulaText(depth + 1)
          y <- genFormulaText(depth + 1)
          op <- Gen.oneOf("+", "-", "*", "/")
        yield s"($x$op$y)"),
        1 -> Gen.lzy(for
          c <- genFormulaText(depth + 1)
          a <- genFormulaText(depth + 1)
          b <- genFormulaText(depth + 1)
        yield s"IF($c>0,$a,$b)")
      )

  private val Alpha = SheetName.unsafe("Alpha")
  private val Beta = SheetName.unsafe("Beta")

  property(
    "renameSheet(renameSheet(e, a, b), b, a) == e, and the renamed AST no longer mentions a"
  ) {
    forAll(genFormulaText(0)) { text =>
      FormulaParser.parse(s"=$text") match
        case Left(err) => fail(s"generator produced an unparsable formula '$text': $err")
        case Right(expr) =>
          val renamed = FormulaShifter.renameSheet(expr, Alpha, Beta)
          assertEquals(FormulaShifter.renameSheet(renamed, Beta, Alpha), expr, text)
          assert(!FormulaShifter.mentionsSheet(renamed, "Alpha"), text)
          assertEquals(
            FormulaShifter.mentionsSheet(renamed, "Beta"),
            FormulaShifter.mentionsSheet(expr, "Alpha"),
            text
          )
          // the string form obeys the same law
          val printed = FormulaPrinter.printFileForm(renamed)
          assertEquals(
            FormulaParser
              .parse(s"=$printed")
              .map(FormulaShifter.renameSheet(_, Beta, Alpha))
              .map(FormulaPrinter.printFileForm),
            Right(FormulaPrinter.printFileForm(expr)): Either[ParseError, String],
            text
          )
    }
  }

  // ===== ADR-017 invariant 1, through a real file =====

  private def zipEntry(path: Path, entryName: String): Array[Byte] =
    val zip = new java.util.zip.ZipFile(path.toFile)
    try
      Option(zip.getEntry(entryName)) match
        case Some(entry) => zip.getInputStream(entry).readAllBytes()
        case None => fail(s"missing $entryName in $path")
    finally zip.close()

  /** Every zip entry whose name starts with `prefix`, as (name, UTF-8 text). */
  private def zipEntriesUnder(path: Path, prefix: String): Vector[(String, String)] =
    import scala.jdk.CollectionConverters.*
    val zip = new java.util.zip.ZipFile(path.toFile)
    try
      zip
        .entries()
        .asScala
        .filter(_.getName.startsWith(prefix))
        .map(e =>
          e.getName -> new String(zip.getInputStream(e).readAllBytes(), StandardCharsets.UTF_8)
        )
        .toVector
    finally zip.close()

  test(
    "SourceContext law: read file → rename → write → re-read rewrites other sheets; unmentioning sheets are byte-identical"
  ) {
    val dir = Files.createTempDirectory("xl-renamer")
    val in = dir.resolve("in.xlsx")
    val out = dir.resolve("out.xlsx")
    val wb = Workbook(
      Sheet("Sheet1").put(ref"A1", num(5)),
      Sheet("Sheet2")
        .put(ref"A1", f("Sheet1!A1*2", Some(num(10))))
        .put(ref"B1", f("SUM(Sheet1!A1:A1)", Some(num(5)))),
      Sheet("Notes")
        .put(ref"A1", CellValue.Text("no formulas here"))
        .put(ref"B1", f("A1&\"!\"", Some(CellValue.Text("no formulas here!"))))
    ).withDefinedName("Total", "Sheet1!$A$1")
    XlsxWriter.write(wb, in).fold(e => fail(e.message), identity)

    val read = XlsxReader.read(in).fold(e => fail(e.message), identity)
    assert(read.sourceContext.exists(_.isClean), "a freshly read workbook must be clean")
    val renamed = rename(read, Sheet1, Data)

    // Invariant 1: every changed sheet went back through Workbook.put, so the tracker names
    // exactly the renamed sheet and the sheet whose formulas were rewritten — not Notes. (The
    // tracker is the law's observable: a rename always modifies workbook.xml, and on a
    // modified-metadata write the surgical writer regenerates every worksheet part from the model
    // rather than copying any verbatim — XlsxWriter's rule, not the renamer's — so the on-disk
    // check below is byte-identity of a deterministic regeneration.)
    val tracker =
      renamed.sourceContext.map(_.modificationTracker).getOrElse(fail("source context lost"))
    assertEquals(tracker.modifiedSheets, Set(0, 1))
    assert(tracker.modifiedMetadata)

    XlsxWriter.write(renamed, out).fold(e => fail(e.message), identity)
    val reread = XlsxReader.read(out).fold(e => fail(e.message), identity)
    assertEquals(reread.sheets.map(_.name.value), Vector("Data", "Sheet2", "Notes"))
    assertEquals(formulaText(reread, "Sheet2", ref"A1"), "Data!A1*2")
    assertEquals(cachedOf(reread, "Sheet2", ref"A1"), Some(num(10)))
    assertEquals(formulaText(reread, "Sheet2", ref"B1"), "SUM(Data!A1:A1)")
    assertEquals(cachedOf(reread, "Sheet2", ref"B1"), Some(num(5)))
    assertEquals(reread.metadata.definedNames.map(_.formula), Vector("Data!$A$1"))

    // What the FILE says — the #559 symptom was a file that still read `Sheet1!A1`.
    val sheet2Xml = new String(zipEntry(out, "xl/worksheets/sheet2.xml"), StandardCharsets.UTF_8)
    assert(sheet2Xml.contains("Data!A1*2"), sheet2Xml)
    assert(!sheet2Xml.contains("Sheet1!"), sheet2Xml)

    // Notes never mentioned Sheet1: its part is byte-identical to the source part, and the model
    // read back from it is unchanged.
    val notesIn = zipEntry(in, "xl/worksheets/sheet3.xml")
    val notesOut = zipEntry(out, "xl/worksheets/sheet3.xml")
    assert(java.util.Arrays.equals(notesIn, notesOut), "Notes part differs from the source part")
    assertEquals(sheetNamed(reread, "Notes").cells, sheetNamed(read, "Notes").cells)
  }
