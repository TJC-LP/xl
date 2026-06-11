package com.tjclp.xl.ooxml.chart

import java.nio.file.{Files, Path}
import java.util.zip.ZipFile

import munit.FunSuite

import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.charts.*
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.{XlsxWriter, XmlUtil}
import com.tjclp.xl.ooxml.writer.WriterConfig
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook

/**
 * GH-222 OoxmlChart emitter: golden emission for all four group shapes, cache resolution rules,
 * GH-263 quoting, backend parity (the house worksheet-parity contract: semantic agreement +
 * canonical sorted declarations + per-backend write-twice stability).
 */
class OoxmlChartSpec extends FunSuite:

  private val data = SheetName.unsafe("Data")

  private val sheet = Sheet(data)
    .put(Cell(ref"A1", CellValue.Text("Hdr")))
    .put(Cell(ref"A2", CellValue.Text("Q1")))
    .put(Cell(ref"A3", CellValue.Text("Q2")))
    .put(Cell(ref"B1", CellValue.Text("Units")))
    .put(Cell(ref"B2", CellValue.Number(BigDecimal(12))))
    .put(Cell(ref"B3", CellValue.Number(BigDecimal("0.5"))))
  private val wb = Workbook(Vector(sheet))

  private val series = Series(
    DataRef(data, ref"B2:B3"),
    Some(DataRef(data, ref"A2:A3")),
    Some(SeriesName.FromCell(data, ref"B1"))
  )

  private def emit(chart: Chart, workbook: Workbook = wb): String =
    XmlUtil
      .compact(OoxmlChart(chart, ChartCaches.resolve(workbook, chart)).toXml)
      .linesIterator
      .drop(1) // XML declaration
      .mkString

  private val serGolden =
    """<ser><idx val="0"/><order val="0"/><tx><strRef><f>Data!$B$1</f><strCache><ptCount val="1"/><pt idx="0"><v>Units</v></pt></strCache></strRef></tx><cat><strRef><f>Data!$A$2:$A$3</f><strCache><ptCount val="2"/><pt idx="0"><v>Q1</v></pt><pt idx="1"><v>Q2</v></pt></strCache></strRef></cat><val><numRef><f>Data!$B$2:$B$3</f><numCache><formatCode>General</formatCode><ptCount val="2"/><pt idx="0"><v>12</v></pt><pt idx="1"><v>0.5</v></pt></numCache></numRef></val></ser>"""

  private def axesGolden(catPos: String, valPos: String) =
    s"""<catAx><axId val="10"/><scaling><orientation val="minMax"/></scaling><delete val="0"/><axPos val="$catPos"/><majorTickMark val="out"/><minorTickMark val="none"/><tickLblPos val="nextTo"/><crossAx val="100"/><crosses val="autoZero"/><auto val="1"/><lblAlgn val="ctr"/><lblOffset val="100"/><noMultiLvlLbl val="0"/></catAx><valAx><axId val="100"/><scaling><orientation val="minMax"/></scaling><delete val="0"/><axPos val="$valPos"/><majorGridlines/><numFmt formatCode="General" sourceLinked="1"/><majorTickMark val="out"/><minorTickMark val="none"/><tickLblPos val="nextTo"/><crossAx val="10"/><crosses val="autoZero"/><crossBetween val="between"/></valAx>"""

  private val prefix =
    """<chartSpace xmlns="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><chart>"""
  private val suffix = """<plotVisOnly val="1"/><dispBlanksAs val="gap"/></chart></chartSpace>"""

  test("golden: clustered column bar with title and legend") {
    val chart = Chart(ChartType.Bar(), Vector(series), Some("T"), Some(Legend()))
    val expected = prefix +
      """<title><tx><rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:t>T</a:t></a:r></a:p></rich></tx><overlay val="0"/></title><autoTitleDeleted val="0"/>""" +
      s"""<plotArea><barChart><barDir val="col"/><grouping val="clustered"/><varyColors val="0"/>$serGolden<gapWidth val="150"/><axId val="10"/><axId val="100"/></barChart>${axesGolden(
          "b",
          "l"
        )}</plotArea>""" +
      """<legend><legendPos val="r"/><overlay val="0"/></legend>""" + suffix
    assertEquals(emit(chart), expected)
  }

  test("golden: stacked horizontal bar emits overlap=100 and swapped axis positions") {
    val chart = Chart(
      ChartType.Bar(BarDirection.Bar, BarGrouping.Stacked),
      Vector(series),
      None,
      None
    )
    val expected = prefix + """<autoTitleDeleted val="1"/>""" +
      s"""<plotArea><barChart><barDir val="bar"/><grouping val="stacked"/><varyColors val="0"/>$serGolden<gapWidth val="150"/><overlap val="100"/><axId val="10"/><axId val="100"/></barChart>${axesGolden(
          "l",
          "b"
        )}</plotArea>""" +
      suffix
    assertEquals(emit(chart), expected)
  }

  test("golden: line chart — per-ser marker BEFORE cat, smooth AFTER val, chart marker=1") {
    val chart = Chart(
      ChartType.Line,
      Vector(series),
      None,
      Some(Legend(LegendPosition.Bottom, overlay = true))
    )
    val lineSer = serGolden
      .replace("</tx><cat>", "</tx><marker><symbol val=\"none\"/></marker><cat>")
      .replace("</val></ser>", "</val><smooth val=\"0\"/></ser>")
    val expected = prefix + """<autoTitleDeleted val="1"/>""" +
      s"""<plotArea><lineChart><grouping val="standard"/><varyColors val="0"/>$lineSer<marker val="1"/><axId val="10"/><axId val="100"/></lineChart>${axesGolden(
          "b",
          "l"
        )}</plotArea>""" +
      """<legend><legendPos val="b"/><overlay val="1"/></legend>""" + suffix
    assertEquals(emit(chart), expected)
  }

  test("golden: pie chart — no axes, varyColors=1, firstSliceAng=0") {
    val chart = Chart(ChartType.Pie, Vector(series), None, Some(Legend()))
    val expected = prefix + """<autoTitleDeleted val="1"/>""" +
      s"""<plotArea><pieChart><varyColors val="1"/>$serGolden<firstSliceAng val="0"/></pieChart></plotArea>""" +
      """<legend><legendPos val="r"/><overlay val="0"/></legend>""" + suffix
    assertEquals(emit(chart), expected)
  }

  // ===== caches =====

  test("cache: referenced sheet absent => cache elements OMITTED entirely (bare c:f)") {
    val chart = Chart(ChartType.Bar(), Vector(series), None, None)
    val xml = emit(chart, Workbook(Vector(Sheet(SheetName.unsafe("Other")))))
    assert(xml.contains("<tx><strRef><f>Data!$B$1</f></strRef></tx>"), xml)
    assert(xml.contains("<cat><strRef><f>Data!$A$2:$A$3</f></strRef></cat>"), xml)
    assert(xml.contains("<val><numRef><f>Data!$B$2:$B$3</f></numRef></val>"), xml)
    assert(!xml.contains("Cache"), "no cache elements expected")
  }

  test("cache: numeric categories canonicalize to numRef; blanks skip pts but count in ptCount") {
    val numSheet = Sheet(data)
      .put(Cell(ref"A2", CellValue.Number(BigDecimal(2001))))
      // A3 blank
      .put(Cell(ref"A4", CellValue.Number(BigDecimal("2002.5"))))
      .put(Cell(ref"B2", CellValue.Bool(true)))
      .put(Cell(ref"B3", CellValue.Formula("1+1", Some(CellValue.Number(BigDecimal(2))))))
      .put(Cell(ref"B4", CellValue.Text("not numeric")))
    val chart = Chart(
      ChartType.Bar(),
      Vector(Series(DataRef(data, ref"B2:B4"), Some(DataRef(data, ref"A2:A4")), None)),
      None,
      None
    )
    val xml = emit(chart, Workbook(Vector(numSheet)))
    // cat: all non-blank cells numeric -> numRef + numCache; blank A3 skips its pt
    assert(
      xml.contains(
        """<cat><numRef><f>Data!$A$2:$A$4</f><numCache><formatCode>General</formatCode><ptCount val="3"/><pt idx="0"><v>2001</v></pt><pt idx="2"><v>2002.5</v></pt></numCache></numRef></cat>"""
      ),
      xml
    )
    // val numCache: Bool -> 1/0, cached formula -> its Number, text skips
    assert(
      xml.contains(
        """<val><numRef><f>Data!$B$2:$B$4</f><numCache><formatCode>General</formatCode><ptCount val="3"/><pt idx="0"><v>1</v></pt><pt idx="1"><v>2</v></pt></numCache></numRef></val>"""
      ),
      xml
    )
  }

  test("cache: strCache renders Number via plainNumber, Bool as TRUE/FALSE, formula via cache") {
    val mixSheet = Sheet(data)
      .put(Cell(ref"A2", CellValue.Text("Q1")))
      .put(Cell(ref"A3", CellValue.Number(BigDecimal("1E2")))) // plainNumber => 100
      .put(Cell(ref"A4", CellValue.Bool(false)))
      .put(Cell(ref"A5", CellValue.Formula("X", Some(CellValue.Text("calc")))))
      .put(Cell(ref"B2", CellValue.Number(BigDecimal(1))))
    val chart = Chart(
      ChartType.Bar(),
      Vector(Series(DataRef(data, ref"B2:B5"), Some(DataRef(data, ref"A2:A5")), None)),
      None,
      None
    )
    val xml = emit(chart, Workbook(Vector(mixSheet)))
    assert(
      xml.contains(
        """<strCache><ptCount val="4"/><pt idx="0"><v>Q1</v></pt><pt idx="1"><v>100</v></pt><pt idx="2"><v>FALSE</v></pt><pt idx="3"><v>calc</v></pt></strCache>"""
      ),
      xml
    )
  }

  // ===== quoting (GH-263) =====

  test("quoting goldens: quoted-apostrophe and cell-shaped sheet names") {
    val q1r = SheetName.unsafe("Q1 'Report")
    val q1 = SheetName.unsafe("Q1")
    val chart = Chart(
      ChartType.Bar(),
      Vector(
        Series(
          DataRef(q1r, ref"A1:A5"),
          Some(DataRef(q1, ref"B1:B5")),
          Some(SeriesName.FromCell(q1r, ref"C1"))
        )
      ),
      None,
      None
    )
    val xml = emit(chart, Workbook(Vector(Sheet(SheetName.unsafe("Other")))))
    assert(xml.contains("<f>'Q1 ''Report'!$A$1:$A$5</f>"), xml)
    assert(xml.contains("<f>'Q1'!$B$1:$B$5</f>"), xml)
    assert(xml.contains("<f>'Q1 ''Report'!$C$1</f>"), xml)
  }

  // ===== backend parity + write-twice (house worksheet-parity contract) =====

  private def write(workbook: Workbook, label: String, config: WriterConfig): Path =
    val out = Files.createTempFile(s"xl-chart-$label-", ".xlsx")
    out.toFile.deleteOnExit()
    XlsxWriter
      .writeWith(workbook, out, config)
      .fold(e => fail(s"write failed: ${e.message}"), _ => ())
    out

  private def entryText(path: Path, name: String): String =
    val zip = new ZipFile(path.toFile)
    try
      Option(zip.getEntry(name)) match
        case Some(e) =>
          new String(zip.getInputStream(e).readAllBytes(), java.nio.charset.StandardCharsets.UTF_8)
        case None => fail(s"zip entry $name not found in $path")
    finally zip.close()

  test("ScalaXml and SaxStax agree on the chart part (semantic parity + canonical decls)") {
    val chart = Chart(ChartType.Bar(), Vector(series), Some("T"), Some(Legend()))
    val workbook = Workbook(Vector(sheet.addChart(chart, ref"D2:K15")))
    val outDom = write(workbook, "par-dom", WriterConfig.scalaXml)
    val outSax = write(workbook, "par-sax", WriterConfig.saxStax)

    val domXml = entryText(outDom, "xl/charts/chart1.xml")
    val saxXml = entryText(outSax, "xl/charts/chart1.xml")
    assertEquals(
      ChartReader.parse(domXml),
      ChartReader.parse(saxXml),
      "backends must agree on the parsed chart model"
    )
    assertEquals(ChartReader.parse(domXml), Some(chart))

    def decls(xml: String): List[String] =
      """xmlns(:[a-z0-9]+)?="[^"]*"""".r.findAllIn(xml).toList
    assertEquals(decls(domXml), decls(saxXml), "namespace declarations must match in order")
    assertEquals(
      decls(domXml).map(_.takeWhile(_ != '=')),
      decls(domXml).map(_.takeWhile(_ != '=')).sorted,
      "declarations print in canonical sorted order"
    )

    // each backend individually write-twice byte-identical
    val outDom2 = write(workbook, "par-dom2", WriterConfig.scalaXml)
    val outSax2 = write(workbook, "par-sax2", WriterConfig.saxStax)
    assertEquals(domXml, entryText(outDom2, "xl/charts/chart1.xml"))
    assertEquals(saxXml, entryText(outSax2, "xl/charts/chart1.xml"))
  }
