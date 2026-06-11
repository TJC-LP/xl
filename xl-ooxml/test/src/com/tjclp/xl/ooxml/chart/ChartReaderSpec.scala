package com.tjclp.xl.ooxml.chart

import java.util.zip.ZipFile

import munit.FunSuite

import com.tjclp.xl.addressing.{CellRange, SheetName}
import com.tjclp.xl.charts.*
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.TestFixtures

/**
 * GH-222 ChartReader: the openpyxl fixture parses to the EXACT typed value; everything outside the
 * strict whitelist returns None (the hosting anchor stays Preserved).
 */
class ChartReaderSpec extends FunSuite:

  private def fixtureChartXml(): String =
    val path = TestFixtures.copyToTemp("chart-bar.xlsx")
    val zip = new ZipFile(path.toFile)
    try
      val entry = Option(zip.getEntry("xl/charts/chart1.xml")).getOrElse(
        fail("chart1.xml missing from fixture")
      )
      new String(zip.getInputStream(entry).readAllBytes(), java.nio.charset.StandardCharsets.UTF_8)
    finally zip.close()

  private val chartData = SheetName.unsafe("ChartData")

  test("fixture chart1.xml parses to the exact typed Chart") {
    val expected = Chart(
      ChartType.Bar(BarDirection.Col, BarGrouping.Clustered),
      Vector(
        Series(
          values = DataRef(chartData, ref"B2:B5"),
          categories = Some(DataRef(chartData, ref"A2:A5")),
          name = Some(SeriesName.FromCell(chartData, ref"B1"))
        )
      ),
      title = Some("Synthetic Units by Quarter"),
      legend = Some(Legend(LegendPosition.Right, overlay = false))
    )
    assertEquals(ChartReader.parse(fixtureChartXml()), Some(expected))
  }

  // A minimal valid chart in the Excel `c:` dialect (the reader is prefix-agnostic; the fixture
  // covers the default-ns dialect). Mutated below for the rejection corpus.
  private val minimalBar: String =
    """<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      |<c:chart><c:plotArea>
      |<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>
      |<c:ser><c:idx val="0"/><c:order val="0"/>
      |<c:val><c:numRef><c:f>Data!$B$2:$B$5</c:f></c:numRef></c:val>
      |</c:ser>
      |<c:gapWidth val="150"/><c:axId val="10"/><c:axId val="100"/></c:barChart>
      |<c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="100"/></c:catAx>
      |<c:valAx><c:axId val="100"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:majorGridlines/><c:crossAx val="10"/></c:valAx>
      |</c:plotArea><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>""".stripMargin

  test("minimal Excel-dialect bar chart parses (prefix-agnostic)") {
    val expected = Chart(
      ChartType.Bar(BarDirection.Col, BarGrouping.Clustered),
      Vector(Series(DataRef(SheetName.unsafe("Data"), ref"B2:B5"), None, None)),
      title = None,
      legend = None
    )
    assertEquals(ChartReader.parse(minimalBar), Some(expected))
  }

  test("stacked bar requires overlap=100; clustered requires its absence (visible state)") {
    val stackedNoOverlap =
      minimalBar.replace("""<c:grouping val="clustered"/>""", """<c:grouping val="stacked"/>""")
    assertEquals(ChartReader.parse(stackedNoOverlap), None)
    val stackedWithOverlap = minimalBar
      .replace("""<c:grouping val="clustered"/>""", """<c:grouping val="stacked"/>""")
      .replace("""<c:gapWidth val="150"/>""", """<c:gapWidth val="150"/><c:overlap val="100"/>""")
    assert(
      ChartReader
        .parse(stackedWithOverlap)
        .exists(_.chartType == ChartType.Bar(BarDirection.Col, BarGrouping.Stacked))
    )
    val clusteredWithOverlap =
      minimalBar.replace(
        """<c:gapWidth val="150"/>""",
        """<c:gapWidth val="150"/><c:overlap val="100"/>"""
      )
    assertEquals(ChartReader.parse(clusteredWithOverlap), None)
  }

  test("rejection corpus: every out-of-fence shape yields None") {
    val cases: Map[String, String] = Map(
      "scatter group" -> minimalBar
        .replace("c:barChart>", "c:scatterChart>")
        .replace("""<c:barDir val="col"/><c:grouping val="clustered"/>""", ""),
      "c:style" -> minimalBar.replace("<c:chart>", """<c:style val="2"/><c:chart>"""),
      "externalData" -> minimalBar.replace(
        "</c:chart>",
        """</c:chart><c:externalData r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>"""
      ),
      "dLbls" -> minimalBar.replace(
        "<c:ser><c:idx",
        "<c:ser><c:dLbls><c:showVal val=\"1\"/></c:dLbls><c:idx"
      ),
      "axis delete=1" -> minimalBar.replace(
        """<c:axPos val="b"/>""",
        """<c:axPos val="b"/><c:delete val="1"/>"""
      ),
      "orientation maxMin" -> minimalBar.replace(
        """<c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling>""",
        """<c:catAx><c:axId val="10"/><c:scaling><c:orientation val="maxMin"/></c:scaling>"""
      ),
      "gapWidth != 150" -> minimalBar.replace(
        """<c:gapWidth val="150"/>""",
        """<c:gapWidth val="80"/>"""
      ),
      "2D val range" -> minimalBar.replace("Data!$B$2:$B$5", "Data!$A$2:$B$5"),
      "unqualified val range" -> minimalBar.replace("Data!$B$2:$B$5", "$B$2:$B$5"),
      "multiLvlStrRef cat" -> minimalBar.replace(
        "<c:val>",
        "<c:cat><c:multiLvlStrRef><c:f>Data!$A$2:$A$5</c:f></c:multiLvlStrRef></c:cat><c:val>"
      ),
      "autoTitleDeleted=0 without title" -> minimalBar.replace(
        "<c:chart>",
        """<c:chart><c:autoTitleDeleted val="0"/>"""
      ),
      "mc:AlternateContent" -> minimalBar.replace(
        "<c:chart>",
        """<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice/></mc:AlternateContent><c:chart>"""
      ),
      "ser spPr with fill" -> minimalBar.replace(
        "<c:val>",
        """<c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></c:spPr><c:val>"""
      ),
      "title without tx" -> minimalBar.replace(
        "<c:chart>",
        """<c:chart><c:title><c:overlay val="0"/></c:title>"""
      ),
      "missing axes" -> minimalBar
        .replace(
          """<c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="100"/></c:catAx>""",
          ""
        ),
      "inconsistent axis wiring" -> minimalBar.replace(
        """<c:crossAx val="100"/></c:catAx>""",
        """<c:crossAx val="999"/></c:catAx>"""
      ),
      "pie with axes" -> minimalBar
        .replace("c:barChart>", "c:pieChart>")
        .replace(
          """<c:barDir val="col"/><c:grouping val="clustered"/>""",
          """<c:varyColors val="1"/>"""
        )
        .replace("""<c:gapWidth val="150"/><c:axId val="10"/><c:axId val="100"/>""", ""),
      "malformed XML" -> "<c:chartSpace><unclosed",
      "smooth=1 line ser" -> minimalBar
        .replace("c:barChart>", "c:lineChart>")
        .replace(
          """<c:barDir val="col"/><c:grouping val="clustered"/>""",
          """<c:grouping val="standard"/>"""
        )
        .replace("""<c:gapWidth val="150"/>""", "")
        .replace("</c:ser>", """<c:smooth val="1"/></c:ser>""")
    )
    cases.foreach { case (label, xml) =>
      assertEquals(ChartReader.parse(xml), None, s"case '$label' must reject:\n$xml")
    }
  }

  test("line chart with markers=none and smooth=0 parses; tx variants parse") {
    val line =
      """<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        |<c:chart><c:plotArea>
        |<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>
        |<c:ser><c:idx val="0"/><c:order val="0"/>
        |<c:tx><c:v>Alpha</c:v></c:tx>
        |<c:marker><c:symbol val="none"/></c:marker>
        |<c:val><c:numRef><c:f>Data!$B$2:$B$5</c:f></c:numRef></c:val>
        |<c:smooth val="0"/>
        |</c:ser>
        |<c:ser><c:idx val="1"/><c:order val="1"/>
        |<c:tx><c:strRef><c:f>Data!$C$1</c:f></c:strRef></c:tx>
        |<c:cat><c:strRef><c:f>Data!$A$2:$A$5</c:f><c:strCache><c:ptCount val="4"/><c:pt idx="0"><c:v>Q1</c:v></c:pt></c:strCache></c:strRef></c:cat>
        |<c:val><c:numRef><c:f>Data!$C$2:$C$5</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="4"/></c:numCache></c:numRef></c:val>
        |</c:ser>
        |<c:marker val="1"/><c:axId val="10"/><c:axId val="100"/></c:lineChart>
        |<c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:crossAx val="100"/></c:catAx>
        |<c:valAx><c:axId val="100"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:crossAx val="10"/></c:valAx>
        |</c:plotArea><c:legend><c:legendPos val="tr"/><c:overlay val="1"/></c:legend>
        |<c:plotVisOnly val="1"/></c:chart></c:chartSpace>""".stripMargin
    val data = SheetName.unsafe("Data")
    val expected = Chart(
      ChartType.Line,
      Vector(
        Series(DataRef(data, ref"B2:B5"), None, Some(SeriesName.Literal("Alpha"))),
        Series(
          DataRef(data, ref"C2:C5"),
          Some(DataRef(data, ref"A2:A5")),
          Some(SeriesName.FromCell(data, ref"C1"))
        )
      ),
      title = None,
      legend = Some(Legend(LegendPosition.TopRight, overlay = true))
    )
    assertEquals(ChartReader.parse(line), Some(expected))
  }

  test("pie chart with exactly one series parses; quoted sheet names resolve") {
    val pie =
      """<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
        |<c:chart><c:plotArea>
        |<c:pieChart><c:varyColors val="1"/>
        |<c:ser><c:idx val="0"/><c:order val="0"/>
        |<c:cat><c:numRef><c:f>'Q1 ''Report'!$A$2:$A$5</c:f></c:numRef></c:cat>
        |<c:val><c:numRef><c:f>'Q1 ''Report'!$B$2:$B$5</c:f></c:numRef></c:val>
        |</c:ser><c:firstSliceAng val="0"/></c:pieChart>
        |</c:plotArea><c:plotVisOnly val="1"/></c:chart></c:chartSpace>""".stripMargin
    val q1 = SheetName.unsafe("Q1 'Report")
    val expected = Chart(
      ChartType.Pie,
      Vector(Series(DataRef(q1, ref"B2:B5"), Some(DataRef(q1, ref"A2:A5")), None)),
      title = None,
      legend = None
    )
    assertEquals(ChartReader.parse(pie), Some(expected))
  }
