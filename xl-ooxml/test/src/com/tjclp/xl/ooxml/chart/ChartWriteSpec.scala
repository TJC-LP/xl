package com.tjclp.xl.ooxml.chart

import java.io.StringReader
import java.nio.file.{Files, Path}
import java.util.zip.ZipFile
import javax.xml.parsers.DocumentBuilderFactory

import munit.FunSuite

import com.tjclp.xl.api.*
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.drawings.TestImages
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.{TestFixtures, XlsxReader, XlsxWriter}
import com.tjclp.xl.ooxml.writer.WriterConfig
import com.tjclp.xl.styles.units.Emu

/**
 * GH-222 chart write planning: equality-match part reuse, same-path regeneration for edited charts,
 * fresh-part allocation on the degenerate reorder+edit, fresh-workbook authoring with the full
 * rels/ContentTypes chain, and mixed chart+picture parts.
 */
class ChartWriteSpec extends FunSuite:

  private val png = ImageData(TestImages.png2x3, ImageFormat.Png)
  private val extent = Extent(Emu(95250L), Emu(190500L))

  private def readFixture(name: String): (Path, Workbook) =
    val path = TestFixtures.copyToTemp(name)
    val wb = XlsxReader
      .read(path)
      .fold(err => fail(s"$name failed to read: ${err.message}"), identity)
    (path, wb)

  private def write(wb: Workbook, label: String): Path =
    val out = Files.createTempFile(s"xl-chartwrite-$label-", ".xlsx")
    out.toFile.deleteOnExit()
    XlsxWriter
      .writeWith(wb, out, WriterConfig())
      .fold(err => fail(s"$label write failed: ${err.message}"), _ => ())
    out

  private def reread(path: Path): Workbook =
    XlsxReader.read(path).fold(err => fail(s"re-read failed: ${err.message}"), identity)

  private def entryBytes(path: Path, name: String): Array[Byte] =
    val zip = new ZipFile(path.toFile)
    try
      Option(zip.getEntry(name)) match
        case Some(e) => zip.getInputStream(e).readAllBytes()
        case None => fail(s"zip entry $name not found in $path")
    finally zip.close()

  private def entryText(path: Path, name: String): String =
    new String(entryBytes(path, name), java.nio.charset.StandardCharsets.UTF_8)

  private def entryNames(path: Path): Set[String] =
    val zip = new ZipFile(path.toFile)
    try
      scala.jdk.CollectionConverters
        .EnumerationHasAsScala(zip.entries())
        .asScala
        .map(_.getName)
        .toSet
    finally zip.close()

  private def fixtureChart(wb: Workbook): Drawing.ChartFrame =
    wb.sheets(0).charts.headOption.getOrElse(fail("fixture chart missing"))

  /**
   * Replace the first sheet's drawings through the tracked update path (the real caller shape — an
   * untracked direct copy on a CLEAN source-backed workbook takes the whole-file verbatim-copy fast
   * path and never reaches the drawing planner).
   */
  private def withDrawings(wb: Workbook, drawings: Vector[Drawing]): Workbook =
    wb.update(wb.sheets(0).name, _.copy(drawings = drawings))
      .fold(e => fail(s"update failed: $e"), identity)

  test("title edit regenerates the chart at the SAME path with the same relId; wiring intact") {
    val (in, wb) = readFixture("chart-bar.xlsx")
    val frame = fixtureChart(wb)
    val edited = frame.copy(chart = frame.chart.copy(title = Some("Edited Title")))
    val out = write(withDrawings(wb, Vector(edited)), "title-edit")

    // same path, regenerated content
    val chartXml = entryText(out, "xl/charts/chart1.xml")
    assert(chartXml.contains("Edited Title"), chartXml)
    assert(!entryNames(out).contains("xl/charts/chart2.xml"), "no part churn for a simple edit")
    // drawing rels keep the source relationship VERBATIM (re-serialized, no appends): the reused
    // rId1 -> chart1.xml resolves by construction and no new chart rel was allocated
    val rels = entryText(out, "xl/drawings/_rels/drawing1.xml.rels")
    assert(rels.contains("""Id="rId1""""), rels)
    assert(rels.contains("chart1.xml"), rels)
    assertEquals("rId".r.findAllIn(rels).size, 1, s"exactly one relationship expected: $rels")
    // worksheet wiring unchanged: the drawing ref and its rel still resolve
    assert(entryText(out, "xl/worksheets/sheet1.xml").contains("<drawing"))
    assert(entryText(out, "xl/worksheets/_rels/sheet1.xml.rels").contains("drawing1.xml"))
    // ContentTypes still carries the chart override (preserved CT already had it)
    assert(entryText(out, "[Content_Types].xml").contains("/xl/charts/chart1.xml"))

    val readBack = reread(out)
    assertEquals(readBack.sheets(0).charts.size, 1)
    assertEquals(readBack.sheets(0).charts(0).chart, edited.chart)
  }

  test("data edit through Sheet.addChart-style update re-reads with regenerated caches") {
    val (_, wb) = readFixture("chart-bar.xlsx")
    val frame = fixtureChart(wb)
    val chartData = SheetName.unsafe("ChartData")
    val widened = frame.chart.copy(series =
      frame.chart.series.map(s => s.copy(values = DataRef(chartData, ref"B2:B4")))
    )
    val out = write(withDrawings(wb, Vector(frame.copy(chart = widened))), "data-edit")
    val chartXml = entryText(out, "xl/charts/chart1.xml")
    // caches regenerated from STORED values: B2=12, B3=19, B4=7 in the fixture
    assert(chartXml.contains("""<f>ChartData!$B$2:$B$4</f>"""), chartXml)
    assert(chartXml.contains("""<pt idx="0"><v>12</v></pt>"""), chartXml)
    assertEquals(reread(out).sheets(0).charts(0).chart, widened)
  }

  test("reorder+edit degenerate allocates a fresh part + appended rel; old part stays on disk") {
    val (_, wb) = readFixture("chart-bar.xlsx")
    val frame = fixtureChart(wb)
    val edited = frame.copy(chart = frame.chart.copy(title = Some("Moved And Edited")))
    val picture = Drawing.Picture(DrawingAnchor.at(ref"A20", extent), png)
    // chart now at anchorIdx 1 (was 0) AND edited: no equality match, no same-anchorIdx snapshot
    val out = write(withDrawings(wb, Vector(picture, edited)), "reorder-edit")

    val names = entryNames(out)
    assert(names.contains("xl/charts/chart2.xml"), s"fresh chart part expected: $names")
    assert(names.contains("xl/charts/chart1.xml"), "old part is never deleted (media policy)")
    assert(entryText(out, "xl/charts/chart2.xml").contains("Moved And Edited"))
    val rels = entryText(out, "xl/drawings/_rels/drawing1.xml.rels")
    assert(rels.contains("chart1.xml") && rels.contains("chart2.xml"), rels)
    assert(entryText(out, "[Content_Types].xml").contains("/xl/charts/chart2.xml"))

    val readBack = reread(out)
    val drawings = readBack.sheets(0).drawings
    assertEquals(drawings.size, 2)
    drawings(1) match
      case Drawing.ChartFrame(_, chart, _) => assertEquals(chart, edited.chart)
      case other => fail(s"expected ChartFrame second, got $other")
  }

  test("pure reorder without edit reuses the source part byte-identically (equality match)") {
    val (in, wb) = readFixture("chart-bar.xlsx")
    val frame = fixtureChart(wb)
    val picture = Drawing.Picture(DrawingAnchor.at(ref"A20", extent), png)
    val out = write(withDrawings(wb, Vector(picture, frame)), "pure-reorder")
    assertEquals(
      entryBytes(out, "xl/charts/chart1.xml").toSeq,
      entryBytes(in, "xl/charts/chart1.xml").toSeq,
      "unchanged chart must ride preservation even after reorder"
    )
    assert(!entryNames(out).contains("xl/charts/chart2.xml"))
  }

  // ===== fresh-workbook authoring =====

  private val nsRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

  private def parseNamespaceAware(xml: String, label: String): org.w3c.dom.Document =
    val factory = DocumentBuilderFactory.newInstance()
    factory.setNamespaceAware(true)
    try factory.newDocumentBuilder().parse(new org.xml.sax.InputSource(new StringReader(xml)))
    catch
      case e: org.xml.sax.SAXException =>
        fail(s"$label is not namespace-well-formed: ${e.getMessage}")

  /** The rels chain: every r:id in the drawing resolves, every target exists, CT registered. */
  private def assertChartStructure(path: Path, drawingPart: String): Unit =
    val names = entryNames(path)
    val doc = parseNamespaceAware(entryText(path, drawingPart), drawingPart)
    val all = doc.getElementsByTagNameNS("*", "*")
    val referenced = (0 until all.getLength).flatMap { i =>
      Option(all.item(i))
        .collect { case e: org.w3c.dom.Element =>
          List(
            Option(e.getAttributeNS(nsRel, "embed")).filter(_.nonEmpty),
            Option(e.getAttributeNS(nsRel, "id")).filter(_.nonEmpty)
          ).flatten
        }
        .toList
        .flatten
    }.toSet
    val relsName =
      val slash = drawingPart.lastIndexOf('/')
      val (dir, name) = drawingPart.splitAt(slash + 1)
      s"${dir}_rels/$name.rels"
    val relsDoc = parseNamespaceAware(entryText(path, relsName), relsName)
    val rels = relsDoc.getElementsByTagNameNS("*", "Relationship")
    val relTargets = (0 until rels.getLength).flatMap { i =>
      Option(rels.item(i)).collect { case e: org.w3c.dom.Element =>
        e.getAttribute("Id") -> e.getAttribute("Target")
      }
    }.toMap
    val unresolved = referenced.filterNot(relTargets.contains)
    assert(unresolved.isEmpty, s"$drawingPart references unresolved rIds: $unresolved")
    relTargets.values.foreach { target =>
      val resolved =
        if target.startsWith("/") then target.drop(1)
        else
          java.nio.file.Paths
            .get(drawingPart)
            .getParent
            .resolve(target)
            .normalize()
            .toString
            .replace('\\', '/')
      assert(names.contains(resolved), s"rel target $target -> $resolved missing from zip")
    }
    val ct = entryText(path, "[Content_Types].xml")
    names.filter(_.startsWith("xl/charts/chart")).foreach { chartPart =>
      assert(ct.contains(s"""PartName="/$chartPart""""), s"no ContentTypes override for $chartPart")
      assert(
        ct.contains("application/vnd.openxmlformats-officedocument.drawingml.chart+xml"),
        "chart content type missing"
      )
    }
    // the drawing root must bind xmlns:c (TRAP-6) when a chart ships
    if names.exists(_.startsWith("xl/charts/")) then
      assert(
        entryText(path, drawingPart)
          .contains("""xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart""""),
        "wsDr must bind xmlns:c"
      )

  test("fresh-workbook authoring wires the full chain: part, rels, override, xmlns:c, rIdDr1") {
    val data = SheetName.unsafe("Data")
    val sheet = Sheet(data)
      .put(Cell(ref"A2", CellValue.Text("Q1")))
      .put(Cell(ref"A3", CellValue.Text("Q2")))
      .put(Cell(ref"B2", CellValue.Number(BigDecimal(3))))
      .put(Cell(ref"B3", CellValue.Number(BigDecimal(4))))
    val chart = Chart
      .bar(
        Vector(Series(DataRef(data, ref"B2:B3"), Some(DataRef(data, ref"A2:A3")), None)),
        title = Some("Fresh")
      )
      .fold(e => fail(s"validated failed: $e"), identity)
    val out = write(Workbook(Vector(sheet.addChart(chart, ref"D2:K15"))), "fresh")

    assertChartStructure(out, "xl/drawings/drawing1.xml")
    assert(entryNames(out).contains("xl/charts/chart1.xml"))
    // worksheet carries the drawing ref; sheet rels carry rIdDr1 -> drawing1.xml
    assert(entryText(out, "xl/worksheets/sheet1.xml").contains("<drawing"))
    assert(entryText(out, "xl/worksheets/_rels/sheet1.xml.rels").contains("drawing1.xml"))

    val readBack = reread(out)
    assertEquals(readBack.sheets(0).charts.size, 1)
    assertEquals(readBack.sheets(0).charts(0).chart, chart)
    assertEquals(readBack.sheets(0).charts(0).name, "Chart 1") // writer default
  }

  test("mixed chart+picture part round-trips with unique cNvPr ids and separate ordinals") {
    val data = SheetName.unsafe("Mix")
    val sheet = Sheet(data)
      .put(Cell(ref"B2", CellValue.Number(BigDecimal(5))))
      .addImage(png, ref"A1", extent)
      .addChart(
        Chart(
          ChartType.Pie,
          Vector(Series(DataRef(data, ref"B2:B4"), None, Some(SeriesName.Literal("Slice")))),
          None,
          Some(Legend(LegendPosition.Left, overlay = false))
        ),
        DrawingAnchor.at(ref"D2", Extent(Emu(5400000L), Emu(2700000L)))
      )
      .addImage(png, ref"H20", extent)
    val out = write(Workbook(Vector(sheet)), "mixed")
    assertChartStructure(out, "xl/drawings/drawing1.xml")

    val drawingXml = entryText(out, "xl/drawings/drawing1.xml")
    assert(drawingXml.contains("""name="Image 1""""), drawingXml)
    assert(drawingXml.contains("""name="Chart 1""""), drawingXml)
    assert(drawingXml.contains("""name="Image 2""""), drawingXml)

    val readBack = reread(out)
    val drawings = readBack.sheets(0).drawings
    assertEquals(drawings.size, 3)
    assertEquals(readBack.sheets(0).pictures.size, 2)
    assertEquals(readBack.sheets(0).charts.size, 1)
    assertEquals(readBack.sheets(0).charts(0).chart.chartType, ChartType.Pie)
    // write-twice stability for the whole file
    val out2 = write(Workbook(Vector(sheet)), "mixed2")
    assertEquals(
      Files.readAllBytes(out).toSeq,
      Files.readAllBytes(out2).toSeq,
      "write-twice must be byte-identical"
    )
  }

  test(
    "chart referencing another sheet resolves caches across sheets; missing sheet omits caches"
  ) {
    val data = SheetName.unsafe("Numbers")
    val host = SheetName.unsafe("Dash")
    val numbers = Sheet(data)
      .put(Cell(ref"A1", CellValue.Number(BigDecimal(7))))
      .put(Cell(ref"A2", CellValue.Number(BigDecimal(9))))
    val chart = Chart(
      ChartType.Line,
      Vector(Series(DataRef(data, ref"A1:A2"), None, None)),
      None,
      None
    )
    val ghost = Chart(
      ChartType.Line,
      Vector(Series(DataRef(SheetName.unsafe("Ghost"), ref"A1:A2"), None, None)),
      None,
      None
    )
    val hostSheet = Sheet(host)
      .addChart(chart, ref"B2:H12")
      .addChart(ghost, ref"B14:H24")
    val out = write(Workbook(Vector(numbers, hostSheet)), "cross-sheet")
    val c1 = entryText(out, "xl/charts/chart1.xml")
    val c2 = entryText(out, "xl/charts/chart2.xml")
    assert(c1.contains("""<pt idx="0"><v>7</v></pt>"""), c1)
    assert(c2.contains("<val><numRef><f>Ghost!$A$1:$A$2</f></numRef></val>"), c2)
    val readBack = reread(out)
    assertEquals(readBack.sheets(1).charts.map(_.chart), Vector(chart, ghost))
  }
