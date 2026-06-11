package com.tjclp.xl.ooxml.drawing

import java.io.StringReader
import java.nio.file.{Files, Path}
import java.util.zip.ZipFile
import javax.xml.parsers.DocumentBuilderFactory
import scala.collection.immutable.ArraySeq

import munit.FunSuite

import com.tjclp.xl.api.*
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.drawings.TestImages
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.{TestFixtures, XlsxReader, XlsxWriter}
import com.tjclp.xl.ooxml.writer.WriterConfig
import com.tjclp.xl.styles.units.Emu

/**
 * GH-221 drawing layer: fixture parsing, authoring round-trips, hybrid Preserved survival, zip
 * structural validity, media dedup, and backend determinism.
 */
class DrawingRoundTripSpec extends FunSuite:

  private val png = ImageData(TestImages.png2x3, ImageFormat.Png)
  private val gif = ImageData(TestImages.gif2x3, ImageFormat.Gif)
  private val jpeg = ImageData(TestImages.jpeg2x3, ImageFormat.Jpeg)
  private val bmp = ImageData(TestImages.bmp2x3, ImageFormat.Bmp)
  private val extent = Extent(Emu(95250L), Emu(190500L))

  private def readFixture(name: String): (Path, Workbook) =
    val path = TestFixtures.copyToTemp(name)
    val wb = XlsxReader
      .read(path)
      .fold(err => fail(s"$name failed to read: ${err.message}"), identity)
    (path, wb)

  private def write(wb: Workbook, label: String, config: WriterConfig = WriterConfig()): Path =
    val out = Files.createTempFile(s"xl-drawing-$label-", ".xlsx")
    out.toFile.deleteOnExit()
    XlsxWriter
      .writeWith(wb, out, config)
      .fold(err => fail(s"$label write failed: ${err.message}"), _ => ())
    out

  private def reread(path: Path): Workbook =
    XlsxReader.read(path).fold(err => fail(s"re-read failed: ${err.message}"), identity)

  private def isPicture(d: Drawing): Boolean = d match
    case _: Drawing.Picture => true
    case _ => false

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

  // ===== fixture parsing =====

  test("image.xlsx parses to exactly the expected typed Picture") {
    val (in, wb) = readFixture("image.xlsx")
    val sheet = wb.sheets(0)
    assertEquals(sheet.drawings.size, 1)
    sheet.drawings(0) match
      case Drawing.Picture(DrawingAnchor.OneCell(from, ext), image, name, descr) =>
        assertEquals(from.cell, ref"B2") // CT_Marker col=1/row=1, 0-based
        assertEquals(from.dx, Emu(0L))
        assertEquals(from.dy, Emu(0L))
        assertEquals(ext, Extent(Emu(28575L), Emu(28575L)))
        assertEquals(image.format, ImageFormat.Png)
        assertEquals(name, "Image 1")
        assertEquals(descr, "Picture")
        assertEquals(
          image.bytes,
          ArraySeq.unsafeWrapArray(entryBytes(in, "xl/media/image1.png")),
          "picture bytes must equal the media part"
        )
      case other => fail(s"expected a OneCell Picture, got $other")
    assertEquals(sheet.pictures.size, 1)
  }

  test("chart-bar.xlsx parses to exactly one typed ChartFrame (GH-222)") {
    val (_, wb) = readFixture("chart-bar.xlsx")
    val sheet = wb.sheets(0)
    assertEquals(sheet.drawings.size, 1)
    val chartData = SheetName.unsafe("ChartData")
    sheet.drawings(0) match
      case Drawing.ChartFrame(DrawingAnchor.OneCell(from, ext), chart, name) =>
        assertEquals(from.cell, ref"D2") // CT_Marker col=3/row=1, 0-based
        assertEquals(ext, Extent(Emu(5400000L), Emu(2700000L)))
        assertEquals(name, "Chart 1")
        assertEquals(chart.chartType, ChartType.Bar(BarDirection.Col, BarGrouping.Clustered))
        assertEquals(chart.title, Some("Synthetic Units by Quarter"))
        assertEquals(chart.legend, Some(Legend(LegendPosition.Right, overlay = false)))
        assertEquals(
          chart.series,
          Vector(
            Series(
              values = DataRef(chartData, ref"B2:B5"),
              categories = Some(DataRef(chartData, ref"A2:A5")),
              name = Some(SeriesName.FromCell(chartData, ref"B1"))
            )
          )
        )
      case other => fail(s"expected a typed OneCell ChartFrame, got $other")
    assert(sheet.pictures.isEmpty)
    assertEquals(sheet.charts.size, 1)
  }

  test("chart-stacked.xlsx parses typed: two series, Stacked grouping, overlap dialect accepted") {
    val (_, wb) = readFixture("chart-stacked.xlsx")
    val sheet = wb.sheets(0)
    assertEquals(sheet.charts.size, 1)
    val chart = sheet.charts(0).chart
    assertEquals(chart.chartType, ChartType.Bar(BarDirection.Col, BarGrouping.Stacked))
    assertEquals(chart.title, Some("Synthetic Stacked Units"))
    val stackData = SheetName.unsafe("StackData")
    assertEquals(chart.series.size, 2)
    assertEquals(chart.series(0).values, DataRef(stackData, ref"B2:B4"))
    assertEquals(chart.series(1).values, DataRef(stackData, ref"C2:C4"))
    assertEquals(chart.series(0).name, Some(SeriesName.FromCell(stackData, ref"B1")))
  }

  test("chart-scatter.xlsx stays a whole-anchor Preserved fragment (outside the typed fence)") {
    val (in, wb) = readFixture("chart-scatter.xlsx")
    val sheet = wb.sheets(0)
    assertEquals(sheet.drawings.size, 1)
    sheet.drawings(0) match
      case Drawing.Preserved(xml) =>
        assert(xml.contains("graphicFrame"), s"expected graphicFrame in preserved xml: $xml")
      case other => fail(s"expected Preserved (scatter is out of fence), got $other")
    assert(sheet.charts.isEmpty)
    // byte round-trip pin: untouched write ships the scatter chart part verbatim
    val out = write(wb, "scatter-pin")
    assertEquals(
      entryBytes(out, "xl/charts/chart1.xml").toSeq,
      entryBytes(in, "xl/charts/chart1.xml").toSeq,
      "scatter chart part must ride byte-preservation"
    )
  }

  test("image-shape.xlsx (mixed wsDr) parses to [Picture, Preserved sp]") {
    val (_, wb) = readFixture("image-shape.xlsx")
    val sheet = wb.sheets(0)
    assertEquals(sheet.drawings.size, 2)
    assert(isPicture(sheet.drawings(0)), "first anchor is the typed picture")
    sheet.drawings(1) match
      case Drawing.Preserved(xml) =>
        assert(xml.contains("<sp>") || xml.contains("<sp "), s"expected sp shape: $xml")
        assert(xml.contains("twoCellAnchor"), "preserved fragment must be the whole anchor")
      case other => fail(s"expected Preserved sp, got $other")
  }

  // ===== zip structural validity (the #291-review template) =====

  private val nsRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

  private def parseNamespaceAware(xml: String, label: String): org.w3c.dom.Document =
    val factory = DocumentBuilderFactory.newInstance()
    factory.setNamespaceAware(true)
    try factory.newDocumentBuilder().parse(new org.xml.sax.InputSource(new StringReader(xml)))
    catch
      case e: org.xml.sax.SAXException =>
        fail(s"$label is not namespace-well-formed: ${e.getMessage}")

  /**
   * Excel-validity proxy for one drawing part: namespace-aware parse, every r:embed/r:id resolves
   * in the part rels, every rel target exists in the zip, ContentTypes carries an Override for the
   * part and a Default for each media extension, and cNvPr ids are unique within the part.
   */
  private def assertDrawingStructure(path: Path, partName: String): Unit =
    val names = entryNames(path)
    assert(names.contains(partName), s"missing $partName in $names")
    val doc = parseNamespaceAware(entryText(path, partName), partName)

    // collect r:embed / r:id references
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
      val slash = partName.lastIndexOf('/')
      val (dir, name) = partName.splitAt(slash + 1)
      s"${dir}_rels/$name.rels"
    val relTargets: Map[String, String] =
      if !names.contains(relsName) then Map.empty
      else
        val relsDoc = parseNamespaceAware(entryText(path, relsName), relsName)
        val rels = relsDoc.getElementsByTagNameNS("*", "Relationship")
        (0 until rels.getLength).flatMap { i =>
          Option(rels.item(i)).collect { case e: org.w3c.dom.Element =>
            e.getAttribute("Id") -> e.getAttribute("Target")
          }
        }.toMap

    val unresolved = referenced.filterNot(relTargets.contains)
    assert(unresolved.isEmpty, s"$partName references unresolved rIds: $unresolved")

    // every rel target exists in the zip (resolve ../ against the part dir, / against root)
    relTargets.values.foreach { target =>
      val resolved =
        if target.startsWith("/") then target.drop(1)
        else
          java.nio.file.Paths
            .get(partName)
            .getParent
            .resolve(target)
            .normalize()
            .toString
            .replace('\\', '/')
      assert(names.contains(resolved), s"rel target $target -> $resolved missing from zip")
    }

    // ContentTypes: Override for the part, Default for each media extension in the zip
    val ct = entryText(path, "[Content_Types].xml")
    assert(ct.contains(s"""PartName="/$partName""""), s"no ContentTypes override for $partName")
    names.filter(_.startsWith("xl/media/")).foreach { media =>
      val ext = media.split('.').lastOption.getOrElse("")
      assert(
        ct.toLowerCase.contains(s"""extension="${ext.toLowerCase}""""),
        s"no ContentTypes default for media extension .$ext"
      )
    }

    // cNvPr ids unique within the part
    val cnvprs = doc.getElementsByTagNameNS("*", "cNvPr")
    val ids = (0 until cnvprs.getLength).flatMap { i =>
      Option(cnvprs.item(i)).collect { case e: org.w3c.dom.Element => e.getAttribute("id") }
    }
    assertEquals(ids.distinct.size, ids.size, s"duplicate cNvPr ids in $partName: $ids")
    assert(ids.forall(id => id.toLongOption.exists(_ > 0)), s"cNvPr ids must be positive: $ids")

  // ===== fresh authoring round-trips: formats x anchors, both backends =====

  private def freshSheet(name: String): Sheet =
    Sheet(SheetName.unsafe(name)).put(com.tjclp.xl.cells.Cell(ref"A1", CellValue.Text("x")))

  List(
    "ScalaXml" -> WriterConfig.scalaXml,
    "SaxStax" -> WriterConfig.saxStax
  ).foreach { case (backend, config) =>
    test(s"fresh authoring round-trips all formats and anchor forms [$backend]") {
      val anchors: Vector[DrawingAnchor] = Vector(
        DrawingAnchor.at(ref"B2", extent),
        DrawingAnchor.over(ref"C3:E6", EditAs.OneCell),
        DrawingAnchor.Absolute(Emu(914400L), Emu(457200L), extent)
      )
      val images = Vector(png, jpeg, gif, bmp)
      val drawingsIn: Vector[Drawing] = for
        (img, i) <- images.zipWithIndex
        (anchor, j) <- anchors.zipWithIndex
      yield Drawing.Picture(
        anchor,
        img,
        name = s"Pic ${i * anchors.size + j + 1}",
        description = "alt"
      )
      val sheet = drawingsIn.foldLeft(freshSheet("Art")) { (s, d) =>
        d match
          case Drawing.Picture(a, img, n, descr) =>
            s.copy(drawings = s.drawings :+ Drawing.Picture(a, img, n, descr))
          case other => s.copy(drawings = s.drawings :+ other)
      }
      val out = write(Workbook(Vector(sheet)), s"fresh-$backend", config)
      val readBack = reread(out)
      assertEquals(readBack.sheets(0).drawings, drawingsIn, "drawings round-trip exactly")
      assertDrawingStructure(out, "xl/drawings/drawing1.xml")
      // worksheet references the part and the sheet rels resolve it
      val sheetXml = entryText(out, "xl/worksheets/sheet1.xml")
      assert(sheetXml.contains("<drawing"), "worksheet must reference its drawing part")
      assert(entryText(out, "xl/worksheets/_rels/sheet1.xml.rels").contains("drawing1.xml"))
    }
  }

  test("fresh authoring dedups identical bytes to one media part (same sheet and across sheets)") {
    val s1 = freshSheet("One").addImage(png, ref"A1", extent).addImage(png, ref"D4", extent)
    val s2 = freshSheet("Two").addImage(png, ref"B2", extent)
    val out = write(Workbook(Vector(s1, s2)), "dedup")
    val media = entryNames(out).filter(_.startsWith("xl/media/"))
    assertEquals(media, Set("xl/media/image1.png"), "identical bytes => exactly one media part")
    val readBack = reread(out)
    assertEquals(readBack.sheets(0).pictures.size, 2)
    assertEquals(readBack.sheets(1).pictures.size, 1)
    assertDrawingStructure(out, "xl/drawings/drawing1.xml")
    assertDrawingStructure(out, "xl/drawings/drawing2.xml")
  }

  // ===== hybrid survival: source-backed addImage / remove =====

  test("image.xlsx + addImage: original media byte-identical at original path, second appended") {
    val (in, wb) = readFixture("image.xlsx")
    val sheetName = wb.sheets(0).name
    val updated = wb
      .update(sheetName, _.addImage(jpeg, ref"D8", extent))
      .fold(err => fail(s"update failed: $err"), identity)
    val out = write(updated, "hybrid-image")

    // original media untouched at its original path
    assertEquals(
      entryBytes(out, "xl/media/image1.png").toSeq,
      entryBytes(in, "xl/media/image1.png").toSeq,
      "source media must be byte-identical"
    )
    // new media appended with the next number
    assert(entryNames(out).contains("xl/media/image2.jpeg"), entryNames(out).toString)

    val readBack = reread(out)
    val pics = readBack.sheets(0).pictures
    assertEquals(pics.size, 2)
    assertEquals(pics(0).image.format, ImageFormat.Png)
    assertEquals(pics(1).image.format, ImageFormat.Jpeg)
    assertEquals(pics(1).image.bytes, jpeg.bytes)
    assertDrawingStructure(out, "xl/drawings/drawing1.xml")
  }

  test("chart-bar.xlsx + addImage: chart bytes identical, typed ChartFrame stays FIRST") {
    val (in, wb) = readFixture("chart-bar.xlsx")
    val sheetName = wb.sheets(0).name
    val updated = wb
      .update(sheetName, _.addImage(png, ref"A10", extent))
      .fold(err => fail(s"update failed: $err"), identity)
    val out = write(updated, "hybrid-chart")

    // pins the GH-222 equality-match preservation path: the untouched chart's part rides
    // byte-preservation even though the drawing part regenerated
    assertEquals(
      entryBytes(out, "xl/charts/chart1.xml").toSeq,
      entryBytes(in, "xl/charts/chart1.xml").toSeq,
      "chart part must be byte-identical"
    )

    val readBack = reread(out)
    val drawings = readBack.sheets(0).drawings
    assertEquals(drawings.size, 2)
    drawings(0) match
      case Drawing.ChartFrame(_, chart, _) =>
        assertEquals(chart.title, Some("Synthetic Units by Quarter"))
      case other => fail(s"expected typed ChartFrame first (z-order law), got $other")
    assert(isPicture(drawings(1)), "added picture appended after")
    assertDrawingStructure(out, "xl/drawings/drawing1.xml")
  }

  test("image-shape.xlsx survives a cell edit and a picture add (mixed pic + sp)") {
    val (in, wb) = readFixture("image-shape.xlsx")
    val sheetName = wb.sheets(0).name
    // cell-only edit: drawing part rides preservation byte-identically
    val cellEdit = wb
      .update(sheetName, _.put(com.tjclp.xl.cells.Cell(ref"A5", CellValue.Text("hi"))))
      .fold(err => fail(s"update failed: $err"), identity)
    val out1 = write(cellEdit, "mixed-cell")
    assertEquals(
      entryBytes(out1, "xl/drawings/drawing1.xml").toSeq,
      entryBytes(in, "xl/drawings/drawing1.xml").toSeq,
      "clean drawings must ride preservation byte-identically"
    )
    // drawing edit: picture + shape + new picture all survive a regeneration
    val drawingEdit = wb
      .update(sheetName, _.addImage(gif, ref"H2", extent))
      .fold(err => fail(s"update failed: $err"), identity)
    val out2 = write(drawingEdit, "mixed-add")
    val readBack = reread(out2)
    val drawings = readBack.sheets(0).drawings
    assertEquals(drawings.size, 3)
    assert(isPicture(drawings(0)))
    assert(!isPicture(drawings(1)), "sp shape survives as Preserved in document order")
    assert(isPicture(drawings(2)))
    assertDrawingStructure(out2, "xl/drawings/drawing1.xml")
  }

  test("image.xlsx remove-all drawings yields a schema-valid empty wsDr that re-reads") {
    val (_, wb) = readFixture("image.xlsx")
    val sheetName = wb.sheets(0).name
    val updated = wb
      .update(sheetName, _.removeDrawing(0))
      .fold(err => fail(s"update failed: $err"), identity)
    val out = write(updated, "remove-all")
    val partXml = entryText(out, "xl/drawings/drawing1.xml")
    assert(partXml.contains("wsDr"), s"empty wsDr expected: $partXml")
    assert(!partXml.contains("oneCellAnchor"), "anchors must be gone")
    // worksheet keeps its (now empty) drawing reference and the rels keep resolving
    assert(entryText(out, "xl/worksheets/sheet1.xml").contains("<drawing"))
    assert(entryText(out, "xl/worksheets/_rels/sheet1.xml.rels").contains("drawing1.xml"))
    // media is never deleted in 6a
    assert(entryNames(out).contains("xl/media/image1.png"))
    val readBack = reread(out)
    assertEquals(readBack.sheets(0).drawings, Vector.empty[Drawing])
  }

  test("adding the same image bytes to a source-backed sheet reuses the source media part") {
    val (_, wb) = readFixture("image.xlsx")
    val sheetName = wb.sheets(0).name
    val fixturePng = wb.sheets(0).pictures(0).image
    val updated = wb
      .update(sheetName, _.addImage(fixturePng, ref"F2", extent))
      .fold(err => fail(s"update failed: $err"), identity)
    val out = write(updated, "reuse-media")
    val media = entryNames(out).filter(_.startsWith("xl/media/"))
    assertEquals(media, Set("xl/media/image1.png"), "sha-identical bytes reuse the source part")
    val readBack = reread(out)
    assertEquals(readBack.sheets(0).pictures.size, 2)
    assertDrawingStructure(out, "xl/drawings/drawing1.xml")
  }

  test("appending a NEW sheet with an image to a source-backed workbook (metadata-modified)") {
    val (in, wb) = readFixture("image.xlsx")
    val extra = freshSheet("Extra").addImage(gif, ref"B2", extent)
    val updated = wb.put(extra) // appends; marks metadata modified -> minimal ContentTypes branch
    val out = write(updated, "new-sheet")

    // original drawing part still ships byte-identically (clean drawings ride preservation even
    // though ALL worksheets regenerate on a metadata-modified write)
    assertEquals(
      entryBytes(out, "xl/drawings/drawing1.xml").toSeq,
      entryBytes(in, "xl/drawings/drawing1.xml").toSeq
    )
    // fresh part numbered after the manifest max
    assertDrawingStructure(out, "xl/drawings/drawing1.xml")
    assertDrawingStructure(out, "xl/drawings/drawing2.xml")
    assert(entryNames(out).contains("xl/media/image2.gif"))
    // the regenerated [Content_Types].xml carries BOTH drawing overrides and BOTH media defaults
    val ct = entryText(out, "[Content_Types].xml")
    assert(ct.contains("/xl/drawings/drawing1.xml"), "manifest drawing override lost")
    assert(ct.contains("/xl/drawings/drawing2.xml"), "fresh drawing override missing")

    val readBack = reread(out)
    assertEquals(readBack.sheets.map(_.pictures.size), Vector(1, 1))
    assertEquals(readBack.sheets(1).pictures(0).image.bytes, gif.bytes)
  }

  // ===== determinism =====

  test("dirty drawing write is write-twice byte-identical (whole file)") {
    val (_, wb) = readFixture("image.xlsx")
    val sheetName = wb.sheets(0).name
    val updated = wb
      .update(sheetName, _.addImage(png, ref"D8", extent))
      .fold(err => fail(s"update failed: $err"), identity)
    val out1 = write(updated, "det1")
    val out2 = write(updated, "det2")
    assertEquals(
      Files.readAllBytes(out1).toSeq,
      Files.readAllBytes(out2).toSeq,
      "write-twice must be byte-identical"
    )
  }

  test("ScalaXml and SaxStax agree on the drawing part (semantic parity + canonical decls)") {
    // NOTE: whole-part byte equality across backends is impossible today for ANY part — the two
    // writers differ globally in the XML declaration (standalone) and empty-element minimization.
    // The contract here matches the house worksheet-parity precedent: identical model after
    // re-read, identical structure, identical namespace declarations in identical order.
    val sheet = freshSheet("Par")
      .addImage(png, ref"B2", extent)
      .addImage(gif, ref"C4:D8", EditAs.Absolute)
    val wb = Workbook(Vector(sheet))
    val outDom = write(wb, "par-dom", WriterConfig.scalaXml)
    val outSax = write(wb, "par-sax", WriterConfig.saxStax)

    assertEquals(
      reread(outDom).sheets(0).drawings,
      reread(outSax).sheets(0).drawings,
      "backends must agree on the parsed drawing model"
    )
    assertDrawingStructure(outDom, "xl/drawings/drawing1.xml")
    assertDrawingStructure(outSax, "xl/drawings/drawing1.xml")

    // canonical sorted xmlns declarations, identical on both backends
    def decls(path: Path): List[String] =
      """xmlns(:[a-z0-9]+)?="[^"]*"""".r
        .findAllIn(entryText(path, "xl/drawings/drawing1.xml"))
        .toList
    assertEquals(decls(outDom), decls(outSax), "namespace declarations must match in order")
    assertEquals(
      decls(outDom).map(_.takeWhile(_ != '=')),
      decls(outDom).map(_.takeWhile(_ != '=')).sorted,
      "declarations print in canonical sorted order"
    )

    // and each backend is individually write-twice byte-identical
    val outDom2 = write(wb, "par-dom2", WriterConfig.scalaXml)
    val outSax2 = write(wb, "par-sax2", WriterConfig.saxStax)
    assertEquals(
      entryText(outDom, "xl/drawings/drawing1.xml"),
      entryText(outDom2, "xl/drawings/drawing1.xml")
    )
    assertEquals(
      entryText(outSax, "xl/drawings/drawing1.xml"),
      entryText(outSax2, "xl/drawings/drawing1.xml")
    )
  }

  test("read(write) is stable for a dirty source-backed workbook (second trip is clean)") {
    val (_, wb) = readFixture("chart-bar.xlsx")
    val sheetName = wb.sheets(0).name
    val updated = wb
      .update(sheetName, _.addImage(png, ref"A10", extent))
      .fold(err => fail(s"update failed: $err"), identity)
    val out1 = write(updated, "stable1")
    val read1 = reread(out1)
    val out2 = write(read1, "stable2")
    val read2 = reread(out2)
    assertEquals(read2.sheets(0).drawings, read1.sheets(0).drawings, "fixpoint after one trip")
  }
