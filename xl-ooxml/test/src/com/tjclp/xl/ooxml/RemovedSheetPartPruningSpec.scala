package com.tjclp.xl.ooxml

import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path, Paths}
import java.util.zip.{ZipEntry, ZipFile, ZipOutputStream}
import scala.jdk.CollectionConverters.*

import munit.FunSuite

import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.api.*
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.drawings.TestImages
import com.tjclp.xl.macros.ref
import com.tjclp.xl.styles.units.Emu

/**
 * GH-417: deleting a sheet that carries charts/images used to leave the whole dependent chain —
 * xl/drawings/drawingN.xml, xl/charts/chartN.xml, their .rels, their media, their
 * [Content_Types].xml overrides — orphaned in the package forever (the preserved-part passthrough
 * re-emits by name with no reachability check against a removal).
 *
 * The fix prunes preserved parts reachable ONLY via a removed sheet's relationship closure.
 * Anything referenced by ANY surviving part survives (a shared image must stay — over-pruning is
 * silent data loss), and parts that were ALREADY unreferenced in the source keep riding through
 * (removal-induced orphans only, no general GC).
 */
class RemovedSheetPartPruningSpec extends FunSuite:

  private val png = ImageData(TestImages.png2x3, ImageFormat.Png)
  private val gif = ImageData(TestImages.gif2x3, ImageFormat.Gif)
  private val extent = Extent(Emu(95250L), Emu(190500L))

  private def name(s: String): SheetName = SheetName.unsafe(s)

  private def write(wb: Workbook, label: String): Path =
    val out = Files.createTempFile(s"xl-417-$label-", ".xlsx")
    out.toFile.deleteOnExit()
    XlsxWriter.write(wb, out).fold(err => fail(s"$label write failed: ${err.message}"), identity)
    out

  private def read(path: Path): Workbook =
    XlsxReader.read(path).fold(err => fail(s"read failed: ${err.message}"), identity)

  private def entryNames(path: Path): Set[String] =
    val zip = new ZipFile(path.toFile)
    try zip.entries().asScala.map(_.getName).toSet
    finally zip.close()

  private def entryBytes(path: Path, entry: String): Array[Byte] =
    val zip = new ZipFile(path.toFile)
    try
      Option(zip.getEntry(entry)) match
        case Some(e) => zip.getInputStream(e).readAllBytes()
        case None => fail(s"zip entry $entry not found in $path (have ${entryNames(path)})")
    finally zip.close()

  private def entryText(path: Path, entry: String): String =
    new String(entryBytes(path, entry), StandardCharsets.UTF_8)

  private def ctOverridePartNames(path: Path): Set[String] =
    val ct = scala.xml.XML.loadString(entryText(path, "[Content_Types].xml"))
    (ct \ "Override").map(o => (o \ "@PartName").text).toSet

  /** Every internal worksheet-rels relationship must resolve to a shipped entry. */
  private def assertWorksheetRelsResolve(path: Path): Unit =
    val names = entryNames(path)
    names.filter(n => n.startsWith("xl/worksheets/_rels/") && n.endsWith(".rels")).foreach {
      relsName =>
        val xml = scala.xml.XML.loadString(entryText(path, relsName))
        (xml \ "Relationship").foreach { rel =>
          if (rel \ "@TargetMode").text != "External" then
            val target = (rel \ "@Target").text
            val cleaned = if target.startsWith("/") then target.drop(1) else target
            val resolved =
              if cleaned.startsWith("xl/") then cleaned
              else
                Paths.get("xl/worksheets").resolve(cleaned).normalize().toString.replace('\\', '/')
            assert(
              names.contains(resolved),
              s"$relsName references $target -> $resolved which is MISSING from the output. " +
                s"Entries: ${names.toVector.sorted.mkString(", ")}"
            )
        }
    }

  /**
   * Copy a zip, appending extra entries and optionally patching existing ones (used to graft
   * chart-colors parts and pre-existing orphans onto writer output).
   */
  private def rewriteZip(
    src: Path,
    label: String,
    extra: Map[String, Array[Byte]],
    patch: Map[String, String => String] = Map.empty
  ): Path =
    val out = Files.createTempFile(s"xl-417-$label-grafted-", ".xlsx")
    out.toFile.deleteOnExit()
    val zin = new ZipFile(src.toFile)
    val zout = new ZipOutputStream(java.nio.file.Files.newOutputStream(out))
    try
      zin.entries().asScala.foreach { e =>
        val bytes = zin.getInputStream(e).readAllBytes()
        val patched = patch.get(e.getName) match
          case Some(f) =>
            f(new String(bytes, StandardCharsets.UTF_8)).getBytes(StandardCharsets.UTF_8)
          case None => bytes
        zout.putNextEntry(new ZipEntry(e.getName))
        zout.write(patched)
        zout.closeEntry()
      }
      extra.foreach { case (entryName, bytes) =>
        zout.putNextEntry(new ZipEntry(entryName))
        zout.write(bytes)
        zout.closeEntry()
      }
    finally
      zout.close()
      zin.close()
    out

  /** Two-sheet workbook: Data (values) + Charts (a bar chart over its own cells). */
  private def chartWorkbook: Workbook =
    val charts = name("Charts")
    val sheet = Sheet(charts)
      .put(Cell(ref"A2", CellValue.Text("a")))
      .put(Cell(ref"A3", CellValue.Text("b")))
      .put(Cell(ref"B2", CellValue.Number(BigDecimal(3))))
      .put(Cell(ref"B3", CellValue.Number(BigDecimal(4))))
    val chart = Chart
      .bar(
        Vector(Series(DataRef(charts, ref"B2:B3"), Some(DataRef(charts, ref"A2:A3")), None)),
        title = Some("Orphan Me")
      )
      .fold(e => fail(s"chart build failed: $e"), identity)
    Workbook(
      Vector(
        Sheet(name("Data")).put(Cell(ref"A1", CellValue.Text("keep"))),
        sheet.addChart(chart, ref"D2:K15")
      )
    )

  test("GH-417: removing a chart sheet prunes the whole chart/drawing chain from the package") {
    val src = write(chartWorkbook, "chart-chain")
    // Fixture sanity: the chain shipped in the source
    val srcNames = entryNames(src)
    assert(srcNames.contains("xl/drawings/drawing1.xml"), s"fixture missing drawing: $srcNames")
    assert(srcNames.contains("xl/charts/chart1.xml"), s"fixture missing chart: $srcNames")

    val removed = read(src).remove(name("Charts")).fold(e => fail(e.message), identity)
    val out = write(removed, "chart-chain-out")

    val names = entryNames(out)
    val leftovers = names.filter(n => n.startsWith("xl/charts/") || n.startsWith("xl/drawings/"))
    assertEquals(
      leftovers,
      Set.empty[String],
      s"chart/drawing parts must fall with their sheet, got ${names.toVector.sorted}"
    )

    // No dangling registrations either
    val overrides = ctOverridePartNames(out)
    assert(
      !overrides.exists(p => p.startsWith("/xl/charts/") || p.startsWith("/xl/drawings/")),
      s"[Content_Types].xml must not register pruned parts: $overrides"
    )

    val result = read(out)
    assertEquals(result.sheetNames.map(_.value), Vector("Data"))
    assertWorksheetRelsResolve(out)
  }

  test("GH-417: chart colors/style parts and their CT overrides fall with the chain") {
    // The writer never emits chartcolorstyle parts itself — graft one onto the package the way
    // Excel ships them (chart rels -> colors1.xml + a [Content_Types].xml override) to prove the
    // pruning walks the chart's own rels and cleans non-writer-owned content types too.
    val plain = write(chartWorkbook, "colors")
    val colorsXml =
      """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" meth="cycle" id="10"><a:schemeClr val="accent1"/></cs:colorStyle>"""
    val chartRels =
      """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors1.xml"/></Relationships>"""
    val ctOverride =
      """<Override PartName="/xl/charts/colors1.xml" ContentType="application/vnd.ms-office.chartcolorstyle+xml"/>"""
    val grafted = rewriteZip(
      plain,
      "colors",
      extra = Map(
        "xl/charts/colors1.xml" -> colorsXml.getBytes(StandardCharsets.UTF_8),
        "xl/charts/_rels/chart1.xml.rels" -> chartRels.getBytes(StandardCharsets.UTF_8)
      ),
      patch = Map(
        "[Content_Types].xml" -> (ct => ct.replace("</Types>", s"$ctOverride</Types>"))
      )
    )
    assert(entryNames(grafted).contains("xl/charts/colors1.xml"))

    val removed = read(grafted).remove(name("Charts")).fold(e => fail(e.message), identity)
    val out = write(removed, "colors-out")

    val names = entryNames(out)
    assert(
      !names.exists(n => n.startsWith("xl/charts/") || n.startsWith("xl/drawings/")),
      s"colors/style chain must fall with the chart: ${names.toVector.sorted}"
    )
    assert(
      !ctOverridePartNames(out).contains("/xl/charts/colors1.xml"),
      s"chartcolorstyle override must be pruned: ${ctOverridePartNames(out)}"
    )
    assertEquals(read(out).sheetNames.map(_.value), Vector("Data"))
  }

  test("GH-417: an image shared by two sheets SURVIVES removal of one of them") {
    val a = Sheet(name("A")).put(Cell(ref"A1", CellValue.Text("a"))).addImage(png, ref"B2", extent)
    val b = Sheet(name("B")).put(Cell(ref"A1", CellValue.Text("b"))).addImage(png, ref"C3", extent)
    val src = write(Workbook(Vector(a, b)), "shared-image")
    val srcNames = entryNames(src)
    // Fixture sanity: media dedup gives ONE media part referenced by BOTH drawing parts
    assertEquals(
      srcNames.count(_.startsWith("xl/media/")),
      1,
      s"fixture should dedup identical bytes to one media part: $srcNames"
    )
    assert(srcNames.contains("xl/drawings/drawing1.xml"))
    assert(srcNames.contains("xl/drawings/drawing2.xml"))

    val removed = read(src).remove(name("A")).fold(e => fail(e.message), identity)
    val out = write(removed, "shared-image-out")

    val names = entryNames(out)
    assertEquals(
      names.count(_.startsWith("xl/media/")),
      1,
      s"the shared image must SURVIVE (B still shows it): ${names.toVector.sorted}"
    )
    assert(
      !names.contains("xl/drawings/drawing1.xml") &&
        !names.contains("xl/drawings/_rels/drawing1.xml.rels"),
      s"A's drawing part must fall with A: ${names.toVector.sorted}"
    )
    assert(names.contains("xl/drawings/drawing2.xml"), "B's drawing part must survive")

    val result = read(out)
    assertEquals(result.sheetNames.map(_.value), Vector("B"))
    val pictures = result(name("B")).fold(e => fail(e.message), identity).pictures
    assertEquals(pictures.size, 1, "B must still render its picture after A is removed")
    assertWorksheetRelsResolve(out)
  }

  test("GH-417: media referenced only by the removed sheet is pruned; the survivor's stays") {
    val a = Sheet(name("A")).put(Cell(ref"A1", CellValue.Text("a"))).addImage(png, ref"B2", extent)
    val b = Sheet(name("B")).put(Cell(ref"A1", CellValue.Text("b"))).addImage(gif, ref"C3", extent)
    val src = write(Workbook(Vector(a, b)), "exclusive-media")
    assertEquals(entryNames(src).count(_.startsWith("xl/media/")), 2)

    val removed = read(src).remove(name("A")).fold(e => fail(e.message), identity)
    val out = write(removed, "exclusive-media-out")

    val media = entryNames(out).filter(_.startsWith("xl/media/"))
    assertEquals(
      media.map(_.split('.').last),
      Set("gif"),
      s"A's exclusive png must be pruned, B's gif must stay: $media"
    )
    val result = read(out)
    val pictures = result(name("B")).fold(e => fail(e.message), identity).pictures
    assertEquals(pictures.map(_.image.format), Vector[ImageFormat](ImageFormat.Gif))
  }

  test("GH-417: pre-existing orphans ride through untouched (removal-induced pruning only)") {
    // A part nothing references (already dead in the source) is preservation-sacred: pruning is
    // scoped to the removed sheet's closure, never a general package GC.
    val base = Workbook(
      Vector(
        Sheet(name("Data")).put(Cell(ref"A1", CellValue.Text("keep"))),
        Sheet(name("Gone")).put(Cell(ref"A1", CellValue.Text("bye")))
      )
    )
    val plain = write(base, "preexisting")
    val grafted = rewriteZip(
      plain,
      "preexisting",
      extra = Map(
        "xl/orphanPayload.xml" -> """<?xml version="1.0"?><payload/>""".getBytes(
          StandardCharsets.UTF_8
        )
      )
    )

    val removed = read(grafted).remove(name("Gone")).fold(e => fail(e.message), identity)
    val out = write(removed, "preexisting-out")

    assert(
      entryNames(out).contains("xl/orphanPayload.xml"),
      s"pre-existing orphan must NOT be GC'd by a sheet removal: ${entryNames(out).toVector.sorted}"
    )
    assertEquals(read(out).sheetNames.map(_.value), Vector("Data"))
  }
