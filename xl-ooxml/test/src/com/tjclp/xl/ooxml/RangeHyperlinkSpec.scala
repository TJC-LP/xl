package com.tjclp.xl.ooxml

import java.nio.file.{Files, Path}
import java.util.zip.ZipFile

import com.tjclp.xl.api.*
import com.tjclp.xl.addressing.CellRange
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ops.ClearWhat
import com.tjclp.xl.ooxml.writer.{WriterConfig, XmlBackend}
import com.tjclp.xl.unsafe.*
import munit.FunSuite

/**
 * #931 / #932: a hyperlink over a range (`<hyperlink ref="D2:E3">`, what Excel writes for a link
 * applied to a selection) is reported for every covered cell, written back unchanged when nothing
 * touches it, split when a covered cell loses its link, and authorable from the model.
 */
class RangeHyperlinkSpec extends FunSuite:

  private val sheetPart = "xl/worksheets/sheet1.xml"
  private val relsPart = "xl/worksheets/_rels/sheet1.xml.rels"
  private val url = "https://example.net/"
  private val link = Hyperlink(url)
  private val backends = Seq(XmlBackend.ScalaXml, XmlBackend.SaxStax)

  private def entryText(path: Path, name: String): Option[String] =
    val zip = new ZipFile(path.toFile)
    try
      Option(zip.getEntry(name)).map { e =>
        val is = zip.getInputStream(e)
        try new String(is.readAllBytes(), "UTF-8")
        finally is.close()
      }
    finally zip.close()

  private def entryBytes(path: Path, name: String): Array[Byte] =
    val zip = new ZipFile(path.toFile)
    try
      val is = zip.getInputStream(zip.getEntry(name))
      try is.readAllBytes()
      finally is.close()
    finally zip.close()

  /**
   * Rewrite one part of a written package (the way the #925 specs build an Excel-shaped source).
   */
  private def patchPart(written: Path, part: String)(patch: String => String): Path =
    val patched = Files.createTempFile("range-hl-source", ".xlsx")
    val in = new ZipFile(written.toFile)
    val out = new java.util.zip.ZipOutputStream(Files.newOutputStream(patched))
    try
      in.entries().asIterator().forEachRemaining { e =>
        val bytes = in.getInputStream(e).readAllBytes()
        val body =
          if e.getName == part then patch(new String(bytes, "UTF-8")).getBytes("UTF-8")
          else bytes
        out.putNextEntry(new java.util.zip.ZipEntry(e.getName))
        out.write(body)
        out.closeEntry()
      }
    finally
      out.close()
      in.close()
      Files.deleteIfExists(written)
    patched

  /**
   * A source with B7 linked on its own and `D2:E3` linked as one range with a tooltip and display,
   * as Excel writes a link applied to a selection. Only D2 holds a value.
   */
  private def rangeSource(): Path =
    val written = Files.createTempFile("range-hl-gen", ".xlsx")
    val wb = Workbook(
      Sheet("Summary")
        .put(Cell(ref"B7", CellValue.Text("Example link")).withHyperlink("https://example.com/"))
        .put(Cell(ref"D2", CellValue.Text("Range link")).withHyperlink(url))
    )
    XlsxWriter.write(wb, written).fold(e => fail(s"write failed: $e"), identity)
    patchPart(written, sheetPart)(
      _.replace("""ref="D2"""", """ref="D2:E3" tooltip="Docs" display="Docs link"""")
    )

  private def read(path: Path): (Workbook, Sheet) =
    val wb = XlsxReader.read(path).fold(e => fail(s"read failed: $e"), identity)
    (wb, wb("Summary").fold(e => fail(e.message), identity))

  private def write(wb: Workbook, config: WriterConfig = WriterConfig()): Path =
    val out = Files.createTempFile("range-hl-out", ".xlsx")
    XlsxWriter.writeWith(wb, out, config).fold(e => fail(s"write failed: $e"), identity)
    out

  /**
   * The `<hyperlink>` children of a worksheet part as attribute maps (a regenerated sheet
   * re-serializes the element, so the comparison is by content, not bytes).
   */
  private def hyperlinksXml(path: Path): List[Map[String, String]] =
    val xml = entryText(path, sheetPart).getOrElse(fail("no worksheet"))
    val root = XmlSecurity.parseSafe(xml, "worksheet").fold(e => fail(e.toString), identity)
    (root \ "hyperlinks" \ "hyperlink").toList.map(_.attributes.asAttrMap)

  /** Each written `<hyperlink>`'s ref with its resolved target (rel target or location). */
  private def writtenLinks(path: Path): List[(String, String)] =
    val xml = entryText(path, sheetPart).getOrElse(fail("no worksheet"))
    val root = XmlSecurity.parseSafe(xml, "worksheet").fold(e => fail(e.toString), identity)
    val targets = entryText(path, relsPart).toList.flatMap { rels =>
      val r = XmlSecurity.parseSafe(rels, "rels").fold(e => fail(e.toString), identity)
      (r \ "Relationship").map(e => (e \@ "Id") -> (e \@ "Target"))
    }.toMap
    (root \ "hyperlinks" \ "hyperlink").toList.map { e =>
      val id = e.attribute(XmlUtil.nsRelationships, "id").map(_.text)
      (e \@ "ref") -> id.fold(e \@ "location")(i => targets.getOrElse(i, fail(s"dangling $i")))
    }

  /** The cells each written link covers, as (cell, target), sorted. */
  private def coverage(path: Path): List[(String, String)] =
    writtenLinks(path).flatMap { (r, target) =>
      val range = CellRange.parse(r).fold(e => fail(e), identity)
      range.cellsRowMajor.map(_.toA1 -> target).toList
    }.sorted

  private def cells(a1: String*): List[(String, String)] = a1.map(_ -> url).toList

  // ===== #931: read =====

  test("#931: every cell a range link covers reports it, without creating cells") {
    val src = rangeSource()
    try
      val (_, sheet) = read(src)
      assertEquals(
        sheet.rangeHyperlinks,
        Vector(
          RangeHyperlink(
            CellRange.parse("D2:E3").toOption.get,
            link,
            Some("Docs"),
            Some("Docs link")
          )
        )
      )
      for a1 <- Seq("D2", "E2", "D3", "E3") do
        val at = ARef.parse(a1).toOption.get
        assertEquals(sheet(at).hyperlink, Some(link), a1)
        assertEquals(sheet.hyperlinkAt(at), Some(link), a1)
      assertEquals(sheet(ref"F2").hyperlink, None)
      assertEquals(sheet(ref"B7").hyperlink, Some(Hyperlink("https://example.com/")))
      assert(!sheet.contains(ref"E2"), "reading the range created E2")
      assertEquals(
        sheet.getRange(CellRange.parse("D2:E3").toOption.get).map(_.hyperlink).toList,
        List(Some(link))
      )
    finally Files.deleteIfExists(src)
  }

  test("#931: an unedited book with a range link writes byte-identical sheet parts") {
    val src = rangeSource()
    try
      val (wb, _) = read(src)
      for backend <- backends do
        val out = write(wb, WriterConfig(backend = backend))
        try
          assert(entryBytes(out, sheetPart).sameElements(entryBytes(src, sheetPart)), backend)
          assert(entryBytes(out, relsPart).sameElements(entryBytes(src, relsPart)), backend)
        finally Files.deleteIfExists(out)
    finally Files.deleteIfExists(src)
  }

  test("#931: an unrelated edit, or a value put through sheet(ref), keeps the source element") {
    val src = rangeSource()
    try
      val (wb, sheet) = read(src)
      val edits = Seq(
        sheet.put(ref"A1", CellValue.Text("edit")),
        sheet.put(sheet(ref"E2").withValue(CellValue.Text("typed"))),
        sheet.put(sheet(ref"D3").withStyle(com.tjclp.xl.styles.units.StyleId(0)))
      )
      for edited <- edits; backend <- backends do
        val out = write(wb.put(edited), WriterConfig(backend = backend))
        try
          assertEquals(hyperlinksXml(out), hyperlinksXml(src), backend)
          val (_, after) = read(out)
          assertEquals(after.rangeHyperlinks, sheet.rangeHyperlinks)
        finally Files.deleteIfExists(out)
    finally Files.deleteIfExists(src)
  }

  // ===== #931: removal splits the range, as in Excel =====

  test("#931: removing the link from one covered cell splits the range around it") {
    val src = rangeSource()
    try
      val (wb, sheet) = read(src)
      val removed = sheet.removeHyperlinks(CellRange.parse("E2:E2").toOption.get)
      assertEquals(removed(ref"E2").hyperlink, None)
      assertEquals(removed(ref"D2").hyperlink, Some(link))
      for backend <- backends do
        val out = write(wb.put(removed), WriterConfig(backend = backend))
        try
          assertEquals(
            coverage(out),
            (cells("D2", "D3", "E3") :+ ("B7" -> "https://example.com/")).sorted,
            backend
          )
          val (_, after) = read(out)
          assertEquals(after(ref"E2").hyperlink, None)
          for a1 <- Seq("D2", "D3", "E3") do
            assertEquals(after(ARef.parse(a1).toOption.get).hyperlink, Some(link), a1)
          // the remnants keep the source tooltip and display
          assert(after.rangeHyperlinks.forall(_.tooltip.contains("Docs")), after.rangeHyperlinks)
          assert(!after.contains(ref"E2"))
        finally Files.deleteIfExists(out)
    finally Files.deleteIfExists(src)
  }

  test("#931: removing the whole range drops its element and relationship") {
    val src = rangeSource()
    try
      val (wb, sheet) = read(src)
      val removed = sheet.removeHyperlinks(CellRange.parse("C1:F4").toOption.get)
      assertEquals(removed.rangeHyperlinks, Vector.empty)
      val out = write(wb.put(removed))
      try
        assertEquals(writtenLinks(out), List("B7" -> "https://example.com/"))
        assert(!entryText(out, relsPart).exists(_.contains(url)), "the range's rel must not ship")
      finally Files.deleteIfExists(out)
    finally Files.deleteIfExists(src)
  }

  test("#931: a contents clear over covered cells takes their link with it") {
    val src = rangeSource()
    try
      val (wb, sheet) = read(src)
      val cleared = sheet.clearRange(CellRange.parse("D3:E3").toOption.get, ClearWhat.contents)
      assertEquals(cleared(ref"E3").hyperlink, None)
      assertEquals(cleared(ref"E2").hyperlink, Some(link))
      val out = write(wb.put(cleared))
      try
        assertEquals(coverage(out), (cells("D2", "E2") :+ ("B7" -> "https://example.com/")).sorted)
      finally Files.deleteIfExists(out)
    finally Files.deleteIfExists(src)
  }

  test("#931: a cell link set inside the range overrides it there (#926 split)") {
    val src = rangeSource()
    try
      val (wb, sheet) = read(src)
      val other = "https://example.org/"
      val edited = sheet.put(sheet(ref"E3").withHyperlink(other))
      assertEquals(edited(ref"E3").hyperlink, Some(Hyperlink(other)))
      val out = write(wb.put(edited))
      try
        assertEquals(
          coverage(out),
          (cells("D2", "E2", "D3") ++ List("E3" -> other, "B7" -> "https://example.com/")).sorted
        )
      finally Files.deleteIfExists(out)
    finally Files.deleteIfExists(src)
  }

  // ===== #932: authoring a range link =====

  test("#932: setHyperlink over a range writes one range element and reads back on every cell") {
    val range = CellRange.parse("A1:A5").toOption.get
    val sheet = Sheet("Summary").put(ref"A1", CellValue.Text("Top")).setHyperlink(range, link)
    assertEquals(sheet.cells.size, 1, "a range link must not create a cell per position")
    for backend <- backends do
      val out = write(Workbook(sheet), WriterConfig(backend = backend))
      try
        assertEquals(writtenLinks(out), List("A1:A5" -> url), backend)
        val (_, after) = read(out)
        for row <- 1 to 5 do
          assertEquals(after(ARef.parse(s"A$row").toOption.get).hyperlink, Some(link), s"A$row")
        assertEquals(after.cells.size, 1)
      finally Files.deleteIfExists(out)
  }

  test("#932: an internal-location range link writes a location and round-trips") {
    val range = CellRange.parse("B2:C3").toOption.get
    val sheet = Sheet("Summary").setHyperlink(range, Hyperlink("Summary!A1"))
    val out = write(Workbook(sheet))
    try
      assertEquals(writtenLinks(out), List("B2:C3" -> "Summary!A1"))
      val (_, after) = read(out)
      assertEquals(after(ref"C3").hyperlink, Some(Hyperlink("Summary!A1")))
    finally Files.deleteIfExists(out)
  }

  test("#932: a range link set over part of a source range replaces it there") {
    val src = rangeSource()
    try
      val (wb, sheet) = read(src)
      val other = "https://example.org/"
      val edited = sheet.setHyperlink(CellRange.parse("E2:F3").toOption.get, Hyperlink(other))
      for backend <- backends do
        val out = write(wb.put(edited), WriterConfig(backend = backend))
        try
          assertEquals(
            coverage(out),
            (cells("D2", "D3") ++ List(
              "E2" -> other,
              "F2" -> other,
              "E3" -> other,
              "F3" -> other,
              "B7" -> "https://example.com/"
            )).sorted,
            backend
          )
          val (_, after) = read(out)
          assertEquals(after(ref"E2").hyperlink, Some(Hyperlink(other)))
          assertEquals(after(ref"D3").hyperlink, Some(link))
        finally Files.deleteIfExists(out)
    finally Files.deleteIfExists(src)
  }

  test("#932: a range link split twice keeps one relationship per link") {
    val range = CellRange.parse("A1:C3").toOption.get
    val sheet = Sheet("Summary")
      .setHyperlink(range, link)
      .removeHyperlinks(
        CellRange.parse("B2:B2").toOption.get
      )
    val out = write(Workbook(sheet))
    try
      assertEquals(
        coverage(out),
        cells("A1", "B1", "C1", "A2", "C2", "A3", "B3", "C3").sorted
      )
      val rels = entryText(out, relsPart).getOrElse(fail("no rels"))
      assertEquals(rels.split(java.util.regex.Pattern.quote(url), -1).length - 1, 1, rels)
      val (_, after) = read(out)
      assertEquals(after(ref"B2").hyperlink, None)
      assertEquals(after(ref"C3").hyperlink, Some(link))
    finally Files.deleteIfExists(out)
  }
