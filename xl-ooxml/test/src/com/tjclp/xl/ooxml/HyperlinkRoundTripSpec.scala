package com.tjclp.xl.ooxml

import java.nio.file.{Files, Path}
import java.util.zip.ZipFile

import com.tjclp.xl.api.*
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.macros.ref
import com.tjclp.xl.unsafe.*
import munit.FunSuite

/**
 * GH-235: Cell.hyperlink was modeled with a public API but never serialized (a silent no-op). These
 * tests prove hyperlinks now write (<hyperlinks> + external-URL relationships) and round-trip back
 * into the model.
 */
class HyperlinkRoundTripSpec extends FunSuite:

  private def entryText(zip: ZipFile, name: String): String =
    val is = zip.getInputStream(zip.getEntry(name))
    try new String(is.readAllBytes(), "UTF-8")
    finally is.close()

  test("GH-235: external hyperlink serializes (<hyperlinks> + External rel) and round-trips") {
    val cell = Cell(ref"A1", CellValue.Text("Anthropic")).withHyperlink("https://anthropic.com")
    val wb = Workbook(Sheet("Sheet1").put(cell))
    val out = Files.createTempFile("hl-ext", ".xlsx")
    XlsxWriter.write(wb, out).fold(e => fail(s"write failed: $e"), identity)

    // Round-trip: Cell.hyperlink populated on read
    val reread = XlsxReader.read(out).fold(e => fail(s"read failed: $e"), identity)
    val sheet = reread("Sheet1").fold(e => fail(s"sheet missing: $e"), identity)
    assertEquals(sheet(ref"A1").hyperlink, Some(Hyperlink("https://anthropic.com")))

    // Raw output: <hyperlinks> in the sheet + an External hyperlink relationship in its .rels
    val zip = new ZipFile(out.toFile)
    val sheetXml = entryText(zip, "xl/worksheets/sheet1.xml")
    val relsXml = entryText(zip, "xl/worksheets/_rels/sheet1.xml.rels")
    zip.close()
    assert(sheetXml.contains("<hyperlink "), s"no <hyperlink> in sheet:\n$sheetXml")
    assert(
      relsXml.contains("anthropic.com") && relsXml.contains("External"),
      s"no external hyperlink rel:\n$relsXml"
    )
    Files.deleteIfExists(out)
  }

  test("GH-235: internal location hyperlink uses the location attribute and round-trips") {
    val cell = Cell(ref"A1", CellValue.Text("Go")).withHyperlink("Sheet1!B2")
    val wb = Workbook(Sheet("Sheet1").put(cell))
    val out = Files.createTempFile("hl-int", ".xlsx")
    XlsxWriter.write(wb, out).fold(e => fail(s"write failed: $e"), identity)

    val reread = XlsxReader.read(out).fold(e => fail(s"read failed: $e"), identity)
    val sheet = reread("Sheet1").fold(e => fail(s"sheet missing: $e"), identity)
    assertEquals(sheet(ref"A1").hyperlink, Some(Hyperlink("Sheet1!B2")))
    Files.deleteIfExists(out)
  }

  test("GH-406: '#'-prefixed internal target normalizes in withHyperlink and round-trips clean") {
    // The '#Sheet1!A1' dialect is common agent/browser spelling; OOXML location wants no '#'.
    val cell = Cell(ref"A1", CellValue.Text("Go")).withHyperlink("#Sheet1!B2")
    assertEquals(cell.hyperlink, Some(Hyperlink("Sheet1!B2")))

    val wb = Workbook(Sheet("Sheet1").put(cell))
    val out = Files.createTempFile("hl-hash", ".xlsx")
    XlsxWriter.write(wb, out).fold(e => fail(s"write failed: $e"), identity)

    val zip = new ZipFile(out.toFile)
    val sheetXml = entryText(zip, "xl/worksheets/sheet1.xml")
    zip.close()
    assert(
      sheetXml.contains("""location="Sheet1!B2""""),
      s"location must not carry the leading '#':\n$sheetXml"
    )

    val reread = XlsxReader.read(out).fold(e => fail(s"read failed: $e"), identity)
    val sheet = reread("Sheet1").fold(e => fail(s"sheet missing: $e"), identity)
    assertEquals(sheet(ref"A1").hyperlink, Some(Hyperlink("Sheet1!B2")))
    Files.deleteIfExists(out)
  }

  test("GH-406: writer strips a leading '#' even when the field bypassed withHyperlink") {
    // Direct construction (copy/generators) skips the chokepoint; the location emit defends.
    val cell = Cell(ref"A1", CellValue.Text("Go"), None, Some(Hyperlink("#Sheet1!C3")))
    val wb = Workbook(Sheet("Sheet1").put(cell))
    val out = Files.createTempFile("hl-hash-direct", ".xlsx")
    XlsxWriter.write(wb, out).fold(e => fail(s"write failed: $e"), identity)

    val zip = new ZipFile(out.toFile)
    val sheetXml = entryText(zip, "xl/worksheets/sheet1.xml")
    zip.close()
    assert(
      sheetXml.contains("""location="Sheet1!C3""""),
      s"location must not carry the leading '#':\n$sheetXml"
    )
    Files.deleteIfExists(out)
  }

  test("GH-406: external URL fragments are untouched (only a LEADING '#' strips)") {
    val target = "https://example.com/page#section"
    val cell = Cell(ref"A1", CellValue.Text("Docs")).withHyperlink(target)
    assertEquals(cell.hyperlink, Some(Hyperlink(target)))

    val wb = Workbook(Sheet("Sheet1").put(cell))
    val out = Files.createTempFile("hl-frag", ".xlsx")
    XlsxWriter.write(wb, out).fold(e => fail(s"write failed: $e"), identity)

    val reread = XlsxReader.read(out).fold(e => fail(s"read failed: $e"), identity)
    val sheet = reread("Sheet1").fold(e => fail(s"sheet missing: $e"), identity)
    assertEquals(sheet(ref"A1").hyperlink, Some(Hyperlink(target)))

    val zip = new ZipFile(out.toFile)
    val relsXml = entryText(zip, "xl/worksheets/_rels/sheet1.xml.rels")
    zip.close()
    assert(relsXml.contains("page#section"), s"external rel must keep its fragment:\n$relsXml")
    Files.deleteIfExists(out)
  }

  // ===== GH-895: removing a sheet's last hyperlink sticks through a read-modify-write =====

  /** Write `cells` on sheet "Summary", read it back: a workbook carrying the source's metadata. */
  private def readFromFile(cells: Cell*): (Workbook, java.nio.file.Path) =
    val src = Files.createTempFile("hl-src", ".xlsx")
    val wb = Workbook(cells.foldLeft(Sheet("Summary"))(_.put(_)))
    XlsxWriter.write(wb, src).fold(e => fail(s"write failed: $e"), identity)
    (XlsxReader.read(src).fold(e => fail(s"read failed: $e"), identity), src)

  private def clearLink(wb: Workbook, at: com.tjclp.xl.addressing.ARef): Workbook =
    val sheet = wb("Summary").fold(e => fail(s"sheet missing: $e"), identity)
    wb.put(sheet.put(sheet(at).clearHyperlink))

  private def writeAndReopen(
    wb: Workbook,
    config: WriterConfig = WriterConfig()
  ): (Sheet, String, Option[String]) =
    val out = Files.createTempFile("hl-out", ".xlsx")
    XlsxWriter.writeWith(wb, out, config).fold(e => fail(s"write failed: $e"), identity)
    val sheet = XlsxReader
      .read(out)
      .flatMap(_("Summary"))
      .fold(e => fail(s"reread failed: $e"), identity)
    val zip = new ZipFile(out.toFile)
    val sheetXml = entryText(zip, "xl/worksheets/sheet1.xml")
    val relsName = "xl/worksheets/_rels/sheet1.xml.rels"
    val relsXml = Option(zip.getEntry(relsName)).map(_ => entryText(zip, relsName))
    zip.close()
    Files.deleteIfExists(out)
    (sheet, sheetXml, relsXml)

  test("GH-895: removing a sheet's only external hyperlink drops <hyperlinks> and its rel") {
    val (wb, src) = readFromFile(
      Cell(ref"B7", CellValue.Text("Example link")).withHyperlink("https://example.com/")
    )
    val (sheet, sheetXml, relsXml) = writeAndReopen(clearLink(wb, ref"B7"))
    assertEquals(sheet(ref"B7").hyperlink, None, "B7's hyperlink must stay removed")
    assertEquals(sheet(ref"B7").value, CellValue.Text("Example link"), "B7 keeps its text")
    assert(!sheetXml.contains("<hyperlinks"), s"no <hyperlinks> element expected:\n$sheetXml")
    relsXml.foreach { rels =>
      assert(!rels.contains("example.com"), s"the removed link's rel must not ship:\n$rels")
    }
    Files.deleteIfExists(src)
  }

  test("GH-895: removing a sheet's only internal (location) hyperlink drops <hyperlinks>") {
    val (wb, src) = readFromFile(
      Cell(ref"B7", CellValue.Text("Go")).withHyperlink("Summary!A1")
    )
    val (sheet, sheetXml, _) = writeAndReopen(clearLink(wb, ref"B7"))
    assertEquals(sheet(ref"B7").hyperlink, None, "B7's hyperlink must stay removed")
    assert(!sheetXml.contains("<hyperlinks"), s"no <hyperlinks> element expected:\n$sheetXml")
    Files.deleteIfExists(src)
  }

  test("GH-895: removing one of two hyperlinks keeps the other (and only the other)") {
    val (wb, src) = readFromFile(
      Cell(ref"B7", CellValue.Text("Example link")).withHyperlink("https://example.com/"),
      Cell(ref"B8", CellValue.Text("Other link")).withHyperlink("https://example.org/")
    )
    val (sheet, sheetXml, relsXml) = writeAndReopen(clearLink(wb, ref"B7"))
    assertEquals(sheet(ref"B7").hyperlink, None)
    assertEquals(sheet(ref"B8").hyperlink, Some(Hyperlink("https://example.org/")))
    assert(!sheetXml.contains("""ref="B7""""), s"B7 must not be linked:\n$sheetXml")
    relsXml.foreach(rels => assert(!rels.contains("example.com"), rels))
    Files.deleteIfExists(src)
  }

  test("GH-895: an untouched read-write keeps the sheet's only hyperlink") {
    val (wb, src) = readFromFile(
      Cell(ref"B7", CellValue.Text("Example link")).withHyperlink("https://example.com/")
    )
    val (sheet, sheetXml, relsXml) = writeAndReopen(wb)
    assertEquals(sheet(ref"B7").hyperlink, Some(Hyperlink("https://example.com/")))
    assert(sheetXml.contains("<hyperlinks"), sheetXml)
    assert(relsXml.exists(_.contains("example.com")), s"the link's rel must ship: $relsXml")
    Files.deleteIfExists(src)
  }

  /** A source whose sheet has a cell link on B7 and a range link (unmodeled) on D2:E3. */
  private def sourceWithRangeLink(): java.nio.file.Path =
    val written = Files.createTempFile("hl-range-gen", ".xlsx")
    val wb = Workbook(
      Sheet("Summary")
        .put(Cell(ref"B7", CellValue.Text("Example link")).withHyperlink("https://example.com/"))
        .put(Cell(ref"D2", CellValue.Text("Range link")).withHyperlink("https://example.net/"))
    )
    XlsxWriter.write(wb, written).fold(e => fail(s"write failed: $e"), identity)
    // Excel writes ref="D2:E3" for a link over a multi-cell selection; xl cannot author one, so
    // the source is patched: the reader keeps it as preserved XML, not on a cell.
    patchPart(written, "xl/worksheets/sheet1.xml")(_.replace("""ref="D2"""", """ref="D2:E3""""))

  private def patchPart(written: Path, part: String)(patch: String => String): Path =
    val patched = Files.createTempFile("hl-source", ".xlsx")
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

  test("GH-895: removing the last cell link keeps a range link and its rel") {
    val src = sourceWithRangeLink()
    val wb = XlsxReader.read(src).fold(e => fail(s"read failed: $e"), identity)
    val (sheet, sheetXml, relsXml) = writeAndReopen(clearLink(wb, ref"B7"))
    assertEquals(sheet(ref"B7").hyperlink, None, "B7's hyperlink must stay removed")
    assert(!sheetXml.contains("""ref="B7""""), s"B7 must not be linked:\n$sheetXml")
    assert(sheetXml.contains("""ref="D2:E3""""), s"the range link must survive:\n$sheetXml")
    val rels = relsXml.getOrElse(fail("the range link's rel must ship"))
    assert(rels.contains("example.net"), s"the range link's rel must ship:\n$rels")
    assert(!rels.contains("example.com"), s"the removed link's rel must not ship:\n$rels")
    Files.deleteIfExists(src)
  }

  test("GH-895: a cell link the reader cannot resolve survives an edit, with its rel") {
    // ISO 29500 Strict names the hyperlink rel type under purl.oclc.org; the reader resolves only
    // the transitional type, so B7's link loads onto no cell and must ride through as XML.
    val written = Files.createTempFile("hl-strict-gen", ".xlsx")
    val seed = Workbook(
      Sheet("Summary")
        .put(Cell(ref"B7", CellValue.Text("Example link")).withHyperlink("https://example.com/"))
    )
    XlsxWriter.write(seed, written).fold(e => fail(s"write failed: $e"), identity)
    val src = patchPart(written, "xl/worksheets/_rels/sheet1.xml.rels")(
      _.replace(XmlUtil.relTypeHyperlink, strictHyperlinkRelType)
    )
    val wb = XlsxReader.read(src).fold(e => fail(s"read failed: $e"), identity)
    val sheet = wb("Summary").fold(e => fail(s"sheet missing: $e"), identity)
    assertEquals(sheet(ref"B7").hyperlink, None, "precondition: the reader leaves B7 unmodeled")

    val edited = wb.put(sheet.put(Cell(ref"A1", CellValue.Text("edit"))))
    val (_, sheetXml, relsXml) = writeAndReopen(edited)
    assert(sheetXml.contains("""<hyperlink ref="B7""""), s"B7's link must survive:\n$sheetXml")
    val rels = relsXml.getOrElse(fail("B7's rel must ship"))
    assert(rels.contains("example.com"), s"B7's rel must ship:\n$rels")
    Files.deleteIfExists(src)
  }

  private val strictHyperlinkRelType =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/hyperlink"

  test("GH-895: MC Choice and Fallback retain their links, namespaces and reserved IDs") {
    val (_, written) = readFromFile(
      Cell(ref"B7", CellValue.Text("Choice")).withHyperlink("https://example.com/")
    )
    val feature = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
    val src = patchPart(written, "xl/worksheets/sheet1.xml") { xml =>
      xml
        .replaceFirst(
          "(<hyperlinks[^>]*>)",
          s"""$$1<mc:AlternateContent xmlns:mc="${XmlUtil.nsMarkupCompatibility}" xmlns:feature="$feature"><mc:Choice Requires="feature">"""
        )
        .replace(
          "</hyperlinks>",
          """</mc:Choice><mc:Fallback><hyperlink ref="B7" location="Summary!A1"/></mc:Fallback></mc:AlternateContent></hyperlinks>"""
        )
    }
    try
      val wb = XlsxReader.read(src).fold(e => fail(e.message), identity)
      val sheet = wb("Summary").fold(e => fail(e.message), identity)
      assertEquals(sheet(ref"B7").hyperlink, None)
      val plain = sheet.put(Cell(ref"A1", CellValue.Text("edit")))
      val linked =
        plain.put(Cell(ref"A2", CellValue.Text("New")).withHyperlink("https://example.org/"))
      for
        backend <- Seq(XmlBackend.ScalaXml, XmlBackend.SaxStax)
        edited <- Seq(plain, linked)
      do
        val (after, xml, rels) = writeAndReopen(wb.put(edited), WriterConfig(backend = backend))
        val root = XmlSecurity.parseSafe(xml, "worksheet").fold(e => fail(e.toString), identity)
        val choice = (root \\ "Choice").headOption.getOrElse(fail("lost MC choice"))
        assertEquals(Option(choice.getNamespace(choice \@ "Requires")), Some(feature))
        assertEquals((root \\ "Fallback" \ "hyperlink") \@ "location", "Summary!A1")
        val choiceId =
          (choice \ "hyperlink").headOption
            .flatMap(_.attribute(XmlUtil.nsRelationships, "id"))
            .map(_.text)
        val relXml = XmlSecurity
          .parseSafe(rels.getOrElse(fail("lost hyperlink rels")), "rels")
          .fold(e => fail(e.toString), identity)
        val entries = (relXml \ "Relationship").map(e => (e \@ "Id") -> (e \@ "Target"))
        assertEquals(entries.map(_._1).distinct.size, entries.size)
        assertEquals(choiceId.flatMap(entries.toMap.get), Some("https://example.com/"))
        assertEquals(after(ref"A2").hyperlink, edited(ref"A2").hyperlink)
    finally Files.deleteIfExists(src)
  }

  test("GH-895: an unrelated edit keeps a worksheet-level MC hyperlink relationship") {
    val (_, written) = readFromFile(
      Cell(ref"B7", CellValue.Text("Choice")).withHyperlink("https://example.com/")
    )
    val src = patchPart(written, "xl/worksheets/sheet1.xml") { xml =>
      xml
        .replaceFirst(
          "(<hyperlinks[^>]*>)",
          s"""<mc:AlternateContent xmlns:mc="${XmlUtil.nsMarkupCompatibility}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><mc:Choice Requires="x14">$$1"""
        )
        .replace("</hyperlinks>", "</hyperlinks></mc:Choice><mc:Fallback/></mc:AlternateContent>")
    }
    try
      val wb = XlsxReader.read(src).fold(e => fail(e.message), identity)
      val edited = wb
        .update("Summary", _.put(ref"A1", CellValue.Text("edit")))
        .fold(e => fail(e.message), identity)
      val (_, xml, rels) = writeAndReopen(edited)
      assert(xml.contains("AlternateContent"), xml)
      assert(rels.exists(_.contains("https://example.com/")), rels.toString)
    finally Files.deleteIfExists(src)
  }

  private case class SourceLink(ref: String, id: String, target: String)

  private val b7 = SourceLink("B7", "sourceCell", "https://example.com/")
  private val d2e3 = SourceLink("D2:E3", "sourceRange", "https://example.net/")

  private def checkMixedLinks(links: SourceLink*): Unit =
    val sheetPart = "xl/worksheets/sheet1.xml"
    val relsPart = "xl/worksheets/_rels/sheet1.xml.rels"
    val (_, seed) = readFromFile(Cell(ref"B7", CellValue.Text("Link")).withHyperlink(b7.target))
    val elements = links.map(l => <hyperlink ref={l.ref} r:id={l.id}/>)
    val hyperlinks = <hyperlinks xmlns:r={XmlUtil.nsRelationships}>{elements}</hyperlinks>
    val withRefs = patchPart(seed, sheetPart)(
      _.replaceFirst("(?s)<hyperlinks\\b[^>]*>.*?</hyperlinks>", hyperlinks.toString)
    )
    val relationships = links.map { link =>
      <Relationship Id={link.id} Type={XmlUtil.relTypeHyperlink} Target={
        link.target
      } TargetMode="External"/>
    }
    val src = patchPart(withRefs, relsPart)(_ =>
      <Relationships xmlns={XmlUtil.nsPackageRels}>{relationships}</Relationships>.toString
    )

    val expected = links.map(l => l.ref -> l.target).toMap
    try
      val zip = new ZipFile(src.toFile)
      try
        assertEquals(
          resolvedLinks(entryText(zip, sheetPart), entryText(zip, relsPart)).sorted,
          expected.toList.sorted
        )
      finally zip.close()
      val wb = XlsxReader.read(src).fold(e => fail(e.message), identity)
      val before = wb("Summary").fold(e => fail(e.message), identity)
      def checkModel(sheet: Sheet): Unit =
        links.foreach { l =>
          val anchor = l.ref.takeWhile(_ != ':')
          val cell = sheet(anchor).fold(e => fail(e.message), identity)
          assertEquals(cell.hyperlink, Option.unless(l.ref.contains(':'))(Hyperlink(l.target)))
        }
      checkModel(before)
      val edited = wb.put(before.put(ref"A1", CellValue.Text("edit")))
      for backend <- Seq(XmlBackend.ScalaXml, XmlBackend.SaxStax) do
        val (after, xml, rels) = writeAndReopen(edited, WriterConfig(backend = backend))
        assertEquals(after(ref"A1").value, CellValue.Text("edit"))
        checkModel(after)
        assertEquals(
          resolvedLinks(xml, rels.getOrElse(fail("missing sheet relationships"))).sorted,
          expected.toList.sorted
        )
    finally Files.deleteIfExists(src)

  test("GH-912: an unrelated edit keeps B7 and the D2:E3 range hyperlink") {
    checkMixedLinks(b7, d2e3)
  }

  test("GH-912: several cell hyperlinks survive alongside a range hyperlink") {
    checkMixedLinks(b7, SourceLink("C9", "otherCell", "https://example.com/other"), d2e3)
  }

  test("GH-912: an unrelated edit keeps both range hyperlinks beside a cell hyperlink") {
    checkMixedLinks(b7, d2e3, SourceLink("G5:H8", "otherRange", "https://example.net/other"))
  }

  test("GH-912: a range holding rIdHL1 cannot redirect a generated cell hyperlink") {
    checkMixedLinks(b7, d2e3.copy(id = "rIdHL1"))
  }

  private def resolvedLinks(
    xml: String,
    rels: String,
    branch: Option[String] = None
  ): List[(String, String)] =
    val root = XmlSecurity.parseSafe(xml, "worksheet").fold(e => fail(e.toString), identity)
    val relRoot = XmlSecurity.parseSafe(rels, "relationships").fold(e => fail(e.toString), identity)
    val entries = relRoot \ "Relationship"
    val ids = entries.map(_ \@ "Id")
    assertEquals(ids.distinct.size, ids.size, "relationship IDs must be unique")
    val targets = entries.collect {
      case e if (e \@ "Type") == XmlUtil.relTypeHyperlink && (e \@ "TargetMode") == "External" =>
        (e \@ "Id") -> (e \@ "Target")
    }.toMap
    val container = root \ "hyperlinks"
    assertEquals(container.size, 1)
    val branchLinks = branch.toList.flatMap { label =>
      val alternate =
        (container \ "AlternateContent").headOption.getOrElse(fail("missing MC wrapper"))
      assertEquals(alternate.namespace, XmlUtil.nsMarkupCompatibility)
      val selected = (alternate \ label).headOption.getOrElse(fail(s"missing MC $label"))
      if label == "Choice" then
        assertEquals(selected.getNamespace(selected \@ "Requires"), XmlUtil.nsRelationships)
      (selected \ "hyperlink").toList
    }
    val links = ((container \ "hyperlink").toList ++ branchLinks).map { e =>
      val id =
        e.attribute(XmlUtil.nsRelationships, "id").map(_.text).getOrElse(fail("missing r:id"))
      (e \@ "ref", id)
    }
    assertEquals(
      links.map(_._2).toSet,
      targets.keySet,
      "unused or unresolved hyperlink relationship"
    )
    links.map((ref, id) => ref -> targets(id))

  private def checkRangeOverrides(sourceRef: String, inAlternateContent: Boolean = false)(
    overrides: (String, String)*
  ): Unit =
    val oldUrl = "https://example.net/"
    val (_, seed) = readFromFile(Cell(ref"D2", CellValue.Text("Range link")).withHyperlink(oldUrl))
    val sheetPart = "xl/worksheets/sheet1.xml"
    val relsPart = "xl/worksheets/_rels/sheet1.xml.rels"
    val branches = if inAlternateContent then List(Some("Choice"), Some("Fallback")) else List(None)
    val src = patchPart(seed, sheetPart) { xml =>
      val ranged = xml.replace("""ref="D2"""", s"""ref="$sourceRef"""")
      if !inAlternateContent then ranged
      else
        ranged.replaceFirst(
          "(<hyperlink [^>]*/>)",
          s"""<mc:AlternateContent xmlns:mc="${XmlUtil.nsMarkupCompatibility}" xmlns:feature="${XmlUtil.nsRelationships}"><mc:Choice Requires="feature">$$1</mc:Choice><mc:Fallback>$$1</mc:Fallback></mc:AlternateContent>"""
        )
    }
    def cells(reference: String): List[String] =
      val range = com.tjclp.xl.addressing.CellRange.parse(reference).fold(e => fail(e), identity)
      assert(range.cellCount <= 9, reference)
      range.cellsRowMajor.map(_.toA1).toList
    try
      val zip = new ZipFile(src.toFile)
      try
        branches.foreach { branch =>
          assertEquals(
            resolvedLinks(entryText(zip, sheetPart), entryText(zip, relsPart), branch),
            List(sourceRef -> oldUrl)
          )
        }
      finally zip.close()
      val wb = XlsxReader.read(src).fold(e => fail(e.message), identity)
      val before = wb("Summary").fold(e => fail(e.message), identity)
      assertEquals(before(ref"D2").hyperlink, None)
      val edited = overrides.foldLeft(before) { case (sheet, (ref, target)) =>
        val cell = sheet(ref).fold(e => fail(e.message), identity)
        sheet.put(cell.withHyperlink(target))
      }
      val expected = cells(sourceRef).map(ref => ref -> overrides.toMap.getOrElse(ref, oldUrl))
      for backend <- Seq(XmlBackend.ScalaXml, XmlBackend.SaxStax) do
        val (after, xml, rels) = writeAndReopen(wb.put(edited), WriterConfig(backend = backend))
        overrides.foreach { (ref, target) =>
          assertEquals(
            after(ref).fold(e => fail(e.message), identity).hyperlink,
            Some(Hyperlink(target))
          )
        }
        val coverage = branches.map { branch =>
          val links =
            resolvedLinks(xml, rels.getOrElse(fail("missing sheet relationships")), branch)
          branch -> links.flatMap((reference, target) => cells(reference).map(_ -> target)).sorted
        }
        assertEquals(coverage, branches.map(_ -> expected.sorted), backend.toString)
    finally Files.deleteIfExists(src)

  test("GH-925: overriding the corner of a range changes only that cell's target") {
    checkRangeOverrides("D2:E3")("D2" -> "https://example.com/")
  }

  test("GH-925: overriding an interior cell keeps the original target on all sides") {
    checkRangeOverrides("D2:F4")("E3" -> "https://example.com/")
  }

  test("GH-925: multiple overrides inside one range each have one target") {
    checkRangeOverrides("D2:F4")(
      "D2" -> "https://example.com/first",
      "F2" -> "https://example.com/second",
      "E4" -> "https://example.com/third"
    )
  }

  test("GH-925: replacing every cell drops the old range and its relationship") {
    checkRangeOverrides("D2:E3")(
      "D2" -> "https://example.com/1",
      "E2" -> "https://example.com/2",
      "D3" -> "https://example.com/3",
      "E3" -> "https://example.com/4"
    )
  }

  test("GH-925: a range without overrides keeps its coverage and relationship") {
    checkRangeOverrides("D2:E3")()
  }

  test("GH-925: a range override excludes the cell from both MC branches") {
    checkRangeOverrides("D2:E3", inAlternateContent = true)("D2" -> "https://example.com/")
  }
