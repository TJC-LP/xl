package com.tjclp.xl.ooxml

import java.nio.file.Files

import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.api.*
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.macros.ref
import com.tjclp.xl.unsafe.*
import com.tjclp.xl.workbooks.DefinedName
import munit.FunSuite

/**
 * GH-236: named ranges (DefinedName) were a read-only model — populated on read but never
 * serialized. These tests prove they now round-trip via both the fresh-write (fromDomain) and the
 * surgical (preserve-on-cell-edit) paths.
 */
class DefinedNameRoundTripSpec extends FunSuite:

  test("GH-236: programmatically authored named range serializes and round-trips") {
    val wb =
      Workbook(Sheet("Sheet1").put(ref"A1" -> 1)).withDefinedName("MyRange", "Sheet1!$A$1:$A$10")
    val out = Files.createTempFile("named-fresh", ".xlsx")
    XlsxWriter.write(wb, out).fold(e => fail(s"write failed: $e"), identity)

    val reread = XlsxReader.read(out).fold(e => fail(s"read failed: $e"), identity)
    val names = reread.metadata.definedNames
    assertEquals(names.map(_.name), Vector("MyRange"))
    assertEquals(names.headOption.map(_.formula), Some("Sheet1!$A$1:$A$10"))
    Files.deleteIfExists(out)
  }

  test("GH-236: surgical write (cell edit) preserves existing named ranges") {
    // Author a file with a defined name, then read it, edit an unrelated cell, and write back.
    val wb0 = Workbook(Sheet("Sheet1").put(ref"A1" -> 1)).withDefinedName("TaxRate", "0.08")
    val src = Files.createTempFile("named-src", ".xlsx")
    XlsxWriter.write(wb0, src).fold(e => fail(s"seed write failed: $e"), identity)

    val edited = for
      wb <- XlsxReader.read(src)
      sheet <- wb("Sheet1")
      updated = sheet.put(ref"B1" -> 2)
    yield wb.put(updated)
    val wb1 = edited.fold(e => fail(s"edit failed: $e"), identity)

    val out = Files.createTempFile("named-out", ".xlsx")
    XlsxWriter.write(wb1, out).fold(e => fail(s"write failed: $e"), identity)

    val reread = XlsxReader.read(out).fold(e => fail(s"reread failed: $e"), identity)
    assertEquals(reread.metadata.definedNames.map(_.name), Vector("TaxRate"))
    assertEquals(reread.metadata.definedNames.headOption.map(_.formula), Some("0.08"))
    Files.deleteIfExists(src)
    Files.deleteIfExists(out)
  }

  test("GH-649: a Name Manager comment with a line break round-trips on both backends") {
    val base = Workbook(Sheet("Sheet1").put(ref"A1" -> 1))
    val comment = "Line one\nLine two\ttabbed"
    val wb0 = base.copy(metadata =
      base.metadata.copy(definedNames =
        Vector(DefinedName("Rate", "0.08", comment = Some(comment)))
      )
    )
    List(
      "dom" -> com.tjclp.xl.ooxml.writer.WriterConfig.scalaXml,
      "sax" -> com.tjclp.xl.ooxml.writer.WriterConfig.saxStax
    ).foreach { case (label, config) =>
      val out = Files.createTempFile(s"named-comment-$label", ".xlsx")
      out.toFile.deleteOnExit()
      XlsxWriter.writeWith(wb0, out, config).fold(e => fail(s"$label write failed: $e"), identity)
      val reread = XlsxReader.read(out).fold(e => fail(s"$label reread failed: $e"), identity)
      assertEquals(reread.metadata.definedNames.headOption.flatMap(_.comment), Some(comment), label)
    }
  }

  test("GH-236: removeDefinedName drops the name on write") {
    val wb = Workbook(Sheet("Sheet1").put(ref"A1" -> 1))
      .withDefinedName("Temp", "Sheet1!$A$1")
      .removeDefinedName("Temp")
    val out = Files.createTempFile("named-rm", ".xlsx")
    XlsxWriter.write(wb, out).fold(e => fail(s"write failed: $e"), identity)
    val reread = XlsxReader.read(out).fold(e => fail(s"read failed: $e"), identity)
    assertEquals(reread.metadata.definedNames, Vector.empty)
    Files.deleteIfExists(out)
  }

  test("GH-462: a name authored with the scoped overload round-trips its localSheetId") {
    val wb = Workbook(Sheet("Cover").put(ref"A1" -> 1), Sheet("Data").put(ref"B2" -> 2))
      .withDefinedName("DataLocal", "Data!$B$2", SheetName.unsafe("Data"))
      .fold(e => fail(e.message), identity)
    val out = Files.createTempFile("named-scoped-overload", ".xlsx")
    XlsxWriter.write(wb, out).fold(e => fail(s"write failed: $e"), identity)
    val reread = XlsxReader.read(out).fold(e => fail(s"read failed: $e"), identity)
    assertEquals(
      reread.metadata.definedNames.map(dn => (dn.name, dn.formula, dn.localSheetId)),
      Vector(("DataLocal", "Data!$B$2", Some(1)))
    )
    Files.deleteIfExists(out)
  }

  test(
    "GH-462: a foreign writer's `_XLNM.PRINT_AREA` is lifted into PageSetup case-insensitively"
  ) {
    // Excel names are case-insensitive identifiers: this IS the sheet's print area to Excel, so
    // the read side lifts it like the canonical spelling and the writer re-derives it canonically.
    val base = Workbook(Sheet("Sheet1").put(ref"A1" -> 1))
    val wb = base.copy(metadata =
      base.metadata.copy(definedNames =
        Vector(DefinedName("_XLNM.PRINT_AREA", "Sheet1!$A$1:$B$2", localSheetId = Some(0)))
      )
    )
    val out = Files.createTempFile("named-print-area-upper", ".xlsx")
    XlsxWriter.write(wb, out).fold(e => fail(s"write failed: $e"), identity)
    val reread = XlsxReader.read(out).fold(e => fail(s"read failed: $e"), identity)
    val sheet = reread("Sheet1").fold(e => fail(s"sheet missing: $e"), identity)
    assertEquals(sheet.pageSetup.flatMap(_.printArea), Some(ref"A1:B2"))
    assertEquals(reread.metadata.definedNames, Vector.empty, "lifted, not left verbatim")
    Files.deleteIfExists(out)
  }

  // ===== GH-434: sheet-scoped names must keep their sheet across order mutations =====

  /** Cover (idx 0) + Data (idx 1), Data carrying a sheet-scoped name, written and read back. */
  private def scopedFixture(label: String): (java.nio.file.Path, Workbook) =
    val base = Workbook(Sheet("Cover").put(ref"A1" -> 1), Sheet("Data").put(ref"B2" -> 2))
    val wb0 = base.copy(metadata =
      base.metadata.copy(definedNames =
        Vector(DefinedName("DataLocal", "Data!$B$2", localSheetId = Some(1)))
      )
    )
    val src = Files.createTempFile(s"named-scope-$label", ".xlsx")
    src.toFile.deleteOnExit()
    XlsxWriter.write(wb0, src).fold(e => fail(s"seed write failed: $e"), identity)
    val read = XlsxReader.read(src).fold(e => fail(s"seed read failed: $e"), identity)
    assertEquals(
      read.metadata.definedNames.map(dn => (dn.name, dn.localSheetId)),
      Vector(("DataLocal", Some(1))),
      "fixture sanity: the scoped name must survive the seed round-trip"
    )
    (src, read)

  private def writeReread(wb: Workbook, label: String): Workbook =
    val out = Files.createTempFile(s"named-scope-$label-out", ".xlsx")
    out.toFile.deleteOnExit()
    XlsxWriter.write(wb, out).fold(e => fail(s"write failed: $e"), identity)
    XlsxReader.read(out).fold(e => fail(s"reread failed: $e"), identity)

  test("GH-434: sheet-scoped name still targets its sheet after removing the sheet above it") {
    val (_, wb) = scopedFixture("remove")
    val removed = wb.remove(SheetName.unsafe("Cover")).fold(e => fail(e.message), identity)
    val reread = writeReread(removed, "remove")
    assertEquals(reread.sheetNames.map(_.value), Vector("Data"))
    assertEquals(
      reread.metadata.definedNames.map(dn => (dn.name, dn.localSheetId, dn.formula)),
      Vector(("DataLocal", Some(0), "Data!$B$2")),
      "the name must follow Data from index 1 to index 0"
    )
  }

  test("GH-434: sheet-scoped name still targets its sheet after insertAt-middle") {
    val (_, wb) = scopedFixture("insert")
    val inserted =
      wb.insertAt(1, Sheet("Inserted")).fold(e => fail(e.message), identity)
    val reread = writeReread(inserted, "insert")
    assertEquals(reread.sheetNames.map(_.value), Vector("Cover", "Inserted", "Data"))
    assertEquals(
      reread.metadata.definedNames.map(dn => (dn.name, dn.localSheetId)),
      Vector(("DataLocal", Some(2))),
      "the name must follow Data from index 1 to index 2"
    )
  }

  test("GH-434: sheet-scoped name still targets its sheet after reorder") {
    val (_, wb) = scopedFixture("reorder")
    val reordered = wb
      .reorder(Vector(SheetName.unsafe("Data"), SheetName.unsafe("Cover")))
      .fold(e => fail(e.message), identity)
    val reread = writeReread(reordered, "reorder")
    assertEquals(reread.sheetNames.map(_.value), Vector("Data", "Cover"))
    assertEquals(
      reread.metadata.definedNames.map(dn => (dn.name, dn.localSheetId)),
      Vector(("DataLocal", Some(0))),
      "the name must follow Data from index 1 to index 0 across the reorder"
    )
  }

  // ===== GH-696: a regenerated <definedNames> keeps every attribute =====

  /** Excel's spelling of names carrying every CT_DefinedName attribute, plus an unknown one. */
  private val attributeRichNames = Vector(
    """<definedName name="Macro1" function="1" vbProcedure="1" shortcutKey="m">Module1.Macro1</definedName>""",
    """<definedName name="Xlm1" comment="note" customMenu="Run it" description="An XLM macro" help="Help topic" statusBar="Running" hidden="1" function="1" xlm="1" functionGroupId="14">Macro1!$A$1</definedName>""",
    """<definedName name="Param" publishToServer="1" workbookParameter="1">Sheet1!$A$1</definedName>""",
    """<definedName name="_xlnm.Print_Area" description="kept, not lifted" localSheetId="0">Sheet1!$A$1:$B$2</definedName>""",
    """<definedName name="Odd" futureAttr="x">0.5</definedName>"""
  )

  /** A workbook xl wrote, its `<definedNames>` swapped for [[attributeRichNames]]. */
  private def attributeRichSource(): java.nio.file.Path =
    import java.util.zip.{ZipEntry, ZipFile, ZipOutputStream}
    val seed = Files.createTempFile("named-attrs-seed", ".xlsx")
    seed.toFile.deleteOnExit()
    val wb = Workbook(Sheet("Sheet1").put(ref"A1" -> 1)).withDefinedName("Placeholder", "1")
    XlsxWriter.write(wb, seed).fold(e => fail(s"seed write failed: $e"), identity)
    val src = Files.createTempFile("named-attrs-src", ".xlsx")
    src.toFile.deleteOnExit()
    val zin = new ZipFile(seed.toFile)
    val zos = new ZipOutputStream(Files.newOutputStream(src))
    try
      zin.entries().asIterator().forEachRemaining { entry =>
        val bytes = zin.getInputStream(entry).readAllBytes()
        val out =
          if entry.getName != "xl/workbook.xml" then bytes
          else
            val xml = new String(bytes, "UTF-8")
            val swapped = xml.replaceFirst(
              "(?s)<definedNames>.*</definedNames>",
              java.util.regex.Matcher.quoteReplacement(
                attributeRichNames.mkString("<definedNames>", "", "</definedNames>")
              )
            )
            assert(swapped != xml, "fixture sanity: the seed carries a <definedNames>")
            swapped.getBytes("UTF-8")
        zos.putNextEntry(new ZipEntry(entry.getName))
        zos.write(out)
        zos.closeEntry()
      }
    finally
      zos.close()
      zin.close()
    src

  private def workbookXml(path: java.nio.file.Path): String =
    val zip = new java.util.zip.ZipFile(path.toFile)
    try
      new String(zip.getInputStream(zip.getEntry("xl/workbook.xml")).readAllBytes(), "UTF-8")
    finally zip.close()

  private val expectedRichModel = Vector(
    DefinedName(
      "Macro1",
      "Module1.Macro1",
      function = true,
      vbProcedure = true,
      shortcutKey = Some("m")
    ),
    DefinedName(
      "Xlm1",
      "Macro1!$A$1",
      hidden = true,
      comment = Some("note"),
      customMenu = Some("Run it"),
      description = Some("An XLM macro"),
      help = Some("Help topic"),
      statusBar = Some("Running"),
      function = true,
      xlm = true,
      functionGroupId = Some(14)
    ),
    DefinedName("Param", "Sheet1!$A$1", publishToServer = true, workbookParameter = true),
    DefinedName(
      "_xlnm.Print_Area",
      "Sheet1!$A$1:$B$2",
      localSheetId = Some(0),
      description = Some("kept, not lifted")
    ),
    DefinedName("Odd", "0.5", otherAttributes = Vector("futureAttr" -> "x"))
  )

  test("GH-696: both readers carry every <definedName> attribute") {
    val src = attributeRichSource()
    val wb = XlsxReader.read(src).fold(e => fail(s"read failed: $e"), identity)
    assertEquals(wb.metadata.definedNames, expectedRichModel)
    val sheet = wb("Sheet1").fold(e => fail(s"sheet missing: $e"), identity)
    assertEquals(
      sheet.pageSetup.flatMap(_.printArea),
      None,
      "an attributed print area stays a name"
    )
    assertEquals(
      metadata.WorkbookMetadataReader.readDefinedNames(src).fold(e => fail(e.message), identity),
      expectedRichModel
    )
  }

  /** Each `<definedName>`'s attributes (as a map) and text, in document order. */
  private def definedNameShapes(xml: String): Vector[(Map[String, String], String)] =
    (scala.xml.XML.loadString(xml) \\ "definedName").collect { case e: scala.xml.Elem =>
      (e.attributes.asAttrMap, e.text)
    }.toVector

  test("GH-696: an unrelated name edit keeps every other <definedName> intact on both backends") {
    val src = attributeRichSource()
    val wb = XlsxReader.read(src).fold(e => fail(s"read failed: $e"), identity)
    val edited = wb.withDefinedName("Added", "Sheet1!$A$1")
    val added = """<definedName name="Added">Sheet1!$A$1</definedName>"""
    List(
      "dom" -> com.tjclp.xl.ooxml.writer.WriterConfig.scalaXml,
      "sax" -> com.tjclp.xl.ooxml.writer.WriterConfig.saxStax
    ).foreach { case (label, config) =>
      val out = Files.createTempFile(s"named-attrs-$label", ".xlsx")
      out.toFile.deleteOnExit()
      XlsxWriter
        .writeWith(edited, out, config)
        .fold(e => fail(s"$label write failed: $e"), identity)
      val xml = workbookXml(out)
      // the SAX backend sorts every element's attributes alphabetically (SaxSupport), so it keeps
      // each name's attributes and text; the DOM backend also keeps Excel's bytes
      assertEquals(
        definedNameShapes(xml),
        definedNameShapes(
          (attributeRichNames :+ added).mkString("<definedNames>", "", "</definedNames>")
        ),
        label
      )
      if label == "dom" then
        (attributeRichNames :+ added).foreach(dn => assert(xml.contains(dn), s"lost $dn in:\n$xml"))
      val reread = XlsxReader.read(out).fold(e => fail(s"$label reread failed: $e"), identity)
      assertEquals(
        reread.metadata.definedNames,
        expectedRichModel :+ DefinedName("Added", "Sheet1!$A$1"),
        label
      )
    }
  }

  test("GH-696: parse(build(names)) is the identity on a fully attributed name") {
    val full = DefinedName(
      "Full",
      "SUM(Sheet1!$A$1:$A$3)",
      localSheetId = Some(2),
      hidden = true,
      comment = Some("c"),
      customMenu = Some("m"),
      description = Some("d"),
      help = Some("h"),
      statusBar = Some("s"),
      function = true,
      vbProcedure = true,
      xlm = true,
      functionGroupId = Some(4294967295L),
      shortcutKey = Some("K"),
      publishToServer = true,
      workbookParameter = true,
      otherAttributes = Vector("zeta" -> "1", "alpha" -> "2")
    )
    val names = Vector(full, DefinedName("Plain", "1"))
    assertEquals(OoxmlWorkbook.parseDefinedNames(OoxmlWorkbook.buildDefinedNames(names)), names)
  }

  test("GH-696: xsd:boolean \"true\" reads as set") {
    val elem =
      <definedNames><definedName name="T" hidden="true" function="true">1</definedName></definedNames>
    val parsed = OoxmlWorkbook.parseDefinedNames(Some(elem))
    assertEquals(parsed.map(dn => (dn.hidden, dn.function)), Vector((true, true)))
  }

  test("GH-696: a \"0\" flag reads as unset and is not written back") {
    val elem =
      <definedNames><definedName name="Z" hidden="0" function="false">1</definedName></definedNames>
    val parsed = OoxmlWorkbook.parseDefinedNames(Some(elem))
    assertEquals(parsed, Vector(DefinedName("Z", "1")))
    assertEquals(
      OoxmlWorkbook.buildDefinedNames(parsed).map(_.toString),
      Some("""<definedNames><definedName name="Z">1</definedName></definedNames>""")
    )
  }

  test("GH-696: hand-built otherAttributes never repeat a typed key, bind no prefix, dedupe") {
    val dn = DefinedName(
      "H",
      "1",
      hidden = true,
      otherAttributes =
        Vector("hidden" -> "0", "x:foo" -> "1", "extra" -> "a", "extra" -> "b", "name" -> "Other")
    )
    val built = OoxmlWorkbook.buildDefinedNames(Vector(dn)).getOrElse(fail("no element"))
    assertEquals(
      built.toString,
      """<definedNames><definedName name="H" hidden="1" extra="a">1</definedName></definedNames>"""
    )
    // and it is well-formed XML that reads back to the typed fields plus the one passthrough
    assertEquals(
      OoxmlWorkbook.parseDefinedNames(Some(scala.xml.XML.loadString(built.toString))),
      Vector(DefinedName("H", "1", hidden = true, otherAttributes = Vector("extra" -> "a")))
    )
  }

  test("GH-696: functionGroupId reads its whole xsd:unsignedInt range, nothing outside it") {
    def groupId(raw: String): Option[Long] =
      OoxmlWorkbook
        .parseDefinedNames(Some(<definedNames><definedName name="F" functionGroupId={
          raw
        }>1</definedName></definedNames>))
        .headOption
        .flatMap(_.functionGroupId)
    assertEquals(groupId("14"), Some(14L))
    assertEquals(groupId("2147483648"), Some(2147483648L))
    assertEquals(groupId("4294967295"), Some(4294967295L))
    assertEquals(groupId("4294967296"), None)
    assertEquals(groupId("-1"), None)
    val built = OoxmlWorkbook.buildDefinedNames(
      Vector(DefinedName("F", "1", functionGroupId = Some(4294967295L)))
    )
    assert(built.exists(_.toString.contains("""functionGroupId="4294967295"""")), built.toString)
  }
