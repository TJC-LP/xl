package com.tjclp.xl.ooxml

import java.nio.file.Files

import munit.FunSuite

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.api.*
import com.tjclp.xl.cells.{ArrayMode, CellValue, FormulaKind}
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ops.FormulaSupport
import com.tjclp.xl.ooxml.writer.{WriterConfig, XmlBackend}

/**
 * GH-714: a dynamic-array record (`cm` on the cell, XLDAPR in `xl/metadata.xml`) survives a
 * regenerated sheet. Without `cm` Excel reads the record as a legacy CSE array `{=SORT(..)}`.
 */
class DynamicArrayPreservationSpec extends FunSuite:
  import DynamicArrayFixtures.*

  private val anchorXml =
    """<c r="B1" t="n" cm="1"><f t="array" ref="B1:B5">_xlfn._xlws.SORT(A1:A5)</f><v>1</v></c>"""

  private val backends = List(
    "ScalaXml" -> WriterConfig(backend = XmlBackend.ScalaXml),
    "SaxStax" -> WriterConfig(backend = XmlBackend.SaxStax)
  )

  private def read(bytes: Array[Byte]): Workbook =
    XlsxReader.readFromBytes(bytes).fold(err => fail(err.message), identity)

  private def write(wb: Workbook, config: WriterConfig = WriterConfig()): Array[Byte] =
    XlsxWriter.writeToBytes(wb, config).fold(err => fail(err.message), identity)

  /** The write's bytes and the report it made from the plan it used. */
  private def reported(
    wb: Workbook,
    config: WriterConfig = WriterConfig()
  ): (Array[Byte], XlsxWriter.WriteReport) =
    XlsxWriter.writeToBytesWithReport(wb, config).fold(err => fail(err.message), identity)

  private def downgradeOf(wb: Workbook): Option[String] = reported(wb)._2.cellMetadataDowngrade

  private def sheet1(bytes: Array[Byte]): String =
    entryText(bytes, "xl/worksheets/sheet1.xml").getOrElse(fail("no sheet1"))

  private def occurrences(bytes: Array[Byte], part: String, needle: String): Int =
    entryText(bytes, part).getOrElse(fail(s"no $part")).split(needle, -1).length - 1

  private def assertWiredOnce(out: Array[Byte]): Unit =
    assertEquals(entries(out).count(_._1 == "xl/metadata.xml"), 1)
    assertEquals(occurrences(out, "[Content_Types].xml", "/xl/metadata.xml"), 1)
    assertEquals(
      occurrences(out, "xl/_rels/workbook.xml.rels", "relationships/sheetMetadata"),
      1
    )

  private def dirty(wb: Workbook): Workbook =
    val s = wb.sheets.headOption.getOrElse(fail("no sheet"))
    wb.put(s.put(ref"D1", CellValue.Number(5)))

  private def kindAt(wb: Workbook, at: ARef): Option[FormulaKind] =
    wb.sheets.headOption.flatMap(_.cells.get(at)).map(_.value).collect {
      case CellValue.Formula(_, _, kind) => kind
    }

  private val sortRange = CellRange(ref"B1", ref"B5")

  // This fixture's sole formula uses an absolute reference; no evaluator is needed to shift it.
  private val anchoredFormulaSupport: FormulaSupport = new FormulaSupport:
    export FormulaSupport.textOnly.{
      validate,
      insertRows,
      deleteRows,
      insertCols,
      deleteCols,
      renameSheet
    }
    def shift(formula: String, colDelta: Int, rowDelta: Int): XLResult[String] =
      if formula == "SEQUENCE($A$1)" then Right(formula)
      else Left(FormulaSupport.unsupported("shift"))

  backends.foreach { case (name, config) =>
    test(s"whole dynamic spill copies preserve their owned cells through OOXML ($name)") {
      val kind = FormulaKind.dynamicArray(ref"B1:B3")
      val source = Sheet("Sheet1")
        .put(ref"A1", CellValue.Number(3))
        .put(ref"B1", CellValue.Formula("SEQUENCE($A$1)", Some(CellValue.Number(1)), kind))
        .put(ref"B2", CellValue.Number(2))
        .put(ref"B3", CellValue.Number(3))
      val sourceWorkbook = read(write(Workbook(source), config))
      val sourceSheet = sourceWorkbook.sheets.headOption.getOrElse(fail("no source sheet"))
      val copied = sourceSheet
        .copyRange(ref"B1:B3", ref"D1:D3", valuesOnly = false)(using anchoredFormulaSupport)
        .fold(e => fail(e.message), identity)
      val out = write(sourceWorkbook.put(copied), config)
      val back = read(out)
      val targetSheet = back.sheets.headOption.getOrElse(fail("no target sheet"))
      assertEquals(kindAt(back, ref"D1"), Some(FormulaKind.dynamicArray(ref"D1:D3")))
      assertEquals(kindAt(back, ref"B1"), Some(kind))
      assertEquals(targetSheet(ref"D2").value, CellValue.Number(2))
      assertEquals(targetSheet(ref"D3").value, CellValue.Number(3))
      assertEquals(
        targetSheet(ref"D1").value,
        CellValue.Formula("SEQUENCE($A$1)", None, FormulaKind.dynamicArray(ref"D1:D3"))
      )
      val xml =
        XmlSecurity.parseSafe(sheet1(out), "sheet1.xml").fold(e => fail(e.message), identity)
      val anchor = (xml \\ "c").find(n => (n \ "@r").text == "D1").getOrElse(fail("no D1"))
      assertEquals((anchor \ "@cm").text, "1")
      assertEquals((anchor \ "f" \ "@ref").text, "D1:D3")
      assertWiredOnce(out)
    }
  }

  backends.foreach { case (name, config) =>
    test(s"put D1 keeps the anchor's cm and the metadata part ($name)") {
      val out = write(dirty(read(xlsx())), config)
      assert(sheet1(out).contains(anchorXml), sheet1(out))
      assertEquals(entryText(out, "xl/metadata.xml"), Some(metadataXml))
      assertWiredOnce(out)
    }
  }

  test("the reader resolves cm=1 to a dynamic array") {
    assertEquals(kindAt(read(xlsx()), ref"B1"), Some(FormulaKind.dynamicArray(sortRange)))
  }

  test("a later dynamic record survives a regenerated sheet") {
    val wb = read(xlsx(parts(metadata = Some(multiRecordMetadataXml))))
    assertEquals(kindAt(wb, ref"B1"), Some(FormulaKind.dynamicArray(sortRange)))
    val out = write(dirty(wb))
    // the shared block (OTHER + XLDAPR) is kept; the anchor is written through a block of its own
    val part = entryText(out, "xl/metadata.xml").getOrElse(fail("no metadata part"))
    assert(
      part.contains("""<bk><rc t="1" v="0"/><rc t="2" v="0"/></bk><bk><rc t="2" v="0"/></bk>"""),
      part
    )
    assert(sheet1(out).contains("""<c r="B1" t="n" cm="2">"""), sheet1(out))
    assertEquals(kindAt(read(out), ref"B1"), Some(FormulaKind.dynamicArray(sortRange)))
  }

  for
    metadataPath <- List("xl/custom/state.xml", "xl/custom/state.bin")
    metadataFirst <- List(false, true)
  do
    test(s"metadata resolves by relationship at $metadataPath, first=$metadataFirst") {
      val relocated = relocatedMetadataParts(metadataPath)
      val bytes = xlsx(if metadataFirst then relocated.reverse else relocated)
      val file = Files.createTempFile("xl-metadata-path-", ".xlsx")
      try
        Files.write(file, bytes)
        val fromFile = XlsxReader.read(file).fold(e => fail(e.message), identity)
        List(read(bytes), fromFile).foreach { wb =>
          assertEquals(kindAt(wb, ref"B1"), Some(FormulaKind.dynamicArray(sortRange)))
          backends.foreach { case (_, config) =>
            val out = write(dirty(wb), config)
            assertEquals(entryText(out, metadataPath), Some(metadataXml))
            assertEquals(entries(out).count(_._1 == metadataPath), 1)
            assertEquals(entryText(out, "xl/metadata.xml"), None)
            assertEquals(kindAt(read(out), ref"B1"), Some(FormulaKind.dynamicArray(sortRange)))
          }
        }
      finally Files.deleteIfExists(file)
    }

  /**
   * A STORED package written by hand: every `locals` entry gets a local header in order, but the
   * central directory lists only the `central` indices — so a name can have a second local entry
   * the sequential scan reads and the directory does not point at.
   */
  private def craftedZip(locals: Vector[(String, String)], central: Vector[Int]): Array[Byte] =
    import java.nio.{ByteBuffer, ByteOrder}
    import java.nio.charset.StandardCharsets.UTF_8
    def le(size: Int)(fill: ByteBuffer => Unit): Array[Byte] =
      val b = ByteBuffer.allocate(size).order(ByteOrder.LITTLE_ENDIAN)
      fill(b)
      b.array()
    val encoded = locals.map((name, text) => (name.getBytes(UTF_8), text.getBytes(UTF_8)))
    def crcOf(data: Array[Byte]): Int =
      val crc = new java.util.zip.CRC32()
      crc.update(data)
      crc.getValue.toInt
    val (localBytes, offsets) =
      encoded.foldLeft((Vector.empty[Array[Byte]], Vector.empty[Int])) {
        case ((acc, offs), (name, data)) =>
          val header = le(30) { b =>
            b.putInt(0x04034b50).putShort(20.toShort).putShort(0.toShort).putShort(0.toShort)
            b.putShort(0.toShort).putShort(0x21.toShort).putInt(crcOf(data))
            b.putInt(data.length).putInt(data.length).putShort(name.length.toShort)
            b.putShort(0.toShort)
          }
          (acc :+ (header ++ name ++ data), offs :+ acc.map(_.length).sum)
      }
    val cdStart = localBytes.map(_.length).sum
    val cd = central.map { i =>
      val (name, data) = encoded(i)
      le(46) { b =>
        b.putInt(0x02014b50).putShort(20.toShort).putShort(20.toShort).putShort(0.toShort)
        b.putShort(0.toShort).putShort(0.toShort).putShort(0x21.toShort).putInt(crcOf(data))
        b.putInt(data.length).putInt(data.length).putShort(name.length.toShort)
        b.putShort(0.toShort).putShort(0.toShort).putShort(0.toShort).putShort(0.toShort)
        b.putInt(0).putInt(offsets(i))
      } ++ name
    }
    val cdSize = cd.map(_.length).sum
    val end = le(22) { b =>
      b.putInt(0x06054b50).putShort(0.toShort).putShort(0.toShort)
      b.putShort(central.size.toShort).putShort(central.size.toShort)
      b.putInt(cdSize).putInt(cdStart).putShort(0.toShort)
    }
    (localBytes ++ cd :+ end).foldLeft(Array.emptyByteArray)(_ ++ _)

  test("a reread metadata part must be the one the scan read (crafted second local entry)") {
    val path = "xl/custom/state.xml"
    val relocated = relocatedMetadataParts(path)
    // the scan meets a second, non-dynamic copy of the part; the directory points at the first
    val forged = metadataXml.replace("fDynamic=\"1\"", "fDynamic=\"0\"")
    val bytes = craftedZip(relocated :+ (path -> forged), relocated.indices.toVector)
    val consistent = craftedZip(relocated, relocated.indices.toVector)
    val file = Files.createTempFile("xl-metadata-crafted-", ".xlsx")
    val clean = Files.createTempFile("xl-metadata-consistent-", ".xlsx")
    try
      Files.write(file, bytes)
      Files.write(clean, consistent)
      def isSecurity(result: XLResult[Workbook]): Boolean = result.left.toOption.exists {
        case XLError.SecurityError(message) => message.contains("differs from the entry")
        case _ => false
      }
      assert(isSecurity(XlsxReader.read(file)), XlsxReader.read(file).toString.take(300))
      assert(isSecurity(XlsxReader.readFromBytes(bytes)))
      // the same hand-written package without the second entry reads, from disk and from bytes
      val fromDisk = XlsxReader.read(clean).fold(e => fail(e.message), identity)
      List(fromDisk, read(consistent)).foreach { wb =>
        assertEquals(kindAt(wb, ref"B1"), Some(FormulaKind.dynamicArray(sortRange)))
      }
    finally
      Files.deleteIfExists(file)
      Files.deleteIfExists(clean)
  }

  test("a renamed metadata part remains subject to the archive size limit") {
    val bytes = xlsx(relocatedMetadataParts())
    val result = XlsxReader.readFromBytes(
      bytes,
      XlsxReader.ReaderConfig(maxUncompressedSize = metadataXml.length.toLong)
    )
    assert(result.left.toOption.exists {
      case XLError.SecurityError(_) => true
      case _ => false
    })
  }

  test("a collapsed-only source part gains the plain variant: both anchors round-trip") {
    val collapsedXml = metadataXml.replace("fCollapsed=\"0\"", "fCollapsed=\"1\"")
    val collapsedB1 =
      FormulaKind.ArrayFormula(sortRange, mode = ArrayMode.Dynamic(collapsed = true))
    val wb = read(xlsx(parts(metadata = Some(collapsedXml))))
    assertEquals(kindAt(wb, ref"B1"), Some(collapsedB1))
    val sheet = wb.sheets.headOption.getOrElse(fail("no sheet"))
    val plainE1 = FormulaKind.dynamicArray(CellRange(ref"E1", ref"E1"))
    val edited = wb.put(sheet.put(ref"E1", CellValue.Formula("SEQUENCE(1)", None, plainE1)))
    backends.foreach { case (_, config) =>
      val out = write(edited, config)
      val back = read(out)
      assertEquals(kindAt(back, ref"B1"), Some(collapsedB1))
      assertEquals(kindAt(back, ref"E1"), Some(plainE1))
      val part = entryText(out, "xl/metadata.xml").getOrElse(fail("no metadata part"))
      assert(part.contains("""<futureMetadata name="XLDAPR" count="2">"""), part)
      assert(sheet1(out).contains("""<c r="B1" t="n" cm="1">"""), sheet1(out))
      assertWiredOnce(out)
    }
  }

  private def withAnchorAt(wb: Workbook, at: ARef): Workbook =
    val sheet = wb.sheets.headOption.getOrElse(fail("no sheet"))
    val kind = FormulaKind.dynamicArray(CellRange(at, at))
    wb.put(sheet.put(at, CellValue.Formula("SEQUENCE(1)", None, kind)))

  test("an unreadable metadata part is kept, its new anchors lose cm, and the writer says so") {
    val broken = "<metadata"
    val wb = read(xlsx(parts(metadata = Some(broken))))
    // nothing regenerated yet, or no dynamic array on a regenerated sheet: nothing is lost
    assertEquals(downgradeOf(wb), None)
    assertEquals(downgradeOf(dirty(wb)), None)
    val edited = withAnchorAt(wb, ref"E1")
    backends.foreach { case (name, config) =>
      val (out, report) = reported(edited, config)
      // the report is read off the plan this very write used
      assertEquals(report.cellMetadataDowngrade, Some("xl/metadata.xml"), name)
      // the part rides through untouched (a second one would orphan whatever else it holds) ...
      assertEquals(entryText(out, "xl/metadata.xml"), Some(broken))
      // ... and the anchor is written in its legacy shape, as the signal warned
      val xml = sheet1(out)
      val e1 =
        xml.substring(xml.indexOf("<c r=\"E1\""), xml.indexOf("</c>", xml.indexOf("<c r=\"E1\"")))
      assert(!e1.contains("cm="), e1)
      assertEquals(
        kindAt(read(out), ref"E1"),
        Some(FormulaKind.ArrayFormula(CellRange(ref"E1", ref"E1")))
      )
    }
  }

  test("a relationship naming a missing part: a part is generated at its target, no downgrade") {
    for (path, source) <- List(
        "xl/metadata.xml" -> parts(),
        "xl/custom/state.xml" -> relocatedMetadataParts("xl/custom/state.xml")
      )
    do
      val dangling = source.filterNot(_._1 == path)
      val wb = read(xlsx(dangling))
      val edited = withAnchorAt(wb, ref"E1")
      backends.foreach { case (name, config) =>
        val (out, report) = reported(edited, config)
        assertEquals(report.cellMetadataDowngrade, None, s"$path ($name)")
        assertEquals(entryText(out, path), Some(metadataXml), s"$path ($name)")
        assertEquals(entries(out).count(_._1 == path), 1, s"$path ($name)")
        if path != "xl/metadata.xml" then
          assertEquals(entryText(out, "xl/metadata.xml"), None, s"$path ($name)")
        assertEquals(occurrences(out, "[Content_Types].xml", s"/$path"), 1, s"$path ($name)")
        assertEquals(
          occurrences(out, "xl/_rels/workbook.xml.rels", "relationships/sheetMetadata"),
          1,
          s"$path ($name)"
        )
        assertEquals(
          kindAt(read(out), ref"E1"),
          Some(FormulaKind.dynamicArray(CellRange(ref"E1", ref"E1")))
        )
      }
  }

  test("recalc on open, pinned: xl adds no calcId; a source's calcPr rides through as it was") {
    def workbookXml(out: Array[Byte]) = entryText(out, "xl/workbook.xml").getOrElse(fail("no wb"))
    // a book xl creates: no <calcPr> at all, so Excel recalculates it (and re-spills) on open
    val kind = FormulaKind.dynamicArray(CellRange(ref"A1", ref"A3"))
    val scratch = Workbook(
      Sheet("S").put(ref"A1", CellValue.Formula("SEQUENCE(3)", Some(CellValue.Number(1)), kind))
    )
    backends.foreach { case (name, config) =>
      val xml = workbookXml(write(scratch, config))
      assert(!xml.contains("calcId"), s"$name: $xml")
      assert(!xml.contains("<calcPr"), s"$name: $xml")
    }
    // a book Excel saved keeps its calcId (and no fullCalcOnLoad is added) when a sheet holding a
    // dynamic extent is regenerated: Excel then trusts the written caches on open, so a spill xl
    // did not re-spill stays as written until Excel recalculates (#737)
    val withCalc = parts().map { (name, xml) =>
      if name == "xl/workbook.xml" then
        name -> xml.replace(
          "</sheets></workbook>",
          """</sheets><calcPr calcId="191029"/></workbook>"""
        )
      else name -> xml
    }
    val edited = dirty(read(xlsx(withCalc)))
    backends.foreach { case (name, config) =>
      val out = write(edited, config)
      assertEquals(kindAt(read(out), ref"B1"), Some(FormulaKind.dynamicArray(sortRange)), name)
      val xml = workbookXml(out)
      assert(xml.contains("""<calcPr calcId="191029""""), s"$name: $xml")
      assert(!xml.contains("fullCalcOnLoad"), s"$name: $xml")
    }
  }

  test("a newly authored variant extends the relationship's renamed metadata part") {
    val path = "xl/custom/state.xml"
    val wb = read(xlsx(relocatedMetadataParts(path)))
    val sheet = wb.sheets.headOption.getOrElse(fail("no sheet"))
    val collapsed = FormulaKind.ArrayFormula(
      CellRange(ref"E1", ref"E1"),
      mode = ArrayMode.Dynamic(collapsed = true)
    )
    val edited = wb.put(sheet.put(ref"E1", CellValue.Formula("SEQUENCE(1)", None, collapsed)))
    backends.foreach { case (_, config) =>
      val out = write(edited, config)
      assertEquals(entries(out).count(_._1 == path), 1)
      assertEquals(entryText(out, "xl/metadata.xml"), None)
      assertEquals(kindAt(read(out), ref"B1"), Some(FormulaKind.dynamicArray(sortRange)))
      assertEquals(kindAt(read(out), ref"E1"), Some(collapsed))
    }
  }

  test("a dangling cm (no metadata part) reads as a legacy array and writes no cm") {
    val wb = read(xlsx(parts(metadata = None)))
    assertEquals(kindAt(wb, ref"B1"), Some(FormulaKind.ArrayFormula(sortRange)))
    val out = write(dirty(wb))
    assert(!sheet1(out).contains("cm="), sheet1(out))
    assertEquals(entryText(out, "xl/metadata.xml"), None)
  }

  private def scratch(kind: FormulaKind): Workbook =
    Workbook(
      Sheet("Sheet1")
        .put(ref"A1", CellValue.Number(3))
        .put(ref"A2", CellValue.Number(1))
        .put(ref"B1", CellValue.Formula("SORT(A1:A2)", Some(CellValue.Number(1)), kind))
        .put(ref"B2", CellValue.Number(3))
    )

  backends.foreach { case (name, config) =>
    test(s"a scratch dynamic array ships the generated part, Override and rel ($name)") {
      val kind = FormulaKind.dynamicArray(CellRange(ref"B1", ref"B2"))
      val wb = scratch(kind)
      val out = write(wb, config)
      assertEquals(entryText(out, "xl/metadata.xml"), Some(metadataXml))
      assert(
        sheet1(out).contains("""<c r="B1" t="n" cm="1"><f t="array" ref="B1:B2">"""),
        sheet1(out)
      )
      val ct = entryText(out, "[Content_Types].xml").getOrElse(fail("no CT"))
      assert(
        ct.contains(
          """ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml""""
        ),
        ct
      )
      assert(ct.contains("""PartName="/xl/metadata.xml""""), ct)
      val rels = entryText(out, "xl/_rels/workbook.xml.rels").getOrElse(fail("no rels"))
      assert(rels.contains("relationships/sheetMetadata\""), rels)
      assert(rels.contains("Target=\"metadata.xml\""), rels)
      assertWiredOnce(out)
      // deterministic, and it reads back as what was written
      assertEquals(write(wb, config).toVector, out.toVector)
      assertEquals(kindAt(read(out), ref"B1"), Some(kind))
    }

    test(s"a scratch book without dynamic arrays ships no metadata part ($name)") {
      val out = write(scratch(FormulaKind.ArrayFormula(CellRange(ref"B1", ref"B2"))), config)
      assertEquals(entryText(out, "xl/metadata.xml"), None)
      assert(!sheet1(out).contains("cm="), sheet1(out))
      val rels = entryText(out, "xl/_rels/workbook.xml.rels").getOrElse(fail("no rels"))
      assert(!rels.contains("sheetMetadata"), rels)
    }

    test(s"a collapsed dynamic array round-trips ($name)") {
      val kind = FormulaKind.ArrayFormula(
        CellRange(ref"B1", ref"B2"),
        mode = ArrayMode.Dynamic(collapsed = true)
      )
      val out = write(scratch(kind), config)
      assertEquals(kindAt(read(out), ref"B1"), Some(kind))
    }
  }

  test("a 1x1 dynamic array writes Excel's bare ref") {
    val wb = Workbook(
      Sheet("Sheet1").put(
        ref"C5",
        CellValue.Formula(
          "SUM(A1:A3*B1:B3)",
          Some(CellValue.Number(14)),
          FormulaKind.dynamicArray(CellRange(ref"C5", ref"C5"))
        )
      )
    )
    val out = write(wb)
    assert(sheet1(out).contains("""<c r="C5" t="n" cm="1"><f t="array" ref="C5">"""), sheet1(out))
  }

  test("an images-in-cells metadata part is extended, never rewritten") {
    val rich =
      """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        |<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"><metadataTypes count="1"><metadataType name="XLRICHVALUE" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1"/></metadataTypes><futureMetadata name="XLRICHVALUE" count="1"><bk><extLst><ext uri="{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}"><xlrd:rvb i="0"/></ext></extLst></bk></futureMetadata><valueMetadata count="1"><bk><rc t="1" v="0"/></bk></valueMetadata></metadata>""".stripMargin
    val plainSheet =
      """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        |<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>5</v></c></row></sheetData></worksheet>""".stripMargin
    val wb = read(xlsx(parts(sheetXml = plainSheet, metadata = Some(rich))))
    val s = wb.sheets.headOption.getOrElse(fail("no sheet"))
    val kind = FormulaKind.dynamicArray(CellRange(ref"B1", ref"B1"))
    val edited =
      wb.put(s.put(ref"B1", CellValue.Formula("SORT(A1:A1)", Some(CellValue.Number(5)), kind)))
    val out = write(edited)
    val part = entryText(out, "xl/metadata.xml").getOrElse(fail("no metadata part"))
    // the rich values' blocks come back byte for byte: every vm index still names the same image
    def between(xml: String, open: String, close: String): String =
      val start = xml.indexOf(open)
      assert(start >= 0, s"$open not in $xml")
      xml.substring(start, xml.indexOf(close, start) + close.length)
    Vector(
      ("<valueMetadata", "</valueMetadata>"),
      ("""<futureMetadata name="XLRICHVALUE"""", "</futureMetadata>"),
      ("""<metadataType name="XLRICHVALUE"""", "/>")
    ).foreach((open, close) => assertEquals(between(part, open, close), between(rich, open, close)))
    assert(
      part.contains("""<valueMetadata count="1"><bk><rc t="1" v="0"/></bk></valueMetadata>"""),
      part
    )
    assert(
      part.contains("""<cellMetadata count="1"><bk><rc t="2" v="0"/></bk></cellMetadata>"""),
      part
    )
    assert(sheet1(out).contains("""<c r="B1" t="n" cm="1">"""), sheet1(out))
    assertWiredOnce(out)
    assertEquals(kindAt(read(out), ref"B1"), Some(kind))
  }

  test("a metadata-modified write (a sheet added) keeps the part and its wiring") {
    val wb = read(xlsx())
    val out = write(wb.put(Sheet("Extra").put(ref"A1", CellValue.Number(1))))
    assert(sheet1(out).contains(anchorXml), sheet1(out))
    assertEquals(entryText(out, "xl/metadata.xml"), Some(metadataXml))
    assertWiredOnce(out)
    assertEquals(kindAt(read(out), ref"B1"), Some(FormulaKind.dynamicArray(sortRange)))
  }

  test("a clean re-write of a dynamic-array book is byte-identical to its first write") {
    val once = write(dirty(read(xlsx())))
    val twice = write(dirty(read(once)))
    assertEquals(sheet1(twice), sheet1(once))
    assertEquals(entryText(twice, "xl/metadata.xml"), entryText(once, "xl/metadata.xml"))
  }
