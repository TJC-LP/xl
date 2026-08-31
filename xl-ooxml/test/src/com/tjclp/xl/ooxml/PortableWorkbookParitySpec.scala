package com.tjclp.xl.ooxml

import java.io.ByteArrayOutputStream
import java.nio.charset.StandardCharsets.UTF_8
import org.scalacheck.Prop.forAll
import com.tjclp.xl.Generators
import com.tjclp.xl.api.Workbook
import com.tjclp.xl.ooxml.style.OoxmlStyles
import com.tjclp.xl.ooxml.worksheet.OoxmlWorksheet

/**
 * W2 (GH-543): workbook-part parity — every SaxSerializable part the XlsxWriter emits under
 * XmlBackend.SaxStax, driven over StaxSaxWriter and PortableSaxWriter for Generators-driven
 * workbooks AND the real-file fixture corpus, must produce identical bytes. The zip layer is
 * writer-independent, so part-level byte equality is per-zip-entry byte equality.
 */
class PortableWorkbookParitySpec extends munit.ScalaCheckSuite:

  override def scalaCheckTestParameters =
    super.scalaCheckTestParameters.withMinSuccessfulTests(30)

  private def both(body: SaxWriter => Unit): (Array[Byte], Array[Byte]) =
    val staxOut = new ByteArrayOutputStream()
    val stax = StaxSaxWriter.create(staxOut)
    body(stax)
    stax.flush()
    val portableOut = new ByteArrayOutputStream()
    val portable = PortableSaxWriter.create(portableOut)
    body(portable)
    portable.flush()
    (staxOut.toByteArray, portableOut.toByteArray)

  private def assertParity(label: String)(body: SaxWriter => Unit): Unit =
    val (stax, portable) = both(body)
    assertEquals(new String(portable, UTF_8), new String(stax, UTF_8), label)
    assert(java.util.Arrays.equals(portable, stax), s"$label: byte-level divergence")

  private def assertWorkbookParts(wb: Workbook, label: String): Unit =
    val sst = SharedStrings.fromWorkbook(wb)
    assertParity(s"$label/sharedStrings")(sst.writeSax)
    val styles = OoxmlStyles.fromWorkbook(wb)
    assertParity(s"$label/styles")(styles.writeSax)
    val ooxmlWb = OoxmlWorkbook.fromDomain(wb)
    assertParity(s"$label/workbook")(ooxmlWb.writeSax)
    wb.sheets.foreach { sheet =>
      val ws = OoxmlWorksheet.fromDomainWithSST(sheet, Some(sst))
      assertParity(s"$label/worksheet(${sheet.name.value})")(ws.writeSax)
      assertParity(s"$label/direct(${sheet.name.value})") { w =>
        DirectSaxEmitter.emitWorksheet(w, sheet, Some(sst), Map.empty)
      }
    }

  property("W2: generated workbooks — all parts byte-identical across writers") {
    forAll(Generators.genWorkbook) { wb =>
      assertWorkbookParts(wb, "genWorkbook")
      true
    }
  }

  property("W2: rich generated workbooks (styles, comments, merges, charts)") {
    forAll(Generators.genRichWorkbook) { wb =>
      assertWorkbookParts(wb, "genRichWorkbook")
      true
    }
  }

  TestFixtures.all.foreach { fixture =>
    test(s"W2 fixture: $fixture — regenerated parts byte-identical across writers") {
      XlsxReader.read(TestFixtures.copyToTemp(fixture)) match
        case Right(wb) => assertWorkbookParts(wb, fixture)
        case Left(err) => fail(s"$fixture failed to read: ${err.message}")
    }
  }

  // --- parts not reachable from a domain Workbook alone -----------------------------------------------

  test("W2: ContentTypes part") {
    val ct = ContentTypes(
      defaults = Map(
        "rels" -> "application/vnd.openxmlformats-package.relationships+xml",
        "xml" -> "application/xml"
      ),
      overrides = Map(
        "/xl/workbook.xml" ->
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
        "/xl/worksheets/sheet1.xml" ->
          "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
      )
    )
    assertParity("contentTypes")(ct.writeSax)
  }

  test("W2: Relationships part") {
    val rels = Relationships(
      Vector(
        Relationship("rId1", XmlUtil.relTypeWorksheet, "worksheets/sheet1.xml"),
        Relationship("rId2", XmlUtil.relTypeStyles, "styles.xml")
      )
    )
    assertParity("relationships")(rels.writeSax)
  }

  test("W2: Comments part") {
    import com.tjclp.xl.macros.ref
    import com.tjclp.xl.richtext.RichText
    val comments = OoxmlComments(
      authors = Vector("Alice", "Böb"),
      comments = Vector(
        OoxmlComment(ref"A1", 0, RichText.plain("First & foremost <note>")),
        OoxmlComment(ref"B2", 1, RichText.plain("  spaced  "))
      )
    )
    assertParity("comments")(comments.writeSax)
  }

  test("W2: Table part") {
    import com.tjclp.xl.addressing.CellRange
    val range = CellRange.parse("A1:B3").getOrElse(fail("bad range"))
    val table = OoxmlTable(
      id = 1L,
      name = "Table1",
      displayName = "Table1",
      ref = range,
      headerRowCount = 1,
      totalsRowCount = 0,
      columns = Vector(
        OoxmlTableColumn(1L, "Region & Zone"),
        OoxmlTableColumn(2L, "Vente été")
      ),
      autoFilter = Some(range),
      styleInfo = Some(OoxmlTableStyleInfo("TableStyleMedium2", false, false, true, false))
    )
    assertParity("table")(table.writeSax)
  }
