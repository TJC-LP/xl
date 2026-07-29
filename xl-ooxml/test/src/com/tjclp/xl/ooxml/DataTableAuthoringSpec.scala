package com.tjclp.xl.ooxml

import java.io.ByteArrayInputStream
import java.nio.charset.StandardCharsets
import java.nio.file.Files
import java.util.zip.ZipInputStream

import munit.ScalaCheckSuite
import org.scalacheck.Prop.forAll

import com.tjclp.xl.Generators
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.api.*
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.writer.{WriterConfig, XmlBackend}
import com.tjclp.xl.sheets.dataTableSyntax.*

/**
 * GH-419 data-table authoring byte laws against the first genuine Excel-authored fixtures in the
 * corpus (datatable-excel.xlsx, datatable-excel-edge.xlsx — see PROVENANCE.md):
 *
 *   - clean read -> write copies the Excel worksheet parts verbatim;
 *   - a sibling put (full sheet regeneration) re-emits every record byte-exactly, including the
 *     bare single-cell `ref="Q2"` dialect (Excel never writes `Q2:Q2` on a 1x1 interior);
 *   - `sheet.dataTable`/`dataTableRow`/`dataTableCol` author corner-only records whose `<f>` bytes
 *     equal Excel's own output on both writer backends.
 *
 * The dtr="0" every-interior records in formula-records.xlsx (the Sentinel preservation dialect)
 * coexist deliberately: reading tolerates every wild shape, while authoring emits only the
 * dtr="1"/ca="1" corner-only dialect these fixtures pin.
 */
class DataTableAuthoringSpec extends ScalaCheckSuite:

  private val backends = List(
    "ScalaXml" -> WriterConfig(backend = XmlBackend.ScalaXml),
    "SaxStax" -> WriterConfig(backend = XmlBackend.SaxStax)
  )

  private def range(a1: String): CellRange = CellRange.parse(a1).fold(fail(_), identity)
  private def aref(a1: String): ARef = ARef.parse(a1).fold(fail(_), identity)
  private def num(n: Int): CellValue = CellValue.Number(BigDecimal(n))

  private def readFixture(name: String): (java.nio.file.Path, Workbook) =
    val in = TestFixtures.copyToTemp(name)
    (in, XlsxReader.read(in).fold(err => fail(err.message), identity))

  private def writeToBytes(wb: Workbook, config: WriterConfig = WriterConfig.default): Array[Byte] =
    val out = Files.createTempFile("xl-dt-authoring-", ".xlsx")
    try
      XlsxWriter.writeWith(wb, out, config).fold(err => fail(err.message), identity)
      Files.readAllBytes(out)
    finally Files.deleteIfExists(out)

  // ===== Fixture byte laws (D13) =====

  test("datatable-excel.xlsx: clean read->write copies the worksheet verbatim") {
    val (in, wb) = readFixture("datatable-excel.xlsx")
    val output = writeToBytes(wb)
    assertEquals(
      zipEntry(output, "xl/worksheets/sheet1.xml"),
      zipEntry(Files.readAllBytes(in), "xl/worksheets/sheet1.xml"),
      "clean write must ride the verbatim-copy loop byte-identically"
    )
  }

  test("datatable-excel-edge.xlsx: clean read->write copies both worksheets verbatim") {
    val (in, wb) = readFixture("datatable-excel-edge.xlsx")
    val output = writeToBytes(wb)
    val inputBytes = Files.readAllBytes(in)
    List("xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml").foreach { part =>
      assertEquals(
        zipEntry(output, part),
        zipEntry(inputBytes, part),
        s"clean write must copy $part byte-identically"
      )
    }
  }

  test("datatable-excel.xlsx: a put on a neighboring cell leaves all three records intact") {
    val (_, wb) = readFixture("datatable-excel.xlsx")
    val sheet = wb.sheets.headOption.getOrElse(fail("missing Sheet1"))
    val edited = wb.put(sheet.put(ref"A30", CellValue.Text("dirty")))
    val sheetXml = zipEntry(writeToBytes(edited), "xl/worksheets/sheet1.xml")
    List(
      """<f t="dataTable" ref="F2:H4" dt2D="1" dtr="1" r1="A1" r2="A2" ca="1"/>""",
      """<f t="dataTable" ref="F10:F12" dt2D="0" dtr="0" r1="A2"/>""",
      """<f t="dataTable" ref="B21:D21" dt2D="0" dtr="1" r1="A1"/>""",
      "<v>1000</v>",
      "<v>9000</v>",
      "<v>300</v>",
      "<v>10</v>"
    ).foreach { needle =>
      assert(sheetXml.contains(needle), s"record degraded on dirty regen: $needle\n$sheetXml")
    }
    // The derived TABLE(...) display text is a model-side convenience, never OOXML bytes.
    assert(!sheetXml.contains("TABLE("), s"derived display text leaked into XML: $sheetXml")
  }

  test("datatable-excel-edge.xlsx: dirty regen re-emits the bare 1x1 ref (never Q2:Q2)") {
    val (_, wb) = readFixture("datatable-excel-edge.xlsx")
    val sheet = wb.sheets.headOption.getOrElse(fail("missing Sheet1"))
    val edited = wb.put(sheet.put(ref"A30", CellValue.Text("dirty")))
    val sheetXml = zipEntry(writeToBytes(edited), "xl/worksheets/sheet1.xml")
    List(
      """<f t="dataTable" ref="Q2" dt2D="1" dtr="1" r1="A1" r2="A2"/>""",
      """<f t="dataTable" ref="L2:M3" dt2D="1" dtr="1" r1="A1" r2="A2"/>""",
      """<f t="dataTable" ref="T2:U4" dt2D="0" dtr="0" r1="A2" ca="1"/>"""
    ).foreach { needle =>
      assert(sheetXml.contains(needle), s"record degraded on dirty regen: $needle\n$sheetXml")
    }
    assert(!sheetXml.contains("Q2:Q2"), s"1x1 interior must render a bare ref: $sheetXml")
  }

  backends.foreach { case (backendName, config) =>
    test(s"a from-model 1x1 dataTable record renders a bare ref ($backendName)") {
      val kind: FormulaKind.DataTable = FormulaKind.DataTable(
        range("Q2"),
        dt2D = true,
        dtr = true,
        r1 = Some(aref("A1")),
        r2 = Some(aref("A2"))
      )
      val sheet = Sheet(SheetName.unsafe("S"))
        .put(ref"A1", num(2))
        .put(ref"Q2", CellValue.dataTable(kind, Some(num(1000))))
      val sheetXml = zipEntry(writeToBytes(Workbook(sheet), config), "xl/worksheets/sheet1.xml")
      assert(
        sheetXml.contains("""<f t="dataTable" ref="Q2" dt2D="1" dtr="1" r1="A1" r2="A2"/>"""),
        s"expected the bare-ref Excel dialect: $sheetXml"
      )
      assert(!sheetXml.contains("Q2:Q2"), s"1x1 interior must render a bare ref: $sheetXml")
    }
  }

  // ===== Authored emission (D1-D6 + D2 dialect) =====

  /** Interior D5:F6 -> corner C4, row axis D4:F4, column axis C5:C6; inputs B1/B2. */
  private def base2D: Sheet =
    Sheet(SheetName.unsafe("S"))
      .put(ref"B1", num(10))
      .put(ref"B2", num(20))
      .put(ref"C4", CellValue.Formula("B1*B2"))
      .put(ref"D4", num(1))
      .put(ref"E4", num(2))
      .put(ref"F4", num(3))
      .put(ref"C5", num(100))
      .put(ref"C6", num(200))

  backends.foreach { case (backendName, config) =>
    test(s"authored 2-D D5:F6: exactly one corner record, plain seeded interiors ($backendName)") {
      val seeds = Seq(
        Seq(num(1000), num(2000), num(3000)),
        Seq(num(2000), num(4000), num(6000))
      )
      val authored = base2D
        .dataTable(range("D5:F6"), aref("B1"), aref("B2"), seeds)
        .fold(err => fail(err.message), identity)
      val sheetXml = zipEntry(writeToBytes(Workbook(authored), config), "xl/worksheets/sheet1.xml")
      assert(
        sheetXml.contains(
          """<c r="D5" t="n"><f t="dataTable" ref="D5:F6" dt2D="1" dtr="1" r1="B1" r2="B2" ca="1"/><v>1000</v></c>"""
        ),
        s"corner record bytes wrong: $sheetXml"
      )
      // The corner is the ONLY record cell; every other interior is a plain cached value.
      assertEquals(
        "t=\"dataTable\"".r.findAllIn(sheetXml).size,
        1,
        s"exactly one record cell expected: $sheetXml"
      )
      List(
        """<c r="E5" t="n"><v>2000</v></c>""",
        """<c r="F5" t="n"><v>3000</v></c>""",
        """<c r="D6" t="n"><v>2000</v></c>""",
        """<c r="E6" t="n"><v>4000</v></c>""",
        """<c r="F6" t="n"><v>6000</v></c>"""
      ).foreach { needle =>
        assert(sheetXml.contains(needle), s"plain interior expected: $needle\n$sheetXml")
      }
    }
  }

  test("authored record equals Excel's own bytes on the fixture geometry (2-D F2:H4)") {
    // Rebuild datatable-excel.xlsx's 2-D table from scratch with the authoring API and compare
    // the record element byte-for-byte against what Excel wrote in the fixture.
    val excelRecord = """<f t="dataTable" ref="F2:H4" dt2D="1" dtr="1" r1="A1" r2="A2" ca="1"/>"""
    val fixtureXml =
      zipEntry(
        Files.readAllBytes(TestFixtures.copyToTemp("datatable-excel.xlsx")),
        "xl/worksheets/sheet1.xml"
      )
    assert(fixtureXml.contains(excelRecord), "fixture must carry the Excel-authored record")

    val sheet = Sheet(SheetName.unsafe("Sheet1"))
      .put(ref"A1", num(2))
      .put(ref"A2", num(3))
      .put(ref"E1", CellValue.Formula("A1*A2", Some(num(6))))
      .put(ref"F1", num(10))
      .put(ref"G1", num(20))
      .put(ref"H1", num(30))
      .put(ref"E2", num(100))
      .put(ref"E3", num(200))
      .put(ref"E4", num(300))
    val seeds = Seq(
      Seq(num(1000), num(2000), num(3000)),
      Seq(num(2000), num(4000), num(6000)),
      Seq(num(3000), num(6000), num(9000))
    )
    val authored = sheet
      .dataTable(range("F2:H4"), aref("A1"), aref("A2"), seeds)
      .fold(err => fail(err.message), identity)
    val sheetXml = zipEntry(writeToBytes(Workbook(authored)), "xl/worksheets/sheet1.xml")
    assert(sheetXml.contains(excelRecord), s"authored record must equal Excel's bytes: $sheetXml")
    assert(
      sheetXml.contains(s"""<c r="F2" t="n">$excelRecord<v>1000</v></c>"""),
      s"corner cell must carry the record plus Excel's cached seed: $sheetXml"
    )
  }

  backends.foreach { case (backendName, config) =>
    test(s"authored 1-D row/col: dt2D=0 with the correct dtr, input on r1 ($backendName)") {
      val rowAuthored = Sheet(SheetName.unsafe("S"))
        .put(ref"A1", num(2))
        .put(ref"B20", num(7))
        .put(ref"C20", num(8))
        .put(ref"A21", CellValue.Formula("A1+1"))
        .dataTableRow(range("B21:C21"), aref("A1"))
        .fold(err => fail(err.message), identity)
      val rowXml = zipEntry(writeToBytes(Workbook(rowAuthored), config), "xl/worksheets/sheet1.xml")
      assert(
        rowXml.contains("""<f t="dataTable" ref="B21:C21" dt2D="0" dtr="1" r1="A1" ca="1"/>"""),
        s"row-oriented record bytes wrong: $rowXml"
      )

      val colAuthored = Sheet(SheetName.unsafe("S"))
        .put(ref"A2", num(3))
        .put(ref"F9", CellValue.Formula("A2*100"))
        .put(ref"E10", num(1))
        .put(ref"E11", num(2))
        .dataTableCol(range("F10:F11"), aref("A2"))
        .fold(err => fail(err.message), identity)
      val colXml = zipEntry(writeToBytes(Workbook(colAuthored), config), "xl/worksheets/sheet1.xml")
      assert(
        colXml.contains("""<f t="dataTable" ref="F10:F11" dt2D="0" dtr="0" r1="A2" ca="1"/>"""),
        s"column-oriented record bytes wrong: $colXml"
      )
    }
  }

  backends.foreach { case (backendName, config) =>
    test(s"authored 1x1 interior emits a bare ref; #NUM! seed emits t=e ($backendName)") {
      val authored = Sheet(SheetName.unsafe("S"))
        .put(ref"A1", num(2))
        .put(ref"A2", num(3))
        .put(ref"P1", CellValue.Formula("A1*A2"))
        .put(ref"Q1", num(10))
        .put(ref"P2", num(100))
        .dataTable(range("Q2:Q2"), aref("A1"), aref("A2"), Seq(Seq(CellValue.Error(CellError.Num))))
        .fold(err => fail(err.message), identity)
      val sheetXml = zipEntry(writeToBytes(Workbook(authored), config), "xl/worksheets/sheet1.xml")
      assert(
        sheetXml.contains(
          """<c r="Q2" t="e"><f t="dataTable" ref="Q2" dt2D="1" dtr="1" r1="A1" r2="A2" ca="1"/><v>#NUM!</v></c>"""
        ),
        s"1x1 #NUM!-seeded record bytes wrong: $sheetXml"
      )
      assert(!sheetXml.contains("Q2:Q2"), s"1x1 interior must render a bare ref: $sheetXml")
    }
  }

  property("author -> write -> read round-trips the corner kind, interiors untouched") {
    forAll(Generators.genAuthoredDataTable) { case (interior, orientation, input1, input2) =>
      val startCol = interior.start.col.index0
      val startRow = interior.start.row.index0
      val sourceRefs: Vector[ARef] = orientation match
        case 0 => Vector(ARef.from0(startCol - 1, startRow - 1))
        case 1 => (startRow to interior.end.row.index0).toVector.map(ARef.from0(startCol - 1, _))
        case _ => (startCol to interior.end.col.index0).toVector.map(ARef.from0(_, startRow - 1))
      val sheet = sourceRefs.foldLeft(Sheet(SheetName.unsafe("G"))) { (s, r) =>
        s.put(r, CellValue.Formula("A1*2"))
      }
      val authored = (orientation match
        case 0 => sheet.dataTable(interior, input1, input2)
        case 1 => sheet.dataTableRow(interior, input1)
        case _ => sheet.dataTableCol(interior, input1)
      ).fold(err => fail(s"author failed: $err"), identity)

      val reread = XlsxReader
        .readFromBytes(writeToBytes(Workbook(authored)))
        .fold(err => fail(err.message), identity)
      val rereadSheet = reread.sheets.headOption.getOrElse(fail("missing sheet"))
      val expectedKind: FormulaKind.DataTable = orientation match
        case 0 =>
          FormulaKind.DataTable(
            interior,
            true,
            true,
            Some(input1),
            Some(input2),
            false,
            false,
            true
          )
        case 1 =>
          FormulaKind.DataTable(interior, false, true, Some(input1), None, false, false, true)
        case _ =>
          FormulaKind.DataTable(interior, false, false, Some(input1), None, false, false, true)
      assertEquals(rereadSheet(interior.start).value, CellValue.dataTable(expectedKind, None))
      interior.cellsRowMajor.filterNot(_ == interior.start).foreach { r =>
        assert(rereadSheet.cells.get(r).isEmpty, s"interior ${r.toA1} must stay untouched")
      }
      true
    }
  }

  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def zipEntry(bytes: Array[Byte], name: String): String =
    val zip = new ZipInputStream(new ByteArrayInputStream(bytes))
    var current = zip.getNextEntry
    var result: Option[String] = None
    try
      while current != null && result.isEmpty do
        if current.getName == name then
          result = Some(new String(zip.readAllBytes(), StandardCharsets.UTF_8))
        else
          zip.closeEntry()
          current = zip.getNextEntry
    finally zip.close()
    result.getOrElse(fail(s"missing ZIP entry $name"))
