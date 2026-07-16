package com.tjclp.xl.ooxml

import munit.FunSuite
import java.io.ByteArrayOutputStream
import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.time.LocalDateTime
import java.util.zip.ZipFile
import scala.xml.{
  Elem,
  MetaData,
  Node,
  Null,
  PrefixedAttribute,
  Text,
  TopScope,
  UnprefixedAttribute,
  XML
}
import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row, SheetName}
import com.tjclp.xl.api.Workbook
import com.tjclp.xl.cells.{Cell, CellValue, Comment}
import com.tjclp.xl.richtext.RichText.*
import com.tjclp.xl.sheets.{
  ColumnProperties,
  HeaderFooter,
  PageMargins,
  PageSetup,
  RowProperties,
  Sheet,
  SheetView
}
// DirectSaxEmitter is in the same package

// Test code uses .get/.head for brevity in assertions
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial", "org.wartremover.warts.IterableOps"))
class DirectSaxEmitterParitySpec extends FunSuite:

  test("direct SAX emission matches worksheet XML structure") {
    val sheet = buildSheet()
    val tableParts = Some(buildTableParts())

    val expected = OoxmlWorksheet.fromDomainWithSST(sheet, None, Map.empty, tableParts)
    val actual = emitDirect(sheet, tableParts)

    assertEquals(normalize(expected.toXml), normalize(actual))
  }

  test("formula without cached value omits type attribute (no Excel corruption)") {
    // Regression test: formulas without cached values were written with t="str"
    // which caused Excel to show a "We found a problem" recovery dialog.
    // The fix is to omit the type attribute entirely.
    val a1 = ARef.from1(1, 1)
    val sheet = Sheet(SheetName.unsafe("Test"))
      .put(a1, CellValue.Formula("SUM(B1:B10)", None)) // No cached value

    val xml = emitDirect(sheet, None)
    val cellElem = (xml \\ "c").head

    // Should NOT have t="str" - either no t attribute or empty
    val typeAttr = cellElem.attribute("t").map(_.text)
    assert(
      typeAttr.isEmpty || typeAttr.contains(""),
      s"Formula without cached value should not have type attribute, got: $typeAttr"
    )

    // Should have formula element
    val formulaElem = (cellElem \ "f").headOption
    assert(formulaElem.isDefined, "Cell should have <f> element")
    assertEquals(formulaElem.get.text, "SUM(B1:B10)")
  }

  test("GH-378: formula with cached DateTime emits t=\"n\" and the Excel serial in <v>") {
    val a1 = ARef.from1(1, 1)
    // Jan 1, 2000 noon = serial 36526.5 exactly
    val cached = LocalDateTime.of(2000, 1, 1, 12, 0, 0)
    val sheet = Sheet(SheetName.unsafe("Test"))
      .put(a1, CellValue.Formula("EOMONTH(A1,0)", Some(CellValue.DateTime(cached))))

    val saxCell = (emitDirect(sheet, None) \\ "c").head
    assertEquals(saxCell.attribute("t").map(_.text), Some("n"))
    assertEquals((saxCell \ "f").text, "EOMONTH(A1,0)")
    assertEquals((saxCell \ "v").map(_.text).toList, List("36526.5"))

    // The DOM writer must emit the identical cached serial (backend parity)
    val domCell =
      (OoxmlWorksheet.fromDomainWithSST(sheet, None, Map.empty, None).toXml \\ "c").head
    assertEquals(domCell.attribute("t").map(_.text), Some("n"))
    assertEquals((domCell \ "v").map(_.text).toList, List("36526.5"))
  }

  test("GH-265: direct SAX emission matches DOM writer for fresh-sheet metadata") {
    // Freeze panes, view settings, page setup (incl. even header + fitToPage), hyperlinks:
    // everything the DOM writer emits for a fresh sheet must come out of the streaming path too.
    val sheet = buildMetadataSheet()

    val expected = OoxmlWorksheet.fromDomainWithSST(sheet, None, Map.empty, None)
    val actual = emitDirect(sheet, None)

    assertEquals(normalize(expected.toXml), normalize(actual))
  }

  test("GH-265: SaxStax streaming write round-trips metadata equal to the DOM writer") {
    val wb = Workbook(Vector(buildMetadataSheet()))
    val saxPath = Files.createTempFile("gh265-sax", ".xlsx")
    val domPath = Files.createTempFile("gh265-dom", ".xlsx")
    XlsxWriter
      .writeWith(wb, saxPath, WriterConfig.saxStax)
      .fold(e => fail(s"SaxStax write failed: $e"), identity)
    XlsxWriter
      .writeWith(wb, domPath, WriterConfig.scalaXml)
      .fold(e => fail(s"ScalaXml write failed: $e"), identity)

    val saxWb = XlsxReader.read(saxPath).fold(e => fail(s"SaxStax read failed: $e"), identity)
    val domWb = XlsxReader.read(domPath).fold(e => fail(s"ScalaXml read failed: $e"), identity)
    assertEquals(saxWb.copy(sourceContext = None), domWb.copy(sourceContext = None))

    val xml = zipEntryString(saxPath, "xl/worksheets/sheet1.xml")
    assert(xml.contains("<pane "), s"freeze pane missing from streaming output: $xml")
    assert(
      xml.contains("topLeftCell=\"D20\""),
      s"scrolled pane topLeftCell (GH-382) missing from streaming output: $xml"
    )
    assert(xml.contains("showGridLines=\"0\""), s"view settings missing: $xml")
    assert(xml.contains("tabSelected=\"1\""), s"tabSelected (GH-372) missing: $xml")
    assert(xml.contains("orientation=\"landscape\""), s"pageSetup missing: $xml")
    assert(xml.contains("<pageMargins "), s"pageMargins missing: $xml")
    assert(xml.contains("<evenHeader>"), s"even header (GH-266) missing: $xml")
    assert(xml.contains("fitToPage=\"1\""), s"sheetPr fitToPage flag missing: $xml")
    // StAX writes <tabColor ...></tabColor>, the DOM writer self-closes — match the open tag only
    assert(xml.contains("<tabColor rgb=\"FF1F4E79\""), s"tabColor (GH-358) missing: $xml")
    assert(
      xml.indexOf("<tabColor ") < xml.indexOf("<pageSetUpPr "),
      s"CT_SheetPr order: tabColor before pageSetUpPr: $xml"
    )
    // Schema order (ECMA-376 18.3.1.99): sheetPr < sheetViews < cols < sheetData < page*
    val order = Seq(
      "<sheetPr>",
      "<dimension ",
      "<sheetViews>",
      "<cols>",
      "<sheetData>",
      "<mergeCells ",
      "<hyperlinks>",
      "<pageMargins ",
      "<pageSetup ",
      "<headerFooter ",
      "<legacyDrawing "
    ).map(tag => tag -> xml.indexOf(tag))
    order.foreach((tag, idx) => assert(idx >= 0, s"$tag missing from streaming output: $xml"))
    assert(
      order.map(_._2) == order.map(_._2).sorted,
      s"streaming output violates schema order: $order"
    )
    Files.deleteIfExists(saxPath)
    Files.deleteIfExists(domPath)
  }

  test("GH-383: rich text inline cell emits <rFont> under rPr and matches DOM writer") {
    val a1 = ARef.from1(1, 1)
    val rich = "Styled".fontFamily("Times New Roman") + " plain"
    val sheet = Sheet(SheetName.unsafe("Sheet1")).put(a1, CellValue.RichText(rich))

    val dom = OoxmlWorksheet.fromDomainWithSST(sheet, None, Map.empty, None).toXml
    val sax = emitDirect(sheet, None)

    for (label, xml) <- List("DOM" -> dom, "SAX" -> sax) do
      val rPr = (xml \\ "rPr").headOption.getOrElse(fail(s"$label: expected <rPr> in: $xml"))
      assertEquals((rPr \ "rFont" \ "@val").text, "Times New Roman", s"$label rPr was: $rPr")
      assert((rPr \ "name").isEmpty, s"$label: CT_RPrElt forbids <name> inside <rPr>: $rPr")

    assertEquals(normalize(dom), normalize(sax))
  }

  private def zipEntryString(path: Path, entry: String): String =
    val zf = new ZipFile(path.toFile)
    try
      val is = zf.getInputStream(zf.getEntry(entry))
      try new String(is.readAllBytes(), StandardCharsets.UTF_8)
      finally is.close()
    finally zf.close()

  /** Fresh sheet exercising every metadata block DirectSaxEmitter must emit (GH-265). */
  private def buildMetadataSheet(): Sheet =
    val b2 = ARef.from1(2, 2)
    buildSheet()
      .put(Cell(ARef.from1(4, 2), CellValue.Text("link")).withHyperlink("https://example.com"))
      // GH-382: scrolled freeze — both backends must emit the scroll target as topLeftCell
      .freezeAt(b2, ARef.from1(4, 20))
      .withViewSettings(
        SheetView(showGridLines = false, zoomScale = Some(85), tabSelected = Some(true))
      )
      // GH-358: tabColor must come out of the streaming path too, first among sheetPr children
      .withTabColor(com.tjclp.xl.styles.color.Color.Rgb(0xff1f4e79))
      .withPageSetup(
        PageSetup(
          orientation = Some("landscape"),
          fitToWidth = Some(1),
          headerFooter = Some(
            HeaderFooter(
              oddHeader = Some("&LXL"),
              oddFooter = Some("&CPage &P of &N"),
              evenHeader = Some("&REven"),
              differentOddEven = true
            )
          ),
          margins = Some(PageMargins.default)
        )
      )

  private def emitDirect(sheet: Sheet, tableParts: Option[Elem]): Elem =
    val output = new ByteArrayOutputStream()
    val writer = StaxSaxWriter.create(output)
    DirectSaxEmitter.emitWorksheet(writer, sheet, None, Map.empty, tableParts)
    val xml = new String(output.toByteArray, StandardCharsets.UTF_8)
    XML.loadString(xml)

  private def buildSheet(): Sheet =
    val a1 = ARef.from1(1, 1)
    val b1 = ARef.from1(2, 1)
    val c1 = ARef.from1(3, 1)
    val d1 = ARef.from1(4, 1)
    val e1 = ARef.from1(5, 1)
    val a2 = ARef.from1(1, 2)
    val b2 = ARef.from1(2, 2)

    val rich = "Bold".bold.red + " plain"

    Sheet(SheetName.unsafe("Sheet1"))
      .put(a1, CellValue.Text("  spaced"))
      .put(b1, CellValue.Number(BigDecimal(42)))
      .put(c1, CellValue.Bool(true))
      .put(
        d1,
        CellValue.Formula("SUM(A1:B1)", Some(CellValue.Number(BigDecimal(3))))
      )
      .put(
        e1,
        // GH-378: cached DateTime must serialize identically on both backends
        CellValue.Formula(
          "EDATE(B2,1)",
          Some(CellValue.DateTime(LocalDateTime.of(2024, 2, 2, 3, 4)))
        )
      )
      .put(a2, CellValue.RichText(rich))
      .put(b2, CellValue.DateTime(LocalDateTime.of(2024, 1, 2, 3, 4)))
      .merge(CellRange(ARef.from1(1, 3), ARef.from1(2, 3)))
      .comment(a1, Comment.plainText("Note", Some("QA")))
      .setRowProperties(
        Row.from1(2),
        RowProperties(height = Some(24.0), hidden = true, outlineLevel = Some(1), collapsed = true)
      )
      // GH-381: property-only row (no cells in row 4) must survive on both backends
      .setRowProperties(Row.from1(4), RowProperties(height = Some(5.25)))
      .setColumnProperties(
        Column.from1(1),
        ColumnProperties(width = Some(12.5), outlineLevel = Some(1))
      )
      .setColumnProperties(
        Column.from1(2),
        ColumnProperties(width = Some(12.5), outlineLevel = Some(1))
      )
      .setColumnProperties(
        Column.from1(3),
        ColumnProperties(
          width = Some(20.0),
          hidden = true,
          outlineLevel = Some(2),
          collapsed = true
        )
      )

  private def buildTableParts(): Elem =
    val tablePart =
      scala.xml.Elem(
        prefix = null,
        label = "tablePart",
        attributes = new PrefixedAttribute("r", "id", "rId1", Null),
        scope = scala.xml.NamespaceBinding("r", XmlUtil.nsRelationships, TopScope),
        minimizeEmpty = true
      )
    XmlUtil.elem("tableParts", "count" -> "1")(tablePart)

  private def normalize(elem: Elem): Elem =
    val attrs = normalizeAttributes(elem.attributes)
    val children = elem.child.flatMap {
      case t: Text if t.text.forall(_.isWhitespace) => None
      case e: Elem => Some(normalize(e))
      case other => Some(other)
    }
    elem.copy(prefix = null, scope = TopScope, attributes = attrs, child = children)

  private def normalizeAttributes(attrs: MetaData): MetaData =
    val sorted = attributePairs(attrs)
      .filterNot { case (key, _) => key == "xmlns" || key.startsWith("xmlns:") }
      .sortBy(_._1)
    sorted.foldLeft(Null: MetaData) { case (acc, (key, value)) =>
      key.split(":", 2) match
        case Array(prefix, local) => new PrefixedAttribute(prefix, local, value, acc)
        case _ => new UnprefixedAttribute(key, value, acc)
    }

  private def attributePairs(attrs: MetaData): List[(String, String)] =
    attrs match
      case Null => Nil
      case PrefixedAttribute(pre, key, value, next) =>
        (s"$pre:$key" -> value.text) :: attributePairs(next)
      case UnprefixedAttribute(key, value, next) =>
        (key -> value.text) :: attributePairs(next)
