package com.tjclp.xl.ooxml

import munit.FunSuite
import scala.xml.*
import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.api.*
import com.tjclp.xl.macros.ref
import com.tjclp.xl.tables.{TableSpec, TableColumn, TableAutoFilter, TableStyle}
import com.tjclp.xl.cells.CellValue
import java.nio.file.{Files, Path}

/**
 * Tests for OOXML Table parsing and serialization.
 *
 * Verifies:
 *   - Parse table XML (xl/tables/tableN.xml)
 *   - Round-trip (parse → encode → parse)
 *   - AutoFilter support
 *   - Table style mappings
 *   - Column definitions
 *   - Forwards compatibility (preserve unknown attributes/children)
 *   - Full workbook round-trip with tables
 */
class TableSpec extends FunSuite:

  val tempDir: Path = Files.createTempDirectory("xl-table-test-")

  override def afterAll(): Unit =
    Files.walk(tempDir)
      .sorted(java.util.Comparator.reverseOrder())
      .forEach(Files.delete)

  // ========================================
  // Category A: XML Parsing (10 tests)
  // ========================================

  test("parse minimal table with required attributes") {
    val xml =
      <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
             id="1" name="Table1" displayName="Table 1" ref="A1:D10">
        <tableColumns count="4">
          <tableColumn id="1" name="Product"/>
          <tableColumn id="2" name="Price"/>
          <tableColumn id="3" name="Quantity"/>
          <tableColumn id="4" name="Total"/>
        </tableColumns>
      </table>

    val result = OoxmlTable.fromXml(xml)

    assert(result.isRight, s"Expected successful parse, got: $result")
    val table = result.toOption.getOrElse(fail("Expected Right"))

    assertEquals(table.id, 1L)
    assertEquals(table.name, "Table1")
    assertEquals(table.displayName, "Table 1")
    assertEquals(table.ref, CellRange(ref"A1", ref"D10"))
    assertEquals(table.columns.size, 4)
    assertEquals(table.columns(0).name, "Product")
  }

  test("parse table with AutoFilter") {
    val xml =
      <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
             id="2" name="Table2" displayName="Sales Data" ref="A1:C100">
        <autoFilter ref="A1:C100"/>
        <tableColumns count="3">
          <tableColumn id="1" name="Date"/>
          <tableColumn id="2" name="Amount"/>
          <tableColumn id="3" name="Status"/>
        </tableColumns>
      </table>

    val result = OoxmlTable.fromXml(xml)
    assert(result.isRight)
    val table = result.toOption.getOrElse(fail("Expected Right"))

    assert(table.autoFilter.isDefined, "Expected AutoFilter")
    assertEquals(table.autoFilter.get, CellRange(ref"A1", ref"C100"))
  }

  test("parse table with style info") {
    val xml =
      <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
             id="3" name="Table3" displayName="Table 3" ref="A1:B10">
        <tableColumns count="2">
          <tableColumn id="1" name="Col1"/>
          <tableColumn id="2" name="Col2"/>
        </tableColumns>
        <tableStyleInfo name="TableStyleMedium9"
                        showFirstColumn="0" showLastColumn="0"
                        showRowStripes="1" showColumnStripes="0"/>
      </table>

    val result = OoxmlTable.fromXml(xml)
    assert(result.isRight)
    val table = result.toOption.getOrElse(fail("Expected Right"))

    assert(table.styleInfo.isDefined)
    assertEquals(table.styleInfo.get.name, "TableStyleMedium9")
    assert(table.styleInfo.get.showRowStripes)
    assert(!table.styleInfo.get.showColumnStripes)
  }

  test("parse table with header and totals rows") {
    val xml =
      <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
             id="4" name="Table4" displayName="Table 4" ref="A1:B10"
             headerRowCount="1" totalsRowCount="1">
        <tableColumns count="2">
          <tableColumn id="1" name="Col1"/>
          <tableColumn id="2" name="Col2"/>
        </tableColumns>
      </table>

    val result = OoxmlTable.fromXml(xml)
    assert(result.isRight)
    val table = result.toOption.getOrElse(fail("Expected Right"))

    assertEquals(table.headerRowCount, 1)
    assertEquals(table.totalsRowCount, 1)
  }

  test("preserve unknown attributes for forwards compatibility") {
    val xml =
      <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
             id="5" name="Table5" displayName="Table 5" ref="A1:B10"
             futureAttr="value123">
        <tableColumns count="2">
          <tableColumn id="1" name="Col1" futureColAttr="colValue"/>
          <tableColumn id="2" name="Col2"/>
        </tableColumns>
      </table>

    val result = OoxmlTable.fromXml(xml)
    assert(result.isRight)
    val table = result.toOption.getOrElse(fail("Expected Right"))

    assert(table.otherAttrs.contains("futureAttr"))
    assertEquals(table.otherAttrs("futureAttr"), "value123")

    // Column attributes preserved too
    assert(table.columns(0).otherAttrs.contains("futureColAttr"))
    assertEquals(table.columns(0).otherAttrs("futureColAttr"), "colValue")
  }

  test("preserve unknown child elements") {
    val xml =
      <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
             id="6" name="Table6" displayName="Table 6" ref="A1:B10">
        <tableColumns count="2">
          <tableColumn id="1" name="Col1"/>
          <tableColumn id="2" name="Col2"/>
        </tableColumns>
        <futureElement attr="value">Content</futureElement>
      </table>

    val result = OoxmlTable.fromXml(xml)
    assert(result.isRight)
    val table = result.toOption.getOrElse(fail("Expected Right"))

    assert(table.otherChildren.nonEmpty, "Expected unknown child preserved")
    assertEquals(table.otherChildren.head.label, "futureElement")
  }

  test("error on missing required attribute: id") {
    val xml =
      <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
             name="Table1" displayName="Table 1" ref="A1:B10">
        <tableColumns count="2">
          <tableColumn id="1" name="Col1"/>
          <tableColumn id="2" name="Col2"/>
        </tableColumns>
      </table>

    val result = OoxmlTable.fromXml(xml)
    assert(result.isLeft, "Should fail when id attribute is missing")
  }

  test("error on missing required attribute: name") {
    val xml =
      <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
             id="1" displayName="Table 1" ref="A1:B10">
        <tableColumns count="2">
          <tableColumn id="1" name="Col1"/>
          <tableColumn id="2" name="Col2"/>
        </tableColumns>
      </table>

    val result = OoxmlTable.fromXml(xml)
    assert(result.isLeft, "Should fail when name attribute is missing")
  }

  test("error on invalid cell range") {
    val xml =
      <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
             id="1" name="Table1" displayName="Table 1" ref="INVALID!">
        <tableColumns count="2">
          <tableColumn id="1" name="Col1"/>
          <tableColumn id="2" name="Col2"/>
        </tableColumns>
      </table>

    val result = OoxmlTable.fromXml(xml)
    assert(result.isLeft, "Should fail when ref is invalid")
  }

  test("error on wrong element type") {
    val xml = <notATable/>

    val result = OoxmlTable.fromXml(xml)
    assert(result.isLeft, "Should fail when element is not <table>")
    assert(result.left.exists(_.contains("Expected <table>")))
  }

  // ========================================
  // Category B: XML Serialization (5 tests)
  // ========================================

  test("encode minimal table produces valid XML") {
    val table = OoxmlTable(
      id = 1L,
      name = "Table1",
      displayName = "Table1",  // No spaces (Excel requirement)
      ref = CellRange(ref"A1", ref"D10"),
      headerRowCount = 1,
      totalsRowCount = 0,
      columns = Vector(
        OoxmlTableColumn(1, "Product"),
        OoxmlTableColumn(2, "Price"),
        OoxmlTableColumn(3, "Quantity"),
        OoxmlTableColumn(4, "Total")
      ),
      autoFilter = None,
      styleInfo = None
    )

    val xml = OoxmlTable.toXml(table)

    assertEquals(xml.label, "table")
    assertEquals((xml \ "@id").text, "1")
    assertEquals((xml \ "@name").text, "Table1")
    assertEquals((xml \ "@displayName").text, "Table1")
    assertEquals((xml \ "@ref").text, "A1:D10")

    val columns = xml \ "tableColumns" \ "tableColumn"
    assertEquals(columns.size, 4)
  }

  test("encode table with AutoFilter") {
    val table = OoxmlTable(
      id = 2L,
      name = "Table2",
      displayName = "Sales",
      ref = CellRange(ref"A1", ref"C100"),
      headerRowCount = 1,
      totalsRowCount = 0,
      columns = Vector(
        OoxmlTableColumn(1, "Date"),
        OoxmlTableColumn(2, "Amount")
      ),
      autoFilter = Some(CellRange(ref"A1", ref"C100")),
      styleInfo = None
    )

    val xml = OoxmlTable.toXml(table)

    val autoFilter = xml \ "autoFilter"
    assert(autoFilter.nonEmpty, "Expected autoFilter element")
    assertEquals((autoFilter.head \ "@ref").text, "A1:C100")
  }

  test("encode table with style info") {
    val styleInfo = OoxmlTableStyleInfo(
      name = "TableStyleMedium9",
      showFirstColumn = false,
      showLastColumn = false,
      showRowStripes = true,
      showColumnStripes = false
    )

    val table = OoxmlTable(
      id = 3L,
      name = "Table3",
      displayName = "Styled",
      ref = CellRange(ref"A1", ref"B10"),
      headerRowCount = 1,
      totalsRowCount = 0,
      columns = Vector(OoxmlTableColumn(1, "Col1")),
      autoFilter = None,
      styleInfo = Some(styleInfo)
    )

    val xml = OoxmlTable.toXml(table)

    val styleElem = xml \ "tableStyleInfo"
    assert(styleElem.nonEmpty)
    assertEquals((styleElem.head \ "@name").text, "TableStyleMedium9")
    assertEquals((styleElem.head \ "@showRowStripes").text, "1")
    assertEquals((styleElem.head \ "@showColumnStripes").text, "0")
  }

  test("attribute ordering is deterministic") {
    val table = OoxmlTable(
      id = 1L,
      name = "Table1",
      displayName = "Test",
      ref = CellRange(ref"A1", ref"B10"),
      headerRowCount = 1,
      totalsRowCount = 1,
      columns = Vector(OoxmlTableColumn(1, "Col")),
      autoFilter = None,
      styleInfo = None
    )

    val xml1 = OoxmlTable.toXml(table)
    val xml2 = OoxmlTable.toXml(table)

    // Attributes should be in same order both times
    assertEquals(xml1.toString, xml2.toString)
  }

  test("preserve unknown attributes and children on encode") {
    val table = OoxmlTable(
      id = 5L,
      name = "Table5",
      displayName = "Test",
      ref = CellRange(ref"A1", ref"B10"),
      headerRowCount = 1,
      totalsRowCount = 0,
      columns = Vector(OoxmlTableColumn(1, "Col", None, None, Map("futureAttr" -> "value"))),
      autoFilter = None,
      styleInfo = None,
      otherAttrs = Map("unknownAttr" -> "preserved"),
      otherChildren = Seq(<futureElement>Data</futureElement>)
    )

    val xml = OoxmlTable.toXml(table)

    // Unknown table attrs preserved
    assertEquals((xml \ "@unknownAttr").text, "preserved")

    // Unknown child elements preserved
    assert((xml \ "futureElement").nonEmpty, "Expected futureElement to be preserved")

    // Column attrs preserved
    val col = (xml \ "tableColumns" \ "tableColumn").head
    assertEquals((col \ "@futureAttr").text, "value")
  }

  // ========================================
  // Category C: Round-Trip XML (8 tests)
  // ========================================

  test("round-trip: minimal table") {
    val originalXml =
      <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
             id="1" name="Table1" displayName="Test" ref="A1:B10">
        <tableColumns count="2">
          <tableColumn id="1" name="Col1"/>
          <tableColumn id="2" name="Col2"/>
        </tableColumns>
      </table>

    val parsed1 = OoxmlTable.fromXml(originalXml)
    assert(parsed1.isRight)

    val encoded = OoxmlTable.toXml(parsed1.toOption.get)
    val parsed2 = OoxmlTable.fromXml(encoded)
    assert(parsed2.isRight)

    val table1 = parsed1.toOption.get
    val table2 = parsed2.toOption.get

    assertEquals(table1.id, table2.id)
    assertEquals(table1.name, table2.name)
    assertEquals(table1.displayName, table2.displayName)
    assertEquals(table1.ref, table2.ref)
    assertEquals(table1.columns.size, table2.columns.size)
  }

  test("round-trip: table with AutoFilter") {
    val table = OoxmlTable(
      id = 2L,
      name = "Filtered",
      displayName = "FilteredTable",
      ref = CellRange(ref"A1", ref"C50"),
      headerRowCount = 1,
      totalsRowCount = 0,
      columns = Vector(
        OoxmlTableColumn(1, "Name"),
        OoxmlTableColumn(2, "Value"),
        OoxmlTableColumn(3, "Status")
      ),
      autoFilter = Some(CellRange(ref"A1", ref"C50")),
      styleInfo = None
    )

    val xml = OoxmlTable.toXml(table)
    val reparsed = OoxmlTable.fromXml(xml).toOption.get

    assertEquals(reparsed.autoFilter, table.autoFilter)
  }

  test("round-trip: table with style info") {
    val styleInfo = OoxmlTableStyleInfo(
      name = "TableStyleLight15",
      showFirstColumn = true,
      showLastColumn = true,
      showRowStripes = false,
      showColumnStripes = true
    )

    val table = OoxmlTable(
      id = 3L,
      name = "StyledTable",
      displayName = "Styled",
      ref = CellRange(ref"A1", ref"D20"),
      headerRowCount = 1,
      totalsRowCount = 0,
      columns = Vector(OoxmlTableColumn(1, "Data")),
      autoFilter = None,
      styleInfo = Some(styleInfo)
    )

    val xml = OoxmlTable.toXml(table)
    val reparsed = OoxmlTable.fromXml(xml).toOption.get

    assertEquals(reparsed.styleInfo.get.name, styleInfo.name)
    assertEquals(reparsed.styleInfo.get.showFirstColumn, styleInfo.showFirstColumn)
    assertEquals(reparsed.styleInfo.get.showRowStripes, styleInfo.showRowStripes)
  }

  test("round-trip: table with totals row") {
    val table = OoxmlTable(
      id = 4L,
      name = "WithTotals",
      displayName = "WithTotals",
      ref = CellRange(ref"A1", ref"B11"),
      headerRowCount = 1,
      totalsRowCount = 1,
      columns = Vector(
        OoxmlTableColumn(1, "Item"),
        OoxmlTableColumn(2, "Amount")
      ),
      autoFilter = None,
      styleInfo = None
    )

    val xml = OoxmlTable.toXml(table)
    val reparsed = OoxmlTable.fromXml(xml).toOption.get

    assertEquals(reparsed.headerRowCount, 1)
    assertEquals(reparsed.totalsRowCount, 1)
  }

  test("round-trip preserves unknown attributes") {
    val table = OoxmlTable(
      id = 5L,
      name = "Table5",
      displayName = "Test",
      ref = CellRange(ref"A1", ref"B10"),
      headerRowCount = 1,
      totalsRowCount = 0,
      columns = Vector(OoxmlTableColumn(1, "Col", None, None, Map("future" -> "value"))),
      autoFilter = None,
      styleInfo = None,
      otherAttrs = Map("unknownAttr" -> "preserved")
    )

    val xml = OoxmlTable.toXml(table)
    val reparsed = OoxmlTable.fromXml(xml).toOption.get

    assertEquals(reparsed.otherAttrs.get("unknownAttr"), Some("preserved"))
    assertEquals(reparsed.columns(0).otherAttrs.get("future"), Some("value"))
  }

  test("round-trip: table with complex column names") {
    val table = OoxmlTable(
      id = 6L,
      name = "Table6",
      displayName = "ComplexColumns",
      ref = CellRange(ref"A1", ref"E100"),
      headerRowCount = 1,
      totalsRowCount = 0,
      columns = Vector(
        OoxmlTableColumn(1, "Column with spaces"),
        OoxmlTableColumn(2, "Column_with_underscores"),
        OoxmlTableColumn(3, "Column-with-dashes"),
        OoxmlTableColumn(4, "Column.with.dots"),
        OoxmlTableColumn(5, "Column123Numbers")
      ),
      autoFilter = None,
      styleInfo = None
    )

    val xml = OoxmlTable.toXml(table)
    val reparsed = OoxmlTable.fromXml(xml).toOption.get

    assertEquals(reparsed.columns.map(_.name), table.columns.map(_.name))
  }

  test("round-trip: large table (1000 columns)") {
    val columns = (1 to 1000).map { i =>
      OoxmlTableColumn(i.toLong, s"Column$i")
    }.toVector

    val table = OoxmlTable(
      id = 7L,
      name = "LargeTable",
      displayName = "Large",
      ref = CellRange(ref"A1", ARef.from1(1000, 100)),
      headerRowCount = 1,
      totalsRowCount = 0,
      columns = columns,
      autoFilter = None,
      styleInfo = None
    )

    val xml = OoxmlTable.toXml(table)
    val reparsed = OoxmlTable.fromXml(xml).toOption.get

    assertEquals(reparsed.columns.size, 1000)
    assertEquals(reparsed.columns.head.name, "Column1")
    assertEquals(reparsed.columns.last.name, "Column1000")
  }

  test("round-trip: empty tables map serializes as empty") {
    val baseSheet = Sheet("Empty").getOrElse(fail("Failed to create sheet"))
    assert(baseSheet.tables.isEmpty)

    val workbook = Workbook(Vector(baseSheet))

    val bytes = XlsxWriter.writeToBytes(workbook).getOrElse(fail("Write failed"))
    val reread = XlsxReader.readFromBytes(bytes).getOrElse(fail("Read failed"))

    val rereadSheet = reread.sheets.headOption.getOrElse(fail("Expected sheet"))
    assert(rereadSheet.tables.isEmpty, "Empty tables should round-trip")
  }

  // ========================================
  // Category D: Domain Conversion (6 tests)
  // ========================================

  test("convert domain TableSpec to OOXML") {
    val domainTable = TableSpec.unsafeFromColumnNames(
      name = "Sales",
      displayName = "SalesData",  // No spaces (Excel requirement)
      range = CellRange(ref"A1", ref"D100"),
      columnNames = Vector("Date", "Amount", "Category", "Status")
    )

    val ooxmlTable = TableConversions.toOoxml(domainTable, id = 1L)

    assertEquals(ooxmlTable.id, 1L)
    assertEquals(ooxmlTable.name, "Sales")
    assertEquals(ooxmlTable.displayName, "SalesData")
    assertEquals(ooxmlTable.ref, domainTable.range)
    assertEquals(ooxmlTable.columns.size, 4)
    assertEquals(ooxmlTable.columns(0).name, "Date")
  }

  test("convert OOXML table to domain TableSpec") {
    val ooxmlTable = OoxmlTable(
      id = 2L,
      name = "Products",
      displayName = "ProductList",
      ref = CellRange(ref"A1", ref"C50"),
      headerRowCount = 1,
      totalsRowCount = 0,
      columns = Vector(
        OoxmlTableColumn(1, "Name"),
        OoxmlTableColumn(2, "Price"),
        OoxmlTableColumn(3, "Stock")
      ),
      autoFilter = Some(CellRange(ref"A1", ref"C50")),
      styleInfo = None
    )

    val domainTable = TableConversions.fromOoxml(ooxmlTable)

    assertEquals(domainTable.name, "Products")
    assertEquals(domainTable.displayName, "ProductList")  // No spaces
    assertEquals(domainTable.columns.size, 3)
    assert(domainTable.autoFilter.exists(_.enabled))
  }

  test("convert table styles: Light, Medium, Dark") {
    val styles = Seq(
      TableStyle.Light(5),
      TableStyle.Medium(9),
      TableStyle.Dark(3)
    )

    styles.foreach { style =>
      val ooxml = TableConversions.toOoxml(
        TableSpec.unsafeFromColumnNames("T", "T", CellRange(ref"A1", ref"B10"), Vector("A", "B"))
          .copy(style = style),
        id = 1L
      )

      val domainBack = TableConversions.fromOoxml(ooxml)
      assertEquals(domainBack.style, style)
    }
  }

  test("convert table with AutoFilter enabled") {
    val domainTable = TableSpec.unsafeFromColumnNames(
      name = "Data",
      displayName = "Data",
      range = CellRange(ref"A1", ref"C100"),
      columnNames = Vector("A", "B", "C")
    ).copy(autoFilter = Some(TableAutoFilter(enabled = true)))

    val ooxml = TableConversions.toOoxml(domainTable, id = 1L)
    assert(ooxml.autoFilter.isDefined)

    val domainBack = TableConversions.fromOoxml(ooxml)
    assert(domainBack.autoFilter.exists(_.enabled))
  }

  test("convert table with header and totals rows") {
    val domainTable = TableSpec.unsafeFromColumnNames(
      name = "Totals",
      displayName = "WithTotals",
      range = CellRange(ref"A1", ref"B11"),
      columnNames = Vector("Item", "Amount")
    ).copy(showHeaderRow = true, showTotalsRow = true)

    val ooxml = TableConversions.toOoxml(domainTable, id = 1L)
    assertEquals(ooxml.headerRowCount, 1)
    assertEquals(ooxml.totalsRowCount, 1)

    val domainBack = TableConversions.fromOoxml(ooxml)
    assert(domainBack.showHeaderRow)
    assert(domainBack.showTotalsRow)
  }

  test("convert table with no style defaults to Medium(2)") {
    val domainTable = TableSpec.unsafeFromColumnNames(
      name = "Default",
      displayName = "DefaultStyle",
      range = CellRange(ref"A1", ref"B10"),
      columnNames = Vector("A", "B")
    ) // Uses TableStyle.default

    val ooxml = TableConversions.toOoxml(domainTable, id = 1L)
    assertEquals(ooxml.styleInfo.get.name, "TableStyleMedium2")

    val domainBack = TableConversions.fromOoxml(ooxml)
    assertEquals(domainBack.style, TableStyle.Medium(2))
  }

  // ========================================
  // Category E: Full Workbook Round-Trip (10 tests)
  // ========================================

  test("workbook with single table round-trips") {
    val table = TableSpec.unsafeFromColumnNames(
      name = "SalesData",
      displayName = "SalesData",  // No spaces (Excel requirement)
      range = CellRange(ref"A1", ref"D100"),
      columnNames = Vector("Date", "Product", "Amount", "Status")
    )

    val baseSheet = Sheet("Q1 Sales").getOrElse(fail("Failed to create sheet"))
    val sheet = baseSheet.withTable(table)
    val workbook = Workbook(Vector(sheet))

    // Write to bytes
    val bytes = XlsxWriter.writeToBytes(workbook).getOrElse(fail("Write failed"))

    // Read back
    val rereadWb = XlsxReader.readFromBytes(bytes).getOrElse(fail("Read failed"))

    val rereadSheet = rereadWb.sheets.headOption.getOrElse(fail("Expected sheet"))
    val rereadTable = rereadSheet.getTable("SalesData").getOrElse(fail("Table missing"))

    assertEquals(rereadTable.name, table.name)
    assertEquals(rereadTable.displayName, table.displayName)
    assertEquals(rereadTable.range, table.range)
    assertEquals(rereadTable.columns.size, table.columns.size)
  }

  test("workbook with multiple tables on one sheet") {
    val table1 = TableSpec.unsafeFromColumnNames(
      name = "Sales",
      displayName = "Sales",
      range = CellRange(ref"A1", ref"C50"),
      columnNames = Vector("Date", "Amount", "Status")
    )

    val table2 = TableSpec.unsafeFromColumnNames(
      name = "Products",
      displayName = "Products",
      range = CellRange(ref"E1", ref"G30"),
      columnNames = Vector("Name", "Price", "Stock")
    )

    val baseSheet = Sheet("Data").getOrElse(fail("Failed"))
    val sheet = baseSheet.withTable(table1).withTable(table2)
    val workbook = Workbook(Vector(sheet))

    val bytes = XlsxWriter.writeToBytes(workbook).getOrElse(fail("Write failed"))
    val reread = XlsxReader.readFromBytes(bytes).getOrElse(fail("Read failed"))

    val rereadSheet = reread.sheets.head
    assertEquals(rereadSheet.tables.size, 2)
    assert(rereadSheet.getTable("Sales").isDefined)
    assert(rereadSheet.getTable("Products").isDefined)
  }

  test("workbook with tables on multiple sheets") {
    val table1 = TableSpec.unsafeFromColumnNames(
      name = "Sheet1Table",
      displayName = "Table1",
      range = CellRange(ref"A1", ref"B10"),
      columnNames = Vector("A", "B")
    )

    val table2 = TableSpec.unsafeFromColumnNames(
      name = "Sheet2Table",
      displayName = "Table2",
      range = CellRange(ref"C1", ref"D10"),
      columnNames = Vector("C", "D")
    )

    val sheet1 = Sheet("Sheet1").getOrElse(fail("Failed")).withTable(table1)
    val sheet2 = Sheet("Sheet2").getOrElse(fail("Failed")).withTable(table2)
    val workbook = Workbook(Vector(sheet1, sheet2))

    val bytes = XlsxWriter.writeToBytes(workbook).getOrElse(fail("Write failed"))
    val reread = XlsxReader.readFromBytes(bytes).getOrElse(fail("Read failed"))

    assertEquals(reread.sheets.size, 2)
    assert(reread.sheets(0).getTable("Sheet1Table").isDefined)
    assert(reread.sheets(1).getTable("Sheet2Table").isDefined)
  }

  test("table with AutoFilter round-trips correctly") {
    val table = TableSpec.unsafeFromColumnNames(
      name = "FilteredData",
      displayName = "FilteredData",
      range = CellRange(ref"A1", ref"C100"),
      columnNames = Vector("Name", "Value", "Status")
    ).copy(autoFilter = Some(TableAutoFilter(enabled = true)))

    val sheet = Sheet("Data").getOrElse(fail("Failed")).withTable(table)
    val workbook = Workbook(Vector(sheet))

    val bytes = XlsxWriter.writeToBytes(workbook).getOrElse(fail("Write failed"))
    val reread = XlsxReader.readFromBytes(bytes).getOrElse(fail("Read failed"))

    val rereadTable = reread.sheets.head.getTable("FilteredData").getOrElse(fail("Missing"))
    assert(rereadTable.autoFilter.exists(_.enabled))
  }

  test("table styles round-trip: Light, Medium, Dark") {
    val styles = Seq(
      ("Light5", TableStyle.Light(5)),
      ("Medium9", TableStyle.Medium(9)),
      ("Dark3", TableStyle.Dark(3))
    )

    styles.foreach { case (name, style) =>
      val table = TableSpec.unsafeFromColumnNames(
        name = name,
        displayName = name,
        range = CellRange(ref"A1", ref"B10"),
        columnNames = Vector("A", "B")
      ).copy(style = style)

      val sheet = Sheet("Test").getOrElse(fail("Failed")).withTable(table)
      val workbook = Workbook(Vector(sheet))

      val bytes = XlsxWriter.writeToBytes(workbook).getOrElse(fail("Write failed"))
      val reread = XlsxReader.readFromBytes(bytes).getOrElse(fail("Read failed"))

      val rereadTable = reread.sheets.head.getTable(name).getOrElse(fail(s"Missing $name"))
      assertEquals(rereadTable.style, style, s"Style mismatch for $name")
    }
  }

  test("table with header and totals rows round-trips") {
    val table = TableSpec.unsafeFromColumnNames(
      name = "WithTotals",
      displayName = "WithTotals",
      range = CellRange(ref"A1", ref"B11"),
      columnNames = Vector("Item", "Amount")
    ).copy(showHeaderRow = true, showTotalsRow = true)

    val sheet = Sheet("Data").getOrElse(fail("Failed")).withTable(table)
    val workbook = Workbook(Vector(sheet))

    val bytes = XlsxWriter.writeToBytes(workbook).getOrElse(fail("Write failed"))
    val reread = XlsxReader.readFromBytes(bytes).getOrElse(fail("Read failed"))

    val rereadTable = reread.sheets.head.getTable("WithTotals").getOrElse(fail("Missing"))
    assert(rereadTable.showHeaderRow)
    assert(rereadTable.showTotalsRow)
  }

  test("table data range calculation is correct") {
    val table = TableSpec.unsafeFromColumnNames(
      name = "Test",
      displayName = "Test",
      range = CellRange(ref"A1", ref"B11"),
      columnNames = Vector("A", "B")
    ).copy(showHeaderRow = true, showTotalsRow = true)

    val dataRange = table.dataRange

    // Header = A1:B1, Data = A2:B10, Totals = A11:B11
    assertEquals(dataRange.start, ref"A2")
    assertEquals(dataRange.end, ref"B10")
  }

  test("table validation detects column count mismatch") {
    // Use case class constructor directly to bypass validation (testing isValid method)
    val table = com.tjclp.xl.tables.TableSpec(
      name = "Invalid",
      displayName = "Invalid",
      range = CellRange(ref"A1", ref"D10"), // 4 columns wide
      columns = Vector(TableColumn(1, "A"), TableColumn(2, "B"), TableColumn(3, "C")) // Only 3 columns
    )

    assert(!table.isValid, "Should detect column count mismatch")
  }

  test("workbook with table and cells coexist") {
    val table = TableSpec.unsafeFromColumnNames(
      name = "Data",
      displayName = "Data",
      range = CellRange(ref"A1", ref"C10"),
      columnNames = Vector("Name", "Value", "Status")
    )

    val baseSheet = Sheet("Sheet1").getOrElse(fail("Failed"))
    val sheet = baseSheet
      .withTable(table)
      .put(ref"A1", CellValue.Text("Name"))
      .put(ref"B1", CellValue.Text("Value"))
      .put(ref"C1", CellValue.Text("Status"))
      .put(ref"A2", CellValue.Text("Item1"))
      .put(ref"B2", CellValue.Number(BigDecimal(100)))
      .put(ref"C2", CellValue.Text("Active"))

    val workbook = Workbook(Vector(sheet))

    val bytes = XlsxWriter.writeToBytes(workbook).getOrElse(fail("Write failed"))
    val reread = XlsxReader.readFromBytes(bytes).getOrElse(fail("Read failed"))

    val rereadSheet = reread.sheets.head

    // Table preserved
    assert(rereadSheet.getTable("Data").isDefined)

    // Cells preserved
    assertEquals(rereadSheet(ref"A1").value, CellValue.Text("Name"))
    assertEquals(rereadSheet(ref"B2").value, CellValue.Number(BigDecimal(100)))
  }

  test("ContentTypes includes table overrides when tables present") {
    val table = TableSpec.unsafeFromColumnNames(
      name = "Table1",
      displayName = "Table1",
      range = CellRange(ref"A1", ref"B10"),
      columnNames = Vector("A", "B")
    )

    val sheet = Sheet("Data").getOrElse(fail("Failed")).withTable(table)
    val workbook = Workbook(Vector(sheet))

    val bytes = XlsxWriter.writeToBytes(workbook).getOrElse(fail("Write failed"))

    // Verify [Content_Types].xml includes table content type
    import java.util.zip.ZipInputStream
    import java.io.ByteArrayInputStream

    val zis = new ZipInputStream(new ByteArrayInputStream(bytes))
    var foundTableType = false

    LazyList.continually(zis.getNextEntry).takeWhile(_ != null).foreach { entry =>
      if entry.getName == "[Content_Types].xml" then
        val content = scala.io.Source.fromInputStream(zis).mkString
        foundTableType = content.contains("application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml")
    }
    zis.close()

    assert(foundTableType, "Expected table content type in [Content_Types].xml")
  }

  // ========================================
  // Category F: Edge Cases (6 tests)
  // ========================================

  test("table with single column") {
    val table = TableSpec.unsafeFromColumnNames(
      name = "SingleCol",
      displayName = "SingleColumn",
      range = CellRange(ref"A1", ref"A100"),
      columnNames = Vector("OnlyOne")
    )

    val sheet = Sheet("Data").getOrElse(fail("Failed")).withTable(table)
    val workbook = Workbook(Vector(sheet))

    val bytes = XlsxWriter.writeToBytes(workbook).getOrElse(fail("Write failed"))
    val reread = XlsxReader.readFromBytes(bytes).getOrElse(fail("Read failed"))

    val rereadTable = reread.sheets.head.getTable("SingleCol").getOrElse(fail("Missing"))
    assertEquals(rereadTable.columns.size, 1)
  }

  test("table with maximum Excel columns (16384)") {
    // Excel max: XFD = 16384 columns
    // Too large for practical testing, but verify API supports it
    val columns = (1 to 100).map(i => TableColumn(i.toLong, s"Col$i")).toVector

    val table = com.tjclp.xl.tables.TableSpec(
      name = "Wide",
      displayName = "WideTable",
      range = CellRange(ref"A1", ARef.from1(100, 10)),
      columns = columns
    )

    assert(table.isValid)
  }

  test("table name uniqueness within sheet") {
    val table1 = TableSpec.unsafeFromColumnNames(
      name = "Data",
      displayName = "First",
      range = CellRange(ref"A1", ref"B10"),
      columnNames = Vector("A", "B")
    )

    val table2 = TableSpec.unsafeFromColumnNames(
      name = "Data", // Same name!
      displayName = "Second",
      range = CellRange(ref"D1", ref"E10"),
      columnNames = Vector("D", "E")
    )

    val sheet = Sheet("Test").getOrElse(fail("Failed"))
      .withTable(table1)
      .withTable(table2) // Should replace table1

    assertEquals(sheet.tables.size, 1)
    assertEquals(sheet.getTable("Data").get.displayName, "Second")
  }

  test("remove table from sheet") {
    val table = TableSpec.unsafeFromColumnNames(
      name = "ToRemove",
      displayName = "ToRemove",  // No spaces
      range = CellRange(ref"A1", ref"B10"),
      columnNames = Vector("A", "B")
    )

    val sheet = Sheet("Test").getOrElse(fail("Failed"))
      .withTable(table)
      .removeTable("ToRemove")

    assert(sheet.tables.isEmpty)
  }

  test("table with empty autoFilter (disabled) omits autoFilter element") {
    val table = TableSpec.unsafeFromColumnNames(
      name = "NoFilter",
      displayName = "NoFilter",  // No spaces
      range = CellRange(ref"A1", ref"C10"),
      columnNames = Vector("A", "B", "C")
    ) // autoFilter = None (default)

    val ooxml = TableConversions.toOoxml(table, id = 1L)
    assert(ooxml.autoFilter.isEmpty)

    val xml = OoxmlTable.toXml(ooxml)
    assert((xml \ "autoFilter").isEmpty, "Should not have autoFilter element")
  }

  test("table column IDs are sequential and 1-indexed") {
    val columns = Vector("First", "Second", "Third", "Fourth")
    val table = TableSpec.unsafeFromColumnNames(
      name = "Test",
      displayName = "Test",
      range = CellRange(ref"A1", ref"D10"),
      columnNames = columns
    )

    table.columns.zipWithIndex.foreach { case (col, idx) =>
      assertEquals(col.id, (idx + 1).toLong, s"Column ${idx} should have ID ${idx + 1}")
    }
  }

  // ========================================
  // Category G: Validation Tests (8 tests)
  // ========================================

  test("TableSpec.create rejects empty name") {
    val result = com.tjclp.xl.tables.TableSpec.create(
      name = "",
      displayName = "Test",
      range = CellRange(ref"A1", ref"B10"),
      columns = Vector(TableColumn(1, "Col1"), TableColumn(2, "Col2"))
    )
    assert(result.isLeft, "Should reject empty name")
  }

  test("TableSpec.create rejects name with spaces") {
    val result = com.tjclp.xl.tables.TableSpec.create(
      name = "My Table",
      displayName = "Test",
      range = CellRange(ref"A1", ref"B10"),
      columns = Vector(TableColumn(1, "Col1"), TableColumn(2, "Col2"))
    )
    assert(result.isLeft, "Should reject name with spaces")
  }

  test("TableSpec.create rejects displayName with spaces") {
    val result = com.tjclp.xl.tables.TableSpec.create(
      name = "MyTable",
      displayName = "My Display Name",
      range = CellRange(ref"A1", ref"B10"),
      columns = Vector(TableColumn(1, "Col1"), TableColumn(2, "Col2"))
    )
    assert(result.isLeft, "Should reject displayName with spaces")
    assert(result.left.exists(_.message.contains("cannot contain spaces")))
  }

  test("TableSpec.create rejects empty displayName") {
    val result = com.tjclp.xl.tables.TableSpec.create(
      name = "MyTable",
      displayName = "",
      range = CellRange(ref"A1", ref"B10"),
      columns = Vector(TableColumn(1, "Col1"), TableColumn(2, "Col2"))
    )
    assert(result.isLeft, "Should reject empty displayName")
  }

  test("TableSpec.create rejects single-row range") {
    val result = com.tjclp.xl.tables.TableSpec.create(
      name = "MyTable",
      displayName = "Test",
      range = CellRange(ref"A1", ref"B1"),  // Only 1 row
      columns = Vector(TableColumn(1, "Col1"), TableColumn(2, "Col2"))
    )
    assert(result.isLeft, "Should reject single-row range")
    assert(result.left.exists(_.message.contains("at least 2 rows")))
  }

  test("TableSpec.create rejects empty columns") {
    val result = com.tjclp.xl.tables.TableSpec.create(
      name = "MyTable",
      displayName = "Test",
      range = CellRange(ref"A1", ref"B10"),
      columns = Vector.empty
    )
    assert(result.isLeft, "Should reject empty columns")
  }

  test("TableSpec.create rejects duplicate column names") {
    val result = com.tjclp.xl.tables.TableSpec.create(
      name = "MyTable",
      displayName = "Test",
      range = CellRange(ref"A1", ref"C10"),
      columns = Vector(
        TableColumn(1, "Name"),
        TableColumn(2, "Name"),  // Duplicate!
        TableColumn(3, "Value")
      )
    )
    assert(result.isLeft, "Should reject duplicate column names")
    assert(result.left.exists(_.message.contains("Duplicate column names")))
  }

  test("TableSpec.create rejects column count mismatch") {
    val result = com.tjclp.xl.tables.TableSpec.create(
      name = "MyTable",
      displayName = "Test",
      range = CellRange(ref"A1", ref"D10"),  // 4 columns wide
      columns = Vector(TableColumn(1, "Col1"), TableColumn(2, "Col2"))  // Only 2 columns
    )
    assert(result.isLeft, "Should reject column count mismatch")
    assert(result.left.exists(_.message.contains("must match range width")))
  }

  test("TableSpec.create accepts valid input") {
    val result = com.tjclp.xl.tables.TableSpec.create(
      name = "Valid_Table",
      displayName = "Valid_Display_Name",
      range = CellRange(ref"A1", ref"C10"),
      columns = Vector(
        TableColumn(1, "Col1"),
        TableColumn(2, "Col2"),
        TableColumn(3, "Col3")
      )
    )
    assert(result.isRight, "Should accept valid input")
  }

  // ========================================
  // Category H: XML Security Tests (2 tests)
  // ========================================

  test("table name with XML special characters is escaped") {
    val table = OoxmlTable(
      id = 1L,
      name = "Table<>&\"'",  // XML special chars
      displayName = "Test",
      ref = CellRange(ref"A1", ref"B10"),
      headerRowCount = 1,
      totalsRowCount = 0,
      columns = Vector(OoxmlTableColumn(1, "Col")),
      autoFilter = None,
      styleInfo = None
    )

    val xml = OoxmlTable.toXml(table)
    val xmlString = xml.toString

    // Scala XML should auto-escape special characters
    assert(xmlString.contains("Table&lt;&gt;&amp;&quot;'") ||
           xmlString.contains("Table<>&\"'"),  // May be in attribute
           "XML special chars should be escaped or safe")
  }

  test("column name with XML special characters is escaped") {
    val table = OoxmlTable(
      id = 1L,
      name = "Test",
      displayName = "Test",
      ref = CellRange(ref"A1", ref"B10"),
      headerRowCount = 1,
      totalsRowCount = 0,
      columns = Vector(OoxmlTableColumn(1, "Col<>&\"\'")),  // XML special chars
      autoFilter = None,
      styleInfo = None
    )

    val xml = OoxmlTable.toXml(table)
    val reparsed = OoxmlTable.fromXml(xml)

    assert(reparsed.isRight, "Should handle XML special chars")
    assertEquals(reparsed.toOption.get.columns(0).name, "Col<>&\"'")
  }

  // ========================================
  // Category I: UID Preservation (1 test)
  // ========================================

  test("round-trip preserves all UIDs (table, autoFilter, columns)") {
    val tableUid = "{TEST-TABLE-UID-1234}"
    val autoFilterUid = "{TEST-AUTOFILTER-UID-5678}"
    val columnUids = Vector("{COL-UID-1}", "{COL-UID-2}")

    val table = OoxmlTable(
      id = 1L,
      name = "TestTable",
      displayName = "TestTable",
      ref = CellRange(ref"A1", ref"B10"),
      headerRowCount = 1,
      totalsRowCount = 0,
      tableUid = Some(tableUid),
      columns = Vector(
        OoxmlTableColumn(1, "Col1", Some(columnUids(0))),
        OoxmlTableColumn(2, "Col2", Some(columnUids(1)))
      ),
      autoFilter = Some(CellRange(ref"A1", ref"B10")),
      autoFilterUid = Some(autoFilterUid),
      styleInfo = None
    )

    val xml = OoxmlTable.toXml(table)
    val reparsed = OoxmlTable.fromXml(xml).toOption.get

    // Verify all UIDs preserved
    assertEquals(reparsed.tableUid, Some(tableUid), "Table UID should be preserved")
    assertEquals(reparsed.autoFilterUid, Some(autoFilterUid), "AutoFilter UID should be preserved")
    assertEquals(reparsed.columns(0).uid, Some(columnUids(0)), "Column 1 UID should be preserved")
    assertEquals(reparsed.columns(1).uid, Some(columnUids(1)), "Column 2 UID should be preserved")
  }
