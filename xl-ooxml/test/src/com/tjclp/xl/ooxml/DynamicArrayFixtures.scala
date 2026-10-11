package com.tjclp.xl.ooxml

import java.io.{ByteArrayInputStream, ByteArrayOutputStream}
import java.nio.charset.StandardCharsets
import java.util.zip.{ZipEntry, ZipInputStream, ZipOutputStream}

/**
 * The package Excel 365 writes for a dynamic-array formula (GH-695/GH-714), shared by the xl-ooxml,
 * xl-cats-effect and xl-cli suites: A1:A5 hold numbers, B1 is a dynamic-array anchor (`<f t="array"
 * ref="B1:B5">_xlfn._xlws.SORT(A1:A5)</f>` with `cm="1"` pointing at `xl/metadata.xml`'s XLDAPR
 * block), its spill is cached in B2:B5, and C1 = `SUM(_xlfn.ANCHORARRAY(B1))` caches 15.
 */
object DynamicArrayFixtures:

  val ctMetadataOverride: String =
    """<Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>"""

  val metadataRel: String =
    """<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/>"""

  /** Excel's metadata part for one plain (non-collapsed) dynamic array, byte for byte. */
  val metadataXml: String =
    """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      |<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"><metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes><futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>""".stripMargin

  /** The spill-anchor sheet: B1 is the dynamic record, C1 sums its spill. */
  val sortSheetXml: String =
    """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      |<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
      |<row r="1"><c r="A1"><v>5</v></c><c r="B1" cm="1"><f t="array" ref="B1:B5">_xlfn._xlws.SORT(A1:A5)</f><v>1</v></c><c r="C1"><f>SUM(_xlfn.ANCHORARRAY(B1))</f><v>15</v></c></row>
      |<row r="2"><c r="A2"><v>3</v></c><c r="B2"><v>2</v></c></row>
      |<row r="3"><c r="A3"><v>1</v></c><c r="B3"><v>3</v></c></row>
      |<row r="4"><c r="A4"><v>4</v></c><c r="B4"><v>4</v></c></row>
      |<row r="5"><c r="A5"><v>2</v></c><c r="B5"><v>5</v></c></row>
      |</sheetData></worksheet>""".stripMargin

  /**
   * The package's parts. `metadata = None` omits the metadata part, its Override and its
   * relationship (a `cm` then dangles); `Some(xml)` ships that part.
   */
  def parts(
    sheetXml: String = sortSheetXml,
    metadata: Option[String] = Some(metadataXml)
  ): Vector[(String, String)] =
    Vector(
      "[Content_Types].xml" ->
        ("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          |<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>""".stripMargin +
          (if metadata.isDefined then ctMetadataOverride else "") + "</Types>"),
      "_rels/.rels" ->
        """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          |<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>""".stripMargin,
      "xl/workbook.xml" ->
        """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          |<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>""".stripMargin,
      "xl/_rels/workbook.xml.rels" ->
        ("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          |<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>""".stripMargin +
          (if metadata.isDefined then metadataRel else "") + "</Relationships>"),
      "xl/worksheets/sheet1.xml" -> sheetXml
    ) ++ metadata.toList.map("xl/metadata.xml" -> _)

  /** Zip the parts into package bytes, in order. */
  def xlsx(parts: Vector[(String, String)] = parts()): Array[Byte] =
    val output = new ByteArrayOutputStream()
    val zip = new ZipOutputStream(output)
    try
      parts.foreach { case (name, text) =>
        zip.putNextEntry(new ZipEntry(name))
        zip.write(text.getBytes(StandardCharsets.UTF_8))
        zip.closeEntry()
      }
    finally zip.close()
    output.toByteArray

  /** Every entry of a package, by name. */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  def entries(bytes: Array[Byte]): Vector[(String, Array[Byte])] =
    val zip = new ZipInputStream(new ByteArrayInputStream(bytes))
    val out = Vector.newBuilder[(String, Array[Byte])]
    try
      var current = zip.getNextEntry
      while current != null do
        out += current.getName -> zip.readAllBytes()
        zip.closeEntry()
        current = zip.getNextEntry
    finally zip.close()
    out.result()

  /** One entry as UTF-8 text, if present. */
  def entryText(bytes: Array[Byte], name: String): Option[String] =
    entries(bytes).collectFirst { case (`name`, b) => new String(b, StandardCharsets.UTF_8) }
