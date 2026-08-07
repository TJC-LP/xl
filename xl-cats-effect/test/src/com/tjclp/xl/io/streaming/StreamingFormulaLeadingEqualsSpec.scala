package com.tjclp.xl.io.streaming

import java.io.ByteArrayOutputStream
import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.ZipFile

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.io.{ExcelIO, RowData}
import com.tjclp.xl.ooxml.StaxSaxWriter

/**
 * GH-456: the streaming writers are `<f>` emission boundaries too — a scripting-authored
 * `CellValue.Formula("=A1*2", ...)` must serialize as `<f>A1*2</f>` (the OOXML store carries the
 * expression without the display form's leading '='), matching the in-memory writers.
 */
class StreamingFormulaLeadingEqualsSpec extends CatsEffectSuite:

  private val excel = ExcelIO.instance[IO]

  private def tempXlsx(label: String): Path =
    val p = Files.createTempFile(s"xl-fixture-leading-eq-$label-", ".xlsx")
    p.toFile.deleteOnExit()
    p

  private def entryText(path: Path, name: String): String =
    val zip = new ZipFile(path.toFile)
    try
      Option(zip.getEntry(name)) match
        case Some(e) => new String(zip.getInputStream(e).readAllBytes(), StandardCharsets.UTF_8)
        case None => fail(s"zip entry $name not found")
    finally zip.close()

  test("writeStream (StreamingXmlWriter) strips the leading '=' from <f> (GH-456)") {
    val path = tempXlsx("write")
    val rows = fs2.Stream
      .emits(
        List(
          RowData(1, Map(0 -> CellValue.Number(2))),
          RowData(2, Map(0 -> CellValue.Formula("=A1*2", None)))
        )
      )
      .covary[IO]
    rows.through(excel.writeStream(path, "Data")).compile.drain.map { _ =>
      val sheetXml = entryText(path, "xl/worksheets/sheet1.xml")
      assert(sheetXml.contains("<f>A1*2</f>"), sheetXml)
      assert(!sheetXml.contains("<f>="), sheetXml)
    }
  }

  test("StreamingTransform.writeCellContent strips the leading '=' from <f> (GH-456)") {
    val out = new ByteArrayOutputStream()
    val writer = StaxSaxWriter.create(out)
    writer.startElement("c")
    StreamingTransform.writeCellContent(writer, CellValue.Formula("=A1*2", None))
    writer.endElement()
    writer.flush()
    val xml = out.toString("UTF-8")
    assert(xml.contains("<f>A1*2</f>"), xml)

    val clean = new ByteArrayOutputStream()
    val cleanWriter = StaxSaxWriter.create(clean)
    cleanWriter.startElement("c")
    StreamingTransform.writeCellContent(cleanWriter, CellValue.Formula("IF(A1=1,2,3)", None))
    cleanWriter.endElement()
    cleanWriter.flush()
    assert(clean.toString("UTF-8").contains("<f>IF(A1=1,2,3)</f>"), clean.toString("UTF-8"))
  }
