package com.tjclp.xl.io.streaming

import java.nio.file.{Files, Path}
import java.time.LocalDateTime
import java.util.zip.ZipFile

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.io.{ExcelIO, RowData}
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.XlsxReader

/**
 * GH-378: streaming writes must serialize formula cached values for every cached type.
 *
 * DateTime caches are written as the Excel serial with t="n" (matching what Excel itself stores
 * after recalculation); Text caches carry t="str". Dropping the cached <v> strands `=TODAY()` /
 * `=EDATE(...)` cells uncached, so openpyxl data_only / pandas / previewers see blanks.
 */
class StreamingFormulaCachedValueSpec extends CatsEffectSuite:

  private val excel = ExcelIO.instance[IO]

  private def tempXlsx(label: String): Path =
    val p = Files.createTempFile(s"xl-stream-cached-$label-", ".xlsx")
    p.toFile.deleteOnExit()
    p

  private def entryText(path: Path, name: String): String =
    val zip = new ZipFile(path.toFile)
    try
      Option(zip.getEntry(name)) match
        case Some(e) => new String(zip.getInputStream(e).readAllBytes(), "UTF-8")
        case None => fail(s"zip entry $name not found")
    finally zip.close()

  test("writeStream serializes formula cached DateTime as Excel serial with t=\"n\" (GH-378)") {
    val path = tempXlsx("datetime")
    // Jan 1, 2000 noon = serial 36526.5 exactly
    val cached = LocalDateTime.of(2000, 1, 1, 12, 0, 0)
    val rows = fs2.Stream
      .emits(
        List(
          RowData(1, Map(0 -> CellValue.DateTime(cached))),
          RowData(2, Map(0 -> CellValue.Formula("EDATE(A1,0)", Some(CellValue.DateTime(cached)))))
        )
      )
      .covary[IO]
    rows.through(excel.writeStream(path, "Data")).compile.drain.map { _ =>
      val sheetXml = entryText(path, "xl/worksheets/sheet1.xml")
      val cellXml = """(?s)<c r="A2".*?</c>""".r
        .findFirstIn(sheetXml)
        .getOrElse(fail(s"A2 cell missing from worksheet XML: $sheetXml"))
      assert(cellXml.contains("""t="n""""), cellXml)
      assert(cellXml.contains("<v>36526.5</v>"), cellXml)

      val wb = XlsxReader.read(path).fold(e => fail(s"read: ${e.message}"), identity)
      wb.sheets(0)(ref"A2").value match
        case CellValue.Formula(expr, Some(CellValue.Number(serial))) =>
          assertEquals(expr, "EDATE(A1,0)")
          assertEquals(serial.toDouble, 36526.5, 0.000001)
        case other => fail(s"Expected Formula with cached serial Number, got $other")
    }
  }

  test("writeStream marks formula cached Text with t=\"str\" (GH-378)") {
    val path = tempXlsx("text")
    val rows = fs2.Stream
      .emits(
        List(
          RowData(
            1,
            Map(0 -> CellValue.Formula("""CONCATENATE("a","b")""", Some(CellValue.Text("ab"))))
          )
        )
      )
      .covary[IO]
    rows.through(excel.writeStream(path, "Data")).compile.drain.map { _ =>
      val sheetXml = entryText(path, "xl/worksheets/sheet1.xml")
      val cellXml = """(?s)<c r="A1".*?</c>""".r
        .findFirstIn(sheetXml)
        .getOrElse(fail(s"A1 cell missing from worksheet XML: $sheetXml"))
      assert(cellXml.contains("""t="str""""), cellXml)
      assert(cellXml.contains("<v>ab</v>"), cellXml)

      val wb = XlsxReader.read(path).fold(e => fail(s"read: ${e.message}"), identity)
      wb.sheets(0)(ref"A1").value match
        case CellValue.Formula(expr, Some(CellValue.Text(t))) =>
          assertEquals(expr, """CONCATENATE("a","b")""")
          assertEquals(t, "ab")
        case other => fail(s"Expected Formula with cached Text, got $other")
    }
  }
