package com.tjclp.xl.io.streaming

import java.nio.file.{Files, Path}
import java.util.zip.ZipFile

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.io.{ExcelIO, RowData, StyledRowData}
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.{XlsxReader, XmlSecurity}
import com.tjclp.xl.ooxml.writer.{SstPolicy, WriterConfig}
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.font.Font
import com.tjclp.xl.styles.numfmt.NumFmt

/**
 * GH-223: two-pass streaming writes — SST deduplication and style registry.
 *
 * Pass 1 accumulates unique strings (first-occurrence order) and the style table while rows
 * stream; pass 2 emits cells as t="s" SST references plus s= style indices, then writes
 * sharedStrings.xml / styles.xml from the accumulators. Memory: O(distinct strings), per the
 * smart-streaming design envelope.
 *
 * The inline-string dialect stays available behind the existing config: SstPolicy.Never.
 */
class StreamingSstSpec extends CatsEffectSuite:

  private val excel = ExcelIO.instance[IO]

  private def tempXlsx(label: String): Path =
    val p = Files.createTempFile(s"xl-stream-sst-$label-", ".xlsx")
    p.toFile.deleteOnExit()
    p

  private def entryText(path: Path, name: String): String =
    val zip = new ZipFile(path.toFile)
    try
      Option(zip.getEntry(name)) match
        case Some(e) => new String(zip.getInputStream(e).readAllBytes(), "UTF-8")
        case None => fail(s"zip entry $name not found in ${zip.entries()}")
    finally zip.close()

  private def hasEntry(path: Path, name: String): Boolean =
    val zip = new ZipFile(path.toFile)
    try Option(zip.getEntry(name)).isDefined
    finally zip.close()

  /** (count, uniqueCount, entries in document order) from xl/sharedStrings.xml. */
  private def parseSst(path: Path): (Int, Int, Vector[String]) =
    val xml = entryText(path, "xl/sharedStrings.xml")
    val elem = XmlSecurity.parseSafe(xml, "sst").fold(e => fail(s"sst parse: ${e.message}"), identity)
    val count = (elem \@ "count").toInt
    val unique = (elem \@ "uniqueCount").toInt
    val entries = (elem \ "si").map(si => (si \ "t").text).toVector
    (count, unique, entries)

  private val dupRows = fs2.Stream
    .emits(
      List(
        RowData(1, Map(0 -> CellValue.Text("alpha"), 1 -> CellValue.Number(BigDecimal(1)))),
        RowData(2, Map(0 -> CellValue.Text("beta"))),
        RowData(3, Map(0 -> CellValue.Text("alpha"))),
        RowData(4, Map(0 -> CellValue.Text("gamma"))),
        RowData(5, Map(0 -> CellValue.Text("beta")))
      )
    )
    .covary[IO]

  test("writeStream (default) emits SST: first-occurrence order, count=refs, uniqueCount=distinct") {
    val path = tempXlsx("dedup")
    dupRows.through(excel.writeStream(path, "Data")).compile.drain.map { _ =>
      assert(hasEntry(path, "xl/sharedStrings.xml"), "sharedStrings.xml missing")
      val (count, unique, entries) = parseSst(path)
      assertEquals(entries, Vector("alpha", "beta", "gamma"), "SST order must be first occurrence")
      assertEquals(unique, 3)
      assertEquals(count, 5, "count attribute counts references, not unique strings")

      val sheetXml = entryText(path, "xl/worksheets/sheet1.xml")
      assert(sheetXml.contains("t=\"s\""), s"cells must reference SST by index: $sheetXml")
      assert(!sheetXml.contains("inlineStr"), s"no inline strings expected in SST mode: $sheetXml")
      // First occurrence of alpha is index 0, beta 1, gamma 2
      assert(sheetXml.contains("<v>0</v>"), "alpha must be referenced as index 0")
      assert(sheetXml.contains("<v>2</v>"), "gamma must be referenced as index 2")
    }
  }

  test("writeStream with SstPolicy.Never keeps the inline-string dialect (no SST part)") {
    val path = tempXlsx("inline")
    dupRows
      .through(excel.writeStream(path, "Data", config = WriterConfig(sstPolicy = SstPolicy.Never)))
      .compile
      .drain
      .map { _ =>
        assert(!hasEntry(path, "xl/sharedStrings.xml"), "SstPolicy.Never must not emit an SST")
        val sheetXml = entryText(path, "xl/worksheets/sheet1.xml")
        assert(sheetXml.contains("inlineStr"), s"inline dialect expected: $sheetXml")
      }
  }

  test("SST-written file reads back identical values via the in-memory reader") {
    val path = tempXlsx("readback")
    dupRows.through(excel.writeStream(path, "Data")).compile.drain.map { _ =>
      val wb = XlsxReader.read(path).fold(e => fail(s"read: ${e.message}"), identity)
      val sheet = wb.sheets(0)
      assertEquals(sheet(ref"A1").value, CellValue.Text("alpha"))
      assertEquals(sheet(ref"A3").value, CellValue.Text("alpha"))
      assertEquals(sheet(ref"A4").value, CellValue.Text("gamma"))
      assertEquals(sheet(ref"B1").value, CellValue.Number(BigDecimal(1)))
    }
  }

  test("RichText cells stay inline while plain text uses the SST (mixed dialect is valid OOXML)") {
    import com.tjclp.xl.richtext.{RichText, TextRun}
    val rich = RichText(Vector(TextRun("bold bit", Some(Font.default.copy(bold = true)))))
    val rows = fs2.Stream
      .emits(
        List(
          RowData(1, Map(0 -> CellValue.Text("plain"), 1 -> CellValue.RichText(rich)))
        )
      )
      .covary[IO]
    val path = tempXlsx("richtext")
    rows.through(excel.writeStream(path, "Data")).compile.drain.map { _ =>
      val sheetXml = entryText(path, "xl/worksheets/sheet1.xml")
      assert(sheetXml.contains("t=\"s\""), "plain text must use SST")
      assert(sheetXml.contains("inlineStr"), "rich text stays inline")
      val (_, unique, entries) = parseSst(path)
      assertEquals(entries, Vector("plain"))
      assertEquals(unique, 1)
      val wb = XlsxReader.read(path).fold(e => fail(s"read: ${e.message}"), identity)
      assertEquals(wb.sheets(0)(ref"A1").value, CellValue.Text("plain"))
    }
  }

  test("formula-injection escaping applies before SST accumulation (secure config)") {
    val rows = fs2.Stream
      .emits(List(RowData(1, Map(0 -> CellValue.Text("=cmd()")))))
      .covary[IO]
    val path = tempXlsx("secure")
    rows
      .through(excel.writeStream(path, "Data", config = WriterConfig.secure))
      .compile
      .drain
      .map { _ =>
        val (_, _, entries) = parseSst(path)
        assertEquals(entries, Vector("'=cmd()"), "SST must store the escaped text")
      }
  }

  test("writeStreamsSeq shares a single SST across sheets (workbook-global dedup)") {
    val sheet1 = fs2.Stream
      .emits(List(RowData(1, Map(0 -> CellValue.Text("shared"), 1 -> CellValue.Text("only-1")))))
      .covary[IO]
    val sheet2 = fs2.Stream
      .emits(List(RowData(1, Map(0 -> CellValue.Text("shared"), 1 -> CellValue.Text("only-2")))))
      .covary[IO]
    val path = tempXlsx("multisheet")
    excel.writeStreamsSeq(path, Seq("S1" -> sheet1, "S2" -> sheet2)).map { _ =>
      val (count, unique, entries) = parseSst(path)
      assertEquals(entries, Vector("shared", "only-1", "only-2"))
      assertEquals(unique, 3)
      assertEquals(count, 4)
      val wb = XlsxReader.read(path).fold(e => fail(s"read: ${e.message}"), identity)
      assertEquals(wb.sheets(0)(ref"A1").value, CellValue.Text("shared"))
      assertEquals(wb.sheets(1)(ref"A1").value, CellValue.Text("shared"))
      assertEquals(wb.sheets(1)(ref"B1").value, CellValue.Text("only-2"))
    }
  }

  test("writeStreamWithAutoDetect emits both the dimension element and the SST") {
    val path = tempXlsx("autodetect")
    dupRows.through(excel.writeStreamWithAutoDetect(path, "Data")).compile.drain.map { _ =>
      val sheetXml = entryText(path, "xl/worksheets/sheet1.xml")
      assert(sheetXml.contains("<dimension ref=\"A1:B5\"/>"), s"dimension missing: $sheetXml")
      val (_, unique, entries) = parseSst(path)
      assertEquals(entries, Vector("alpha", "beta", "gamma"))
      assertEquals(unique, 3)
    }
  }

  test("writeStreamStyled emits styles.xml from the style table and cells round-trip styles") {
    val bold = CellStyle.default.copy(font = Font.default.copy(bold = true))
    val percent = CellStyle.default.withNumFmt(NumFmt.Percent)
    val rows = fs2.Stream
      .emits(
        List(
          StyledRowData(
            1,
            Map(0 -> CellValue.Text("Header"), 1 -> CellValue.Number(BigDecimal("0.25"))),
            Map(0 -> 0, 1 -> 1) // indices into the style table below
          ),
          StyledRowData(2, Map(0 -> CellValue.Text("Plain")), Map.empty)
        )
      )
      .covary[IO]
    val path = tempXlsx("styled")
    rows
      .through(excel.writeStreamStyled(path, "Data", Vector(bold, percent)))
      .compile
      .drain
      .map { _ =>
        val wb = XlsxReader.read(path).fold(e => fail(s"read: ${e.message}"), identity)
        val sheet = wb.sheets(0)

        def resolved(r: com.tjclp.xl.addressing.ARef): CellStyle =
          sheet(r).styleId
            .flatMap(sheet.styleRegistry.get)
            .getOrElse(CellStyle.default)

        assert(resolved(ref"A1").font.bold, "A1 must resolve to the bold style")
        assertEquals(resolved(ref"B1").numFmt, NumFmt.Percent, "B1 must resolve to percent format")
        assertEquals(sheet(ref"A2").styleId, None, "unstyled cell stays default")

        // Values survive alongside styles (SST still applies)
        assertEquals(sheet(ref"A1").value, CellValue.Text("Header"))
        assertEquals(sheet(ref"B1").value, CellValue.Number(BigDecimal("0.25")))
        val (_, unique, _) = parseSst(path)
        assertEquals(unique, 2)
      }
  }

  test("memory envelope: 100k rows with 100 distinct strings -> SST has exactly 100 entries") {
    val path = tempXlsx("bigdedup")
    val rows = fs2.Stream
      .range(1, 100_001)
      .map(i => RowData(i, Map(0 -> CellValue.Text(s"Company-${i % 100}"))))
      .covary[IO]
    rows.through(excel.writeStream(path, "Big")).compile.drain.map { _ =>
      val (count, unique, _) = parseSst(path)
      assertEquals(unique, 100, "SST must contain only the distinct strings")
      assertEquals(count, 100_000, "count attribute counts every reference")
    }
  }
