package com.tjclp.xl.io

import java.io.{ByteArrayInputStream, FileOutputStream}
import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipOutputStream}

import cats.effect.IO
import cats.syntax.all.*
import fs2.Stream
import munit.CatsEffectSuite
import org.scalacheck.Gen
import org.scalacheck.rng.Seed

import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.ooxml.{SharedStringIndex, SharedStrings}

/**
 * Every reader reads a `t="s"` value with SharedStringIndex, the rule the writer's reference scan
 * and the lint use: a value names entry n when the rule reads it as n and n is below the table's
 * size. A value that names no entry reads as each reader's form for nothing: empty in the streaming
 * readers and the single-cell reader, `#REF!` (an index past the table) or `#VALUE!` (any other
 * value) in XlsxReader.
 */
class SharedStringIndexStreamingSpec extends CatsEffectSuite:

  private val table = SharedStrings.fromStrings(Vector("e0", "e1", "e2", "e3"))

  private val fixed = Vector(
    "0",
    "3",
    "+3",
    "+03",
    "03",
    " 3 ",
    " +3 ",
    "\t3",
    "3\n",
    "3.0",
    "3.",
    "4",
    "-0",
    "-00",
    "-1",
    "+-3",
    "+ 3",
    "3 3",
    "3e0",
    "3.5",
    "٣",
    "３",
    "",
    ("0" * 30) + "2",
    "2147483648"
  )

  private val sampled: Vector[String] =
    val char = Gen.frequency(
      12 -> Gen.choose('0', '5'),
      2 -> Gen.oneOf(' ', '\t', '\n'),
      2 -> Gen.oneOf('+', '-'),
      1 -> Gen.oneOf('.', 'x', '٣', '３')
    )
    Gen
      .listOfN(300, Gen.choose(0, 6).flatMap(n => Gen.listOfN(n, char)).map(_.mkString))
      .apply(Gen.Parameters.default, Seed(768L))
      .getOrElse(Vector.empty)
      .toVector

  private val values = fixed ++ sampled

  private val sheetXml: String =
    val rows = values.zipWithIndex
      .map((v, i) => s"""<row r="${i + 1}"><c r="A${i + 1}" t="s"><v>$v</v></c></row>""")
      .mkString
    "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
      s"<sheetData>$rows</sheetData></worksheet>"

  private def expected(v: String): CellValue =
    SharedStringIndex
      .parse(v)
      .flatMap(table.apply)
      .map(table.toCellValue)
      .getOrElse(CellValue.Empty)

  private def byRow(rows: List[RowData]): Map[Int, CellValue] =
    rows.flatMap(r => r.cells.get(0).map(r.rowIndex -> _)).toMap

  private def check(name: String, read: IO[List[RowData]]): IO[Unit] =
    read.map { rows =>
      val got = byRow(rows)
      values.zipWithIndex.foreach { (v, i) =>
        assertEquals(got.getOrElse(i + 1, CellValue.Empty), expected(v), s"$name [$v]")
      }
    }

  test("SaxStreamingReader reads each t=\"s\" value as the shared rule does") {
    check(
      "sax",
      SaxStreamingReader
        .parseWorksheetStream[IO](
          new ByteArrayInputStream(sheetXml.getBytes(StandardCharsets.UTF_8)),
          Some(table)
        )
        .compile
        .toList
    )
  }

  test("StreamingXmlReader reads each t=\"s\" value as the shared rule does") {
    check(
      "fs2-xml",
      StreamingXmlReader
        .parseWorksheetStream[IO](
          Stream.emits(sheetXml.getBytes(StandardCharsets.UTF_8)).covary[IO],
          Some(table)
        )
        .compile
        .toList
    )
  }

  // ========== every reader agrees ==========

  private val strings = Vector("e0", "e1", "e2", "e3")

  /** A one-sheet package whose A1, A2, ... are `t="s"` cells holding `vs` over `strings`. */
  private def writeBook(vs: Vector[String]): IO[Path] = IO.blocking {
    val ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    val rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    val pkg = "http://schemas.openxmlformats.org/package/2006/relationships"
    val ct = "application/vnd.openxmlformats-officedocument.spreadsheetml"
    val path = Files.createTempFile("xl-sst-index-readers-", ".xlsx")
    val out = new ZipOutputStream(new FileOutputStream(path.toFile))
    def put(name: String, content: String): Unit =
      out.putNextEntry(new ZipEntry(name))
      out.write(content.getBytes(StandardCharsets.UTF_8))
      out.closeEntry()
    put(
      "[Content_Types].xml",
      """<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">""" +
        """<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>""" +
        """<Default Extension="xml" ContentType="application/xml"/>""" +
        s"""<Override PartName="/xl/workbook.xml" ContentType="$ct.sheet.main+xml"/>""" +
        s"""<Override PartName="/xl/worksheets/sheet1.xml" ContentType="$ct.worksheet+xml"/>""" +
        s"""<Override PartName="/xl/sharedStrings.xml" ContentType="$ct.sharedStrings+xml"/>""" +
        "</Types>"
    )
    put(
      "_rels/.rels",
      s"""<?xml version="1.0"?><Relationships xmlns="$pkg">""" +
        s"""<Relationship Id="rId1" Type="$rel/officeDocument" Target="xl/workbook.xml"/></Relationships>"""
    )
    put(
      "xl/workbook.xml",
      s"""<?xml version="1.0"?><workbook xmlns="$ns" xmlns:r="$rel"><sheets>""" +
        """<sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"""
    )
    put(
      "xl/_rels/workbook.xml.rels",
      s"""<?xml version="1.0"?><Relationships xmlns="$pkg">""" +
        s"""<Relationship Id="rId1" Type="$rel/worksheet" Target="worksheets/sheet1.xml"/>""" +
        s"""<Relationship Id="rId2" Type="$rel/sharedStrings" Target="sharedStrings.xml"/>""" +
        "</Relationships>"
    )
    put(
      "xl/sharedStrings.xml",
      s"""<?xml version="1.0" encoding="UTF-8"?><sst xmlns="$ns" count="${vs.size}" """ +
        s"""uniqueCount="${strings.size}">${strings
            .map(t => s"<si><t>$t</t></si>")
            .mkString}</sst>"""
    )
    val rows = vs.zipWithIndex
      .map((v, i) => s"""<row r="${i + 1}"><c r="A${i + 1}" t="s"><v>$v</v></c></row>""")
      .mkString
    put(
      "xl/worksheets/sheet1.xml",
      s"""<worksheet xmlns="$ns"><sheetData>$rows</sheetData></worksheet>"""
    )
    out.close()
    path
  }

  test("XlsxReader, readStream and streamCellDetails resolve each value to the same entry") {
    val vs = fixed ++ sampled.take(40)
    val excel = ExcelIO.instance[IO]
    writeBook(vs).flatMap { path =>
      for
        wb <- excel.read(path)
        rows <- excel.readStream(path).compile.toList
        details <- vs.indices.toVector.traverse { i =>
          excel.streamCellDetails(path, "Sheet1", ARef.from0(0, i)).map(_.value)
        }
        _ <- IO.blocking(Files.deleteIfExists(path))
      yield
        val sheet = wb.sheets.headOption.getOrElse(fail("no sheet"))
        val streamed = byRow(rows)
        vs.zipWithIndex.foreach { (v, i) =>
          val named = SharedStringIndex.parse(v).filter(_ < strings.size).map(strings)
          val nothing = SharedStringIndex.parse(v) match
            case Some(_) => CellValue.Error(CellError.Ref)
            case None => CellValue.Error(CellError.Value)
          val whole = sheet.cells.get(ARef.from0(0, i)).map(_.value)
          assertEquals(whole, Some(named.fold(nothing)(CellValue.Text(_))), s"XlsxReader [$v]")
          val blank = named.fold(CellValue.Empty)(CellValue.Text(_))
          assertEquals(streamed.getOrElse(i + 1, CellValue.Empty), blank, s"readStream [$v]")
          assertEquals(details(i), blank, s"streamCellDetails [$v]")
        }
    }
  }
