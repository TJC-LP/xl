package com.tjclp.xl.io

import java.io.InputStream
import java.nio.file.{Files, Path, StandardCopyOption}

import cats.effect.IO
import cats.syntax.all.*
import munit.CatsEffectSuite

import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.ooxml.XlsxReader
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook

/**
 * Streaming/in-memory parity law over the real-file fixture corpus (GH-240).
 *
 * The SAX streaming reader (SaxStreamingReader + SaxSharedStringsReader) and the in-memory reader
 * (XlsxReader + WorksheetReader) are fully independent parse paths; a prior SST-indexing regression
 * proved they can drift. For every committed fixture - openpyxl inline-string dialect AND
 * LibreOffice SST dialect - every sheet must yield the same (row, column) -> CellValue map from
 * both paths.
 *
 * Scope of the law: cell VALUES and refs. Streaming legitimately omits workbook-level detail the
 * in-memory model carries: resolved CellStyle objects (only raw s= indices are surfaced via
 * RowData.cellStyles), merged ranges, comments, hyperlinks, column widths/row heights, and sheet
 * metadata. Styled-but-valueless cells (e.g. a border on an empty cell) are surfaced by the
 * in-memory reader as CellValue.Empty cells but skipped entirely by the streaming reader, so the
 * comparison excludes Empty values on both sides.
 */
class StreamingParitySpec extends CatsEffectSuite:

  private val excel = ExcelIO.instance[IO]

  /** Mirror of TestFixtures.all in xl-ooxml/test (test classpaths are not shared). */
  private val fixtures: List[String] = List(
    "small-values.xlsx",
    "styled.xlsx",
    "formulas.xlsx",
    "autofilter.xlsx",
    "chart-bar.xlsx",
    "image.xlsx",
    "comments-hyperlinks.xlsx",
    "small-values-lo.xlsx",
    "styled-lo.xlsx",
    "formulas-lo.xlsx"
  )

  private def fixturePath(name: String): Path =
    val resource = s"/fixtures/$name"
    val stream: InputStream = Option(getClass.getResourceAsStream(resource)) match
      case Some(s) => s
      case None => fail(s"Fixture not on classpath: $resource (xl-cats-effect.test.resources)")
    try
      val target = Files.createTempFile(s"xl-parity-${name.stripSuffix(".xlsx")}-", ".xlsx")
      target.toFile.deleteOnExit()
      Files.copy(stream, target, StandardCopyOption.REPLACE_EXISTING)
      target
    finally stream.close()

  /** (1-based row, 0-based col) -> value, excluding Empty (streaming never emits Empty). */
  private def inMemoryValues(sheet: Sheet): Map[(Int, Int), CellValue] =
    sheet.cells.collect {
      case (ref, cell) if cell.value != CellValue.Empty =>
        (ref.row.index1, ref.col.index0) -> cell.value
    }

  private def streamedValues(path: Path, sheetName: String): IO[Map[(Int, Int), CellValue]] =
    excel
      .readSheetStream(path, sheetName)
      .compile
      .toVector
      .map { rows =>
        rows.flatMap { row =>
          row.cells.map { case (colIdx, value) => (row.rowIndex, colIdx) -> value }
        }.toMap
      }

  private def loadInMemory(name: String, path: Path): IO[Workbook] =
    IO.fromEither(
      XlsxReader
        .read(path)
        .left
        .map(err => new Exception(s"$name in-memory read failed: ${err.message}"))
    )

  fixtures.foreach { fixture =>
    test(s"parity: $fixture - streaming rows agree with in-memory cells on every sheet") {
      val path = fixturePath(fixture)
      loadInMemory(fixture, path).flatMap { wb =>
        assert(wb.sheets.nonEmpty, s"$fixture has no sheets")
        wb.sheets.toList.traverse_ { sheet =>
          val expected = inMemoryValues(sheet)
          streamedValues(path, sheet.name.value).map { streamed =>
            assertEquals(
              streamed,
              expected,
              s"streaming vs in-memory drift in $fixture!${sheet.name.value} - " +
                s"only-streaming=${(streamed.keySet -- expected.keySet).toList.sorted}, " +
                s"only-memory=${(expected.keySet -- streamed.keySet).toList.sorted}"
            )
          }
        }
      }
    }
  }

  test("parity: readStream (first sheet) agrees with in-memory first sheet for every fixture") {
    fixtures.traverse_ { fixture =>
      val path = fixturePath(fixture)
      loadInMemory(fixture, path).flatMap { wb =>
        val first = wb.sheets(0)
        val expected = inMemoryValues(first)
        excel
          .readStream(path)
          .compile
          .toVector
          .map { rows =>
            val streamed = rows.flatMap { row =>
              row.cells.map { case (colIdx, value) => (row.rowIndex, colIdx) -> value }
            }.toMap
            assertEquals(streamed, expected, s"readStream drift on $fixture (first sheet)")
          }
      }
    }
  }

  // ========== GH-288: _xHHHH_ escape decoding parity ==========

  private val crText = "line1\rline2\r\nline3"
  private val literalEscape = "_x000D_ literal"

  private def crWorkbook: Workbook =
    import com.tjclp.xl.cells.Cell
    import com.tjclp.xl.macros.ref
    Workbook(
      Vector(
        Sheet(SheetName.unsafe("Data"))
          .put(Cell(ref"A1", CellValue.Text(crText)))
          .put(Cell(ref"B2", CellValue.Text(literalEscape)))
      )
    )

  List(
    "inline" -> com.tjclp.xl.ooxml.writer.SstPolicy.Never,
    "sst" -> com.tjclp.xl.ooxml.writer.SstPolicy.Always
  ).foreach { case (dialect, sstPolicy) =>
    test(s"GH-288 parity: CR (_x000D_) decodes identically in streaming and in-memory ($dialect)") {
      val path = Files.createTempFile(s"xl-parity-cr-$dialect-", ".xlsx")
      path.toFile.deleteOnExit()
      IO.fromEither(
        com.tjclp.xl.ooxml.XlsxWriter
          .writeWith(
            crWorkbook,
            path,
            com.tjclp.xl.ooxml.writer.WriterConfig(sstPolicy = sstPolicy)
          )
          .left
          .map(err => new Exception(s"write failed: ${err.message}"))
      ) >> loadInMemory(s"cr-$dialect", path).flatMap { wb =>
        val expected = inMemoryValues(wb.sheets(0))
        assertEquals(
          expected.get((1, 0)),
          Some(CellValue.Text(crText)),
          s"$dialect: in-memory reader lost the CR"
        )
        assertEquals(
          expected.get((2, 1)),
          Some(CellValue.Text(literalEscape)),
          s"$dialect: in-memory reader corrupted the literal _xHHHH_ text"
        )
        streamedValues(path, "Data").map { streamed =>
          assertEquals(streamed, expected, s"$dialect: streaming vs in-memory CR drift")
        }
      }
    }
  }

  test("GH-293: streaming surfaces s= style indices for inlineStr cells (styled.xlsx)") {
    // Every cell in styled.xlsx (openpyxl inline-string dialect) carries an s=
    // attribute in the XML (s=1..8). The streaming reader used to commit
    // inline-string values inside the </is> handler, bypassing the </c> branch
    // that records currentCellStyleId - so the s= index was silently dropped for
    // every t="inlineStr" cell while t="n" cells kept theirs. Both readers must
    // surface the same raw style index for every styled, non-empty cell.
    val path = fixturePath("styled.xlsx")
    val expectedRawStyleIdx = Map(
      (1, 0) -> 1, // A1 inlineStr "bold red on yellow"
      (1, 1) -> 2, // B1 n
      (1, 2) -> 3, // C1 inlineStr
      (1, 3) -> 4, // D1 inlineStr
      (2, 0) -> 5, // A2 inlineStr
      (2, 1) -> 6, // B2 n
      (3, 1) -> 7, // B3 n
      (4, 0) -> 8 // A4 inlineStr
    )
    loadInMemory("styled.xlsx", path).flatMap { wb =>
      val sheet = wb.sheets(0)
      // .iterator: collecting pairs straight from the Map would build a
      // Map[Int, Int] and collapse rows
      val inMemoryStyleIdx = sheet.cells.iterator.flatMap { case (ref, cell) =>
        cell.styleId
          .filter(_ => cell.value != CellValue.Empty)
          .map(sid => (ref.row.index1, ref.col.index0) -> sid.value)
      }.toMap
      assertEquals(
        inMemoryStyleIdx,
        expectedRawStyleIdx,
        "in-memory reader should keep the raw s= index for all 8 cells"
      )
      excel
        .readSheetStream(path, sheet.name.value)
        .compile
        .toVector
        .map { rows =>
          val streamedStyles = rows
            .flatMap { row =>
              row.cellStyles.map { case (colIdx, sid) => (row.rowIndex, colIdx) -> sid }
            }
            .toMap
          assertEquals(
            streamedStyles,
            inMemoryStyleIdx,
            "streaming must surface the same s= index as the in-memory reader for " +
              "every styled cell, including t=\"inlineStr\" (GH-293)"
          )
        }
    }
  }

  // ========== GH-293 / GH-305: raw-XML dialect probes ==========
  // XL's own writer never emits whitespace-padded formulas or escape sequences
  // split across rich runs, so these parity probes are built from raw part XML.

  private val minimalWorkbookXml =
    """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      |<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
      |          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      |  <sheets>
      |    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
      |  </sheets>
      |</workbook>
      |""".stripMargin

  private val minimalWorkbookRelsXml =
    """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      |<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      |  <Relationship Id="rId1"
      |                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
      |                Target="worksheets/sheet1.xml"/>
      |</Relationships>
      |""".stripMargin

  private val minimalContentTypesXml =
    """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      |<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      |  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      |  <Default Extension="xml" ContentType="application/xml"/>
      |  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
      |  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
      |</Types>
      |""".stripMargin

  private val minimalRootRelsXml =
    """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      |<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      |  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
      |</Relationships>
      |""".stripMargin

  /** Hand-built single-sheet xlsx from raw worksheet XML. */
  private def rawXlsx(name: String, sheetXml: String): IO[Path] = IO {
    import java.nio.charset.StandardCharsets
    import java.util.zip.{ZipEntry, ZipOutputStream}
    val path = Files.createTempFile(s"xl-parity-raw-$name-", ".xlsx")
    path.toFile.deleteOnExit()
    val zipOut = new ZipOutputStream(Files.newOutputStream(path))
    try
      List(
        "[Content_Types].xml" -> minimalContentTypesXml,
        "_rels/.rels" -> minimalRootRelsXml,
        "xl/workbook.xml" -> minimalWorkbookXml,
        "xl/_rels/workbook.xml.rels" -> minimalWorkbookRelsXml,
        "xl/worksheets/sheet1.xml" -> sheetXml
      ).foreach { case (entryName, content) =>
        val entry = new ZipEntry(entryName)
        entry.setTime(0L)
        zipOut.putNextEntry(entry)
        zipOut.write(content.getBytes(StandardCharsets.UTF_8))
        zipOut.closeEntry()
      }
    finally zipOut.close()
    path
  }

  /**
   * Text content regardless of rich formatting. The streaming reader surfaces inline rich text as
   * plain CellValue.Text (it carries no run formatting), so the cross-reader law for rich inline
   * strings is over text content, not CellValue equality.
   */
  private def plainTextOf(cv: CellValue): Option[String] = cv match
    case CellValue.Text(t) => Some(t)
    case CellValue.RichText(rt) => Some(rt.toPlainText)
    case _ => None

  test("GH-293: streaming trims <f> formula text exactly like the in-memory reader") {
    // The in-memory WorksheetReader reads formulas via `.text.trim`; pretty-printed
    // part XML (newline + indent inside <f>) must not leak whitespace into the
    // formula expression on the streaming path either.
    val sheetXml =
      """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        |<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        |  <sheetData>
        |    <row r="1">
        |      <c r="A1"><f>
        |        SUM(B1:C1)  </f><v>5</v></c>
        |      <c r="B1"><v>2</v></c>
        |      <c r="C1"><v>3</v></c>
        |    </row>
        |  </sheetData>
        |</worksheet>
        |""".stripMargin
    rawXlsx("formula-trim", sheetXml).flatMap { path =>
      loadInMemory("formula-trim", path).flatMap { wb =>
        val expected = inMemoryValues(wb.sheets(0))
        assertEquals(
          expected.get((1, 0)),
          Some(CellValue.Formula("SUM(B1:C1)", Some(CellValue.Number(BigDecimal(5))))),
          "in-memory reader trims <f> text"
        )
        streamedValues(path, "Sheet1").map { streamed =>
          assertEquals(
            streamed,
            expected,
            "whitespace-padded formulas must trim identically on both paths (GH-293)"
          )
        }
      }
    }
  }

  test("GH-305: _xHHHH_ escape split across rich inline-string runs decodes per-run") {
    // A1: run1 ends "_x00", run2 starts "0D_". Neither run contains a complete
    // escape, so the decoded value is the literal concatenation "seg_x000D_tail"
    // - NOT a CR. Decoding after concatenation would fabricate a CR; the DOM
    // reader decodes per-run. B1 pins multi-run concatenation itself. C1 pins
    // phonetic-run (rPh) exclusion: furigana text is not part of the cell value
    // on either path.
    val sheetXml =
      """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        |<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        |  <sheetData>
        |    <row r="1">
        |      <c r="A1" t="inlineStr"><is><r><t>seg_x00</t></r><r><t>0D_tail</t></r></is></c>
        |      <c r="B1" t="inlineStr"><is><r><rPr><b/></rPr><t>Alpha</t></r><r><t>Beta</t></r></is></c>
        |      <c r="C1" t="inlineStr"><is><r><t>kanji</t></r><rPh sb="0" eb="1"><t>kana</t></rPh></is></c>
        |    </row>
        |  </sheetData>
        |</worksheet>
        |""".stripMargin
    rawXlsx("split-escape", sheetXml).flatMap { path =>
      loadInMemory("split-escape", path).flatMap { wb =>
        val expected = inMemoryValues(wb.sheets(0))
        assertEquals(
          expected.get((1, 0)).flatMap(plainTextOf),
          Some("seg_x000D_tail"),
          "in-memory reader decodes per-run: literal pieces, no CR"
        )
        assertEquals(expected.get((1, 1)).flatMap(plainTextOf), Some("AlphaBeta"))
        assertEquals(expected.get((1, 2)).flatMap(plainTextOf), Some("kanji"))
        streamedValues(path, "Sheet1").map { streamed =>
          assertEquals(
            streamed.get((1, 0)).flatMap(plainTextOf),
            Some("seg_x000D_tail"),
            "streaming must decode each run before concatenation (GH-305)"
          )
          assert(
            streamed.get((1, 0)).flatMap(plainTextOf).forall(!_.contains('\r')),
            "a CR must not materialize from an escape split across runs (GH-305)"
          )
          assertEquals(
            streamed.get((1, 1)).flatMap(plainTextOf),
            Some("AlphaBeta"),
            "streaming must concatenate ALL runs, not keep the last one (GH-305)"
          )
          assertEquals(
            streamed.get((1, 2)).flatMap(plainTextOf),
            Some("kanji"),
            "streaming must exclude phonetic <rPh> runs like the in-memory reader (GH-305)"
          )
        }
      }
    }
  }
