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

  test("KNOWN GAP: streaming drops s= style indices for inlineStr cells (styled.xlsx)") {
    // Every cell in styled.xlsx carries an s= attribute in the XML (s=1..8) and
    // the in-memory reader assigns a styleId to all 8 cells. The streaming
    // reader, however, commits inline-string values inside the </is> handler,
    // which bypasses the </c> branch that records currentCellStyleId - so the
    // s= index is silently dropped for every t="inlineStr" cell. openpyxl
    // writes ALL strings as inlineStr, so streaming style preservation sees no
    // style for any openpyxl text cell. Number cells (t="n") keep theirs.
    // This pins the current behavior; when SaxStreamingReader is fixed, flip
    // this test to assert full agreement with the in-memory styleIds.
    val path = fixturePath("styled.xlsx")
    loadInMemory("styled.xlsx", path).flatMap { wb =>
      val sheet = wb.sheets(0)
      // .iterator: collecting (Int, Int) pairs straight from the Map would build
      // a Map[Int, Int] and collapse rows
      val inMemoryStyled = sheet.cells.iterator.collect {
        case (ref, cell) if cell.styleId.isDefined && cell.value != CellValue.Empty =>
          (ref.row.index1, ref.col.index0)
      }.toSet
      assertEquals(
        inMemoryStyled,
        Set((1, 0), (1, 1), (1, 2), (1, 3), (2, 0), (2, 1), (3, 1), (4, 0)),
        "in-memory reader should style all 8 cells"
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
            Map((1, 1) -> 2, (2, 1) -> 6, (3, 1) -> 7),
            "streaming currently surfaces s= only for the three t=\"n\" cells; if " +
              "inlineStr cells now appear, the gap was fixed - flip this test to " +
              "assert parity with the in-memory styleIds"
          )
        }
    }
  }
