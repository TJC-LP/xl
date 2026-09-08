package com.tjclp.xl.io.streaming

import java.nio.file.{Files, Path}

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.{Sheet, Workbook, given}
import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.{CellValue, Comment}
import com.tjclp.xl.error.{XLError, XLException}
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.XlsxReader
import com.tjclp.xl.ooxml.XlsxReader.ReaderConfig
import com.tjclp.xl.ooxml.writer.{SstPolicy, WriterConfig}
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.numfmt.NumFmt

/**
 * GH-640: the shared-string table is loaded ONCE, through the streaming reader's SAX parser, under
 * the reader's ZIP-bomb limits, and handed to every streaming read of the same file — the row
 * streams and `streamCellDetails` alike — instead of each read parsing it again (the single-cell
 * read used to parse it as a DOM, three times the memory of the table itself).
 */
class SharedStringsLoadSpec extends CatsEffectSuite:

  private val excel = ExcelIO.instance[IO]

  /** The default policy inlines a small book's strings; these tests are about the table. */
  private val withSst = WriterConfig(sstPolicy = SstPolicy.Always)

  private def tempXlsx(label: String): Path =
    val p = Files.createTempFile(s"xl-sst-load-$label-", ".xlsx")
    p.toFile.deleteOnExit()
    p

  /**
   * Two rows of text and numbers, one styled, one commented: what every read path must agree on.
   */
  private val book: Workbook =
    val sheet = Sheet("Data")
      .put(ref"A1", "alpha")
      .put(ref"B1", 10)
      .put(ref"A2", "beta")
      .put(ref"B2", BigDecimal("0.25"))
      .put(ref"A3", "alpha")
      .comment(ref"A1", Comment.plainText("input", Some("qa")))
    val styled = sheet.styleAt("B2", CellStyle.default.withNumFmt(NumFmt.Percent)).getOrElse(sheet)
    Workbook(Vector(styled))

  /** A table wide enough to trip a small byte cap and repetitive enough to trip a ratio cap. */
  private val wide: Workbook =
    val sheet = (1 to 2000).foldLeft(Sheet("Wide")) { (s, i) =>
      s.put(ARef.from0(0, i - 1), f"row-$i%06d-xxxxxxxxxxxxxxxxxxxxxxxx")
    }
    Workbook(Vector(sheet))

  test("loadSharedStrings is the table the in-memory reader sees; None when the file has none") {
    val path = tempXlsx("table")
    val inline = tempXlsx("inline")
    for
      _ <- excel.writeWith(book, path, withSst)
      _ <- excel.writeWith(wide, inline, WriterConfig(sstPolicy = SstPolicy.Never))
      loaded <- excel.loadSharedStrings(path, ReaderConfig.default)
      none <- excel.loadSharedStrings(inline, ReaderConfig.default)
    yield
      val expected = XlsxReader.read(path).fold(e => fail(e.message), identity)
      assertEquals(
        loaded.map(_.strings.collect { case Left(text) => text }.sorted),
        Some(Vector("alpha", "beta"))
      )
      // the same two strings the loaded sheet holds (three cells, two distinct texts)
      assertEquals(
        expected.sheets
          .flatMap(_.cells.values.map(_.value))
          .collect { case CellValue.Text(t) =>
            t
          }
          .distinct
          .sorted,
        Vector("alpha", "beta")
      )
      assertEquals(none, None)
  }

  test("the byte cap: inflation stops one byte past --max-size with SECURITY_ERROR, 0 lifts it") {
    val path = tempXlsx("cap")
    for
      _ <- excel.writeWith(wide, path, withSst)
      capped <- excel
        .loadSharedStrings(path, ReaderConfig.default.copy(maxUncompressedSize = 64))
        .attempt
      lifted <- excel.loadSharedStrings(path, ReaderConfig.permissive)
    yield
      capped match
        case Left(x: XLException) =>
          x.error match
            case XLError.SecurityError(reason) =>
              assert(reason.contains("xl/sharedStrings.xml"), reason)
              assert(reason.contains("64 bytes"), reason)
            case other => fail(s"expected SecurityError, got $other")
          assertEquals(x.error.code, "SECURITY_ERROR")
        case other => fail(s"expected an XLException, got $other")
      assertEquals(lifted.map(_.strings.size), Some(2000))
  }

  test(
    "the ratio cap: a table inflating past maxCompressionRatio × its compressed size is refused"
  ) {
    val path = tempXlsx("ratio")
    for
      _ <- excel.writeWith(wide, path, withSst)
      refused <- excel
        .loadSharedStrings(path, ReaderConfig.default.copy(maxCompressionRatio = 1))
        .attempt
      allowed <- excel.loadSharedStrings(path, ReaderConfig.default)
    yield
      refused match
        case Left(x: XLException) =>
          x.error match
            case XLError.SecurityError(reason) =>
              assert(reason.contains("Compression ratio"), reason)
              assert(reason.contains("xl/sharedStrings.xml"), reason)
            case other => fail(s"expected SecurityError, got $other")
        case other => fail(s"expected an XLException, got $other")
      assertEquals(allowed.map(_.strings.size), Some(2000))
  }

  test("a pre-loaded table streams the same rows as a read that loads its own") {
    val path = tempXlsx("rows")
    val window = CellRange(ref"A1", ref"B3")
    for
      _ <- excel.writeWith(book, path, withSst)
      sst <- excel.loadSharedStrings(path, ReaderConfig.default)
      own <- excel.readSheetStreamRange(path, "Data", window).compile.toVector
      shared <- excel.readSheetStreamRange(path, "Data", window, sst).compile.toVector
      whole <- excel.readSheetStream(path, "Data", sst).compile.toVector
    yield
      assertEquals(shared, own)
      assertEquals(whole, own)
      assertEquals(
        own.flatMap(_.cells.get(0)),
        Vector("alpha", "beta", "alpha").map(CellValue.Text(_))
      )
  }

  test("streamCellDetails with the table and styles pre-loaded answers exactly as the plain form") {
    val path = tempXlsx("cell")
    for
      _ <- excel.writeWith(book, path, withSst)
      sst <- excel.loadSharedStrings(path, ReaderConfig.default)
      styles <- excel.loadStyles(path)
      text <- excel.streamCellDetails(path, "Data", ref"A1")
      textShared <- excel.streamCellDetails(path, "Data", ref"A1", sst, styles)
      styled <- excel.streamCellDetails(path, "Data", ref"B2")
      styledShared <- excel.streamCellDetails(path, "Data", ref"B2", sst, styles)
      absent <- excel.streamCellDetails(path, "Data", ref"D9", sst, styles)
    yield
      assertEquals(textShared, text)
      assertEquals(text.value, CellValue.Text("alpha"))
      assertEquals(text.comment.map(_.text.toPlainText), Some("input"))
      assertEquals(styledShared, styled)
      assertEquals(styled.style.map(_.numFmt), Some(NumFmt.Percent))
      assertEquals(absent.value, CellValue.Empty)
  }
