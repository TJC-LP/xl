package com.tjclp.xl.ooxml

import com.tjclp.xl.api.*
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.writer.{SstPolicy, WriterConfig, XmlBackend}
import munit.FunSuite

/**
 * GH-289: the SST must store strings byte-faithfully — no Unicode normalization.
 *
 * The old SST deduplicated by NFC key but stored only the FIRST occurrence's codepoints, so with
 * SST active an NFD "Café" (..., 'e', U+0301) read back as the first-stored NFC "Café" (...,
 * U+00E9): the canonically-equivalent-but-different string silently changed bytes, and which bytes
 * you got depended on cell iteration order. Excel deduplicates EXACT strings; so does xl now. NFC
 * and NFD spellings are distinct SST entries and every cell reads back its own codepoints.
 */
class SstUnicodeFidelitySpec extends FunSuite:

  private val nfc = "Caf\u00e9" // 4 codepoints, precomposed e-acute
  private val nfd = "Cafe\u0301" // 5 codepoints, e + combining acute

  private def codepoints(s: String): List[Int] = s.codePoints().toArray.toList

  test("GH-289: fromStrings keeps NFC and NFD spellings as DISTINCT entries (exact dedup)") {
    val sst = SharedStrings.fromStrings(List(nfc, nfd, nfc))
    assertEquals(sst.strings, Vector(Left(nfc), Left(nfd)), "exact dedup: two distinct entries")
    assertEquals(sst.totalCount, 3)
    assertEquals(sst.indexOf(nfc), Some(0), "exact lookup of the NFC spelling")
    assertEquals(sst.indexOf(nfd), Some(1), "exact lookup of the NFD spelling")
  }

  test("GH-289: fromEntries dedups exact strings only") {
    val sst = SharedStrings.fromEntries(List(Left(nfd), Left(nfc), Left(nfd)))
    assertEquals(sst.strings, Vector(Left(nfd), Left(nfc)))
    assertEquals(sst.totalCount, 3)
  }

  List(
    "ScalaXml" -> WriterConfig(backend = XmlBackend.ScalaXml, sstPolicy = SstPolicy.Always),
    "SaxStax" -> WriterConfig(backend = XmlBackend.SaxStax, sstPolicy = SstPolicy.Always)
  ).foreach { case (label, config) =>
    test(s"GH-289: NFC and NFD cells round-trip their own codepoints through the SST ($label)") {
      val sheet = Sheet("Data")
        .put(Cell(ref"A1", CellValue.Text(nfc)))
        .put(Cell(ref"A2", CellValue.Text(nfd)))
      val tempPath = java.nio.file.Files.createTempFile("xl-sst-nfc-", ".xlsx")
      try
        XlsxWriter
          .writeWith(Workbook(Vector(sheet)), tempPath, config)
          .fold(err => fail(s"write failed: ${err.message}"), identity)
        val back = XlsxReader
          .read(tempPath)
          .fold(err => fail(s"read failed: ${err.message}"), identity)
        def cellText(r: com.tjclp.xl.addressing.ARef): String =
          back("Data").toOption
            .flatMap(_.cells.get(r))
            .map(_.value)
            .collect { case CellValue.Text(s) => s }
            .getOrElse(fail(s"${r.toA1} is not Text"))
        assertEquals(
          codepoints(cellText(ref"A1")),
          codepoints(nfc),
          s"$label: NFC cell changed codepoints"
        )
        assertEquals(
          codepoints(cellText(ref"A2")),
          codepoints(nfd),
          s"$label: NFD cell collapsed to the first-stored NFC spelling"
        )
      finally java.nio.file.Files.deleteIfExists(tempPath)
    }
  }

  test("GH-289: a lone NFD cell keeps its decomposed codepoints through the SST (issue probe)") {
    val sheet = Sheet("Data").put(Cell(ref"A1", CellValue.Text(nfd)))
    val tempPath = java.nio.file.Files.createTempFile("xl-sst-nfd-", ".xlsx")
    try
      XlsxWriter
        .writeWith(
          Workbook(Vector(sheet)),
          tempPath,
          WriterConfig(sstPolicy = SstPolicy.Always)
        )
        .fold(err => fail(s"write failed: ${err.message}"), identity)
      val back = XlsxReader
        .read(tempPath)
        .fold(err => fail(s"read failed: ${err.message}"), identity)
      val got = back("Data").toOption
        .flatMap(_.cells.get(ref"A1"))
        .map(_.value)
        .collect { case CellValue.Text(s) => s }
        .getOrElse(fail("A1 is not Text"))
      assertEquals(codepoints(got), List(67, 97, 102, 101, 0x301))
    finally java.nio.file.Files.deleteIfExists(tempPath)
  }
