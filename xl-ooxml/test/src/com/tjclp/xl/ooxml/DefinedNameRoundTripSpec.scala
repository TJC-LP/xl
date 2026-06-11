package com.tjclp.xl.ooxml

import java.nio.file.Files

import com.tjclp.xl.api.*
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.macros.ref
import com.tjclp.xl.unsafe.*
import munit.FunSuite

/**
 * GH-236: named ranges (DefinedName) were a read-only model — populated on read but never
 * serialized. These tests prove they now round-trip via both the fresh-write (fromDomain) and the
 * surgical (preserve-on-cell-edit) paths.
 */
class DefinedNameRoundTripSpec extends FunSuite:

  test("GH-236: programmatically authored named range serializes and round-trips") {
    val wb =
      Workbook(Sheet("Sheet1").put(ref"A1" -> 1)).withDefinedName("MyRange", "Sheet1!$A$1:$A$10")
    val out = Files.createTempFile("named-fresh", ".xlsx")
    XlsxWriter.write(wb, out).fold(e => fail(s"write failed: $e"), identity)

    val reread = XlsxReader.read(out).fold(e => fail(s"read failed: $e"), identity)
    val names = reread.metadata.definedNames
    assertEquals(names.map(_.name), Vector("MyRange"))
    assertEquals(names.headOption.map(_.formula), Some("Sheet1!$A$1:$A$10"))
    Files.deleteIfExists(out)
  }

  test("GH-236: surgical write (cell edit) preserves existing named ranges") {
    // Author a file with a defined name, then read it, edit an unrelated cell, and write back.
    val wb0 = Workbook(Sheet("Sheet1").put(ref"A1" -> 1)).withDefinedName("TaxRate", "0.08")
    val src = Files.createTempFile("named-src", ".xlsx")
    XlsxWriter.write(wb0, src).fold(e => fail(s"seed write failed: $e"), identity)

    val edited = for
      wb <- XlsxReader.read(src)
      sheet <- wb("Sheet1")
      updated = sheet.put(ref"B1" -> 2)
    yield wb.put(updated)
    val wb1 = edited.fold(e => fail(s"edit failed: $e"), identity)

    val out = Files.createTempFile("named-out", ".xlsx")
    XlsxWriter.write(wb1, out).fold(e => fail(s"write failed: $e"), identity)

    val reread = XlsxReader.read(out).fold(e => fail(s"reread failed: $e"), identity)
    assertEquals(reread.metadata.definedNames.map(_.name), Vector("TaxRate"))
    assertEquals(reread.metadata.definedNames.headOption.map(_.formula), Some("0.08"))
    Files.deleteIfExists(src)
    Files.deleteIfExists(out)
  }

  test("GH-236: removeDefinedName drops the name on write") {
    val wb = Workbook(Sheet("Sheet1").put(ref"A1" -> 1))
      .withDefinedName("Temp", "Sheet1!$A$1")
      .removeDefinedName("Temp")
    val out = Files.createTempFile("named-rm", ".xlsx")
    XlsxWriter.write(wb, out).fold(e => fail(s"write failed: $e"), identity)
    val reread = XlsxReader.read(out).fold(e => fail(s"read failed: $e"), identity)
    assertEquals(reread.metadata.definedNames, Vector.empty)
    Files.deleteIfExists(out)
  }
