package com.tjclp.xl.ooxml

import java.nio.file.Files

import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.api.*
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.macros.ref
import com.tjclp.xl.unsafe.*
import com.tjclp.xl.workbooks.DefinedName
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

  // ===== GH-434: sheet-scoped names must keep their sheet across order mutations =====

  /** Cover (idx 0) + Data (idx 1), Data carrying a sheet-scoped name, written and read back. */
  private def scopedFixture(label: String): (java.nio.file.Path, Workbook) =
    val base = Workbook(Sheet("Cover").put(ref"A1" -> 1), Sheet("Data").put(ref"B2" -> 2))
    val wb0 = base.copy(metadata =
      base.metadata.copy(definedNames =
        Vector(DefinedName("DataLocal", "Data!$B$2", localSheetId = Some(1)))
      )
    )
    val src = Files.createTempFile(s"named-scope-$label", ".xlsx")
    src.toFile.deleteOnExit()
    XlsxWriter.write(wb0, src).fold(e => fail(s"seed write failed: $e"), identity)
    val read = XlsxReader.read(src).fold(e => fail(s"seed read failed: $e"), identity)
    assertEquals(
      read.metadata.definedNames.map(dn => (dn.name, dn.localSheetId)),
      Vector(("DataLocal", Some(1))),
      "fixture sanity: the scoped name must survive the seed round-trip"
    )
    (src, read)

  private def writeReread(wb: Workbook, label: String): Workbook =
    val out = Files.createTempFile(s"named-scope-$label-out", ".xlsx")
    out.toFile.deleteOnExit()
    XlsxWriter.write(wb, out).fold(e => fail(s"write failed: $e"), identity)
    XlsxReader.read(out).fold(e => fail(s"reread failed: $e"), identity)

  test("GH-434: sheet-scoped name still targets its sheet after removing the sheet above it") {
    val (_, wb) = scopedFixture("remove")
    val removed = wb.remove(SheetName.unsafe("Cover")).fold(e => fail(e.message), identity)
    val reread = writeReread(removed, "remove")
    assertEquals(reread.sheetNames.map(_.value), Vector("Data"))
    assertEquals(
      reread.metadata.definedNames.map(dn => (dn.name, dn.localSheetId, dn.formula)),
      Vector(("DataLocal", Some(0), "Data!$B$2")),
      "the name must follow Data from index 1 to index 0"
    )
  }

  test("GH-434: sheet-scoped name still targets its sheet after insertAt-middle") {
    val (_, wb) = scopedFixture("insert")
    val inserted =
      wb.insertAt(1, Sheet("Inserted")).fold(e => fail(e.message), identity)
    val reread = writeReread(inserted, "insert")
    assertEquals(reread.sheetNames.map(_.value), Vector("Cover", "Inserted", "Data"))
    assertEquals(
      reread.metadata.definedNames.map(dn => (dn.name, dn.localSheetId)),
      Vector(("DataLocal", Some(2))),
      "the name must follow Data from index 1 to index 2"
    )
  }

  test("GH-434: sheet-scoped name still targets its sheet after reorder") {
    val (_, wb) = scopedFixture("reorder")
    val reordered = wb
      .reorder(Vector(SheetName.unsafe("Data"), SheetName.unsafe("Cover")))
      .fold(e => fail(e.message), identity)
    val reread = writeReread(reordered, "reorder")
    assertEquals(reread.sheetNames.map(_.value), Vector("Data", "Cover"))
    assertEquals(
      reread.metadata.definedNames.map(dn => (dn.name, dn.localSheetId)),
      Vector(("DataLocal", Some(0))),
      "the name must follow Data from index 1 to index 0 across the reorder"
    )
  }
