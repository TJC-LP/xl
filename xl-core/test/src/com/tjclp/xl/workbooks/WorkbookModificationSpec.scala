package com.tjclp.xl.workbooks

import java.nio.file.Files

import com.tjclp.xl.Workbook
import com.tjclp.xl.context.{SourceContext, SourceFingerprint}
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.api.*
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.PartManifest
import com.tjclp.xl.sheets.{PageSetup, Sheet}
import munit.FunSuite

class WorkbookModificationSpec extends FunSuite:

  private val path = Files.createTempFile("workbook-mod", ".xlsx")

  override def afterAll(): Unit =
    Files.deleteIfExists(path)

  private val ctx =
    SourceContext.fromFile(path, PartManifest.empty, SourceFingerprint.fromPath(path))
  private val baseSheet = Sheet("Sheet1")
  private val workbook = Workbook(Vector(baseSheet), sourceContext = Some(ctx))

  test("update marks tracker") {
    val updated =
      workbook.update("Sheet1", identity).fold(err => fail(s"Update failed: $err"), identity)
    val tracker =
      updated.sourceContext.fold(fail("Missing source context"))(identity).modificationTracker
    assertEquals(tracker.modifiedSheets, Set(0))
  }

  test("delete tracks deletions") {
    val sheet2 = Sheet("Sheet2")
    val wb = workbook.copy(sheets = Vector(baseSheet, sheet2))
    val updated =
      wb.delete(SheetName.unsafe("Sheet2")).fold(err => fail(s"Delete failed: $err"), identity)
    val tracker =
      updated.sourceContext.fold(fail("Missing source context"))(identity).modificationTracker
    assertEquals(tracker.deletedSheets, Set(1))
  }

  test("GH-462: the scoped withDefinedName / removeDefinedName mark metadata modified") {
    val other = SheetName.unsafe("Other")
    val wb = workbook.copy(sheets = Vector(baseSheet, Sheet("Other")))
    def modifiedMetadata(result: XLResult[Workbook]): Boolean =
      result
        .fold(err => fail(s"scoped name edit failed: $err"), identity)
        .sourceContext
        .fold(fail("Missing source context"))(identity)
        .modificationTracker
        .modifiedMetadata
    assert(modifiedMetadata(wb.withDefinedName("Local", "Other!$A$1", other)))
    assert(modifiedMetadata(wb.removeDefinedName("Local", other)))
  }

  test("GH-462: clearing a print field the read lifted into PageSetup is a tracked sheet update") {
    val sheet1 = SheetName.unsafe("Sheet1")
    val lifted = workbook.copy(sheets =
      Vector(baseSheet.withPageSetup(PageSetup(printArea = Some(ref"A1:B2"))))
    )
    def tracker(result: XLResult[Workbook]) =
      result
        .fold(err => fail(s"scoped print-name edit failed: $err"), identity)
        .sourceContext
        .fold(fail("Missing source context"))(identity)
        .modificationTracker
    val removed = tracker(lifted.removeDefinedName("_xlnm.Print_Area", sheet1))
    assertEquals(removed.modifiedSheets, Set(0))
    assert(removed.modifiedMetadata)
    val authored = tracker(lifted.withDefinedName("_xlnm.Print_Area", "Sheet1!$C$1:$D$2", sheet1))
    assertEquals(authored.modifiedSheets, Set(0))
    // no lifted field to clear: the metadata edit alone, no sheet touched
    assertEquals(tracker(lifted.removeDefinedName("Local", sheet1)).modifiedSheets, Set.empty)
  }

  test("reorder marks reorder flag without marking sheets modified") {
    val sheet2 = Sheet("Sheet2")
    val wb = workbook.copy(sheets = Vector(baseSheet, sheet2))
    val reordered = wb
      .reorder(Vector(SheetName.unsafe("Sheet2"), SheetName.unsafe("Sheet1")))
      .fold(err => fail(s"Reorder failed: $err"), identity)
    val tracker =
      reordered.sourceContext.fold(fail("Missing source context"))(identity).modificationTracker
    assert(tracker.reorderedSheets)
    assertEquals(tracker.modifiedSheets, Set.empty)
  }

  test("rename marks both sheet and metadata as modified") {
    val renamed = workbook
      .rename(SheetName.unsafe("Sheet1"), SheetName.unsafe("Sales"))
      .fold(err => fail(s"Rename failed: $err"), identity)
    val tracker =
      renamed.sourceContext.fold(fail("Missing source context"))(identity).modificationTracker
    assert(
      tracker.modifiedMetadata,
      "Metadata should be marked modified (workbook.xml has sheet names)"
    )
    assertEquals(
      tracker.modifiedSheets,
      Set(0),
      "Sheet should be marked modified to preserve styles"
    )
    assertEquals(renamed.sheets(0).name.value, "Sales")
  }

  test("rename refuses a new name another sheet carries in any case (Excel sheet-name rule)") {
    val wb = Workbook(Vector(Sheet(SheetName.unsafe("S")), Sheet(SheetName.unsafe("T"))))
    assertEquals(
      wb.rename(SheetName.unsafe("T"), SheetName.unsafe("s")),
      Left(XLError.DuplicateSheet("s")): XLResult[Workbook]
    )
    assertEquals(
      wb.rename(SheetName.unsafe("T"), SheetName.unsafe("S")),
      Left(XLError.DuplicateSheet("S")): XLResult[Workbook]
    )
    // a sheet may change the case of its OWN name
    val recased = wb
      .rename(SheetName.unsafe("T"), SheetName.unsafe("t"))
      .fold(err => fail(s"Rename failed: $err"), identity)
    assertEquals(recased.sheets.map(_.name.value), Vector("S", "t"))
    // and an unrelated new name is a plain rename
    val renamed = wb
      .rename(SheetName.unsafe("T"), SheetName.unsafe("Data"))
      .fold(err => fail(s"Rename failed: $err"), identity)
    assertEquals(renamed.sheets.map(_.name.value), Vector("S", "Data"))
  }

  test("insertAt and addSheet refuse a name already used in any case") {
    val wb = Workbook(Vector(Sheet(SheetName.unsafe("Data"))))
    assertEquals(
      wb.insertAt(1, Sheet(SheetName.unsafe("data"))),
      Left(XLError.DuplicateSheet("data")): XLResult[Workbook]
    )
    assertEquals(
      wb.insertAt(0, Sheet(SheetName.unsafe("DATA"))),
      Left(XLError.DuplicateSheet("DATA")): XLResult[Workbook]
    )
    assert(wb.insertAt(1, Sheet(SheetName.unsafe("Notes"))).isRight)
    @annotation.nowarn("cat=deprecation")
    val added = wb.addSheet(Sheet(SheetName.unsafe("data")))
    assertEquals(added, Left(XLError.DuplicateSheet("data")): XLResult[Workbook])
  }

  test("put marks sheet as modified when replacing existing sheet") {
    val modifiedSheet = baseSheet.put(ref"A1" -> "New Value")
    val updated = workbook.put(modifiedSheet)
    val tracker =
      updated.sourceContext.fold(fail("Missing source context"))(identity).modificationTracker
    assertEquals(tracker.modifiedSheets, Set(0))
  }

  test("put marks metadata as modified when adding new sheet") {
    val newSheet = Sheet("Sheet2")
    val updated = workbook.put(newSheet)
    val tracker =
      updated.sourceContext.fold(fail("Missing source context"))(identity).modificationTracker
    assert(tracker.modifiedMetadata, "Adding new sheet should mark metadata as modified")
    assertEquals(
      tracker.modifiedSheets,
      Set.empty
    ) // Individual sheets not modified, just workbook.xml
  }
