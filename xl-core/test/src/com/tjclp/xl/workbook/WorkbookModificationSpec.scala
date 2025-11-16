package com.tjclp.xl.workbook

import java.nio.file.Files

import com.tjclp.xl.{SourceContext, Workbook}
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.ooxml.{PartManifest, PreservedPartStore}
import com.tjclp.xl.sheet.Sheet
import munit.FunSuite

class WorkbookModificationSpec extends FunSuite:

  private val path = Files.createTempFile("workbook-mod", ".xlsx")

  override def afterAll(): Unit =
    Files.deleteIfExists(path)

  private val ctx = SourceContext.fromFile(path, PartManifest.empty, PreservedPartStore.empty)
  private val baseSheet = Sheet("Sheet1").fold(err => fail(s"Failed to create sheet: $err"), identity)
  private val workbook = Workbook(Vector(baseSheet), sourceContext = Some(ctx))

  test("update marks tracker") {
    val updated = workbook.update("Sheet1", identity).fold(err => fail(s"Update failed: $err"), identity)
    val tracker = updated.sourceContext.fold(fail("Missing source context"))(identity).modificationTracker
    assertEquals(tracker.modifiedSheets, Set(0))
  }

  test("delete tracks deletions") {
    val sheet2 = Sheet("Sheet2").fold(err => fail(s"Failed to create sheet: $err"), identity)
    val wb = workbook.copy(sheets = Vector(baseSheet, sheet2))
    val updated = wb.delete(SheetName.unsafe("Sheet2")).fold(err => fail(s"Delete failed: $err"), identity)
    val tracker = updated.sourceContext.fold(fail("Missing source context"))(identity).modificationTracker
    assertEquals(tracker.deletedSheets, Set(1))
  }

  test("reorder marks reorder and all sheets modified") {
    val sheet2 = Sheet("Sheet2").fold(err => fail(s"Failed to create sheet: $err"), identity)
    val wb = workbook.copy(sheets = Vector(baseSheet, sheet2))
    val reordered = wb.reorder(Vector(SheetName.unsafe("Sheet2"), SheetName.unsafe("Sheet1"))).fold(err => fail(s"Reorder failed: $err"), identity)
    val tracker = reordered.sourceContext.fold(fail("Missing source context"))(identity).modificationTracker
    assert(tracker.reorderedSheets)
    assertEquals(tracker.modifiedSheets, Set(0, 1))
  }
