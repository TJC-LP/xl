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
  private val baseSheet = Sheet("Sheet1").toOption.get
  private val workbook = Workbook(Vector(baseSheet), sourceContext = Some(ctx))

  test("updateSheet marks tracker") {
    val updated = workbook.updateSheet("Sheet1", identity).toOption.get
    val tracker = updated.sourceContext.get.modificationTracker
    assertEquals(tracker.modifiedSheets, Set(0))
  }

  test("deleteSheet tracks deletions") {
    val sheet2 = Sheet("Sheet2").toOption.get
    val wb = workbook.copy(sheets = Vector(baseSheet, sheet2))
    val updated = wb.deleteSheet(SheetName.unsafe("Sheet2")).toOption.get
    val tracker = updated.sourceContext.get.modificationTracker
    assertEquals(tracker.deletedSheets, Set(1))
  }

  test("reorderSheets marks reorder and all sheets modified") {
    val sheet2 = Sheet("Sheet2").toOption.get
    val wb = workbook.copy(sheets = Vector(baseSheet, sheet2))
    val reordered = wb.reorderSheets(Vector(SheetName.unsafe("Sheet2"), SheetName.unsafe("Sheet1"))).toOption.get
    val tracker = reordered.sourceContext.get.modificationTracker
    assert(tracker.reorderedSheets)
    assertEquals(tracker.modifiedSheets, Set(0, 1))
  }
