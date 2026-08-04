package com.tjclp.xl

import java.nio.file.Files

import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.context.{SourceContext, SourceFingerprint}
import com.tjclp.xl.ooxml.PartManifest
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.{DefinedName, Workbook}
import munit.FunSuite

class SourceContextSpec extends FunSuite:

  private val tempPath = Files.createTempFile("source-context", ".xlsx")

  override def afterAll(): Unit =
    Files.deleteIfExists(tempPath)

  private def name(s: String): SheetName = SheetName.unsafe(s)

  test("isClean reflects tracker state") {
    val ctx =
      SourceContext.fromFile(tempPath, PartManifest.empty, SourceFingerprint.fromPath(tempPath))
    assert(ctx.isClean)
    val dirty = ctx.markSheetModified(0)
    assert(!dirty.isClean)
  }

  test("mark helpers delegate to tracker") {
    val ctx =
      SourceContext.fromFile(tempPath, PartManifest.empty, SourceFingerprint.fromPath(tempPath))
    val updated = ctx
      .markSheetModified(1)
      .markSheetDeleted(2, name("Gone"))
      .markMetadataModified
      .markReordered

    assertEquals(updated.modificationTracker.modifiedSheets, Set(1))
    assertEquals(updated.modificationTracker.deletedSheets, Set(2))
    assert(updated.modificationTracker.modifiedMetadata)
    assert(updated.modificationTracker.reorderedSheets)
  }

  test("GH-315: deletion marks metadata modified (workbook.xml restructures)") {
    val ctx =
      SourceContext.fromFile(tempPath, PartManifest.empty, SourceFingerprint.fromPath(tempPath))
    val deleted = ctx.markSheetDeleted(0, name("Gone"))
    assert(
      deleted.modificationTracker.modifiedMetadata,
      "a deletion must force the metadata-modified write path"
    )
  }

  test("GH-315: markSheetDeleted drops the sheet's identity-keyed mappings") {
    val ctx = SourceContext
      .fromFile(
        tempPath,
        PartManifest.empty,
        SourceFingerprint.fromPath(tempPath),
        commentPathMapping = Map(name("A") -> "xl/comments1.xml", name("B") -> "xl/comments2.xml"),
        drawingPathMapping = Map(name("A") -> "xl/drawings/drawing1.xml"),
        drawingSnapshots = Map(name("A") -> Vector.empty),
        sheetPathMapping =
          Map(name("A") -> "xl/worksheets/sheet1.xml", name("B") -> "xl/worksheets/sheet2.xml")
      )
    val deleted = ctx.markSheetDeleted(0, name("A"))
    assertEquals(deleted.commentPathMapping, Map(name("B") -> "xl/comments2.xml"))
    assertEquals(deleted.drawingPathMapping, Map.empty[SheetName, String])
    assertEquals(deleted.drawingSnapshots.keySet, Set.empty[SheetName])
    assertEquals(deleted.sheetPathMapping, Map(name("B") -> "xl/worksheets/sheet2.xml"))
  }

  test("GH-470: reconciledWith marks an untracked sheet-set reduction like a delete") {
    val ctx = SourceContext
      .fromFile(
        tempPath,
        PartManifest.empty,
        SourceFingerprint.fromPath(tempPath),
        commentPathMapping = Map(name("B") -> "xl/comments1.xml"),
        sheetPathMapping =
          Map(name("A") -> "xl/worksheets/sheet1.xml", name("B") -> "xl/worksheets/sheet2.xml")
      )
    // "B" dropped via Workbook.copy(sheets = ...) — no tracked delete ever ran
    val reduced = Workbook(Vector(Sheet(name("A"))))
    val reconciled = ctx.reconciledWith(reduced)

    assert(!reconciled.isClean, "an untracked reduction must dirty the context")
    assertEquals(reconciled.modificationTracker.deletedSheets, Set(1))
    assert(reconciled.modificationTracker.modifiedMetadata)
    assertEquals(reconciled.sheetPathMapping, Map(name("A") -> "xl/worksheets/sheet1.xml"))
    assertEquals(reconciled.commentPathMapping, Map.empty[SheetName, String])
  }

  test("GH-470: reconciledWith does NOT shift live modified-sheet marks") {
    val ctx = SourceContext
      .fromFile(
        tempPath,
        PartManifest.empty,
        SourceFingerprint.fromPath(tempPath),
        sheetPathMapping =
          Map(name("A") -> "xl/worksheets/sheet1.xml", name("B") -> "xl/worksheets/sheet2.xml")
      )
      .markSheetModified(0) // live index 0 = "B" AFTER the untracked drop of "A"
    val reduced = Workbook(Vector(Sheet(name("B"))))
    val reconciled = ctx.reconciledWith(reduced)

    assertEquals(reconciled.modificationTracker.deletedSheets, Set(0))
    assertEquals(
      reconciled.modificationTracker.modifiedSheets,
      Set(0),
      "the surviving sheet's modification mark must not be shifted or dropped"
    )
  }

  test("GH-470: reconciledWith marks untracked additions and defined-name drift as metadata") {
    val ctx = SourceContext
      .fromFile(
        tempPath,
        PartManifest.empty,
        SourceFingerprint.fromPath(tempPath),
        sheetPathMapping = Map(name("A") -> "xl/worksheets/sheet1.xml")
      )
    val appended = Workbook(Vector(Sheet(name("A")), Sheet(name("New"))))
    assert(ctx.reconciledWith(appended).modificationTracker.modifiedMetadata)
    assertEquals(ctx.reconciledWith(appended).modificationTracker.deletedSheets, Set.empty[Int])

    val snapshot =
      ctx.copy(definedNamesAsRead = Some(Vector(DefinedName(name = "SECRET", formula = "A!$A$1"))))
    val redacted = Workbook(Vector(Sheet(name("A"))))
    assert(
      snapshot.reconciledWith(redacted).modificationTracker.modifiedMetadata,
      "dropping a defined name via copy(metadata = ...) must dirty the context"
    )
  }

  test("GH-470: reconciledWith is identity when nothing diverges") {
    val ctx = SourceContext
      .fromFile(
        tempPath,
        PartManifest.empty,
        SourceFingerprint.fromPath(tempPath),
        sheetPathMapping = Map(name("A") -> "xl/worksheets/sheet1.xml")
      )
    val untouched = Workbook(Vector(Sheet(name("A"))))
    assert(
      ctx.reconciledWith(untouched) eq ctx,
      "a clean workbook must keep the verbatim fast path"
    )
  }

  test("GH-315: markSheetRenamed re-keys identity mappings (identity follows the sheet)") {
    val ctx = SourceContext
      .fromFile(
        tempPath,
        PartManifest.empty,
        SourceFingerprint.fromPath(tempPath),
        commentPathMapping = Map(name("A") -> "xl/comments1.xml"),
        drawingPathMapping = Map(name("A") -> "xl/drawings/drawing1.xml"),
        sheetPathMapping = Map(name("A") -> "xl/worksheets/sheet1.xml")
      )
    val renamed = ctx.markSheetRenamed(name("A"), name("Z"))
    assertEquals(renamed.commentPathMapping, Map(name("Z") -> "xl/comments1.xml"))
    assertEquals(renamed.drawingPathMapping, Map(name("Z") -> "xl/drawings/drawing1.xml"))
    assertEquals(renamed.sheetPathMapping, Map(name("Z") -> "xl/worksheets/sheet1.xml"))
    // Renaming a sheet with no recorded mappings is a no-op
    assertEquals(
      ctx.markSheetRenamed(name("Nope"), name("X")).commentPathMapping.keySet.map(_.value),
      Set("A")
    )
  }
