package com.tjclp.xl

import java.nio.file.Files

import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.context.{SourceContext, SourceFingerprint}
import com.tjclp.xl.ooxml.PartManifest
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
