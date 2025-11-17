package com.tjclp.xl

import java.nio.file.Files

import com.tjclp.xl.ooxml.PartManifest
import munit.FunSuite

class SourceContextSpec extends FunSuite:

  private val tempPath = Files.createTempFile("source-context", ".xlsx")

  override def afterAll(): Unit =
    Files.deleteIfExists(tempPath)

  test("isClean reflects tracker state") {
    val ctx = SourceContext.fromFile(tempPath, PartManifest.empty)
    assert(ctx.isClean)
    val dirty = ctx.markSheetModified(0)
    assert(!dirty.isClean)
  }

  test("mark helpers delegate to tracker") {
    val ctx = SourceContext.fromFile(tempPath, PartManifest.empty)
    val updated = ctx
      .markSheetModified(1)
      .markSheetDeleted(2)
      .markMetadataModified
      .markReordered

    assertEquals(updated.modificationTracker.modifiedSheets, Set(1))
    assertEquals(updated.modificationTracker.deletedSheets, Set(2))
    assert(updated.modificationTracker.modifiedMetadata)
    assert(updated.modificationTracker.reorderedSheets)
  }
