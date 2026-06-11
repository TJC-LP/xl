package com.tjclp.xl.ooxml

import java.nio.file.Files
import java.time.LocalDateTime
import java.util.zip.{ZipEntry, ZipInputStream}

import munit.FunSuite

import com.tjclp.xl.api.*
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.macros.ref
import com.tjclp.xl.workbooks.WorkbookMetadata

/**
 * GH-242: document properties (docProps/core.xml + docProps/app.xml) are emitted on save.
 *
 * The contract:
 *   - core.xml carries creator/created/modified/lastModifiedBy; emitted iff at least one is set
 *   - app.xml carries Application/AppVersion; emitted iff at least one is set
 *   - both parts are registered in [Content_Types].xml and the package-level _rels/.rels
 *   - output is deterministic: no GUIDs, no wall-clock timestamps — only model fields, with
 *     created/modified in W3CDTF (UTC, second precision)
 *   - reading a written file restores the metadata (round-trip)
 *   - surgically re-writing a read file emits the MODEL's docProps exactly once (no duplicate-entry
 *     collision with preserved unknown parts)
 */
class DocPropsSpec extends FunSuite:

  private val fullMetadata = WorkbookMetadata(
    creator = Some("Jane Analyst"),
    created = Some(LocalDateTime.of(2024, 1, 15, 10, 30, 0)),
    modified = Some(LocalDateTime.of(2025, 6, 1, 8, 0, 0)),
    lastModifiedBy = Some("Bob Reviewer"),
    application = Some("XL Test Suite"),
    appVersion = Some("1.0")
  )

  private def workbookWith(meta: WorkbookMetadata): Workbook =
    val sheet = Sheet(com.tjclp.xl.addressing.SheetName.unsafe("Data"))
      .put(ref"A1", CellValue.Text("hello"))
    Workbook(Vector(sheet), metadata = meta)

  private def writeBytes(wb: Workbook): Array[Byte] =
    XlsxWriter.writeToBytes(wb).fold(err => fail(s"write failed: ${err.message}"), identity)

  /** All zip entry names -> UTF-8 content. */
  private def entryMap(bytes: Array[Byte]): Map[String, String] =
    val zis = new ZipInputStream(new java.io.ByteArrayInputStream(bytes))
    try
      Iterator
        .continually(zis.getNextEntry)
        .takeWhile(_ != null)
        .collect {
          case e: ZipEntry if !e.isDirectory => e.getName -> new String(zis.readAllBytes(), "UTF-8")
        }
        .toMap
    finally zis.close()

  test("write emits docProps/core.xml and docProps/app.xml with model fields") {
    val parts = entryMap(writeBytes(workbookWith(fullMetadata)))

    val core = parts.getOrElse("docProps/core.xml", fail("docProps/core.xml missing from output"))
    assert(core.contains("<dc:creator>Jane Analyst</dc:creator>"), s"creator missing: $core")
    assert(
      core.contains("<cp:lastModifiedBy>Bob Reviewer</cp:lastModifiedBy>"),
      s"lastModifiedBy missing: $core"
    )
    assert(core.contains("2024-01-15T10:30:00Z"), s"created W3CDTF missing: $core")
    assert(core.contains("2025-06-01T08:00:00Z"), s"modified W3CDTF missing: $core")
    assert(core.contains("dcterms:W3CDTF"), s"xsi:type W3CDTF marker missing: $core")

    val app = parts.getOrElse("docProps/app.xml", fail("docProps/app.xml missing from output"))
    assert(app.contains("<Application>XL Test Suite</Application>"), s"Application missing: $app")
    assert(app.contains("<AppVersion>1.0</AppVersion>"), s"AppVersion missing: $app")
  }

  test("content types and package rels register both docProps parts") {
    val parts = entryMap(writeBytes(workbookWith(fullMetadata)))

    val ct = parts.getOrElse("[Content_Types].xml", fail("no content types"))
    assert(
      ct.contains("/docProps/core.xml") &&
        ct.contains("application/vnd.openxmlformats-package.core-properties+xml"),
      s"core.xml override missing: $ct"
    )
    assert(
      ct.contains("/docProps/app.xml") &&
        ct.contains("application/vnd.openxmlformats-officedocument.extended-properties+xml"),
      s"app.xml override missing: $ct"
    )

    val rels = parts.getOrElse("_rels/.rels", fail("no package rels"))
    assert(
      rels.contains("relationships/metadata/core-properties") && rels.contains("docProps/core.xml"),
      s"core props relationship missing: $rels"
    )
    assert(
      rels.contains("relationships/extended-properties") && rels.contains("docProps/app.xml"),
      s"app props relationship missing: $rels"
    )
  }

  test("core.xml omitted when no core fields set; app.xml still written from defaults") {
    // Default WorkbookMetadata: creator/created/modified/lastModifiedBy = None,
    // application/appVersion default to the XL identifiers
    val parts = entryMap(writeBytes(workbookWith(WorkbookMetadata())))
    assert(
      !parts.contains("docProps/core.xml"),
      "core.xml should be omitted when all fields are None"
    )
    assert(
      parts.contains("docProps/app.xml"),
      "app.xml should be emitted from default application/appVersion"
    )
    val ct = parts.getOrElse("[Content_Types].xml", fail("no content types"))
    assert(!ct.contains("/docProps/core.xml"), "stale core.xml content-type override")
    val rels = parts.getOrElse("_rels/.rels", fail("no package rels"))
    assert(!rels.contains("core-properties"), "stale core props relationship")
  }

  test("app.xml omitted when application and appVersion are both None") {
    val meta = WorkbookMetadata(application = None, appVersion = None)
    val parts = entryMap(writeBytes(workbookWith(meta)))
    assert(
      !parts.contains("docProps/app.xml"),
      "app.xml should be omitted when both fields are None"
    )
  }

  test("round-trip: write metadata -> read -> metadata equal") {
    val bytes = writeBytes(workbookWith(fullMetadata))
    val readBack =
      XlsxReader.readFromBytes(bytes).fold(err => fail(s"read failed: ${err.message}"), identity)
    assertEquals(readBack.metadata.creator, fullMetadata.creator)
    assertEquals(readBack.metadata.created, fullMetadata.created)
    assertEquals(readBack.metadata.modified, fullMetadata.modified)
    assertEquals(readBack.metadata.lastModifiedBy, fullMetadata.lastModifiedBy)
    assertEquals(readBack.metadata.application, fullMetadata.application)
    assertEquals(readBack.metadata.appVersion, fullMetadata.appVersion)
  }

  test("round-trip: XML-escapable characters in creator survive") {
    val meta = WorkbookMetadata(creator = Some("""R&D <Team> "quoted""""))
    val bytes = writeBytes(workbookWith(meta))
    val readBack =
      XlsxReader.readFromBytes(bytes).fold(err => fail(s"read failed: ${err.message}"), identity)
    assertEquals(readBack.metadata.creator, meta.creator)
  }

  test("reading a file without docProps yields empty docProps metadata fields") {
    // Write a workbook with NO app/core fields at all, read it back
    val meta = WorkbookMetadata(application = None, appVersion = None)
    val bytes = writeBytes(workbookWith(meta))
    val readBack =
      XlsxReader.readFromBytes(bytes).fold(err => fail(s"read failed: ${err.message}"), identity)
    assertEquals(readBack.metadata.creator, None)
    assertEquals(readBack.metadata.application, None)
    assertEquals(readBack.metadata.appVersion, None)
  }

  test("deterministic: two writes of the same workbook produce identical docProps bytes") {
    val wb = workbookWith(fullMetadata)
    val parts1 = entryMap(writeBytes(wb))
    val parts2 = entryMap(writeBytes(wb))
    assertEquals(parts1.get("docProps/core.xml"), parts2.get("docProps/core.xml"))
    assertEquals(parts1.get("docProps/app.xml"), parts2.get("docProps/app.xml"))
  }

  test("surgical rewrite of a read file emits docProps exactly once (no duplicate collision)") {
    // 1. Write a file with metadata, 2. read it from disk (SourceContext present),
    // 3. modify a cell, 4. write again — docProps must come from the model, exactly once.
    val src = Files.createTempFile("xl-docprops-src-", ".xlsx")
    val dst = Files.createTempFile("xl-docprops-dst-", ".xlsx")
    src.toFile.deleteOnExit()
    dst.toFile.deleteOnExit()

    XlsxWriter
      .write(workbookWith(fullMetadata), src)
      .fold(err => fail(s"seed write failed: ${err.message}"), identity)

    val wb = XlsxReader.read(src).fold(err => fail(s"read failed: ${err.message}"), identity)
    assertEquals(wb.metadata.creator, fullMetadata.creator, "reader must surface docProps creator")

    val modified = wb
      .update(wb.sheets(0).name, _.put(ref"B2", CellValue.Number(BigDecimal(42))))
      .fold(err => fail(s"update failed: $err"), identity)

    XlsxWriter
      .write(modified, dst)
      .fold(err => fail(s"surgical write failed: ${err.message}"), identity)

    // Count occurrences of each docProps entry: duplicate zip entries would corrupt the file
    val zis = new ZipInputStream(new java.io.FileInputStream(dst.toFile))
    val names =
      try Iterator.continually(zis.getNextEntry).takeWhile(_ != null).map(_.getName).toList
      finally zis.close()
    assertEquals(names.count(_ == "docProps/core.xml"), 1, s"core.xml count in $names")
    assertEquals(names.count(_ == "docProps/app.xml"), 1, s"app.xml count in $names")

    val readBack =
      XlsxReader.read(dst).fold(err => fail(s"re-read failed: ${err.message}"), identity)
    assertEquals(readBack.metadata.creator, fullMetadata.creator)
    assertEquals(readBack.metadata.application, fullMetadata.application)
  }
