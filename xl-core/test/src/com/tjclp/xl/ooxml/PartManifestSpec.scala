package com.tjclp.xl.ooxml

import munit.FunSuite

class PartManifestSpec extends FunSuite:

  test("builder tracks parsed and unparsed parts") {
    val manifest = new PartManifestBuilder()
      .recordParsed("xl/workbook.xml")
      .recordUnparsed("xl/charts/chart1.xml", sheetIndex = Some(0))
      .build()

    assertEquals(manifest.parsedParts, Set("xl/workbook.xml"))
    assertEquals(manifest.unparsedParts, Set("xl/charts/chart1.xml"))
    assertEquals(manifest.dependentSheets("xl/charts/chart1.xml"), Set(0))
  }

  test("builder captures zip metadata via recordZipMetadata") {
    val manifest = new PartManifestBuilder()
      .recordZipMetadata(
        "xl/sharedStrings.xml",
        size = 128L,
        compressedSize = 64L,
        crc = 12345L,
        method = 8
      )
      .build()

    val stored = manifest.entries("xl/sharedStrings.xml")
    assertEquals(stored.size, Some(128L))
    assertEquals(stored.compressedSize, Some(64L))
    assertEquals(stored.crc, Some(12345L))
    assertEquals(stored.method, Some(8))
  }

  test("recordZipMetadata maps unknown (-1) size metadata to None") {
    // Streamed entries can carry -1 until the data descriptor is read — same mapping the
    // ZipEntry-based builder overload applied.
    val stored = new PartManifestBuilder()
      .recordZipMetadata(
        "xl/media/image1.png",
        size = -1L,
        compressedSize = -1L,
        crc = -1L,
        method = 0
      )
      .build()
      .entries("xl/media/image1.png")
    assertEquals(stored.size, None)
    assertEquals(stored.compressedSize, None)
    assertEquals(stored.crc, None)
    assertEquals(stored.method, Some(0))
  }
