package com.tjclp.xl.ooxml

import java.util.zip.ZipEntry

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

  test("builder captures zip metadata via +=") {
    val entry = ZipEntry("xl/sharedStrings.xml")
    entry.setSize(128L)
    entry.setCompressedSize(64L)
    entry.setCrc(12345L)
    val manifest = new PartManifestBuilder()
      .+=(entry)
      .build()

    val stored = manifest.entries("xl/sharedStrings.xml")
    assertEquals(stored.size, Some(128L))
    assertEquals(stored.compressedSize, Some(64L))
    assertEquals(stored.crc, Some(12345L))
  }
