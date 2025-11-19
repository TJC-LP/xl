package com.tjclp.xl.ooxml

import munit.FunSuite

class RelationshipGraphSpec extends FunSuite:

  test("empty graph has no dependencies") {
    val graph = RelationshipGraph.empty
    assertEquals(graph.dependenciesFor("xl/charts/chart1.xml"), Set.empty)
    assertEquals(graph.pathForSheet(0), "xl/worksheets/sheet1.xml") // Default path
  }

  test("fromManifest extracts sheets paths") {
    val builder = PartManifestBuilder.empty
      .recordParsed("xl/worksheets/sheet1.xml", sheetIndex = Some(0))
      .recordParsed("xl/worksheets/sheet2.xml", sheetIndex = Some(1))
    val manifest = builder.build()

    val graph = RelationshipGraph.fromManifest(manifest)

    assertEquals(graph.pathForSheet(0), "xl/worksheets/sheet1.xml")
    assertEquals(graph.pathForSheet(1), "xl/worksheets/sheet2.xml")
    assertEquals(graph.pathForSheet(99), "xl/worksheets/sheet100.xml") // Default fallback
  }

  test("fromManifest extracts dependencies from sheetIndex") {
    val builder = PartManifestBuilder.empty
      .recordUnparsed("xl/charts/chart1.xml", sheetIndex = Some(0))
      .recordUnparsed("xl/drawings/drawing1.xml", sheetIndex = Some(1))
      .recordUnparsed("xl/comments/comment1.xml", sheetIndex = Some(0))
    val manifest = builder.build()

    val graph = RelationshipGraph.fromManifest(manifest)

    // chart1 depends on sheets 0
    assertEquals(graph.dependenciesFor("xl/charts/chart1.xml"), Set(0))

    // drawing1 depends on sheets 1
    assertEquals(graph.dependenciesFor("xl/drawings/drawing1.xml"), Set(1))

    // comment1 depends on sheets 0
    assertEquals(graph.dependenciesFor("xl/comments/comment1.xml"), Set(0))

    // Unknown part has no dependencies
    assertEquals(graph.dependenciesFor("xl/unknown.xml"), Set.empty)
  }

  test("fromManifest handles parts with no sheets dependencies") {
    val builder = PartManifestBuilder.empty
      .recordUnparsed("xl/media/image1.png") // No sheetIndex
      .recordParsed("xl/workbook.xml") // No sheetIndex
    val manifest = builder.build()

    val graph = RelationshipGraph.fromManifest(manifest)

    // Parts with no sheetIndex have empty dependencies
    assertEquals(graph.dependenciesFor("xl/media/image1.png"), Set.empty)
    assertEquals(graph.dependenciesFor("xl/workbook.xml"), Set.empty)
  }

  test("dependenciesFor returns empty set for unknown paths") {
    val graph = RelationshipGraph.empty

    assertEquals(graph.dependenciesFor("xl/nonexistent.xml"), Set.empty)
  }

  test("pathForSheet returns default path for unmapped indices") {
    val graph = RelationshipGraph.empty

    // Should return default path format
    assertEquals(graph.pathForSheet(0), "xl/worksheets/sheet1.xml")
    assertEquals(graph.pathForSheet(5), "xl/worksheets/sheet6.xml")
  }
