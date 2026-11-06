package com.tjclp.xl.workbooks

import scala.collection.immutable.ArraySeq

import com.tjclp.xl.{*, given}
import com.tjclp.xl.context.{ModificationTracker, SourceContent, SourceContext, SourceFingerprint}
import com.tjclp.xl.ooxml.{PartManifest, PartManifestEntry}
import munit.FunSuite

class SheetAdoptionSpec extends FunSuite:
  private def context(externalLinks: Int = 0): SourceContext =
    val entries = (1 to externalLinks)
      .map(i =>
        val path = s"xl/externalLinks/externalLink$i.xml"
        path -> PartManifestEntry.unparsed(path)
      )
      .toMap
    SourceContext(
      SourceContent.InMemory(ArraySeq.empty[Byte]),
      PartManifest(entries),
      ModificationTracker.clean,
      SourceFingerprint(0L, ArraySeq.empty[Byte])
    )

  private def failure(result: XLResult[Workbook]): String =
    result.fold(_.toString, _ => fail("Expected adoption refusal"))

  test("GH-526: adopts source-free modeled sheet, preserving styles and tracked destination") {
    val portable = Sheet("Portable").put(ref"A1", 42).style(ref"A1", CellStyle.default.bold)
    val source = Workbook(portable)
    val destination = Workbook("Destination").copy(sourceContext = Some(context()))
    val result = destination.adopt(portable, from = source).fold(e => fail(e.toString), identity)
    assertEquals(result.sheets, destination.sheets :+ portable)
    assertEquals(result.sheets.lastOption.map(_.styleRegistry), Some(portable.styleRegistry))
    assert(result.sourceContext.exists(!_.isClean))
    assert(destination.sourceContext.exists(_.isClean))
    assertEquals(source.sheets, Vector(portable))
  }

  test("GH-526: refuses dangling and valid-but-rebinding ordinals before changing either book") {
    val sheet = Sheet("Imported")
      .put(ref"B2", CellValue.Formula("[1]Other!A1"))
      .put(ref"A1", CellValue.Formula("'[3]Other Data'!B2"))
    val source = Workbook(sheet).copy(sourceContext = Some(context(3)))
    val destination = Workbook("Destination").copy(sourceContext = Some(context(1)))
    val message = failure(destination.adopt(sheet, from = source))
    assert(message.contains("could dangle or silently rebind"))
    assert(message.contains("A1: '[3]Other Data'!B2"))
    assert(message.contains("B2: [1]Other!A1"))
    assert(message.indexOf("A1:") < message.indexOf("B2:"))
    assertEquals(destination.sheets.size, 1)
    assert(destination.sourceContext.exists(_.isClean))
    assert(source.sourceContext.exists(_.isClean))
  }

  test("GH-526: refuses external ordinals even on source-free raw formula models") {
    val sheet = Sheet("Imported").put(ref"A1", CellValue.Formula("[999999999999]Other!A1"))
    assert(failure(Workbook("Destination").adopt(sheet, Workbook(sheet))).contains("external"))
  }

  test("GH-526: formula strings, quoted names and structured brackets are not external ordinals") {
    val formulas = Vector(
      "\"see [3]\"&A1",
      "\"a\"\"[7]\"\"b\"",
      "'It''s [2]'!A1",
      "SUM(Table1[3])",
      "SUM(Table1[[#This Row],[3]])",
      "[0]!ThisBookName"
    )
    formulas.zipWithIndex.foreach { case (formula, i) =>
      val sheet = Sheet("Portable").put(ARef.from0(0, i), CellValue.Formula(formula))
      assert(Workbook("Destination").adopt(sheet, Workbook(sheet)).isRight, formula)
    }
  }

  test("GH-526: source-backed transfers are conservatively refused without losing provenance") {
    val sheet = Sheet("Imported").put(ref"A1", 1)
    val source = Workbook(sheet).copy(sourceContext = Some(context()))
    assert(failure(Workbook("Destination").adopt(sheet, source)).contains("package migration"))
    assert(source.sourceContext.exists(_.isClean))
  }

  test("GH-526: exact same-workbook adoption is identity, edits must match supplied source") {
    val sheet = Sheet("Imported").put(ref"A1", CellValue.Formula("[1]External!A1"))
    val source = Workbook(sheet).copy(sourceContext = Some(context(1)))
    assertEquals(source.adopt(sheet, source), Right(source))
    val edited = sheet.put(ref"B1", 2)
    assert(failure(source.adopt(edited, source)).contains("does not match"))
    val other = Workbook("Other")
    assert(failure(other.adopt(sheet, other)).contains("does not match"))
  }

  test("GH-526: add-only adoption refuses case-insensitive collisions rather than replacing") {
    val sheet = Sheet("DATA").put(ref"A1", 42)
    val destination = Workbook("Data")
    assertEquals(destination.adopt(sheet, Workbook(sheet)), Left(XLError.DuplicateSheet("DATA")))
    assertEquals(destination.sheets.size, 1)
  }

  test("GH-526: refuses unsupported relationship-bearing modeled payloads") {
    val sheet =
      Sheet("Imported").withDataValidation(ref"A1:A2", DataValidation.custom("[1]Data!A1"))
    assert(failure(Workbook("Destination").adopt(sheet, Workbook(sheet))).contains("validations"))
  }

  test("GH-526: raw put retains its low-level replacement and external-formula behavior") {
    val sheet = Sheet("Data").put(ref"A1", CellValue.Formula("[1]External!A1"))
    val destination = Workbook("Data")
    assertEquals(destination.put(sheet).sheets, Vector(sheet))
  }

  test("GH-526: escaped brackets in structured headers cannot hide a later external ordinal") {
    val sheet = Sheet("Imported").put(ref"A1", CellValue.Formula("SUM(Table1['[Col])+[1]Data!A1"))
    assert(failure(Workbook("Destination").adopt(sheet, Workbook(sheet))).contains("A1:"))
  }

  test("GH-526: default font and theme mismatches are refused; equivalent defaults are allowed") {
    val sheet = Sheet("Portable").put(ref"A1", 1)
    val source = Workbook(sheet)
    val destination = Workbook("Destination")
    assert(
      failure(destination.adopt(sheet, source.withDefaultFont(Font("Times New Roman", 10.0))))
        .contains("default font or theme")
    )
    val themed =
      source.copy(metadata = source.metadata.copy(theme = source.metadata.theme.copy(accent1 = 1)))
    assert(failure(destination.adopt(sheet, themed)).contains("default font or theme"))
    assert(destination.adopt(sheet, source.withDefaultFont(Font.default)).isRight)
  }
