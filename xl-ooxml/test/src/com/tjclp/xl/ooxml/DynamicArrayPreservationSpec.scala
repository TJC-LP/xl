package com.tjclp.xl.ooxml

import munit.FunSuite

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.api.*
import com.tjclp.xl.cells.{ArrayMode, CellValue, FormulaKind}
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.writer.{WriterConfig, XmlBackend}

/**
 * GH-714: a dynamic-array record (`cm` on the cell, XLDAPR in `xl/metadata.xml`) survives a
 * regenerated sheet. Without `cm` Excel reads the record as a legacy CSE array `{=SORT(..)}`.
 */
class DynamicArrayPreservationSpec extends FunSuite:
  import DynamicArrayFixtures.*

  private val anchorXml =
    """<c r="B1" t="n" cm="1"><f t="array" ref="B1:B5">_xlfn._xlws.SORT(A1:A5)</f><v>1</v></c>"""

  private val backends = List(
    "ScalaXml" -> WriterConfig(backend = XmlBackend.ScalaXml),
    "SaxStax" -> WriterConfig(backend = XmlBackend.SaxStax)
  )

  private def read(bytes: Array[Byte]): Workbook =
    XlsxReader.readFromBytes(bytes).fold(err => fail(err.message), identity)

  private def write(wb: Workbook, config: WriterConfig = WriterConfig()): Array[Byte] =
    XlsxWriter.writeToBytes(wb, config).fold(err => fail(err.message), identity)

  private def sheet1(bytes: Array[Byte]): String =
    entryText(bytes, "xl/worksheets/sheet1.xml").getOrElse(fail("no sheet1"))

  private def occurrences(bytes: Array[Byte], part: String, needle: String): Int =
    entryText(bytes, part).getOrElse(fail(s"no $part")).split(needle, -1).length - 1

  private def assertWiredOnce(out: Array[Byte]): Unit =
    assertEquals(entries(out).count(_._1 == "xl/metadata.xml"), 1)
    assertEquals(occurrences(out, "[Content_Types].xml", "/xl/metadata.xml"), 1)
    assertEquals(
      occurrences(out, "xl/_rels/workbook.xml.rels", "relationships/sheetMetadata"),
      1
    )

  private def dirty(wb: Workbook): Workbook =
    val s = wb.sheets.headOption.getOrElse(fail("no sheet"))
    wb.put(s.put(ref"D1", CellValue.Number(5)))

  private def kindAt(wb: Workbook, at: ARef): Option[FormulaKind] =
    wb.sheets.headOption.flatMap(_.cells.get(at)).map(_.value).collect {
      case CellValue.Formula(_, _, kind) => kind
    }

  private val sortRange = CellRange(ref"B1", ref"B5")

  backends.foreach { case (name, config) =>
    test(s"put D1 keeps the anchor's cm and the metadata part ($name)") {
      val out = write(dirty(read(xlsx())), config)
      assert(sheet1(out).contains(anchorXml), sheet1(out))
      assertEquals(entryText(out, "xl/metadata.xml"), Some(metadataXml))
      assertWiredOnce(out)
    }
  }

  test("the reader resolves cm=1 to a dynamic array") {
    assertEquals(kindAt(read(xlsx()), ref"B1"), Some(FormulaKind.dynamicArray(sortRange)))
  }

  test("a dangling cm (no metadata part) reads as a legacy array and writes no cm") {
    val wb = read(xlsx(parts(metadata = None)))
    assertEquals(kindAt(wb, ref"B1"), Some(FormulaKind.ArrayFormula(sortRange)))
    val out = write(dirty(wb))
    assert(!sheet1(out).contains("cm="), sheet1(out))
    assertEquals(entryText(out, "xl/metadata.xml"), None)
  }

  private def scratch(kind: FormulaKind): Workbook =
    Workbook(
      Sheet("Sheet1")
        .put(ref"A1", CellValue.Number(3))
        .put(ref"A2", CellValue.Number(1))
        .put(ref"B1", CellValue.Formula("SORT(A1:A2)", Some(CellValue.Number(1)), kind))
        .put(ref"B2", CellValue.Number(3))
    )

  backends.foreach { case (name, config) =>
    test(s"a scratch dynamic array ships the generated part, Override and rel ($name)") {
      val kind = FormulaKind.dynamicArray(CellRange(ref"B1", ref"B2"))
      val wb = scratch(kind)
      val out = write(wb, config)
      assertEquals(entryText(out, "xl/metadata.xml"), Some(metadataXml))
      assert(
        sheet1(out).contains("""<c r="B1" t="n" cm="1"><f t="array" ref="B1:B2">"""),
        sheet1(out)
      )
      val ct = entryText(out, "[Content_Types].xml").getOrElse(fail("no CT"))
      assert(
        ct.contains(
          """ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml""""
        ),
        ct
      )
      assert(ct.contains("""PartName="/xl/metadata.xml""""), ct)
      val rels = entryText(out, "xl/_rels/workbook.xml.rels").getOrElse(fail("no rels"))
      assert(rels.contains("relationships/sheetMetadata\""), rels)
      assert(rels.contains("Target=\"metadata.xml\""), rels)
      assertWiredOnce(out)
      // deterministic, and it reads back as what was written
      assertEquals(write(wb, config).toVector, out.toVector)
      assertEquals(kindAt(read(out), ref"B1"), Some(kind))
    }

    test(s"a scratch book without dynamic arrays ships no metadata part ($name)") {
      val out = write(scratch(FormulaKind.ArrayFormula(CellRange(ref"B1", ref"B2"))), config)
      assertEquals(entryText(out, "xl/metadata.xml"), None)
      assert(!sheet1(out).contains("cm="), sheet1(out))
      val rels = entryText(out, "xl/_rels/workbook.xml.rels").getOrElse(fail("no rels"))
      assert(!rels.contains("sheetMetadata"), rels)
    }

    test(s"a collapsed dynamic array round-trips ($name)") {
      val kind = FormulaKind.ArrayFormula(
        CellRange(ref"B1", ref"B2"),
        mode = ArrayMode.Dynamic(collapsed = true)
      )
      val out = write(scratch(kind), config)
      assertEquals(kindAt(read(out), ref"B1"), Some(kind))
    }
  }

  test("a 1x1 dynamic array writes Excel's bare ref") {
    val wb = Workbook(
      Sheet("Sheet1").put(
        ref"C5",
        CellValue.Formula(
          "SUM(A1:A3*B1:B3)",
          Some(CellValue.Number(14)),
          FormulaKind.dynamicArray(CellRange(ref"C5", ref"C5"))
        )
      )
    )
    val out = write(wb)
    assert(sheet1(out).contains("""<c r="C5" t="n" cm="1"><f t="array" ref="C5">"""), sheet1(out))
  }

  test("an images-in-cells metadata part is extended, never rewritten") {
    val rich =
      """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        |<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"><metadataTypes count="1"><metadataType name="XLRICHVALUE" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1"/></metadataTypes><futureMetadata name="XLRICHVALUE" count="1"><bk><extLst><ext uri="{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}"><xlrd:rvb i="0"/></ext></extLst></bk></futureMetadata><valueMetadata count="1"><bk><rc t="1" v="0"/></bk></valueMetadata></metadata>""".stripMargin
    val plainSheet =
      """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        |<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>5</v></c></row></sheetData></worksheet>""".stripMargin
    val wb = read(xlsx(parts(sheetXml = plainSheet, metadata = Some(rich))))
    val s = wb.sheets.headOption.getOrElse(fail("no sheet"))
    val kind = FormulaKind.dynamicArray(CellRange(ref"B1", ref"B1"))
    val edited =
      wb.put(s.put(ref"B1", CellValue.Formula("SORT(A1:A1)", Some(CellValue.Number(5)), kind)))
    val out = write(edited)
    val part = entryText(out, "xl/metadata.xml").getOrElse(fail("no metadata part"))
    assert(
      part.contains("""<valueMetadata count="1"><bk><rc t="1" v="0"/></bk></valueMetadata>"""),
      part
    )
    assert(
      part.contains("""<cellMetadata count="1"><bk><rc t="2" v="0"/></bk></cellMetadata>"""),
      part
    )
    assert(sheet1(out).contains("""<c r="B1" t="n" cm="1">"""), sheet1(out))
    assertWiredOnce(out)
    assertEquals(kindAt(read(out), ref"B1"), Some(kind))
  }

  test("a metadata-modified write (a sheet added) keeps the part and its wiring") {
    val wb = read(xlsx())
    val out = write(wb.put(Sheet("Extra").put(ref"A1", CellValue.Number(1))))
    assert(sheet1(out).contains(anchorXml), sheet1(out))
    assertEquals(entryText(out, "xl/metadata.xml"), Some(metadataXml))
    assertWiredOnce(out)
    assertEquals(kindAt(read(out), ref"B1"), Some(FormulaKind.dynamicArray(sortRange)))
  }

  test("a clean re-write of a dynamic-array book is byte-identical to its first write") {
    val once = write(dirty(read(xlsx())))
    val twice = write(dirty(read(once)))
    assertEquals(sheet1(twice), sheet1(once))
    assertEquals(entryText(twice, "xl/metadata.xml"), entryText(once, "xl/metadata.xml"))
  }
