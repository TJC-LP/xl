package com.tjclp.xl.io

import java.io.{ByteArrayInputStream, ByteArrayOutputStream}
import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.addressing.CellRange
import com.tjclp.xl.cells.{CellValue, FormulaKind}
import com.tjclp.xl.io.streaming.StreamingTransform
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.DynamicArrayFixtures

/**
 * GH-714 streaming parity: both SAX readers resolve a dynamic-array anchor's `cm` through the
 * metadata part like the DOM reader, and a streamed value patch over an anchor drops the cell
 * metadata (`cm`/`vm`) that described the replaced content.
 */
class DynamicArrayStreamingSpec extends CatsEffectSuite:

  private val dynamicSort: FormulaKind = FormulaKind.dynamicArray(CellRange(ref"B1", ref"B5"))

  private def fixture(bytes: Array[Byte]): IO[Path] = IO.blocking {
    val p = Files.createTempFile("xl-gh714-stream-", ".xlsx")
    Files.write(p, bytes)
    p
  }

  test("the row stream reads B1 as a dynamic array") {
    for
      path <- fixture(DynamicArrayFixtures.xlsx())
      rows <- ExcelIO.instance[IO].readStream(path).compile.toVector
    yield
      val b1 = rows.find(_.rowIndex == 1).flatMap(_.cells.get(1))
      assertEquals(kindOf(b1), Some(dynamicSort))
  }

  test("the single-cell reader reads B1 as a dynamic array") {
    for
      path <- fixture(DynamicArrayFixtures.xlsx())
      details <- ExcelIO.instance[IO].streamCellDetails(path, "Sheet1", ref"B1")
    yield assertEquals(kindOf(Some(details.value)), Some(dynamicSort))
  }

  test("without the metadata part both readers report a legacy array") {
    for
      path <- fixture(DynamicArrayFixtures.xlsx(DynamicArrayFixtures.parts(metadata = None)))
      rows <- ExcelIO.instance[IO].readStream(path).compile.toVector
      details <- ExcelIO.instance[IO].streamCellDetails(path, "Sheet1", ref"B1")
    yield
      val legacy = FormulaKind.ArrayFormula(CellRange(ref"B1", ref"B5"))
      assertEquals(kindOf(rows.find(_.rowIndex == 1).flatMap(_.cells.get(1))), Some(legacy))
      assertEquals(kindOf(Some(details.value)), Some(legacy))
  }

  private def transform(patch: StreamingTransform.CellPatch, sheetXml: String): String =
    val out = new ByteArrayOutputStream()
    StreamingTransform.transformWorksheet(
      new ByteArrayInputStream(sheetXml.getBytes(StandardCharsets.UTF_8)),
      out,
      Map(ref"B1" -> patch)
    )
    out.toString("UTF-8")

  test("a streamed value patch over a dynamic anchor drops cm (and vm)") {
    val sheet = DynamicArrayFixtures.sortSheetXml.replace(
      """<c r="B1" cm="1">""",
      """<c r="B1" cm="1" vm="2">"""
    )
    val putf = transform(
      StreamingTransform.CellPatch.SetValue(CellValue.Formula("A1*2", None)),
      sheet
    )
    assert(putf.contains("""<c r="B1" t="str"><f>A1*2</f></c>"""), putf)
    val put =
      transform(StreamingTransform.CellPatch.SetValue(CellValue.Number(BigDecimal(7))), sheet)
    assert(put.contains("""<c r="B1"><v>7</v></c>"""), put)
    // untouched neighbours, and a style-only patch, pass through with their attributes
    val styled = transform(StreamingTransform.CellPatch.SetStyle(3), sheet)
    assert(styled.contains("""<c r="B1" cm="1" vm="2" s="3">"""), styled)
  }

  private def kindOf(v: Option[CellValue]): Option[FormulaKind] = v.collect {
    case CellValue.Formula(_, _, kind) => kind
  }

  test("the row-stream writer degrades a dynamic array to a legacy record (no cm, no part)") {
    // ExcelIO.writeStream ships no metadata part, so it cannot point a `cm` anywhere: the record
    // is written as a CSE array. Pinned until streaming authoring lands (GH-714 follow-up).
    val row = RowData(
      1,
      Map(
        0 -> CellValue.Number(BigDecimal(2)),
        1 -> CellValue.Formula("A1*2", None, FormulaKind.dynamicArray(CellRange(ref"B1", ref"B1")))
      )
    )
    for
      path <- IO.blocking(Files.createTempFile("xl-gh714-writestream-", ".xlsx"))
      _ <- fs2.Stream
        .emit(row)
        .covary[IO]
        .through(ExcelIO.instance[IO].writeStream(path, "S"))
        .compile
        .drain
      bytes <- IO.blocking(Files.readAllBytes(path))
      wb <- ExcelIO.instance[IO].read(path)
    yield
      assertEquals(DynamicArrayFixtures.entryText(bytes, "xl/metadata.xml"), None)
      val b1 = wb.sheets.headOption.flatMap(_.cells.get(ref"B1")).map(_.value)
      assertEquals(kindOf(b1), Some(FormulaKind.ArrayFormula(CellRange(ref"B1", ref"B1"))))
  }
