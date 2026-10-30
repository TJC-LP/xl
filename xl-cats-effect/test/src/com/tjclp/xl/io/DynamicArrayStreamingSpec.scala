package com.tjclp.xl.io

import java.io.{ByteArrayInputStream, ByteArrayOutputStream}
import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}

import cats.effect.IO
import cats.syntax.all.*
import munit.CatsEffectSuite

import com.tjclp.xl.addressing.CellRange
import com.tjclp.xl.cells.{ArrayMode, CellValue, FormulaKind}
import com.tjclp.xl.io.streaming.{StreamingTransform, ZipTransformer}
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.{DynamicArrayFixtures, SheetMetadataPart}
import com.tjclp.xl.styles.CellStyle

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

  test("both streaming readers follow a renamed metadata part and find later block records") {
    val parts = DynamicArrayFixtures
      .relocatedMetadataParts(
        metadata = DynamicArrayFixtures.multiRecordMetadataXml
      )
      .reverse
    for
      path <- fixture(DynamicArrayFixtures.xlsx(parts))
      rows <- ExcelIO.instance[IO].readStream(path).compile.toVector
      details <- ExcelIO.instance[IO].streamCellDetails(path, "Sheet1", ref"B1")
    yield
      assertEquals(kindOf(rows.find(_.rowIndex == 1).flatMap(_.cells.get(1))), Some(dynamicSort))
      assertEquals(kindOf(Some(details.value)), Some(dynamicSort))
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

  test("GH-749: the row-stream writer authors a dynamic anchor and metadata") {
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
      assert(DynamicArrayFixtures.entryText(bytes, "xl/metadata.xml").exists(_.contains("XLDAPR")))
      val b1 = wb.sheets.headOption.flatMap(_.cells.get(ref"B1")).map(_.value)
      assertEquals(kindOf(b1), Some(FormulaKind.dynamicArray(CellRange(ref"B1", ref"B1"))))
  }

  test("GH-749: streamed read-write preserves dynamic arrays through every row writer") {
    val excel = ExcelIO.instance[IO]
    val collapsed = FormulaKind.ArrayFormula(
      CellRange(ref"C1", ref"C1"),
      mode = ArrayMode.Dynamic(collapsed = true)
    )
    for
      source <- fixture(DynamicArrayFixtures.xlsx())
      output <- fixture(Array.emptyByteArray)
      _ <- (0 until 6).toList.traverse_ { writer =>
        val rows = excel.readStream(source).map { row =>
          if row.rowIndex == 1 then
            row.copy(cells = row.cells + (2 -> CellValue.Formula("A1", None, collapsed)))
          else row
        }
        val styled = rows.map(row => StyledRowData(row.rowIndex, row.cells, Map.empty))
        val write = writer match
          case 0 => rows.through(excel.writeStream(output, "Sheet1")).compile.drain
          case 1 => rows.through(excel.writeStreamWithAutoDetect(output, "Sheet1")).compile.drain
          case 2 =>
            styled
              .through(excel.writeStreamStyled(output, "Sheet1", Vector(CellStyle.default)))
              .compile
              .drain
          case 3 =>
            styled
              .through(
                excel.writeStreamStyledWithAutoDetect(output, "Sheet1", Vector(CellStyle.default))
              )
              .compile
              .drain
          case 4 => excel.writeStreamsSeq(output, Seq("Sheet1" -> rows))
          case _ => excel.writeStreamsSeqWithAutoDetect(output, Seq("Sheet1" -> rows))
        for
          _ <- write
          wb <- excel.read(output)
          bytes <- IO.blocking(Files.readAllBytes(output))
        yield
          val sheet = wb.sheets.headOption.getOrElse(fail("missing sheet"))
          assertEquals(
            kindOf(sheet.cells.get(ref"B1").map(_.value)),
            Some(dynamicSort),
            s"writer $writer"
          )
          assertEquals(
            kindOf(sheet.cells.get(ref"C1").map(_.value)),
            Some(collapsed),
            s"writer $writer"
          )
          val metadata = DynamicArrayFixtures
            .entryText(bytes, "xl/metadata.xml")
            .getOrElse(fail("missing metadata"))
          val index = SheetMetadataPart.parse(metadata)
          assertEquals(index.byCm.size, 2)
      }
    yield ()
  }

  test("GH-749: package streaming patches retain source metadata indices and bytes") {
    // A source cm=2 must not be renumbered when a style or an unrelated cell changes.
    val sheet = DynamicArrayFixtures.sortSheetXml.replace("cm=\"1\"", "cm=\"2\"")
    val metadata = DynamicArrayFixtures.metadataXml.replace(
      "<cellMetadata count=\"1\"><bk>",
      "<cellMetadata count=\"2\"><bk><rc t=\"1\" v=\"0\"/></bk><bk>"
    )
    val parts = DynamicArrayFixtures.parts().map { (name, xml) =>
      name -> (if name == "xl/worksheets/sheet1.xml" then sheet
               else if name == "xl/metadata.xml" then metadata
               else xml)
    }
    for
      source <- fixture(DynamicArrayFixtures.xlsx(parts))
      output <- fixture(Array.emptyByteArray)
      _ <- ZipTransformer.transform[IO](
        source,
        output,
        "xl/worksheets/sheet1.xml",
        Map(ref"B1" -> StreamingTransform.CellPatch.SetStyle(0))
      )
      bytes <- IO.blocking(Files.readAllBytes(output))
      _ <- IO {
        assertEquals(DynamicArrayFixtures.entryText(bytes, "xl/metadata.xml"), Some(metadata))
        assert(
          DynamicArrayFixtures
            .entryText(bytes, "xl/worksheets/sheet1.xml")
            .exists(_.contains("cm=\"2\""))
        )
      }
      _ <- ZipTransformer.transformValues[IO](
        source,
        output,
        "xl/worksheets/sheet1.xml",
        Map(ref"B1" -> CellValue.Number(BigDecimal(7)))
      )
      overwritten <- IO.blocking(Files.readAllBytes(output))
    yield
      assertEquals(DynamicArrayFixtures.entryText(overwritten, "xl/metadata.xml"), Some(metadata))
      assert(
        !DynamicArrayFixtures
          .entryText(overwritten, "xl/worksheets/sheet1.xml")
          .exists(_.contains("cm="))
      )
  }

  test("GH-749: ordinary workbook bytes match the pre-fix row writers") {
    val excel = ExcelIO.instance[IO]
    val row = RowData(
      1,
      Map(
        0 -> CellValue.Text("same"),
        1 -> CellValue.Number(BigDecimal(3)),
        2 -> CellValue.Formula("B1*2", Some(CellValue.Number(BigDecimal(6))))
      )
    )
    val rows = fs2.Stream.emit(row).covary[IO]
    val styled = rows.map(r => StyledRowData(r.rowIndex, r.cells, Map.empty))
    for
      output <- fixture(Array.emptyByteArray)
      hashes <- (0 until 6).toList.traverse { writer =>
        val write = writer match
          case 0 => rows.through(excel.writeStream(output, "S")).compile.drain
          case 1 => rows.through(excel.writeStreamWithAutoDetect(output, "S")).compile.drain
          case 2 => styled.through(excel.writeStreamStyled(output, "S", Vector.empty)).compile.drain
          case 3 =>
            styled
              .through(excel.writeStreamStyledWithAutoDetect(output, "S", Vector.empty))
              .compile
              .drain
          case 4 => excel.writeStreamsSeq(output, Seq("S" -> rows))
          case _ => excel.writeStreamsSeqWithAutoDetect(output, Seq("S" -> rows))
        write *> IO.blocking {
          val bytes = Files.readAllBytes(output)
          assertEquals(DynamicArrayFixtures.entryText(bytes, "xl/metadata.xml"), None)
          java.security.MessageDigest
            .getInstance("SHA-256")
            .digest(bytes)
            .map(b => f"${b & 0xff}%02x")
            .mkString
        }
      }
    yield
      // Captured from a19dbb6 before changing either production writer (JDK pinned by Mill).
      val plain = "8becbf6722fa46e3b38c56665ec764c592e4397f11e1070a641e928784ca8948"
      val dimensioned = "c511b4b8e002bb15362f34012cd6b2ad5605e548a7c618b5d69bc676831f8821"
      assertEquals(hashes, List(plain, dimensioned, plain, dimensioned, plain, dimensioned))
  }

  test("GH-749: metadata indices remain valid across sheets and repeated runs") {
    val collapsed =
      FormulaKind.ArrayFormula(CellRange(ref"A1", ref"A1"), mode = ArrayMode.Dynamic(true))
    val normal = FormulaKind.dynamicArray(CellRange(ref"A1", ref"A1"))
    val legacy = FormulaKind.ArrayFormula(CellRange(ref"B1", ref"B1"))
    val first = RowData(1, Map(0 -> CellValue.Formula("1", None, collapsed)))
    val second = RowData(
      1,
      Map(0 -> CellValue.Formula("2", None, normal), 1 -> CellValue.Formula("3", None, legacy))
    )
    val excel = ExcelIO.instance[IO]
    for
      output <- fixture(Array.emptyByteArray)
      _ <- List(false, true).traverse_ { auto =>
        val sheets = Seq(
          "Collapsed" -> fs2.Stream.emit(first).covary[IO],
          "Normal" -> fs2.Stream.emit(second).covary[IO]
        )
        val write =
          if auto then excel.writeStreamsSeqWithAutoDetect(output, sheets)
          else excel.writeStreamsSeq(output, sheets)
        for
          _ <- write
          before <- IO.blocking(Files.readAllBytes(output).toVector)
          _ <- write
          after <- IO.blocking(Files.readAllBytes(output).toVector)
          wb <- excel.read(output)
        yield
          assertEquals(after, before)
          assertEquals(
            kindOf(wb.sheets.lift(0).flatMap(_.cells.get(ref"A1")).map(_.value)),
            Some(collapsed)
          )
          assertEquals(
            kindOf(wb.sheets.lift(1).flatMap(_.cells.get(ref"A1")).map(_.value)),
            Some(normal)
          )
          assertEquals(
            kindOf(wb.sheets.lift(1).flatMap(_.cells.get(ref"B1")).map(_.value)),
            Some(legacy)
          )
      }
    yield ()
  }
