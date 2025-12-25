package com.tjclp.xl.ooxml

import munit.FunSuite
import java.nio.file.{Files, Path}
import java.time.LocalDateTime
import com.tjclp.xl.addressing.{CellRange, Column, Row}
import com.tjclp.xl.api.{Sheet, Workbook}
import com.tjclp.xl.cells.{CellError, CellValue, Comment}
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.macros.ref
import com.tjclp.xl.richtext.RichText.*
import com.tjclp.xl.sheets.{ColumnProperties, RowProperties}
import com.tjclp.xl.sheets.styleSyntax.*
import com.tjclp.xl.styles.{CellStyle, Font}

class SaxStaxRoundTripSpec extends FunSuite:

  val tempDir: Path = Files.createTempDirectory("xl-sax-roundtrip-")

  override def afterAll(): Unit =
    Files.walk(tempDir)
      .sorted(java.util.Comparator.reverseOrder())
      .forEach(Files.delete)

  test("SaxStax output round-trips and matches ScalaXml") {
    val workbook = buildWorkbook()
    val saxPath = tempDir.resolve("sax.xlsx")
    val scalaPath = tempDir.resolve("scala.xml.xlsx")

    val saxConfig = WriterConfig.default.copy(sstPolicy = SstPolicy.Always)
    val scalaConfig = WriterConfig.scalaXml.copy(sstPolicy = SstPolicy.Always)

    XlsxWriter.writeWith(workbook, saxPath, saxConfig).getOrElse(fail("SaxStax write failed"))
    XlsxWriter.writeWith(workbook, scalaPath, scalaConfig).getOrElse(fail("ScalaXml write failed"))

    val saxWb = XlsxReader.read(saxPath).getOrElse(fail("SaxStax read failed"))
    val scalaWb = XlsxReader.read(scalaPath).getOrElse(fail("ScalaXml read failed"))

    assertEquals(stripSource(saxWb), stripSource(scalaWb))

    val saxSheet = saxWb.sheets.headOption.getOrElse(fail("Expected worksheet"))
    assertEquals(saxSheet(ref"A1").value, CellValue.Text("  spaced"))
    assertEquals(saxSheet(ref"B1").value, CellValue.Number(BigDecimal(42)))
    assertEquals(saxSheet(ref"C1").value, CellValue.Bool(true))
    assertEquals(
      saxSheet(ref"D1").value,
      CellValue.Formula("SUM(A1:B1)", Some(CellValue.Number(BigDecimal(3))))
    )
    val richValue = saxSheet(ref"A2").value match
      case CellValue.RichText(rt) => rt
      case other => fail(s"Expected rich text in A2, got: $other")
    assertEquals(richValue.toPlainText, "Bold plain")
    val richRun = richValue.runs.headOption.getOrElse(fail("Expected rich text run"))
    val hasBold =
      richRun.font.exists(_.bold) || richRun.rawRPrXml.exists(_.contains("<b"))
    val hasRed =
      richRun.font.flatMap(_.color) match
        case Some(com.tjclp.xl.styles.color.Color.Rgb(argb)) => argb == 0xffff0000
        case _ => richRun.rawRPrXml.exists(_.contains("FFFF0000"))
    assert(hasBold && hasRed, "Rich text formatting should preserve bold red styling")
    val expectedSerial = CellValue.dateTimeToExcelSerial(LocalDateTime.of(2024, 1, 2, 3, 4))
    saxSheet(ref"B2").value match
      case CellValue.Number(serial) =>
        assertEquals(serial.toDouble, expectedSerial, 0.0001)
      case other => fail(s"Expected DateTime serial number, got: $other")
    assertEquals(saxSheet(ref"C2").value, CellValue.Error(CellError.Div0))
    assert(saxSheet.mergedRanges.contains(CellRange(ref"A3", ref"B3")))
    assertEquals(saxSheet.comments(ref"A1"), Comment.plainText("Note", Some("QA")))
    assertEquals(
      saxSheet.rowProperties(Row.from1(2)).height,
      Some(24.0)
    )
    assertEquals(
      saxSheet.columnProperties(Column.from1(1)).width,
      Some(12.5)
    )
    assert(saxSheet(ref"A1").styleId.isDefined, "Styled cell should preserve styleId")
  }

  private def stripSource(wb: Workbook): Workbook =
    wb.copy(sourceContext = None)

  private def buildWorkbook(): Workbook =
    val rich = "Bold".bold.red + " plain"
    val sheet = Sheet("Sheet1")
      .put(ref"A1", CellValue.Text("  spaced"))
      .put(ref"B1", CellValue.Number(BigDecimal(42)))
      .put(ref"C1", CellValue.Bool(true))
      .put(ref"D1", CellValue.Formula("SUM(A1:B1)", Some(CellValue.Number(BigDecimal(3)))))
      .put(ref"A2", CellValue.RichText(rich))
      .put(ref"B2", CellValue.DateTime(LocalDateTime.of(2024, 1, 2, 3, 4)))
      .put(ref"C2", CellValue.Error(CellError.Div0))
      .merge(CellRange(ref"A3", ref"B3"))
      .comment(ref"A1", Comment.plainText("Note", Some("QA")))
      .setRowProperties(
        Row.from1(2),
        RowProperties(height = Some(24.0), hidden = true, outlineLevel = Some(1), collapsed = true)
      )
      .setColumnProperties(
        Column.from1(1),
        ColumnProperties(width = Some(12.5), outlineLevel = Some(1))
      )
      .setColumnProperties(
        Column.from1(2),
        ColumnProperties(width = Some(12.5), outlineLevel = Some(1))
      )
      .setColumnProperties(
        Column.from1(3),
        ColumnProperties(width = Some(20.0), hidden = true, outlineLevel = Some(2), collapsed = true)
      )
      .withCellStyle(
        ref"A1",
        CellStyle.default.withFont(Font("Arial", 12.0, bold = true))
      )

    Workbook(Vector(sheet))
