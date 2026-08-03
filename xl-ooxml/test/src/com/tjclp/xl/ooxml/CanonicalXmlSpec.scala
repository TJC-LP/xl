package com.tjclp.xl.ooxml

import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.ZipFile

import com.tjclp.xl.api.*
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.macros.ref
import com.tjclp.xl.sheets.syntax.withCellStyle
import com.tjclp.xl.sheets.{ColumnProperties, RowProperties}
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.color.{Color, ThemeSlot}
import com.tjclp.xl.styles.font.Font
import munit.FunSuite

/**
 * GH-448: Excel-canonical XML forms, so an xl-authored sheet is textually identical to the same
 * sheet authored by Excel:
 *   - integral font sizes without a decimal point (`sz="10"`, not `sz="10.0"`)
 *   - theme tints in Excel's ≤17-significant-digit plain form, omitted entirely when 0
 *   - the mandatory gray125 placeholder fill written bare
 *   - `sheetFormatPr` outline summary attributes stamped from row/col outline levels
 *   - theme slot indices via themeSlotToIndex on EVERY path (the ThemeSlot enum order swaps
 *     dark/light pairs relative to OOXML indices, so `slot.ordinal` wrote Dark2 as Light2)
 */
class CanonicalXmlSpec extends FunSuite:

  private def zipEntryString(path: Path, entry: String): String =
    val zf = new ZipFile(path.toFile)
    try
      val is = zf.getInputStream(zf.getEntry(entry))
      try new String(is.readAllBytes(), StandardCharsets.UTF_8)
      finally is.close()
    finally zf.close()

  private def written(wb: Workbook, prefix: String): Path =
    val out = Files.createTempFile(prefix, ".xlsx")
    XlsxWriter.write(wb, out).fold(e => fail(s"write failed: $e"), identity)
    out

  test("GH-448: integral font sizes emit without a decimal point") {
    val wb = Workbook(
      Sheet("Sheet1")
        .put(ref"A1" -> 1)
        .withCellStyle(ref"A1", CellStyle.default.withFont(Font("Times New Roman", 10.0)))
    )
    val out = written(wb, "canon-sz")
    val styles = zipEntryString(out, "xl/styles.xml")
    assert(styles.contains("<sz val=\"10\"/>"), s"integral sz must drop the decimal: $styles")
    assert(!styles.contains("val=\"10.0\""), s"sz=10.0 must not appear: $styles")
    Files.deleteIfExists(out)
  }

  test("GH-448: fractional font sizes keep their decimals") {
    val wb = Workbook(
      Sheet("Sheet1")
        .put(ref"A1" -> 1)
        .withCellStyle(ref"A1", CellStyle.default.withFont(Font("Arial", 10.5)))
    )
    val out = written(wb, "canon-szfrac")
    val styles = zipEntryString(out, "xl/styles.xml")
    assert(styles.contains("<sz val=\"10.5\"/>"), s"fractional sz preserved: $styles")
    Files.deleteIfExists(out)
  }

  test("GH-448: theme tint uses Excel's 17-significant-digit plain form") {
    val wb = Workbook(
      Sheet("Sheet1")
        .put(ref"A1" -> 1)
        .withTabColor(Color.Theme(ThemeSlot.Accent1, 0.79998168889431442))
    )
    val out = written(wb, "canon-tint")
    val xml = zipEntryString(out, "xl/worksheets/sheet1.xml")
    assert(
      xml.contains("tint=\"0.79998168889431442\""),
      s"tint must be the 17-digit Excel form, not the shortest round-trip: $xml"
    )
    Files.deleteIfExists(out)
  }

  test("GH-448: a zero tint is omitted, and Dark2 maps to theme index 3 (not its ordinal)") {
    val wb = Workbook(
      Sheet("Sheet1").put(ref"A1" -> 1).withTabColor(Color.Theme(ThemeSlot.Dark2, 0.0))
    )
    val out = written(wb, "canon-tint0")
    val xml = zipEntryString(out, "xl/worksheets/sheet1.xml")
    assert(xml.contains("<tabColor theme=\"3\"/>"), s"bare theme=3 tabColor expected: $xml")
    assert(!xml.contains("tint="), s"tint=0 must be omitted: $xml")
    Files.deleteIfExists(out)
  }

  test("GH-448: the mandatory gray125 placeholder fill serializes bare") {
    val wb = Workbook(Sheet("Sheet1").put(ref"A1" -> 1))
    val out = written(wb, "canon-gray")
    val styles = zipEntryString(out, "xl/styles.xml")
    assert(
      styles.contains("<patternFill patternType=\"gray125\"/>"),
      s"gray125 must be bare like Excel writes it: $styles"
    )
    Files.deleteIfExists(out)
  }

  test("GH-448: outline summary attributes stamp onto sheetFormatPr from row/col levels") {
    val sheet = Sheet("Sheet1")
      .put(ref"A1" -> 1, ref"A2" -> 2)
      .setRowProperties(ref"A2".row, RowProperties(outlineLevel = Some(1)))
      .setColumnProperties(ref"C1".col, ColumnProperties(outlineLevel = Some(2)))
    val out = written(Workbook(sheet), "canon-outline")
    val xml = zipEntryString(out, "xl/worksheets/sheet1.xml")
    assert(xml.contains("outlineLevelRow=\"1\""), s"outlineLevelRow missing: $xml")
    assert(xml.contains("outlineLevelCol=\"2\""), s"outlineLevelCol missing: $xml")
    Files.deleteIfExists(out)
  }
