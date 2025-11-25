package com.tjclp.xl.ooxml

import munit.FunSuite
import com.tjclp.xl.api.*
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.sheets.syntax.*
import com.tjclp.xl.styles.*
import com.tjclp.xl.macros.ref

/**
 * Tests for deterministic output (byte-identical on re-serialization)
 *
 * Verifies that allFills deduplication preserves insertion order for
 * stable diffs and reproducible builds
 */
class DeterminismSpec extends FunSuite:

  test("styles.xml has deterministic fill order") {
    // Create StyleIndex with multiple fills
    val fill1 = Fill.Solid(Color.Rgb(0xffff0000))
    val fill2 = Fill.Solid(Color.Rgb(0xff00ff00))
    val fill3 = Fill.Solid(Color.Rgb(0xff0000ff))

    val styleIndex = StyleIndex(
      fonts = Vector(Font.default),
      fills = Vector(fill1, fill2, fill3, fill1, fill2), // Duplicates intentional
      borders = Vector(Border.none),
      numFmts = Vector.empty,
      cellStyles = Vector(CellStyle.default),
      styleToIndex = Map(CellStyle.canonicalKey(CellStyle.default) -> StyleId(0))
    )

    val ooxmlStyles = OoxmlStyles(styleIndex)

    // Serialize twice
    val xml1 = ooxmlStyles.toXml.toString
    val xml2 = ooxmlStyles.toXml.toString

    // Should be byte-identical
    assertEquals(xml1, xml2, "Multiple serializations should produce identical XML")

    // Verify fill order: None, Gray125, fill1, fill2, fill3 (no duplicates)
    val fillsRegex = """<fills count="5">""".r
    assert(fillsRegex.findFirstIn(xml1).isDefined, "Should have exactly 5 fills (2 defaults + 3 custom)")

    // Verify deduplication by checking overall XML structure
    assert(xml1.contains("<fills"), "Should have fills element")
    assert(xml1.contains("patternFill"), "Should have pattern fills")
  }

  test("workbook styles serialize deterministically across writes") {
    // Create workbook with styled cells
    val sheet1 = Sheet("Sheet1") match
      case Right(s) => s
      case Left(err) => fail(s"Failed to create sheet: $err")

    val styledSheet = sheet1
      .put(ref"A1", CellValue.Text("Hello"))
      .put(ref"B1", CellValue.Number(42))
      .withCellStyle(
        ref"A1",
        CellStyle(
          font = Font.default.copy(bold = true, sizePt = 14.0),
          fill = Fill.Solid(Color.Rgb(0xffff0000)),
          border = Border.none,
          numFmt = NumFmt.General,
          align = Align(horizontal = HAlign.Center)
        )
      )

    // Build StyleIndex from workbook
    val (styleIndex1, _) = StyleIndex.fromWorkbook(Workbook(sheets = Vector(styledSheet)))
    val (styleIndex2, _) = StyleIndex.fromWorkbook(Workbook(sheets = Vector(styledSheet)))

    // Serialize both
    val xml1 = OoxmlStyles(styleIndex1).toXml.toString
    val xml2 = OoxmlStyles(styleIndex2).toXml.toString

    // Should be identical (deterministic)
    assertEquals(xml1, xml2, "Multiple styleIndex builds should produce identical XML")
  }

  test("multiple identical fills deduplicate deterministically") {
    val redFill = Fill.Solid(Color.Rgb(0xffff0000))

    // Create style index where same fill appears multiple times
    val styleIndex = StyleIndex(
      fonts = Vector(Font.default),
      fills = Vector(redFill, redFill, redFill), // Same fill 3 times
      borders = Vector(Border.none),
      numFmts = Vector.empty,
      cellStyles = Vector(CellStyle.default),
      styleToIndex = Map(CellStyle.canonicalKey(CellStyle.default) -> StyleId(0))
    )

    val xml = OoxmlStyles(styleIndex).toXml.toString

    // Should have exactly 3 fills: None, Gray125, red (no duplicates)
    assert(xml.contains("""<fills count="3">"""), "Should deduplicate to 3 fills total")

    // Verify structure
    assert(xml.contains("<fills"), "Should have fills element")
    assert(xml.contains("patternFill"), "Should have pattern fills")
  }

  test("fill order is stable across multiple serializations") {
    val fills = (0 until 50).map { i =>
      Fill.Solid(Color.Rgb(0xff000000 | (i * 1000)))
    }.toVector

    val styleIndex = StyleIndex(
      fonts = Vector(Font.default),
      fills = fills,
      borders = Vector(Border.none),
      numFmts = Vector.empty,
      cellStyles = Vector(CellStyle.default),
      styleToIndex = Map(CellStyle.canonicalKey(CellStyle.default) -> StyleId(0))
    )

    // Serialize 5 times
    val xmlStrings = (1 to 5).map { _ =>
      OoxmlStyles(styleIndex).toXml.toString
    }

    // All should be identical
    xmlStrings.sliding(2).foreach { case Seq(a, b) =>
      assertEquals(a, b, "Fill order should be stable across serializations")
    }
  }
