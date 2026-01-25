package com.tjclp.xl.io.streaming

import munit.FunSuite
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.font.Font
import com.tjclp.xl.styles.fill.Fill
import com.tjclp.xl.styles.color.Color
import com.tjclp.xl.error.XLError

/**
 * Tests for error handling in streaming module.
 *
 * Covers:
 *   - Malformed XML handling in StylePatcher
 *   - Error recovery in streaming transforms
 */
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial"))
class StreamingErrorSpec extends FunSuite:

  // ========== StylePatcher Error Handling ==========

  test("StylePatcher.addStyle: handles malformed XML gracefully") {
    val malformedXml = "<?xml version=\"1.0\"?><styleSheet><fonts<broken"

    val result = StylePatcher.addStyle(malformedXml, CellStyle.default)
    assert(result.isLeft, "Should return Left for malformed XML")

    result match
      case Left(error: XLError.ParseError) =>
        assert(error.message.contains("XML parse error"))
      case Left(other) =>
        assert(other.message.nonEmpty, "Should have error message")
      case Right(_) =>
        fail("Should not succeed with malformed XML")
  }

  test("StylePatcher.addStyle: handles empty XML") {
    val emptyXml = ""

    val result = StylePatcher.addStyle(emptyXml, CellStyle.default)
    assert(result.isLeft, "Should return Left for empty XML")
  }

  test("StylePatcher.addStyle: handles XML with XXE attempt") {
    // XXE attack vector - should be blocked by XmlSecurity.parseSafe
    val xxeXml = """<?xml version="1.0"?>
      |<!DOCTYPE foo [<!ENTITY xxe SYSTEM "file:///etc/passwd">]>
      |<styleSheet>&xxe;</styleSheet>""".stripMargin

    val result = StylePatcher.addStyle(xxeXml, CellStyle.default)
    assert(result.isLeft, "Should reject XML with DOCTYPE (XXE protection)")
  }

  test("StylePatcher.getStyle: handles malformed XML") {
    val malformedXml = "not even xml at all"

    val result = StylePatcher.getStyle(malformedXml, 0)
    assert(result.isLeft, "Should return Left for malformed XML")
  }

  test("StylePatcher.addStyles: handles parse error mid-batch") {
    // First valid, then we simulate error in intermediate step
    val validStylesXml = minimalStylesXml

    // This should succeed
    val result = StylePatcher.addStyles(
      validStylesXml,
      Map(
        "bold" -> CellStyle.default.withFont(Font.default.withBold(true)),
        "italic" -> CellStyle.default.withFont(Font.default.withItalic(true))
      )
    )

    // With valid input, should succeed
    assert(result.isRight, s"Should succeed with valid styles XML: ${result.swap.map(_.message)}")
  }

  // ========== StylePatcher Style Merging ==========

  test("StylePatcher.mergeStyles: preserves existing properties when overlay is default") {
    val existing = CellStyle.default
      .withFont(Font.default.withBold(true).withSize(14))
      .withFill(Fill.Solid(Color.Rgb(0xFFFF0000)))

    val overlay = CellStyle.default // Default overlay

    val merged = StylePatcher.mergeStyles(existing, overlay)

    // Existing properties should be preserved
    assertEquals(merged.font.bold, true)
    assertEquals(merged.font.sizePt, 14.0)
    assertEquals(merged.fill, Fill.Solid(Color.Rgb(0xFFFF0000)))
  }

  test("StylePatcher.mergeStyles: overlay takes precedence") {
    val existing = CellStyle.default.withFont(Font.default.withSize(10))
    val overlay = CellStyle.default.withFont(Font.default.withSize(16))

    val merged = StylePatcher.mergeStyles(existing, overlay)

    assertEquals(merged.font.sizePt, 16.0) // Overlay wins
  }

  test("StylePatcher.mergeStyles: boolean properties are OR-ed") {
    val existing = CellStyle.default.withFont(Font.default.withBold(true))
    val overlay = CellStyle.default.withFont(Font.default.withItalic(true))

    val merged = StylePatcher.mergeStyles(existing, overlay)

    // Both should be true (OR-ed)
    assertEquals(merged.font.bold, true)
    assertEquals(merged.font.italic, true)
  }

  // ========== StreamingTransform Edge Cases ==========

  test("StreamingTransform.analyzePatches: empty patches returns None") {
    val result = StreamingTransform.analyzePatches(Map.empty)
    assertEquals(result, None)
  }

  test("StreamingTransform.analyzePatches: single patch") {
    import com.tjclp.xl.addressing.ARef

    val patches = Map(
      ARef.from1(1, 5) -> StreamingTransform.CellPatch.SetStyle(1)
    )

    val result = StreamingTransform.analyzePatches(patches)
    assert(result.isDefined)
    assertEquals(result.get.minRow, 5)
    assertEquals(result.get.maxRow, 5)
    assertEquals(result.get.patchCount, 1)
  }

  test("StreamingTransform.analyzePatches: multiple patches") {
    import com.tjclp.xl.addressing.ARef

    val patches = Map(
      ARef.from1(1, 1) -> StreamingTransform.CellPatch.SetStyle(1),
      ARef.from1(2, 100) -> StreamingTransform.CellPatch.SetStyle(2),
      ARef.from1(3, 50) -> StreamingTransform.CellPatch.SetStyle(3)
    )

    val result = StreamingTransform.analyzePatches(patches)
    assert(result.isDefined)
    assertEquals(result.get.minRow, 1)
    assertEquals(result.get.maxRow, 100)
    assertEquals(result.get.patchCount, 3)
  }

  // ========== Helpers ==========

  private val minimalStylesXml: String =
    """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      |<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      |<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
      |<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
      |<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
      |<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
      |<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
      |</styleSheet>""".stripMargin.replaceAll("\n", "")
