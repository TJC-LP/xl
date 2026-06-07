package com.tjclp.xl.ooxml

import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.worksheet.OoxmlCell
import munit.FunSuite

/**
 * GH-237 (control-char sanitization) + GH-238 (plain-decimal numerics).
 *
 * Both write paths (ScalaXml `toXml` + SaxStax `writeCharacters`) share `XmlUtil.sanitizeXmlText`
 * and `XmlUtil.plainNumber`, so output is backend-independent and OOXML-valid.
 */
class XmlValueSanitizationSpec extends FunSuite:

  // ===== plainNumber: never scientific notation in <v> (GH-238) =====

  test("plainNumber renders large/small magnitudes as plain decimals") {
    assertEquals(XmlUtil.plainNumber(BigDecimal("1E20")), "100000000000000000000")
    assertEquals(XmlUtil.plainNumber(BigDecimal("1E-7")), "0.0000001")
    assertEquals(XmlUtil.plainNumber(BigDecimal(30)), "30")
    assertEquals(XmlUtil.plainNumber(BigDecimal("1234.50")), "1234.50")
  }

  test("plainNumber(Double) keeps normal serials identical to the old toString") {
    assertEquals(XmlUtil.plainNumber(45292.0), "45292.0")
    assertEquals(XmlUtil.plainNumber(45292.5), "45292.5")
  }

  test("OoxmlCell number cell emits a plain decimal in <v> (no E)") {
    val cell = OoxmlCell(ref"A1", CellValue.Number(BigDecimal("1E20")), None, "n")
    val v = (cell.toXml \\ "v").text
    assertEquals(v, "100000000000000000000")
    assert(!v.toUpperCase.contains("E"), s"scientific notation leaked into <v>: $v")
  }

  // ===== sanitizeXmlText: strip XML-illegal control chars (GH-237) =====

  test("sanitizeXmlText strips C0 control chars but keeps tab/newline/CR") {
    // U+0001 and U+001F are XML-illegal; tab/newline/CR are legal and must survive.
    val withControls = "a" + 1.toChar + "b" + 31.toChar + "c"
    assertEquals(XmlUtil.sanitizeXmlText(withControls), "abc")
    val withWhitespace = "a" + 9.toChar + "b" + 10.toChar + "c" + 13.toChar + "d"
    assertEquals(XmlUtil.sanitizeXmlText(withWhitespace), withWhitespace)
  }

  test("sanitizeXmlText returns the same instance for already-clean text (fast path)") {
    val s = "clean text"
    assert(XmlUtil.sanitizeXmlText(s) eq s)
  }

  test("sanitizeXmlText preserves astral characters (emoji surrogate pairs)") {
    val emoji = "hi " + new String(Character.toChars(0x1F600)) // grinning face
    assertEquals(XmlUtil.sanitizeXmlText(emoji), emoji)
  }

  test("OoxmlCell text cell strips control chars from <t>") {
    val cell = OoxmlCell(ref"A1", CellValue.Text("a" + 1.toChar + "b"), None, "inlineStr")
    assertEquals((cell.toXml \\ "t").text, "ab")
  }
