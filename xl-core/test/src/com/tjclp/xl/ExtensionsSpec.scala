package com.tjclp.xl

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.error.XLError
import com.tjclp.xl.extensions.given // For CellWriter given instances
import com.tjclp.xl.richtext.RichText
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.unsafe.* // For .unsafe extension
import java.time.{LocalDate, LocalDateTime}
import munit.FunSuite

/**
 * Comprehensive tests for Easy Mode string-based extensions.
 *
 * Covers all extension methods in xl-core/src/com/tjclp/xl/extensions.scala:
 *   - String-based put() for 9 primitive types (unstyled and styled)
 *   - Template style() operations (cell and range)
 *   - Safe lookup methods (cell, range, get)
 *   - Merge operations
 *   - Chainable XLResult[Sheet] extensions
 *
 * These tests establish the API contract baseline for future type class refactor.
 */
class ExtensionsSpec extends FunSuite:

  // Test fixture
  val baseSheet = Sheet("Test").unsafe
  val testStyle = CellStyle.default.bold

  // ========== String-Based put() - Unstyled (9 types) ==========

  test("put String with valid ref succeeds") {
    val result = baseSheet.put("A1", "Hello")
    assert(result.isRight)
    assertEquals(result.unsafe.cell("A1").map(_.value), Some(CellValue.Text("Hello")))
  }

  test("put Int with valid ref succeeds") {
    val result = baseSheet.put("B1", 42)
    assert(result.isRight)
    assertEquals(result.unsafe.cell("B1").map(_.value), Some(CellValue.Number(BigDecimal(42))))
  }

  test("put Long with valid ref succeeds") {
    val result = baseSheet.put("C1", 42L)
    assert(result.isRight)
    assertEquals(result.unsafe.cell("C1").map(_.value), Some(CellValue.Number(BigDecimal(42))))
  }

  test("put Double with valid ref succeeds") {
    val result = baseSheet.put("D1", 3.14)
    assert(result.isRight)
    assertEquals(result.unsafe.cell("D1").map(_.value), Some(CellValue.Number(BigDecimal(3.14))))
  }

  test("put BigDecimal with valid ref succeeds") {
    val result = baseSheet.put("E1", BigDecimal("123.45"))
    assert(result.isRight)
    assertEquals(
      result.unsafe.cell("E1").map(_.value),
      Some(CellValue.Number(BigDecimal("123.45")))
    )
  }

  test("put Boolean with valid ref succeeds") {
    val result = baseSheet.put("F1", true)
    assert(result.isRight)
    assertEquals(result.unsafe.cell("F1").map(_.value), Some(CellValue.Bool(true)))
  }

  test("put LocalDate with valid ref succeeds") {
    val date = LocalDate.of(2025, 11, 19)
    val result = baseSheet.put("G1", date)
    assert(result.isRight)
    assertEquals(
      result.unsafe.cell("G1").map(_.value),
      Some(CellValue.DateTime(date.atStartOfDay))
    )
  }

  test("put LocalDateTime with valid ref succeeds") {
    val dt = LocalDateTime.of(2025, 11, 19, 14, 30)
    val result = baseSheet.put("H1", dt)
    assert(result.isRight)
    assertEquals(result.unsafe.cell("H1").map(_.value), Some(CellValue.DateTime(dt)))
  }

  test("put RichText with valid ref succeeds") {
    val richText = "Bold".bold + " normal"
    val result = baseSheet.put("I1", richText)
    assert(result.isRight)
    val cellValue = result.unsafe.cell("I1").map(_.value)
    assert(cellValue.exists {
      case _: CellValue.RichText => true
      case _ => false
    })
  }

  test("put with invalid ref returns Left") {
    val result = baseSheet.put("INVALID", "value")
    assert(result.isLeft)
    result match
      case Left(XLError.InvalidCellRef(ref, msg)) =>
        assertEquals(ref, "INVALID")
        assert(msg.contains("Invalid cell reference"))
      case _ => fail("Expected InvalidCellRef error")
  }

  test("put with out-of-range column returns Left") {
    val result = baseSheet.put("XFE1", "value") // XFE > XFD (max column)
    assert(result.isLeft)
  }

  test("put with out-of-range row returns Left") {
    val result = baseSheet.put("A1048577", "value") // Row > 1048576 (max row)
    assert(result.isLeft)
  }

  // ========== String-Based put() - Styled (9 types) ==========

  test("put String with style applies both value and style") {
    val result = baseSheet.put("A1", "Bold Text", testStyle)
    assert(result.isRight)
    val sheet = result.unsafe
    assertEquals(sheet.cell("A1").map(_.value), Some(CellValue.Text("Bold Text")))
    // Verify style applied (styleId should be set)
    assert(sheet.cell("A1").flatMap(_.styleId).isDefined)
  }

  test("put Int with style applies both") {
    val result = baseSheet.put("A1", 42, testStyle)
    assert(result.isRight)
    assert(result.unsafe.cell("A1").flatMap(_.styleId).isDefined)
  }

  test("put BigDecimal with style applies both") {
    val result = baseSheet.put("A1", BigDecimal("123.45"), testStyle)
    assert(result.isRight)
    assert(result.unsafe.cell("A1").flatMap(_.styleId).isDefined)
  }

  test("put LocalDate with style applies both") {
    val result = baseSheet.put("A1", LocalDate.of(2025, 11, 19), testStyle)
    assert(result.isRight)
    assert(result.unsafe.cell("A1").flatMap(_.styleId).isDefined)
  }

  test("put styled with invalid ref returns Left") {
    val result = baseSheet.put("INVALID", 42, testStyle)
    assert(result.isLeft)
  }

  // ========== Template style() Operations ==========

  test("style applies to single cell") {
    val result = baseSheet.put("A1", "Text").style("A1", testStyle)
    assert(result.isRight)
    assert(result.unsafe.cell("A1").flatMap(_.styleId).isDefined)
  }

  test("style applies to range") {
    val result = baseSheet
      .put("A1", "A")
      .put("B1", "B")
      .style("A1:B1", testStyle)
    assert(result.isRight)
    val sheet = result.unsafe
    assert(sheet.cell("A1").flatMap(_.styleId).isDefined)
    assert(sheet.cell("B1").flatMap(_.styleId).isDefined)
  }

  test("style with invalid cell ref returns Left") {
    val result = baseSheet.style("INVALID", testStyle)
    assert(result.isLeft)
  }

  test("style with invalid range returns Left") {
    val result = baseSheet.style("A1:INVALID", testStyle)
    assert(result.isLeft)
  }

  // ========== Safe Lookup Methods ==========

  test("cell returns Some for existing cell") {
    val sheet = baseSheet.put("A1", "Value").unsafe
    val cell = sheet.cell("A1")
    assert(cell.isDefined)
    assertEquals(cell.map(_.value), Some(CellValue.Text("Value")))
  }

  test("cell returns None for missing cell") {
    assertEquals(baseSheet.cell("Z99"), None)
  }

  test("cell returns None for invalid ref") {
    assertEquals(baseSheet.cell("INVALID"), None)
  }

  test("range returns all cells in range") {
    val sheet = baseSheet
      .put("A1", "A")
      .put("B1", "B")
      .put("C1", "C")
      .unsafe
    val cells = sheet.range("A1:C1")
    assertEquals(cells.length, 3)
  }

  test("range returns empty for invalid range") {
    assertEquals(baseSheet.range("A1:INVALID"), List.empty)
  }

  test("range returns empty for range with no cells") {
    assertEquals(baseSheet.range("Z1:Z10"), List.empty)
  }

  test("get auto-detects single cell") {
    val sheet = baseSheet.put("A1", "Value").unsafe
    val cells = sheet.get("A1")
    assertEquals(cells.length, 1)
    cells.headOption match
      case Some(cell) => assertEquals(cell.value, CellValue.Text("Value"))
      case None => fail("Expected one cell at A1")
  }

  test("get auto-detects range") {
    val sheet = baseSheet.put("A1", "A").put("B1", "B").unsafe
    val cells = sheet.get("A1:B1")
    assertEquals(cells.length, 2)
  }

  test("get returns empty for invalid ref") {
    assertEquals(baseSheet.get("INVALID"), List.empty)
  }

  // ========== Merge Operations ==========

  test("merge with valid range succeeds") {
    val result = baseSheet.put("A1", "Merged").merge("A1:B1")
    assert(result.isRight)
    val sheet = result.unsafe
    assert(sheet.mergedRanges.nonEmpty)
  }

  test("merge with invalid range returns Left") {
    val result = baseSheet.merge("A1:INVALID")
    assert(result.isLeft)
  }

  // ========== Chainable XLResult[Sheet] Operations (CRITICAL) ==========

  test("chain multiple put operations") {
    val result = baseSheet
      .put("A1", "Title")
      .put("A2", 42)
      .put("A3", true)
      .put("A4", LocalDate.of(2025, 11, 19))

    assert(result.isRight)
    val sheet = result.unsafe
    assertEquals(sheet.cells.size, 4)
    assertEquals(sheet.cell("A1").map(_.value), Some(CellValue.Text("Title")))
    assertEquals(sheet.cell("A2").map(_.value), Some(CellValue.Number(BigDecimal(42))))
  }

  test("chain short-circuits on first error") {
    val result = baseSheet
      .put("A1", "Valid")
      .put("INVALID", "Fail")
      .put("A3", "Never reached")

    assert(result.isLeft)
    result match
      case Left(XLError.InvalidCellRef("INVALID", _)) => // Success
      case _ => fail("Expected InvalidCellRef for INVALID")
  }

  test("chain style after put") {
    val result = baseSheet
      .put("A1", "Bold Text")
      .style("A1", testStyle)

    assert(result.isRight)
    assert(result.unsafe.cell("A1").flatMap(_.styleId).isDefined)
  }

  test("chain put with style") {
    val result = baseSheet
      .put("A1", "First")
      .put("A2", "Styled", testStyle)

    assert(result.isRight)
    val sheet = result.unsafe
    assertEquals(sheet.cells.size, 2)
    assert(sheet.cell("A2").flatMap(_.styleId).isDefined)
  }

  test("chain merge after puts") {
    val result = baseSheet
      .put("A1", "Start")
      .put("B1", "End")
      .merge("A1:B1")

    assert(result.isRight)
    assert(result.unsafe.mergedRanges.nonEmpty)
  }

  test("chain mixed operations") {
    val result = baseSheet
      .put("A1", "Title")
      .style("A1", testStyle)
      .put("A2", 100)
      .put("A3", "Footer")

    assert(result.isRight)
    assertEquals(result.unsafe.cells.size, 3)
  }

  // ========== Template Pattern (style before data) ==========

  test("style before put preserves style (template pattern)") {
    val result = baseSheet
      .style("A1", testStyle)
      .put("A1", "Data")

    assert(result.isRight)
    val sheet = result.unsafe
    sheet.cell("A1") match
      case Some(cell) =>
        // Verify data added and style preserved
        assertEquals(cell.value, CellValue.Text("Data"))
        assert(cell.styleId.isDefined)
      case None => fail("Expected styled cell at A1")
  }

  test("style range then put individual cells preserves formatting") {
    val result = baseSheet
      .style("A1:B1", testStyle)
      .put("A1", "First")
      .put("B1", "Second")

    assert(result.isRight)
    val sheet = result.unsafe
    // Both cells should have style
    assert(sheet.cell("A1").flatMap(_.styleId).isDefined)
    assert(sheet.cell("B1").flatMap(_.styleId).isDefined)
  }

  // ========== Edge Cases ==========

  test("put empty string succeeds") {
    val result = baseSheet.put("A1", "")
    assert(result.isRight)
    assertEquals(result.unsafe.cell("A1").map(_.value), Some(CellValue.Text("")))
  }

  test("put overwrites existing cell") {
    val result = baseSheet
      .put("A1", "First")
      .put("A1", "Second")

    assert(result.isRight)
    assertEquals(result.unsafe.cell("A1").map(_.value), Some(CellValue.Text("Second")))
  }

  test("put preserves other cells") {
    val result = baseSheet
      .put("A1", "Keep")
      .put("B1", "This")

    assert(result.isRight)
    val sheet = result.unsafe
    assertEquals(sheet.cells.size, 2)
    assert(sheet.cell("A1").isDefined)
    assert(sheet.cell("B1").isDefined)
  }

  test("style on empty cell creates styled empty cell") {
    val result = baseSheet.style("A1", testStyle)
    assert(result.isRight)
    val sheet = result.unsafe
    val cell = sheet.cell("A1")
    assert(cell.isDefined)
    assert(cell.flatMap(_.styleId).isDefined)
  }

  test("range returns cells in order") {
    val sheet = baseSheet
      .put("A1", "A")
      .put("A2", "B")
      .put("A3", "C")
      .unsafe

    val cells = sheet.range("A1:A3")
    assertEquals(cells.length, 3)
    assertEquals(cells.map(_.value), List(
      CellValue.Text("A"),
      CellValue.Text("B"),
      CellValue.Text("C")
    ))
  }

  test("range includes only existing cells") {
    val sheet = baseSheet
      .put("A1", "First")
      .put("A3", "Third") // A2 missing
      .unsafe

    val cells = sheet.range("A1:A3")
    assertEquals(cells.length, 2) // Only A1 and A3
  }

  // ========== String-Based put() - Styled (Representative Tests) ==========

  test("put styled String applies both value and style") {
    val result = baseSheet.put("A1", "Styled", testStyle)
    assert(result.isRight)
    result.unsafe.cell("A1") match
      case Some(cell) =>
        assertEquals(cell.value, CellValue.Text("Styled"))
        assert(cell.styleId.isDefined)
      case None => fail("Expected styled cell at A1")
  }

  test("put styled Int applies both") {
    val result = baseSheet.put("A1", 999, testStyle)
    assert(result.isRight)
    result.unsafe.cell("A1") match
      case Some(cell) =>
        assertEquals(cell.value, CellValue.Number(BigDecimal(999)))
        assert(cell.styleId.isDefined)
      case None => fail("Expected styled cell at A1")
  }

  test("put styled LocalDate applies both") {
    val date = LocalDate.of(2025, 11, 19)
    val result = baseSheet.put("A1", date, testStyle)
    assert(result.isRight)
    assert(result.unsafe.cell("A1").flatMap(_.styleId).isDefined)
  }

  test("put styled with invalid ref returns Left") {
    val result = baseSheet.put("INVALID", "value", testStyle)
    assert(result.isLeft)
  }

  // ========== Chainable put(Patch) Extension ==========

  test("chainable put(Patch) applies patch") {
    import com.tjclp.xl.patch.Patch
    val bRef = ARef.parse("B1").toOption.fold(fail("Failed to parse B1"))(identity)
    val patch = Patch.Put(bRef, CellValue.Text("Patched"))
    val result = baseSheet
      .put("A1", "First")
      .put(patch)

    assert(result.isRight)
    val sheet = result.unsafe
    assertEquals(sheet.cells.size, 2)
    assertEquals(sheet.cell("B1").map(_.value), Some(CellValue.Text("Patched")))
  }

  // ========== Error Message Quality ==========

  test("invalid ref error message is clear") {
    val result = baseSheet.put("123ABC", "value")
    result match
      case Left(XLError.InvalidCellRef(ref, msg)) =>
        assertEquals(ref, "123ABC")
        assert(msg.contains("Invalid cell reference"), s"Message was: $msg")
      case _ => fail("Expected InvalidCellRef")
  }

  test("invalid range error message is clear") {
    val result = baseSheet.style("A1:XYZ", testStyle)
    result match
      case Left(XLError.InvalidCellRef(ref, msg)) =>
        assertEquals(ref, "A1:XYZ")
        assert(msg.contains("Invalid range"), s"Message was: $msg")
      case _ => fail("Expected InvalidCellRef for range")
  }

  // ========== Integration: Complex Chains ==========

  test("complex chain with all operation types") {
    val result = baseSheet
      .put("A1", "Title")
      .style("A1", testStyle)
      .put("A2", 100)
      .put("A3", LocalDate.of(2025, 11, 19))
      .put("A4", "Footer", CellStyle.default.italic)
      .merge("A1:B1")

    assert(result.isRight)
    val sheet = result.unsafe
    assertEquals(sheet.cells.size, 4)
    assert(sheet.mergedRanges.nonEmpty)
  }

  test("real-world example: financial report header") {
    val headerStyle = CellStyle.default.bold.size(14.0).center
    val result = baseSheet
      .style("A1:D1", headerStyle)
      .put("A1", "Q1 2025 Sales Report")
      .put("A2", "Product")
      .put("B2", "Revenue")
      .put("C2", "Units")
      .put("D2", "Profit")

    assert(result.isRight)
    // 8 cells: A1,B1,C1,D1 (styled from range) + A2,B2,C2,D2 (data)
    assertEquals(result.unsafe.cells.size, 8)
  }

  // ========== NumFmt Auto-Application Tests (Bug Fix Verification) ==========

  test("put LocalDate auto-applies Date format (BUG FIX)") {
    val result = baseSheet.put("A1", LocalDate.of(2025, 11, 19))

    result match
      case Right(updated) =>
        val ref = ARef.parse("A1").fold(err => fail(s"parse failed: $err"), identity)
        val cell = updated.cells(ref)
        assert(cell.styleId.isDefined, "Cell should have style")
        val styleId = cell.styleId.getOrElse(fail("Missing styleId"))
        val style = updated.styleRegistry.get(styleId).getOrElse(fail("Missing style"))
        assertEquals(style.numFmt, com.tjclp.xl.styles.numfmt.NumFmt.Date)
      case Left(err) => fail(s"Unexpected error: $err")
  }

  test("put BigDecimal auto-applies Decimal format") {
    val result = baseSheet.put("A1", BigDecimal("123.45"))

    result match
      case Right(updated) =>
        val ref = ARef.parse("A1").fold(err => fail(s"parse failed: $err"), identity)
        val cell = updated.cells(ref)
        assert(cell.styleId.isDefined)
        val styleId = cell.styleId.getOrElse(fail("Missing styleId"))
        val style = updated.styleRegistry.get(styleId).getOrElse(fail("Missing style"))
        assertEquals(style.numFmt, com.tjclp.xl.styles.numfmt.NumFmt.Decimal)
      case Left(err) => fail(s"Unexpected error: $err")
  }

  test("put LocalDate with style merges auto NumFmt") {
    val boldStyle = CellStyle.default.bold
    val result = baseSheet.put("A1", LocalDate.of(2025, 11, 19), boldStyle)

    result match
      case Right(updated) =>
        val ref = ARef.parse("A1").fold(err => fail(s"parse failed: $err"), identity)
        val cell = updated.cells(ref)
        val styleId = cell.styleId.getOrElse(fail("Missing styleId"))
        val style = updated.styleRegistry.get(styleId).getOrElse(fail("Missing style"))
        // Verify both bold and Date format applied
        assert(style.font.bold)
        assertEquals(style.numFmt, com.tjclp.xl.styles.numfmt.NumFmt.Date)
      case Left(err) => fail(s"Unexpected error: $err")
  }

  test("put preserves template style while applying auto NumFmt") {
    val templateStyle = CellStyle.default.bold
    val templated = baseSheet.style("A1", templateStyle).unsafe
    val result = templated.put("A1", BigDecimal("123.45"))

    result match
      case Right(updated) =>
        val ref = ARef.parse("A1").fold(err => fail(s"parse failed: $err"), identity)
        val cell = updated.cells(ref)
        val styleId = cell.styleId.getOrElse(fail("Missing styleId"))
        val style = updated.styleRegistry.get(styleId).getOrElse(fail("Missing style"))
        assert(style.font.bold, "Template bold style should remain")
        assertEquals(style.numFmt, com.tjclp.xl.styles.numfmt.NumFmt.Decimal)
      case Left(err) => fail(s"Unexpected error: $err")
  }
