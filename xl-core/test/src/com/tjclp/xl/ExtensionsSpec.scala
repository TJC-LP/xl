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

  test("put String with valid ref succeeds (string literal returns Sheet)") {
    // String literals now return Sheet directly (compile-time validated)
    val sheet = baseSheet.put("A1", "Hello")
    assertEquals(sheet.cell("A1").map(_.value), Some(CellValue.Text("Hello")))
  }

  test("put Int with valid ref succeeds") {
    val sheet = baseSheet.put("B1", 42)
    assertEquals(sheet.cell("B1").map(_.value), Some(CellValue.Number(BigDecimal(42))))
  }

  test("put Long with valid ref succeeds") {
    val sheet = baseSheet.put("C1", 42L)
    assertEquals(sheet.cell("C1").map(_.value), Some(CellValue.Number(BigDecimal(42))))
  }

  test("put Double with valid ref succeeds") {
    val sheet = baseSheet.put("D1", 3.14)
    assertEquals(sheet.cell("D1").map(_.value), Some(CellValue.Number(BigDecimal(3.14))))
  }

  test("put BigDecimal with valid ref succeeds") {
    val sheet = baseSheet.put("E1", BigDecimal("123.45"))
    assertEquals(
      sheet.cell("E1").map(_.value),
      Some(CellValue.Number(BigDecimal("123.45")))
    )
  }

  test("put Boolean with valid ref succeeds") {
    val sheet = baseSheet.put("F1", true)
    assertEquals(sheet.cell("F1").map(_.value), Some(CellValue.Bool(true)))
  }

  test("put LocalDate with valid ref succeeds") {
    val date = LocalDate.of(2025, 11, 19)
    val sheet = baseSheet.put("G1", date)
    assertEquals(
      sheet.cell("G1").map(_.value),
      Some(CellValue.DateTime(date.atStartOfDay))
    )
  }

  test("put LocalDateTime with valid ref succeeds") {
    val dt = LocalDateTime.of(2025, 11, 19, 14, 30)
    val sheet = baseSheet.put("H1", dt)
    assertEquals(sheet.cell("H1").map(_.value), Some(CellValue.DateTime(dt)))
  }

  test("put RichText with valid ref succeeds") {
    val richText = "Bold".bold + " normal"
    val sheet = baseSheet.put("I1", richText)
    val cellValue = sheet.cell("I1").map(_.value)
    assert(cellValue.exists {
      case _: CellValue.RichText => true
      case _ => false
    })
  }

  test("put with invalid ref returns Left (runtime string)") {
    // Use variable to force runtime string evaluation
    val invalidRef = "INVALID"
    val result = baseSheet.put(invalidRef, "value")
    assert(result.isLeft)
    result match
      case Left(XLError.InvalidCellRef(ref, _)) =>
        assertEquals(ref, "INVALID")
      case _ => fail("Expected InvalidCellRef error")
  }

  test("put with out-of-range column returns Left (runtime string)") {
    val outOfRangeRef = "XFE1" // XFE > XFD (max column)
    val result = baseSheet.put(outOfRangeRef, "value")
    assert(result.isLeft)
  }

  test("put with out-of-range row returns Left (runtime string)") {
    val outOfRangeRef = "A1048577" // Row > 1048576 (max row)
    val result = baseSheet.put(outOfRangeRef, "value")
    assert(result.isLeft)
  }

  // ========== String-Based put() - Styled (9 types) ==========

  test("put String with style applies both value and style") {
    val sheet = baseSheet.put("A1", "Bold Text", testStyle)
    assertEquals(sheet.cell("A1").map(_.value), Some(CellValue.Text("Bold Text")))
    // Verify style applied (styleId should be set)
    assert(sheet.cell("A1").flatMap(_.styleId).isDefined)
  }

  test("put Int with style applies both") {
    val sheet = baseSheet.put("A1", 42, testStyle)
    assert(sheet.cell("A1").flatMap(_.styleId).isDefined)
  }

  test("put BigDecimal with style applies both") {
    val sheet = baseSheet.put("A1", BigDecimal("123.45"), testStyle)
    assert(sheet.cell("A1").flatMap(_.styleId).isDefined)
  }

  test("put LocalDate with style applies both") {
    val sheet = baseSheet.put("A1", LocalDate.of(2025, 11, 19), testStyle)
    assert(sheet.cell("A1").flatMap(_.styleId).isDefined)
  }

  test("put styled with invalid ref returns Left (runtime string)") {
    val invalidRef = "INVALID"
    val result = baseSheet.put(invalidRef, 42, testStyle)
    assert(result.isLeft)
  }

  // ========== Template style() Operations ==========

  test("style applies to single cell (string literal returns Sheet)") {
    // String literals are compile-time validated, returns Sheet directly
    val sheet = baseSheet.put("A1", "Text").style("A1", testStyle)
    assert(sheet.cell("A1").flatMap(_.styleId).isDefined)
  }

  test("style applies to range (string literal returns Sheet)") {
    // String literals are compile-time validated
    val sheet = baseSheet
      .put("A1", "A")
      .put("B1", "B")
      .style("A1:B1", testStyle)
    assert(sheet.cell("A1").flatMap(_.styleId).isDefined)
    assert(sheet.cell("B1").flatMap(_.styleId).isDefined)
  }

  test("style with invalid cell ref returns Left (runtime string)") {
    // Use variable to force runtime string evaluation
    val invalidRef = "INVALID"
    val result = baseSheet.style(invalidRef, testStyle)
    assert(result.isLeft)
  }

  test("style with invalid range returns Left (runtime string)") {
    // Use variable to force runtime string evaluation
    val invalidRange = "A1:INVALID"
    val result = baseSheet.style(invalidRange, testStyle)
    assert(result.isLeft)
  }

  // ========== Typed ref style() Operations ==========

  test("style applies to single cell (ARef)") {
    // put("A1", ...) now returns Sheet directly, so just chain style
    val sheet = baseSheet.put("A1", "Text").style(ref"A1", testStyle)
    assert(sheet.cell("A1").flatMap(_.styleId).isDefined)
  }

  test("style applies to range (CellRange)") {
    val sheet = baseSheet
      .put("A1", "A")
      .put("B1", "B")
      .style(ref"A1:B1", testStyle)
    assert(sheet.cell("A1").flatMap(_.styleId).isDefined)
    assert(sheet.cell("B1").flatMap(_.styleId).isDefined)
  }

  test("style with ARef returns Sheet directly") {
    val sheet = baseSheet.put("A1", "Text")
    val styled = sheet.style(ref"A1", testStyle)
    assertEquals(styled.cells.get(ref"A1").flatMap(_.styleId).isDefined, true)
  }

  test("style with CellRange returns Sheet directly") {
    val sheet = baseSheet.put("A1", "A").put("B1", "B")
    val styled = sheet.style(ref"A1:B1", testStyle)
    assert(styled.cells.get(ref"A1").flatMap(_.styleId).isDefined)
    assert(styled.cells.get(ref"B1").flatMap(_.styleId).isDefined)
  }

  test("style chainable with ARef on Sheet") {
    // put("A1", ...) returns Sheet, style(ARef) returns Sheet
    val sheet = baseSheet
      .put("A1", "Text")
      .style(ref"A1", testStyle)
    assert(sheet.cell("A1").flatMap(_.styleId).isDefined)
  }

  test("style chainable with CellRange on Sheet") {
    val sheet = baseSheet
      .put("A1", "A")
      .put("B1", "B")
      .style(ref"A1:B1", testStyle)
    assert(sheet.cell("A1").flatMap(_.styleId).isDefined)
    assert(sheet.cell("B1").flatMap(_.styleId).isDefined)
  }

  // ========== Safe Lookup Methods ==========

  test("cell returns Some for existing cell") {
    val sheet = baseSheet.put("A1", "Value")
    val cell = sheet.cell("A1")
    assert(cell.isDefined)
    assertEquals(cell.map(_.value), Some(CellValue.Text("Value")))
  }

  test("cell returns None for missing cell") {
    assertEquals(baseSheet.cell("Z99"), None)
  }

  test("cell returns None for invalid ref (runtime string)") {
    // Invalid literals now fail at compile time; use runtime string for testing
    val invalidRef = "INVALID"
    assertEquals(baseSheet.cell(invalidRef), None)
  }

  test("range returns all cells in range") {
    val sheet = baseSheet
      .put("A1", "A")
      .put("B1", "B")
      .put("C1", "C")
    val cells = sheet.range("A1:C1")
    assertEquals(cells.length, 3)
  }

  test("range returns empty for invalid range (runtime string)") {
    // Invalid literals now fail at compile time; use runtime string for testing
    val invalidRange = "A1:INVALID"
    assertEquals(baseSheet.range(invalidRange), List.empty)
  }

  test("range returns empty for range with no cells") {
    assertEquals(baseSheet.range("Z1:Z10"), List.empty)
  }

  test("get auto-detects single cell") {
    val sheet = baseSheet.put("A1", "Value")
    val cells = sheet.get("A1")
    assertEquals(cells.length, 1)
    cells.headOption match
      case Some(cell) => assertEquals(cell.value, CellValue.Text("Value"))
      case None => fail("Expected one cell at A1")
  }

  test("get auto-detects range") {
    val sheet = baseSheet.put("A1", "A").put("B1", "B")
    val cells = sheet.get("A1:B1")
    assertEquals(cells.length, 2)
  }

  test("get returns empty for invalid ref (runtime string)") {
    // Invalid literals now fail at compile time; use runtime string for testing
    val invalidRef = "INVALID"
    assertEquals(baseSheet.get(invalidRef), List.empty)
  }

  // ========== Merge Operations ==========

  test("merge with valid range succeeds (string literal returns Sheet)") {
    // String literals are compile-time validated, returns Sheet directly
    val sheet = baseSheet.put("A1", "Merged").merge("A1:B1")
    assert(sheet.mergedRanges.nonEmpty)
  }

  test("merge with invalid range returns Left (runtime string)") {
    // Use variable to force runtime string evaluation
    val invalidRange = "A1:INVALID"
    val result = baseSheet.merge(invalidRange)
    assert(result.isLeft)
  }

  // ========== Chainable Operations (CRITICAL) ==========

  test("chain multiple put operations (string literals return Sheet)") {
    // All literal strings - returns Sheet directly
    val sheet = baseSheet
      .put("A1", "Title")
      .put("A2", 42)
      .put("A3", true)
      .put("A4", LocalDate.of(2025, 11, 19))

    assertEquals(sheet.cells.size, 4)
    assertEquals(sheet.cell("A1").map(_.value), Some(CellValue.Text("Title")))
    assertEquals(sheet.cell("A2").map(_.value), Some(CellValue.Number(BigDecimal(42))))
  }

  test("chain short-circuits on first error (runtime strings)") {
    // Use variables to force runtime evaluation
    val validRef = "A1"
    val invalidRef = "INVALID"
    val ref3 = "A3"
    val result = baseSheet
      .put(validRef, "Valid")
      .flatMap(_.put(invalidRef, "Fail"))
      .flatMap(_.put(ref3, "Never reached"))

    assert(result.isLeft)
    result match
      case Left(XLError.InvalidCellRef("INVALID", _)) => // Success
      case _ => fail("Expected InvalidCellRef for INVALID")
  }

  test("chain style after put (string literals return Sheet)") {
    // All literals - chain returns Sheet directly
    val sheet = baseSheet
      .put("A1", "Bold Text")
      .style("A1", testStyle)

    assert(sheet.cell("A1").flatMap(_.styleId).isDefined)
  }

  test("chain put with style (string literals)") {
    // String literals return Sheet directly
    val sheet = baseSheet
      .put("A1", "First")
      .put("A2", "Styled", testStyle)

    assertEquals(sheet.cells.size, 2)
    assert(sheet.cell("A2").flatMap(_.styleId).isDefined)
  }

  test("chain merge after puts (string literals return Sheet)") {
    // All literals - chain returns Sheet directly
    val sheet = baseSheet
      .put("A1", "Start")
      .put("B1", "End")
      .merge("A1:B1")

    assert(sheet.mergedRanges.nonEmpty)
  }

  test("chain mixed operations (string literals return Sheet)") {
    // All literals - chain returns Sheet directly
    val sheet = baseSheet
      .put("A1", "Title")
      .style("A1", testStyle)
      .put("A2", 100)
      .put("A3", "Footer")

    assertEquals(sheet.cells.size, 3)
  }

  // ========== Template Pattern (style before data) ==========

  test("style before put preserves style (template pattern)") {
    // All literals - chain returns Sheet directly
    val sheet = baseSheet
      .style("A1", testStyle)
      .put("A1", "Data")

    sheet.cell("A1") match
      case Some(cell) =>
        // Verify data added and style preserved
        assertEquals(cell.value, CellValue.Text("Data"))
        assert(cell.styleId.isDefined)
      case None => fail("Expected styled cell at A1")
  }

  test("style range then put individual cells preserves formatting") {
    // All literals - chain returns Sheet directly
    val sheet = baseSheet
      .style("A1:B1", testStyle)
      .put("A1", "First")
      .put("B1", "Second")

    // Both cells should have style
    assert(sheet.cell("A1").flatMap(_.styleId).isDefined)
    assert(sheet.cell("B1").flatMap(_.styleId).isDefined)
  }

  // ========== Edge Cases ==========

  test("put empty string succeeds (string literal returns Sheet)") {
    val sheet = baseSheet.put("A1", "")
    assertEquals(sheet.cell("A1").map(_.value), Some(CellValue.Text("")))
  }

  test("put overwrites existing cell (string literal returns Sheet)") {
    val sheet = baseSheet
      .put("A1", "First")
      .put("A1", "Second")

    assertEquals(sheet.cell("A1").map(_.value), Some(CellValue.Text("Second")))
  }

  test("put preserves other cells (string literal returns Sheet)") {
    val sheet = baseSheet
      .put("A1", "Keep")
      .put("B1", "This")

    assertEquals(sheet.cells.size, 2)
    assert(sheet.cell("A1").isDefined)
    assert(sheet.cell("B1").isDefined)
  }

  test("style on empty cell creates styled empty cell (string literal returns Sheet)") {
    val sheet = baseSheet.style("A1", testStyle)
    val cell = sheet.cell("A1")
    assert(cell.isDefined)
    assert(cell.flatMap(_.styleId).isDefined)
  }

  test("range returns cells in order") {
    val sheet = baseSheet
      .put("A1", "A")
      .put("A2", "B")
      .put("A3", "C")

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

    val cells = sheet.range("A1:A3")
    assertEquals(cells.length, 2) // Only A1 and A3
  }

  // ========== String-Based put() - Styled (Representative Tests) ==========

  test("put styled String applies both value and style (string literal returns Sheet)") {
    val sheet = baseSheet.put("A1", "Styled", testStyle)
    sheet.cell("A1") match
      case Some(cell) =>
        assertEquals(cell.value, CellValue.Text("Styled"))
        assert(cell.styleId.isDefined)
      case None => fail("Expected styled cell at A1")
  }

  test("put styled Int applies both (string literal returns Sheet)") {
    val sheet = baseSheet.put("A1", 999, testStyle)
    sheet.cell("A1") match
      case Some(cell) =>
        assertEquals(cell.value, CellValue.Number(BigDecimal(999)))
        assert(cell.styleId.isDefined)
      case None => fail("Expected styled cell at A1")
  }

  test("put styled LocalDate applies both (string literal returns Sheet)") {
    val date = LocalDate.of(2025, 11, 19)
    val sheet = baseSheet.put("A1", date, testStyle)
    assert(sheet.cell("A1").flatMap(_.styleId).isDefined)
  }

  test("put styled with invalid ref returns Left (runtime string)") {
    // Use variable to force runtime string evaluation
    val invalidRef = "INVALID"
    val result = baseSheet.put(invalidRef, "value", testStyle)
    assert(result.isLeft)
  }

  // ========== Chainable put(Patch) Extension ==========

  test("put(Patch) applies patch (string literal returns Sheet)") {
    import com.tjclp.xl.patch.Patch
    val bRef = ARef.parse("B1").toOption.fold(fail("Failed to parse B1"))(identity)
    val patch = Patch.Put(bRef, CellValue.Text("Patched"))
    // put("A1", ...) returns Sheet (literal), .put(patch) returns Sheet (infallible)
    val sheet = baseSheet
      .put("A1", "First")
      .put(patch)

    assertEquals(sheet.cells.size, 2)
    assertEquals(sheet.cell("B1").map(_.value), Some(CellValue.Text("Patched")))
  }

  // ========== Error Message Quality ==========

  test("invalid ref error message is clear (runtime string)") {
    // Use variable to force runtime string evaluation
    val invalidRef = "123ABC"
    val result = baseSheet.put(invalidRef, "value")
    result match
      case Left(XLError.InvalidCellRef(ref, _)) =>
        assertEquals(ref, "123ABC")
      case _ => fail("Expected InvalidCellRef")
  }

  test("invalid range error message is clear (runtime string)") {
    // Use variable to force runtime string evaluation
    val invalidRange = "A1:XYZ"
    val result = baseSheet.style(invalidRange, testStyle)
    result match
      case Left(XLError.InvalidCellRef(ref, msg)) =>
        assertEquals(ref, "A1:XYZ")
        assert(msg.contains("Invalid range"), s"Message was: $msg")
      case _ => fail("Expected InvalidCellRef for range")
  }

  // ========== Integration: Complex Chains ==========

  test("complex chain with all operation types (string literals return Sheet)") {
    // All literals - chain returns Sheet directly
    val sheet = baseSheet
      .put("A1", "Title")
      .style("A1", testStyle)
      .put("A2", 100)
      .put("A3", LocalDate.of(2025, 11, 19))
      .put("A4", "Footer", CellStyle.default.italic)
      .merge("A1:B1")

    assertEquals(sheet.cells.size, 4)
    assert(sheet.mergedRanges.nonEmpty)
  }

  test("real-world example: financial report header (string literals return Sheet)") {
    val headerStyle = CellStyle.default.bold.size(14.0).center
    // All literals - chain returns Sheet directly
    val sheet = baseSheet
      .style("A1:D1", headerStyle)
      .put("A1", "Q1 2025 Sales Report")
      .put("A2", "Product")
      .put("B2", "Revenue")
      .put("C2", "Units")
      .put("D2", "Profit")

    // 8 cells: A1,B1,C1,D1 (styled from range) + A2,B2,C2,D2 (data)
    assertEquals(sheet.cells.size, 8)
  }

  // ========== NumFmt Auto-Application Tests (Bug Fix Verification) ==========

  test("put LocalDate auto-applies Date format (BUG FIX)") {
    // String literals return Sheet directly
    val sheet = baseSheet.put("A1", LocalDate.of(2025, 11, 19))

    val aRef = ref"A1"
    val cell = sheet.cells(aRef)
    assert(cell.styleId.isDefined, "Cell should have style")
    val styleId = cell.styleId.getOrElse(fail("Missing styleId"))
    val style = sheet.styleRegistry.get(styleId).getOrElse(fail("Missing style"))
    assertEquals(style.numFmt, com.tjclp.xl.styles.numfmt.NumFmt.Date)
  }

  test("put BigDecimal auto-applies Decimal format") {
    // String literals return Sheet directly
    val sheet = baseSheet.put("A1", BigDecimal("123.45"))

    val aRef = ref"A1"
    val cell = sheet.cells(aRef)
    assert(cell.styleId.isDefined)
    val styleId = cell.styleId.getOrElse(fail("Missing styleId"))
    val style = sheet.styleRegistry.get(styleId).getOrElse(fail("Missing style"))
    assertEquals(style.numFmt, com.tjclp.xl.styles.numfmt.NumFmt.Decimal)
  }

  test("put LocalDate with style merges auto NumFmt") {
    val boldStyle = CellStyle.default.bold
    // String literals return Sheet directly
    val sheet = baseSheet.put("A1", LocalDate.of(2025, 11, 19), boldStyle)

    val aRef = ref"A1"
    val cell = sheet.cells(aRef)
    val styleId = cell.styleId.getOrElse(fail("Missing styleId"))
    val style = sheet.styleRegistry.get(styleId).getOrElse(fail("Missing style"))
    // Verify both bold and Date format applied
    assert(style.font.bold)
    assertEquals(style.numFmt, com.tjclp.xl.styles.numfmt.NumFmt.Date)
  }

  test("put preserves template style while applying auto NumFmt") {
    val templateStyle = CellStyle.default.bold
    // style("A1", ...) returns Sheet directly for literals
    val templated = baseSheet.style("A1", templateStyle)
    // put("A1", ...) returns Sheet directly for literals
    val sheet = templated.put("A1", BigDecimal("123.45"))

    val aRef = ref"A1"
    val cell = sheet.cells(aRef)
    val styleId = cell.styleId.getOrElse(fail("Missing styleId"))
    val style = sheet.styleRegistry.get(styleId).getOrElse(fail("Missing style"))
    assert(style.font.bold, "Template bold style should remain")
    assertEquals(style.numFmt, com.tjclp.xl.styles.numfmt.NumFmt.Decimal)
  }
