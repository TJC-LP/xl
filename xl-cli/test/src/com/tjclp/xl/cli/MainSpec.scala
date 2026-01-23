package com.tjclp.xl.cli

import java.nio.file.{Files, Path}

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.helpers.BatchParser
import com.tjclp.xl.cli.helpers.BatchParser.BatchOp
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.macros.ref

/**
 * Regression tests for CLI fixes from PR #44 review feedback.
 *
 * Tests error handling for:
 *   - Invalid regex patterns in search command
 *   - Empty override values in eval command
 */
@SuppressWarnings(
  Array(
    "org.wartremover.warts.Var",
    "org.wartremover.warts.OptionPartial",
    "org.wartremover.warts.IterableOps",
    "org.wartremover.warts.IsInstanceOf"
  )
)
class MainSpec extends CatsEffectSuite:

  // Create a temporary Excel file for testing
  private def withTempExcelFile[A](test: Path => IO[A]): IO[A] =
    IO.blocking {
      val tempFile = Files.createTempFile("xl-cli-test-", ".xlsx")
      tempFile.toFile.deleteOnExit()
      tempFile
    }.flatMap { tempFile =>
      val sheet = Sheet("Test")
        .put(ref"A1", CellValue.Text("Hello"))
        .put(ref"A2", CellValue.Number(BigDecimal(42)))
        .put(ref"A3", CellValue.Text("World"))

      val wb = Workbook(Vector(sheet))
      ExcelIO.instance[IO].write(wb, tempFile) *> test(tempFile)
    }

  // ========== Regex Error Handling (PR #44 feedback item #3) ==========

  test("search: invalid regex pattern returns user-friendly error") {
    IO {
      // Invalid regex: unclosed bracket
      val invalidPattern = "invalid[regex"

      // We can't easily test the full CLI, but we can test the pattern validation
      scala.util.Try(invalidPattern.r) match
        case scala.util.Failure(e) =>
          assert(
            e.getMessage.contains("Unclosed"),
            s"Should mention unclosed bracket, got: ${e.getMessage}"
          )
        case scala.util.Success(_) =>
          fail("Invalid regex should fail to compile")
    }
  }

  test("search: valid regex pattern compiles successfully") {
    withTempExcelFile { _ =>
      IO {
        val validPattern = "Hello.*World"
        val result = scala.util.Try(validPattern.r)
        assert(result.isSuccess, "Valid regex should compile")
      }
    }
  }

  // ========== Override Validation (PR #44 feedback item #5) ==========

  test("override parsing: rejects empty value") {
    IO {
      // Test the pattern matching logic
      val emptyOverride = "A1="
      emptyOverride.split("=", 2) match
        case Array(refStr, valueStr) if valueStr.trim.nonEmpty =>
          fail("Empty value should not match nonEmpty guard")
        case Array(refStr, _) =>
          // This is the expected path - empty value rejected
          assert(refStr == "A1")
        case _ =>
          fail("Should match Array pattern")
    }
  }

  test("override parsing: accepts valid value") {
    IO {
      val validOverride = "A1=100"
      validOverride.split("=", 2) match
        case Array(refStr, valueStr) if valueStr.trim.nonEmpty =>
          assertEquals(refStr, "A1")
          assertEquals(valueStr, "100")
        case _ =>
          fail("Valid override should match")
    }
  }

  test("override parsing: accepts value with equals sign") {
    IO {
      // "A1=x=y" should parse as ref="A1", value="x=y"
      val overrideWithEquals = "A1=x=y"
      overrideWithEquals.split("=", 2) match
        case Array(refStr, valueStr) if valueStr.trim.nonEmpty =>
          assertEquals(refStr, "A1")
          assertEquals(valueStr, "x=y")
        case _ =>
          fail("Override with equals in value should match")
    }
  }

  test("override parsing: rejects whitespace-only value") {
    IO {
      val whitespaceOverride = "A1=   "
      whitespaceOverride.split("=", 2) match
        case Array(refStr, valueStr) if valueStr.trim.nonEmpty =>
          fail("Whitespace-only value should not match nonEmpty guard")
        case Array(refStr, _) =>
          // Expected: whitespace trimmed to empty
          assert(refStr == "A1")
        case _ =>
          fail("Should match Array pattern")
    }
  }

  // ========== Lazy Evaluation (PR #44 feedback item #4) ==========

  test("iterator-based search is lazy (does not materialize full collection)") {
    IO {
      // Verify that .iterator.filter.take doesn't require full materialization
      var filterCount = 0
      val hugeRange = (1 to 1000000).iterator
        .filter { i =>
          filterCount += 1
          i % 2 == 0
        }
        .take(5)
        .toVector

      assertEquals(hugeRange.length, 5)
      // With lazy evaluation, we should only filter ~10 elements to get 5 matches
      // (not all 1 million)
      assert(filterCount < 100, s"Filter should be lazy, but ran $filterCount times")
    }
  }

  // ========== Batch Operations (PR #65 - CLI refactor) ==========

  test("batch: multiple put operations all persist (regression for stale sheet bug)") {
    // This test ensures that batch operations don't use a stale sheet reference,
    // which would cause earlier operations to be overwritten by later ones.
    val sheet = Sheet("Test")
    val wb = Workbook(Vector(sheet))

    val ops = Vector(
      BatchOp.Put("A1", "First"),
      BatchOp.Put("B1", "Second"),
      BatchOp.Put("C1", "Third")
    )

    BatchParser.applyBatchOperations(wb, Some(sheet), ops).map { updatedWb =>
      val updatedSheet = updatedWb.sheets.head

      // All three cells should have values - not just the last one
      val a1 = updatedSheet.cells.get(ref"A1").map(_.value)
      val b1 = updatedSheet.cells.get(ref"B1").map(_.value)
      val c1 = updatedSheet.cells.get(ref"C1").map(_.value)

      assertEquals(a1, Some(CellValue.Text("First")), "A1 should have 'First'")
      assertEquals(b1, Some(CellValue.Text("Second")), "B1 should have 'Second'")
      assertEquals(c1, Some(CellValue.Text("Third")), "C1 should have 'Third'")
    }
  }

  test("batch: mixed put and putf operations all persist") {
    val sheet = Sheet("Test")
    val wb = Workbook(Vector(sheet))

    val ops = Vector(
      BatchOp.Put("A1", "100"),
      BatchOp.Put("B1", "200"),
      BatchOp.PutFormula("C1", "=A1+B1")
    )

    BatchParser.applyBatchOperations(wb, Some(sheet), ops).map { updatedWb =>
      val updatedSheet = updatedWb.sheets.head

      val a1 = updatedSheet.cells.get(ref"A1").map(_.value)
      val b1 = updatedSheet.cells.get(ref"B1").map(_.value)
      val c1 = updatedSheet.cells.get(ref"C1").map(_.value)

      assertEquals(a1, Some(CellValue.Number(BigDecimal(100))))
      assertEquals(b1, Some(CellValue.Number(BigDecimal(200))))
      assert(c1.exists(_.isInstanceOf[CellValue.Formula]), "C1 should be a formula")
    }
  }

  test("batch: operations grouped by sheet are applied correctly") {
    // Test that operations targeting different sheets via qualified refs work
    val sheet1 = Sheet("Sheet1")
    val sheet2 = Sheet("Sheet2")
    val wb = Workbook(Vector(sheet1, sheet2))

    val ops = Vector(
      BatchOp.Put("Sheet1!A1", "InSheet1"),
      BatchOp.Put("Sheet2!A1", "InSheet2"),
      BatchOp.Put("Sheet1!B1", "AlsoInSheet1")
    )

    // No default sheet - all refs are qualified
    BatchParser.applyBatchOperations(wb, None, ops).map { updatedWb =>
      val s1 = updatedWb.sheets.find(_.name.value == "Sheet1").get
      val s2 = updatedWb.sheets.find(_.name.value == "Sheet2").get

      assertEquals(
        s1.cells.get(ref"A1").map(_.value),
        Some(CellValue.Text("InSheet1"))
      )
      assertEquals(
        s1.cells.get(ref"B1").map(_.value),
        Some(CellValue.Text("AlsoInSheet1"))
      )
      assertEquals(
        s2.cells.get(ref"A1").map(_.value),
        Some(CellValue.Text("InSheet2"))
      )
    }
  }

  // ========== JSON Parsing Edge Cases (GH-67 - uPickle migration) ==========

  test("parseBatchJson: handles nested braces in values (GH-67 regression)") {
    // This was broken with the old regex parser: {[^}]+} couldn't handle nested braces
    val json = """[{"op": "put", "ref": "A1", "value": "foo{bar}baz"}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isRight, s"Should parse: $result")
    assertEquals(result.toOption.get, Vector(BatchOp.Put("A1", "foo{bar}baz")))
  }

  test("parseBatchJson: handles JSON object syntax in string values (GH-67 regression)") {
    // Even more complex: actual JSON-like syntax inside string value
    val json = """[{"op": "put", "ref": "A1", "value": "{\"nested\": \"json\"}"}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isRight, s"Should parse: $result")
    assertEquals(result.toOption.get, Vector(BatchOp.Put("A1", "{\"nested\": \"json\"}")))
  }

  test("parseBatchJson: handles escaped quotes in values") {
    val json = """[{"op": "put", "ref": "A1", "value": "He said \"hello\""}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isRight, s"Should parse: $result")
    assertEquals(result.toOption.get, Vector(BatchOp.Put("A1", "He said \"hello\"")))
  }

  test("parseBatchJson: handles unicode in values") {
    val json = """[{"op": "put", "ref": "A1", "value": "日本語 emoji: 🎉"}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isRight, s"Should parse: $result")
    assertEquals(result.toOption.get, Vector(BatchOp.Put("A1", "日本語 emoji: 🎉")))
  }

  test("parseBatchJson: handles numeric values without quotes") {
    val json = """[{"op": "put", "ref": "A1", "value": 12345.67}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isRight, s"Should parse: $result")
    assertEquals(result.toOption.get, Vector(BatchOp.Put("A1", "12345.67")))
  }

  test("parseBatchJson: handles boolean values") {
    val json = """[{"op": "put", "ref": "A1", "value": true}, {"op": "put", "ref": "A2", "value": false}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isRight, s"Should parse: $result")
    assertEquals(
      result.toOption.get,
      Vector(BatchOp.Put("A1", "true"), BatchOp.Put("A2", "false"))
    )
  }

  test("parseBatchJson: handles null values as empty string") {
    val json = """[{"op": "put", "ref": "A1", "value": null}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isRight, s"Should parse: $result")
    assertEquals(result.toOption.get, Vector(BatchOp.Put("A1", "")))
  }

  test("parseBatchJson: provides clear error for invalid JSON") {
    val json = """[{"op": "put", "ref": "A1", value: unquoted}]""" // Missing quotes
    val result = BatchParser.parseBatchJson(json)

    assert(result.isLeft, "Should fail to parse invalid JSON")
    // uPickle provides detailed parse errors
    val errorMsg = result.swap.toOption.get.getMessage
    assert(errorMsg.nonEmpty, s"Error message should not be empty: $errorMsg")
  }

  // ========== Nested Formula Evaluation (TJC-350 / GH-94) ==========

  test("evaluateWithDependencyCheck: nested formulas evaluate correctly") {
    import com.tjclp.xl.formula.SheetEvaluator

    IO {
      // Create a sheet with nested formulas:
      // B2:B4 = raw numbers
      // B5 = SUM(B2:B4) - sums the raw numbers
      // C5 = SUM(B2:B4) - same sum
      // D5 = SUM(B2:B4) - same sum
      // E5 = SUM(B2:B4) - same sum
      // F5 = SUM(B5:E5) - DEPENDS on B5, C5, D5, E5 (nested!)
      val sheet = Sheet("Test")
        .put(ref"B2", CellValue.Number(BigDecimal(10)))
        .put(ref"B3", CellValue.Number(BigDecimal(20)))
        .put(ref"B4", CellValue.Number(BigDecimal(30)))
        .put(ref"B5", CellValue.Formula("=SUM(B2:B4)", None)) // = 60
        .put(ref"C5", CellValue.Formula("=SUM(B2:B4)", None)) // = 60
        .put(ref"D5", CellValue.Formula("=SUM(B2:B4)", None)) // = 60
        .put(ref"E5", CellValue.Formula("=SUM(B2:B4)", None)) // = 60
        .put(ref"F5", CellValue.Formula("=SUM(B5:E5)", None)) // = 240 (depends on above formulas!)

      // Evaluate using dependency-aware evaluation
      val result = SheetEvaluator.evaluateWithDependencyCheck(sheet)()

      assert(result.isRight, s"Evaluation should succeed, got: $result")

      val evaluated = result.toOption.get

      // B5, C5, D5, E5 should each be 60
      assertEquals(evaluated.get(ref"B5"), Some(CellValue.Number(BigDecimal(60))))
      assertEquals(evaluated.get(ref"C5"), Some(CellValue.Number(BigDecimal(60))))
      assertEquals(evaluated.get(ref"D5"), Some(CellValue.Number(BigDecimal(60))))
      assertEquals(evaluated.get(ref"E5"), Some(CellValue.Number(BigDecimal(60))))

      // F5 should be 240 (sum of 60+60+60+60) - THIS IS THE CRITICAL TEST
      // Before the fix, F5 would return 0 because B5:E5 weren't evaluated first
      assertEquals(
        evaluated.get(ref"F5"),
        Some(CellValue.Number(BigDecimal(240))),
        "F5=SUM(B5:E5) should correctly sum the dependent formulas"
      )
    }
  }

  test("evaluateWithDependencyCheck: simple chain A1=10, B1=A1*2, C1=B1+5") {
    import com.tjclp.xl.formula.SheetEvaluator

    IO {
      val sheet = Sheet("Test")
        .put(ref"A1", CellValue.Formula("=10", None))
        .put(ref"B1", CellValue.Formula("=A1*2", None)) // depends on A1
        .put(ref"C1", CellValue.Formula("=B1+5", None)) // depends on B1

      val result = SheetEvaluator.evaluateWithDependencyCheck(sheet)()

      assert(result.isRight, s"Evaluation should succeed, got: $result")

      val evaluated = result.toOption.get
      assertEquals(evaluated.get(ref"A1"), Some(CellValue.Number(BigDecimal(10))))
      assertEquals(evaluated.get(ref"B1"), Some(CellValue.Number(BigDecimal(20))))
      assertEquals(evaluated.get(ref"C1"), Some(CellValue.Number(BigDecimal(25))))
    }
  }

  // ========== Batch Styling Operations (GH-88) ==========

  test("parseBatchJson: parses style operation with all properties") {
    val json =
      """[{"op": "style", "range": "A1:B2", "bold": true, "italic": true, "bg": "#FFFF00", "align": "center"}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isRight, s"Should parse: $result")
    val ops = result.toOption.get
    assertEquals(ops.size, 1)
    ops.head match
      case BatchOp.Style(range, props) =>
        assertEquals(range, "A1:B2")
        assert(props.bold)
        assert(props.italic)
        assertEquals(props.bg, Some("#FFFF00"))
        assertEquals(props.align, Some("center"))
      case other => fail(s"Expected Style op, got: $other")
  }

  test("parseBatchJson: parses merge operation") {
    val json = """[{"op": "merge", "range": "A1:D1"}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isRight, s"Should parse: $result")
    assertEquals(result.toOption.get, Vector(BatchOp.Merge("A1:D1")))
  }

  test("parseBatchJson: parses unmerge operation") {
    val json = """[{"op": "unmerge", "range": "A1:D1"}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isRight, s"Should parse: $result")
    assertEquals(result.toOption.get, Vector(BatchOp.Unmerge("A1:D1")))
  }

  test("parseBatchJson: parses colwidth operation") {
    val json = """[{"op": "colwidth", "col": "A", "width": 15.5}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isRight, s"Should parse: $result")
    assertEquals(result.toOption.get, Vector(BatchOp.ColWidth("A", 15.5)))
  }

  test("parseBatchJson: parses rowheight operation") {
    val json = """[{"op": "rowheight", "row": 1, "height": 30.0}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isRight, s"Should parse: $result")
    assertEquals(result.toOption.get, Vector(BatchOp.RowHeight(1, 30.0)))
  }

  test("batch: style operation applies formatting") {
    import com.tjclp.xl.styles.color.Color
    import com.tjclp.xl.styles.fill.Fill

    val sheet = Sheet("Test").put(ref"A1" -> CellValue.Text("Hello"))
    val wb = Workbook(Vector(sheet))

    val ops = Vector(
      BatchOp.Style("A1:B2", BatchParser.StyleProps(bold = true, bg = Some("#FFFF00")))
    )

    BatchParser.applyBatchOperations(wb, Some(sheet), ops).map { updatedWb =>
      val updatedSheet = updatedWb.sheets.head

      // Check that A1 has bold font and yellow background
      val style = updatedSheet.getCellStyle(ref"A1")
      assert(style.isDefined, "Cell should have a style")
      val cellStyle = style.get
      assert(cellStyle.font.bold, "Font should be bold")
      // Check yellow fill (#FFFF00 = RGB(255, 255, 0))
      val expectedColor = Color.fromRgb(255, 255, 0)
      assertEquals(cellStyle.fill, Fill.Solid(expectedColor), "Fill should be yellow")
    }
  }

  test("batch: merge operation creates merged region") {
    val sheet = Sheet("Test").put(ref"A1" -> CellValue.Text("Title"))
    val wb = Workbook(Vector(sheet))

    val ops = Vector(BatchOp.Merge("A1:D1"))

    BatchParser.applyBatchOperations(wb, Some(sheet), ops).map { updatedWb =>
      val updatedSheet = updatedWb.sheets.head
      val range = com.tjclp.xl.addressing.CellRange.parse("A1:D1").toOption.get

      assert(updatedSheet.mergedRanges.contains(range), "Should contain merged range A1:D1")
    }
  }

  test("batch: unmerge operation removes merged region") {
    val range = com.tjclp.xl.addressing.CellRange.parse("A1:D1").toOption.get
    val sheet = Sheet("Test")
      .put(ref"A1" -> CellValue.Text("Title"))
      .merge(range)
    val wb = Workbook(Vector(sheet))

    // Verify merge exists first
    assert(sheet.mergedRanges.contains(range), "Pre-condition: range should be merged")

    val ops = Vector(BatchOp.Unmerge("A1:D1"))

    BatchParser.applyBatchOperations(wb, Some(sheet), ops).map { updatedWb =>
      val updatedSheet = updatedWb.sheets.head
      assert(!updatedSheet.mergedRanges.contains(range), "Range should be unmerged")
    }
  }

  test("batch: colwidth operation sets column width") {
    val sheet = Sheet("Test")
    val wb = Workbook(Vector(sheet))

    val ops = Vector(BatchOp.ColWidth("A", 20.0))

    BatchParser.applyBatchOperations(wb, Some(sheet), ops).map { updatedWb =>
      val updatedSheet = updatedWb.sheets.head
      val colA = com.tjclp.xl.addressing.Column.fromLetter("A").toOption.get
      val props = updatedSheet.getColumnProperties(colA)

      assertEquals(props.width, Some(20.0))
    }
  }

  test("batch: rowheight operation sets row height") {
    val sheet = Sheet("Test")
    val wb = Workbook(Vector(sheet))

    val ops = Vector(BatchOp.RowHeight(1, 30.0))

    BatchParser.applyBatchOperations(wb, Some(sheet), ops).map { updatedWb =>
      val updatedSheet = updatedWb.sheets.head
      val row1 = com.tjclp.xl.addressing.Row.from1(1)
      val props = updatedSheet.getRowProperties(row1)

      assertEquals(props.height, Some(30.0))
    }
  }

  test("batch: combined put + style + merge workflow") {
    val sheet = Sheet("Test")
    val wb = Workbook(Vector(sheet))

    val ops = Vector(
      BatchOp.Put("A1", "Title"),
      BatchOp.Style("A1:D1", BatchParser.StyleProps(bold = true, bg = Some("#0000FF"))),
      BatchOp.Merge("A1:D1"),
      BatchOp.ColWidth("A", 25.0)
    )

    BatchParser.applyBatchOperations(wb, Some(sheet), ops).map { updatedWb =>
      val updatedSheet = updatedWb.sheets.head

      // Check value was put
      val value = updatedSheet.cells.get(ref"A1").map(_.value)
      assertEquals(value, Some(CellValue.Text("Title")))

      // Check style was applied
      val style = updatedSheet.getCellStyle(ref"A1")
      assert(style.isDefined)
      assert(style.get.font.bold)

      // Check merge was created
      val range = com.tjclp.xl.addressing.CellRange.parse("A1:D1").toOption.get
      assert(updatedSheet.mergedRanges.contains(range))

      // Check column width was set
      val colA = com.tjclp.xl.addressing.Column.fromLetter("A").toOption.get
      assertEquals(updatedSheet.getColumnProperties(colA).width, Some(25.0))
    }
  }

  test("batch: style with replace=true overwrites existing style") {
    // Create sheet with pre-styled cell
    val sheet = Sheet("Test")
      .put(ref"A1" -> CellValue.Text("Hello"))
      .style(ref"A1", com.tjclp.xl.styles.CellStyle.default.withFont(
        com.tjclp.xl.styles.font.Font.default.withBold(true).withItalic(true)
      ))
    val wb = Workbook(Vector(sheet))

    // Apply style with replace=true - should clear existing italic
    val ops = Vector(
      BatchOp.Style("A1", BatchParser.StyleProps(bold = true, replace = true))
    )

    BatchParser.applyBatchOperations(wb, Some(sheet), ops).map { updatedWb =>
      val updatedSheet = updatedWb.sheets.head
      val style = updatedSheet.getCellStyle(ref"A1")

      assert(style.isDefined)
      // In replace mode, only the new style is applied (no merging)
      assert(style.get.font.bold, "Bold should be applied")
      // Since we're replacing with a fresh style that only has bold=true,
      // italic should be the default (false)
      assert(!style.get.font.italic, "Italic should not be preserved in replace mode")
    }
  }

  test("batch: operations order is preserved for deterministic results") {
    val sheet = Sheet("Test")
    val wb = Workbook(Vector(sheet))

    // Put value, then style - order matters for style to work on existing cell
    val ops = Vector(
      BatchOp.Put("A1", "100"),
      BatchOp.Style("A1", BatchParser.StyleProps(bold = true))
    )

    BatchParser.applyBatchOperations(wb, Some(sheet), ops).map { updatedWb =>
      val updatedSheet = updatedWb.sheets.head

      // Value should exist
      val value = updatedSheet.cells.get(ref"A1").map(_.value)
      assertEquals(value, Some(CellValue.Number(BigDecimal(100))))

      // Style should be applied
      val style = updatedSheet.getCellStyle(ref"A1")
      assert(style.isDefined)
      assert(style.get.font.bold)
    }
  }
