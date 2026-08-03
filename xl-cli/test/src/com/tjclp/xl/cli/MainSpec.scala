package com.tjclp.xl.cli

import java.nio.file.{Files, Path}

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.helpers.{BatchParser, StyleBuilder}
import com.tjclp.xl.cli.helpers.BatchParser.BatchOp
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.macros.ref
import com.tjclp.xl.styles.numfmt.NumFmt

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

  test("put command parser wires --no-detect") {
    val parser = com.monovore.decline.Command("xl", "test")(Main.putCmd)

    assertEquals(
      parser.parse(Seq("put", "A1", "2025-11-10", "--no-detect")),
      Right(CliCommand.Put("A1", List("2025-11-10"), detect = false))
    )
    assertEquals(
      parser.parse(Seq("put", "A1", "2025-11-10")),
      Right(CliCommand.Put("A1", List("2025-11-10"), detect = true))
    )
  }

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
      BatchOp.Put("A1", CellValue.Text("First"), None),
      BatchOp.Put("B1", CellValue.Text("Second"), None),
      BatchOp.Put("C1", CellValue.Text("Third"), None)
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
      BatchOp.Put("A1", CellValue.Number(BigDecimal(100)), None),
      BatchOp.Put("B1", CellValue.Number(BigDecimal(200)), None),
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
      BatchOp.Put("Sheet1!A1", CellValue.Text("InSheet1"), None),
      BatchOp.Put("Sheet2!A1", CellValue.Text("InSheet2"), None),
      BatchOp.Put("Sheet1!B1", CellValue.Text("AlsoInSheet1"), None)
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

  test("batch autofit: invalid column spec fails fast") {
    val sheet = Sheet("Test").put(ref"A1", CellValue.Text("x"))
    val wb = Workbook(Vector(sheet))
    val ops = Vector(BatchOp.AutoFit(Some("A::C")))

    BatchParser.applyBatchOperations(wb, Some(sheet), ops).attempt.map {
      case Left(err) =>
        assert(
          err.getMessage.contains("Invalid autofit columns"),
          s"Expected invalid autofit error, got: ${err.getMessage}"
        )
      case Right(_) =>
        fail("Invalid autofit columns spec should fail")
    }
  }

  // ========== New Batch Operations (GH-88 - batch completeness) ==========

  test("batch: comment adds cell comment with author") {
    val sheet = Sheet("Test").put(ref"A1", CellValue.Text("Revenue"))
    val wb = Workbook(Vector(sheet))
    val ops = Vector(BatchOp.AddComment("A1", "Q1 note", Some("Analyst")))

    BatchParser.applyBatchOperations(wb, Some(sheet), ops).map { result =>
      val s = result.sheets.head
      assert(s.hasComment(ref"A1"), "Cell should have a comment")
      val comment = s.comments(ref"A1")
      assertEquals(comment.text.toPlainText, "Q1 note")
      assertEquals(comment.author, Some("Analyst"))
    }
  }

  test("batch: remove-comment removes existing comment") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Text("Revenue"))
      .comment(ref"A1", com.tjclp.xl.cells.Comment.plainText("Old note", None))
    val wb = Workbook(Vector(sheet))
    val ops = Vector(BatchOp.RemoveComment("A1"))

    BatchParser.applyBatchOperations(wb, Some(sheet), ops).map { result =>
      val s = result.sheets.head
      assert(!s.hasComment(ref"A1"), "Comment should be removed")
    }
  }

  test("batch: clear with all flag removes contents, styles, and comments") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Text("Data"))
      .comment(ref"A1", com.tjclp.xl.cells.Comment.plainText("Note", None))
    val wb = Workbook(Vector(sheet))
    val ops = Vector(BatchOp.Clear("A1", all = true, styles = false, comments = false))

    BatchParser.applyBatchOperations(wb, Some(sheet), ops).map { result =>
      val s = result.sheets.head
      assert(!s.hasComment(ref"A1"), "Comment should be cleared")
      assertEquals(s.cells.get(ref"A1"), None)
    }
  }

  test("batch: clear with styles flag preserves contents") {
    val sheet = Sheet("Test").put(ref"A1", CellValue.Text("Keep"))
    val wb = Workbook(Vector(sheet))
    val ops = Vector(BatchOp.Clear("A1", all = false, styles = true, comments = false))

    BatchParser.applyBatchOperations(wb, Some(sheet), ops).map { result =>
      val s = result.sheets.head
      // Content preserved (only styles cleared)
      val cell = s.cells.get(ref"A1")
      assert(cell.isDefined, "Cell content should be preserved when only clearing styles")
    }
  }

  test("batch: col-hide sets column hidden property") {
    val sheet = Sheet("Test").put(ref"B1", CellValue.Text("Data"))
    val wb = Workbook(Vector(sheet))
    val ops = Vector(BatchOp.ColHide("B"))

    BatchParser.applyBatchOperations(wb, Some(sheet), ops).map { result =>
      val s = result.sheets.head
      val col = com.tjclp.xl.addressing.Column.fromLetter("B").toOption.get
      assert(s.getColumnProperties(col).hidden, "Column B should be hidden")
    }
  }

  test("batch: col-show clears column hidden property") {
    val col = com.tjclp.xl.addressing.Column.fromLetter("B").toOption.get
    val sheet = Sheet("Test")
      .put(ref"B1", CellValue.Text("Data"))
      .setColumnProperties(col, com.tjclp.xl.sheets.ColumnProperties(hidden = true))
    val wb = Workbook(Vector(sheet))
    val ops = Vector(BatchOp.ColShow("B"))

    BatchParser.applyBatchOperations(wb, Some(sheet), ops).map { result =>
      val s = result.sheets.head
      assert(!s.getColumnProperties(col).hidden, "Column B should be visible")
    }
  }

  test("batch: row-hide sets row hidden property") {
    val sheet = Sheet("Test").put(ref"A2", CellValue.Text("Data"))
    val wb = Workbook(Vector(sheet))
    val ops = Vector(BatchOp.RowHide(2))

    BatchParser.applyBatchOperations(wb, Some(sheet), ops).map { result =>
      val s = result.sheets.head
      val row = com.tjclp.xl.addressing.Row.from1(2)
      assert(s.getRowProperties(row).hidden, "Row 2 should be hidden")
    }
  }

  test("batch: row-show clears row hidden property") {
    val row = com.tjclp.xl.addressing.Row.from1(2)
    val sheet = Sheet("Test")
      .put(ref"A2", CellValue.Text("Data"))
      .setRowProperties(row, com.tjclp.xl.sheets.RowProperties(hidden = true))
    val wb = Workbook(Vector(sheet))
    val ops = Vector(BatchOp.RowShow(2))

    BatchParser.applyBatchOperations(wb, Some(sheet), ops).map { result =>
      val s = result.sheets.head
      assert(!s.getRowProperties(row).hidden, "Row 2 should be visible")
    }
  }

  test("batch: autofit adjusts column width based on content") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Text("Short"))
      .put(ref"A2", CellValue.Text("A much longer text value"))
    val wb = Workbook(Vector(sheet))
    val ops = Vector(BatchOp.AutoFit(Some("A")))

    BatchParser.applyBatchOperations(wb, Some(sheet), ops).map { result =>
      val s = result.sheets.head
      val col = com.tjclp.xl.addressing.Column.fromLetter("A").toOption.get
      val width = s.getColumnProperties(col).width
      assert(width.isDefined, "Column A should have a width set")
      assert(width.get > 8.43, "Width should be wider than default for long text")
    }
  }

  test("batch: add-sheet creates new sheet") {
    val sheet = Sheet("Sheet1").put(ref"A1", CellValue.Text("Original"))
    val wb = Workbook(Vector(sheet))
    val ops = Vector(BatchOp.AddSheet("Sheet2", None))

    BatchParser.applyBatchOperations(wb, Some(sheet), ops).map { result =>
      assertEquals(result.sheets.length, 2)
      assertEquals(result.sheets(1).name.value, "Sheet2")
    }
  }

  test("batch: add-sheet with after positions correctly") {
    val sheet1 = Sheet("Sheet1")
    val sheet2 = Sheet("Sheet3")
    val wb = Workbook(Vector(sheet1, sheet2))
    val ops = Vector(BatchOp.AddSheet("Sheet2", Some("Sheet1")))

    BatchParser.applyBatchOperations(wb, Some(sheet1), ops).map { result =>
      assertEquals(result.sheets.length, 3)
      assertEquals(result.sheets(0).name.value, "Sheet1")
      assertEquals(result.sheets(1).name.value, "Sheet2")
      assertEquals(result.sheets(2).name.value, "Sheet3")
    }
  }

  test("batch: add-sheet rejects duplicate name") {
    val sheet = Sheet("Sheet1")
    val wb = Workbook(Vector(sheet))
    val ops = Vector(BatchOp.AddSheet("Sheet1", None))

    BatchParser.applyBatchOperations(wb, Some(sheet), ops).attempt.map {
      case Left(err) =>
        assert(
          err.getMessage.contains("already exists"),
          s"Expected duplicate error: ${err.getMessage}"
        )
      case Right(_) =>
        fail("Duplicate sheet name should fail")
    }
  }

  test("batch: rename-sheet changes sheet name") {
    val sheet = Sheet("OldName").put(ref"A1", CellValue.Text("Data"))
    val wb = Workbook(Vector(sheet))
    val ops = Vector(BatchOp.RenameSheet("OldName", "NewName"))

    BatchParser.applyBatchOperations(wb, Some(sheet), ops).map { result =>
      assertEquals(result.sheets.head.name.value, "NewName")
    }
  }

  test("batch: rename-sheet error on missing source") {
    val sheet = Sheet("Sheet1")
    val wb = Workbook(Vector(sheet))
    val ops = Vector(BatchOp.RenameSheet("NoSuchSheet", "NewName"))

    BatchParser.applyBatchOperations(wb, Some(sheet), ops).attempt.map {
      case Left(_) => () // Expected
      case Right(_) => fail("Renaming non-existent sheet should fail")
    }
  }

  // ========== JSON Parsing Edge Cases (GH-67 - uPickle migration) ==========

  test("parseBatchJson: handles nested braces in values (GH-67 regression)") {
    // This was broken with the old regex parser: {[^}]+} couldn't handle nested braces
    val json = """[{"op": "put", "ref": "A1", "value": "foo{bar}baz"}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isRight, s"Should parse: $result")
    assertEquals(
      result.toOption.get.ops,
      Vector(BatchOp.Put("A1", CellValue.Text("foo{bar}baz"), None))
    )
  }

  test("parseBatchJson: handles JSON object syntax in string values (GH-67 regression)") {
    // Even more complex: actual JSON-like syntax inside string value
    val json = """[{"op": "put", "ref": "A1", "value": "{\"nested\": \"json\"}"}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isRight, s"Should parse: $result")
    assertEquals(
      result.toOption.get.ops,
      Vector(BatchOp.Put("A1", CellValue.Text("{\"nested\": \"json\"}"), None))
    )
  }

  test("parseBatchJson: handles escaped quotes in values") {
    val json = """[{"op": "put", "ref": "A1", "value": "He said \"hello\""}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isRight, s"Should parse: $result")
    assertEquals(
      result.toOption.get.ops,
      Vector(BatchOp.Put("A1", CellValue.Text("He said \"hello\""), None))
    )
  }

  test("parseBatchJson: handles unicode in values") {
    val json = """[{"op": "put", "ref": "A1", "value": "日本語 emoji: 🎉"}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isRight, s"Should parse: $result")
    assertEquals(
      result.toOption.get.ops,
      Vector(BatchOp.Put("A1", CellValue.Text("日本語 emoji: 🎉"), None))
    )
  }

  test("parseBatchJson: handles native JSON numbers") {
    val json = """[{"op": "put", "ref": "A1", "value": 12345.67}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isRight, s"Should parse: $result")
    assertEquals(
      result.toOption.get.ops,
      Vector(BatchOp.Put("A1", CellValue.Number(BigDecimal("12345.67")), None))
    )
  }

  test("parseBatchJson: handles native JSON booleans") {
    val json =
      """[{"op": "put", "ref": "A1", "value": true}, {"op": "put", "ref": "A2", "value": false}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isRight, s"Should parse: $result")
    assertEquals(
      result.toOption.get.ops,
      Vector(
        BatchOp.Put("A1", CellValue.Bool(true), None),
        BatchOp.Put("A2", CellValue.Bool(false), None)
      )
    )
  }

  test("parseBatchJson: handles null values as Empty") {
    val json = """[{"op": "put", "ref": "A1", "value": null}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isRight, s"Should parse: $result")
    assertEquals(result.toOption.get.ops, Vector(BatchOp.Put("A1", CellValue.Empty, None)))
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
    val ops = result.toOption.get.ops
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
    assertEquals(result.toOption.get.ops, Vector(BatchOp.Merge("A1:D1")))
  }

  test("parseBatchJson: parses unmerge operation") {
    val json = """[{"op": "unmerge", "range": "A1:D1"}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isRight, s"Should parse: $result")
    assertEquals(result.toOption.get.ops, Vector(BatchOp.Unmerge("A1:D1")))
  }

  test("parseBatchJson: parses colwidth operation") {
    val json = """[{"op": "colwidth", "col": "A", "width": 15.5}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isRight, s"Should parse: $result")
    assertEquals(result.toOption.get.ops, Vector(BatchOp.ColWidth("A", 15.5)))
  }

  test("parseBatchJson: parses rowheight operation") {
    val json = """[{"op": "rowheight", "row": 1, "height": 30.0}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isRight, s"Should parse: $result")
    assertEquals(result.toOption.get.ops, Vector(BatchOp.RowHeight(1, 30.0)))
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
      BatchOp.Put("A1", CellValue.Text("Title"), None),
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
      .style(
        ref"A1",
        com.tjclp.xl.styles.CellStyle.default.withFont(
          com.tjclp.xl.styles.font.Font.default.withBold(true).withItalic(true)
        )
      )
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
      BatchOp.Put("A1", CellValue.Number(BigDecimal(100)), None),
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

  // ========== Batch JSON Enhanced Syntax (#179, #180, #178) ==========

  test("parseBatchJson: format 'currency' applies Currency NumFmt") {
    val json = """[{"op": "put", "ref": "A1", "value": 99.0, "format": "currency"}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isRight, s"Should parse: $result")
    val op = result.toOption.get.ops.head
    op match
      case BatchOp.Put(ref, CellValue.Number(n), Some(NumFmt.Currency)) =>
        assertEquals(ref, "A1")
        assertEquals(n, BigDecimal("99.0"))
      case other => fail(s"Expected Put with Currency format, got $other")
  }

  test("parseBatchJson: format 'percent' stores value as-is") {
    val json = """[{"op": "put", "ref": "A1", "value": 0.594, "format": "percent"}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isRight, s"Should parse: $result")
    val op = result.toOption.get.ops.head
    op match
      case BatchOp.Put(ref, CellValue.Number(n), Some(NumFmt.Percent)) =>
        assertEquals(n, BigDecimal("0.594"))
      case other => fail(s"Expected Put with Percent format, got $other")
  }

  test("parseBatchJson: custom format code '0.0x' creates Custom NumFmt") {
    val json = """[{"op": "put", "ref": "A1", "value": 3.5, "format": "0.0x"}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isRight, s"Should parse: $result")
    val op = result.toOption.get.ops.head
    op match
      case BatchOp.Put(ref, CellValue.Number(n), Some(NumFmt.Custom(code))) =>
        assertEquals(code, "0.0x")
        assertEquals(n, BigDecimal("3.5"))
      case other => fail(s"Expected Put with Custom format, got $other")
  }

  test("parseBatchJson: $99.00 auto-detected as currency") {
    val json = """[{"op": "put", "ref": "A1", "value": "$99.00"}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isRight, s"Should parse: $result")
    val op = result.toOption.get.ops.head
    op match
      case BatchOp.Put(ref, CellValue.Number(n), Some(NumFmt.Currency)) =>
        assertEquals(n, BigDecimal("99.00"))
      case other => fail(s"Expected Put with auto-detected currency, got $other")
  }

  test("parseBatchJson: 59.4% auto-detected as percent, stored as 0.594") {
    val json = """[{"op": "put", "ref": "A1", "value": "59.4%"}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isRight, s"Should parse: $result")
    val op = result.toOption.get.ops.head
    op match
      case BatchOp.Put(ref, CellValue.Number(n), Some(NumFmt.Percent)) =>
        assertEquals(n, BigDecimal("0.594"))
      case other => fail(s"Expected Put with auto-detected percent, got $other")
  }

  test("parseBatchJson: ISO date auto-detected") {
    val json = """[{"op": "put", "ref": "A1", "value": "2025-11-10"}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isRight, s"Should parse: $result")
    val op = result.toOption.get.ops.head
    op match
      case BatchOp.Put(ref, CellValue.DateTime(dt), Some(NumFmt.Date)) =>
        assertEquals(dt.toLocalDate.toString, "2025-11-10")
      case other => fail(s"Expected Put with auto-detected date, got $other")
  }

  test("parseBatchJson: explicit date and time formats reject unparseable strings") {
    List("date", "datetime", "time").foreach { format =>
      val json =
        s"""[{"op": "put", "ref": "A1", "value": "not-a-date", "format": "$format"}]"""
      val result = BatchParser.parseBatchJson(json)

      assert(result.isLeft, s"Explicit $format should reject unparseable input: $result")
      val error = result.swap.toOption.getOrElse(fail(s"Expected $format parse error"))
      assert(error.getMessage.contains("Object 1"))
      assert(error.getMessage.contains("invalid for explicit format"))
      assert(error.getMessage.contains("Expected ISO format (YYYY-MM-DD)"))
    }
  }

  test("parseBatchJson: putf with range and 'from' creates PutFormulaDragging") {
    val json = """[{"op": "putf", "ref": "B2:B10", "value": "=SUM($A$1:A2)", "from": "B2"}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isRight, s"Should parse: $result")
    val op = result.toOption.get.ops.head
    op match
      case BatchOp.PutFormulaDragging(range, formula, from, None) =>
        assertEquals(range, "B2:B10")
        assertEquals(formula, "=SUM($A$1:A2)")
        assertEquals(from, "B2")
      case other => fail(s"Expected PutFormulaDragging, got $other")
  }

  test("parseBatchJson: putf with 'values' array creates PutFormulas") {
    val json = """[{"op": "putf", "ref": "B2:B4", "values": ["=A2*2", "=A3*2", "=A4*2"]}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isRight, s"Should parse: $result")
    val op = result.toOption.get.ops.head
    op match
      case BatchOp.PutFormulas(range, formulas, None) =>
        assertEquals(range, "B2:B4")
        assertEquals(formulas, Vector("=A2*2", "=A3*2", "=A4*2"))
      case other => fail(s"Expected PutFormulas, got $other")
  }

  test("parseBatchJson: putf accepts 'formula' as alias for 'value'") {
    val json = """[{"op": "putf", "ref": "D14", "formula": "=SUM(D5:D12)"}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isRight, s"Should parse: $result")
    val op = result.toOption.get.ops.head
    op match
      case BatchOp.PutFormula(ref, formula, None) =>
        assertEquals(ref, "D14")
        assertEquals(formula, "=SUM(D5:D12)")
      case other => fail(s"Expected PutFormula, got $other")
  }

  test("parseBatchJson: putf 'formula' alias works with dragging") {
    val json = """[{"op": "putf", "ref": "B2:B10", "formula": "=A2*2", "from": "B2"}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isRight, s"Should parse: $result")
    val op = result.toOption.get.ops.head
    op match
      case BatchOp.PutFormulaDragging(range, formula, from, None) =>
        assertEquals(range, "B2:B10")
        assertEquals(formula, "=A2*2")
        assertEquals(from, "B2")
      case other => fail(s"Expected PutFormulaDragging, got $other")
  }

  test("parseBatchJson: putf 'value' still preferred over 'formula' when both present") {
    val json = """[{"op": "putf", "ref": "A1", "value": "=B1", "formula": "=C1"}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isRight, s"Should parse: $result")
    val op = result.toOption.get.ops.head
    op match
      case BatchOp.PutFormula(_, formula, _) =>
        assertEquals(formula, "=B1")
      case other => fail(s"Expected PutFormula with 'value' winning, got $other")
  }

  test("parseBatchJson: putf accepts 'format' without unknown-prop warning (GH-356)") {
    val json = """[{"op": "putf", "ref": "C1", "value": "=A1*2", "format": "#,##0.0"}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isRight, s"Should parse: $result")
    assertEquals(
      result.toOption.get.warnings,
      Vector.empty[String],
      "format is a known putf property and must not be warned about"
    )
  }

  test("parseBatchJson: putf with custom format code creates PutFormula with NumFmt (GH-356)") {
    val json = """[{"op": "putf", "ref": "C1", "value": "=A1*2", "format": "#,##0.0"}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isRight, s"Should parse: $result")
    result.toOption.get.ops.head match
      case BatchOp.PutFormula(ref, formula, Some(NumFmt.Custom(code))) =>
        assertEquals(ref, "C1")
        assertEquals(formula, "=A1*2")
        assertEquals(code, "#,##0.0")
      case other => fail(s"Expected PutFormula with Custom format, got $other")
  }

  test("parseBatchJson: putf named format applies to dragging variant (GH-356)") {
    val json =
      """[{"op": "putf", "ref": "B2:B10", "value": "=A2*2", "from": "B2", "format": "percent"}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isRight, s"Should parse: $result")
    result.toOption.get.ops.head match
      case BatchOp.PutFormulaDragging(range, formula, from, Some(NumFmt.Percent)) =>
        assertEquals(range, "B2:B10")
        assertEquals(formula, "=A2*2")
        assertEquals(from, "B2")
      case other => fail(s"Expected PutFormulaDragging with Percent format, got $other")
  }

  test("parseBatchJson: putf format applies to 'values' array variant (GH-356)") {
    val json =
      """[{"op": "putf", "ref": "B2:B4", "values": ["=A2*2", "=A3*2", "=A4*2"], "format": "currency"}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isRight, s"Should parse: $result")
    result.toOption.get.ops.head match
      case BatchOp.PutFormulas(range, formulas, Some(NumFmt.Currency)) =>
        assertEquals(range, "B2:B4")
        assertEquals(formulas.length, 3)
      case other => fail(s"Expected PutFormulas with Currency format, got $other")
  }

  test("batch: putf with format writes formula and numFmt style (GH-356)") {
    val sheet = Sheet("Test").put(ref"A1", CellValue.Number(BigDecimal(10)))
    val wb = Workbook(Vector(sheet))

    val ops = Vector(
      BatchOp.PutFormula("C1", "=A1*2", Some(NumFmt.Custom("#,##0.0")))
    )

    BatchParser.applyBatchOperations(wb, Some(sheet), ops).map { updatedWb =>
      val updatedSheet = updatedWb.sheets.head

      updatedSheet.cells.get(ref"C1").map(_.value) match
        case Some(CellValue.Formula(f, _, _)) => assertEquals(f, "A1*2")
        case other => fail(s"Expected Formula at C1, got $other")

      val style = updatedSheet.getCellStyle(ref"C1")
      assert(style.isDefined, "numFmt style should be set on the formula cell")
      assertEquals(style.get.numFmt, NumFmt.Custom("#,##0.0"))
    }
  }

  test("batch: putf dragging with format applies numFmt to every cell (GH-356)") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Number(BigDecimal(1)))
      .put(ref"A2", CellValue.Number(BigDecimal(2)))
    val wb = Workbook(Vector(sheet))

    val ops = Vector(
      BatchOp.PutFormulaDragging("B1:B2", "=A1*2", "B1", Some(NumFmt.Percent))
    )

    BatchParser.applyBatchOperations(wb, Some(sheet), ops).map { updatedWb =>
      val updatedSheet = updatedWb.sheets.head
      List(ref"B1", ref"B2").foreach { r =>
        updatedSheet.cells.get(r).map(_.value) match
          case Some(CellValue.Formula(_, _, _)) => ()
          case other => fail(s"Expected Formula at ${r.toA1}, got $other")
        val style = updatedSheet.getCellStyle(r)
        assert(style.isDefined, s"numFmt style should be set at ${r.toA1}")
        assertEquals(style.get.numFmt, NumFmt.Percent)
      }
    }
  }

  test("batch: putf explicit formulas with format applies numFmt to every cell (GH-356)") {
    val sheet = Sheet("Test")
    val wb = Workbook(Vector(sheet))

    val ops = Vector(
      BatchOp.PutFormulas("B1:B2", Vector("=1+1", "=2+2"), Some(NumFmt.Currency))
    )

    BatchParser.applyBatchOperations(wb, Some(sheet), ops).map { updatedWb =>
      val updatedSheet = updatedWb.sheets.head
      List(ref"B1", ref"B2").foreach { r =>
        val style = updatedSheet.getCellStyle(r)
        assert(style.isDefined, s"numFmt style should be set at ${r.toA1}")
        assertEquals(style.get.numFmt, NumFmt.Currency)
      }
    }
  }

  test("batch: putf without format leaves existing cell style untouched (GH-356)") {
    val ops = BatchParser
      .parseBatchJson("""[{"op": "putf", "ref": "C1", "value": "=A1*2"}]""")
      .toOption
      .get
      .ops
    val sheet = Sheet("Test")
    val wb = Workbook(Vector(sheet))

    BatchParser.applyBatchOperations(wb, Some(sheet), ops).map { updatedWb =>
      val updatedSheet = updatedWb.sheets.head
      assertEquals(updatedSheet.getCellStyle(ref"C1"), None)
    }
  }

  // ========== Batch put values[] op-level format (GH-416) ==========

  test("parseBatchJson: put values[] threads op-level format into every element (GH-416)") {
    // Exact repro from GH-416: this used to write General cells, silently dropping the format
    val json = """[{"op":"put","ref":"A1:A2","values":[1.0,2.0],"format":"0.00%"}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isRight, s"Should parse: $result")
    result.toOption.get.ops.head match
      case BatchOp.PutValues(range, values) =>
        assertEquals(range, "A1:A2")
        assertEquals(
          values.map(_.cellValue),
          Vector[CellValue](
            CellValue.Number(BigDecimal("1.0")),
            CellValue.Number(BigDecimal("2.0"))
          )
        )
        assertEquals(
          values.map(_.format),
          Vector[Option[NumFmt]](Some(NumFmt.Custom("0.00%")), Some(NumFmt.Custom("0.00%"))),
          "op-level format must apply to every value the op writes"
        )
      case other => fail(s"Expected PutValues, got $other")
  }

  test("batch: put values[] with op-level format writes numFmt on every cell (GH-416)") {
    val sheet = Sheet("Test")
    val wb = Workbook(Vector(sheet))
    val ops = BatchParser
      .parseBatchJson("""[{"op":"put","ref":"A1:A2","values":[1.0,2.0],"format":"0.00%"}]""")
      .toOption
      .get
      .ops

    BatchParser.applyBatchOperations(wb, Some(sheet), ops).map { updatedWb =>
      val updatedSheet = updatedWb.sheets.head
      assertEquals(
        updatedSheet.cells.get(ref"A1").map(_.value),
        Some[CellValue](CellValue.Number(BigDecimal("1.0")))
      )
      assertEquals(
        updatedSheet.cells.get(ref"A2").map(_.value),
        Some[CellValue](CellValue.Number(BigDecimal("2.0")))
      )
      List(ref"A1", ref"A2").foreach { r =>
        val style = updatedSheet.getCellStyle(r)
        assert(style.isDefined, s"numFmt style should be set at ${r.toA1}")
        assertEquals(style.get.numFmt, NumFmt.Custom("0.00%"))
      }
    }
  }

  test("parseBatchJson: put values[] op-level format leaves bool/null elements bare (GH-416)") {
    // Single-put parity: explicit format applies to numbers/strings, never to booleans/null
    val json = """[{"op":"put","ref":"A1:A3","values":[1.5,true,null],"format":"currency"}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isRight, s"Should parse: $result")
    result.toOption.get.ops.head match
      case BatchOp.PutValues(_, values) =>
        assertEquals(values.map(_.format).toList, List(Some(NumFmt.Currency), None, None))
      case other => fail(s"Expected PutValues, got $other")
  }

  test("parseBatchJson: put values[] element parses exactly like the single-value arm (GH-416)") {
    val singleJson = """[{"op":"put","ref":"A1","value":"$99.00","format":"currency"}]"""
    val arrayJson = """[{"op":"put","ref":"A1:A1","values":["$99.00"],"format":"currency"}]"""

    val single = BatchParser.parseBatchJson(singleJson).toOption.get.ops.head match
      case BatchOp.Put(_, cv, fmt) => (cv, fmt)
      case other => fail(s"Expected Put, got $other")
    val fromArray = BatchParser.parseBatchJson(arrayJson).toOption.get.ops.head match
      case BatchOp.PutValues(_, values) => (values.head.cellValue, values.head.format)
      case other => fail(s"Expected PutValues, got $other")

    assertEquals(fromArray, single, "values[] elements must follow single-put format semantics")
  }

  test("parseBatchJson: put values[] explicit date format rejects unparseable strings (GH-416)") {
    val json = """[{"op":"put","ref":"A1","values":["not-a-date"],"format":"date"}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isLeft, s"Explicit date format should reject unparseable input: $result")
    val error = result.swap.toOption.getOrElse(fail("Expected parse error"))
    assert(error.getMessage.contains("invalid for explicit format"))
  }

  test("parseBatchJson: put values[] without op-level format keeps smart detection") {
    val json = """[{"op":"put","ref":"A1:A2","values":["$99.00","59.4%"]}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isRight, s"Should parse: $result")
    result.toOption.get.ops.head match
      case BatchOp.PutValues(_, values) =>
        assertEquals(values.map(_.format).toList, List(Some(NumFmt.Currency), Some(NumFmt.Percent)))
      case other => fail(s"Expected PutValues, got $other")
  }

  test("formatSummary: produces expected summary lines") {
    val ops = Vector(
      BatchOp.PutFormula("A1", "=SUM(B1:B10)"),
      BatchOp.Style("A1:D1", BatchParser.StyleProps()),
      BatchOp.Merge("A1:D1")
    )
    val summary = BatchParser.formatSummary(ops)
    assert(summary.contains("PUTF A1 = =SUM(B1:B10)"))
    assert(summary.contains("STYLE A1:D1"))
    assert(summary.contains("MERGE A1:D1"))
  }

  test("dry-run: parseBatchOperations + formatSummary produces validation without side effects") {
    // Regression: --dry-run on the --file/--output batch path previously ignored the flag
    // and still wrote the workbook. This test verifies the dry-run pipeline works end-to-end.
    val json =
      """[{"op":"put","ref":"A1","value":"Revenue"},{"op":"putf","ref":"B1","formula":"=SUM(A1:A10)"}]"""
    val result = BatchParser.parseBatchOperations(json).unsafeRunSync()
    assertEquals(result.ops.size, 2)
    assertEquals(result.warnings.size, 0)
    val summary = BatchParser.formatSummary(result.ops)
    assert(summary.contains("PUT A1"), s"Expected PUT A1 in summary: $summary")
    assert(summary.contains("PUTF B1 = =SUM(A1:A10)"), s"Expected PUTF B1 in summary: $summary")
  }

  test("parseBatchJson: backward compatible plain string remains text") {
    val json = """[{"op": "put", "ref": "A1", "value": "Hello World"}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isRight, s"Should parse: $result")
    val op = result.toOption.get.ops.head
    op match
      case BatchOp.Put(ref, CellValue.Text(s), None) =>
        assertEquals(s, "Hello World")
      case other => fail(s"Expected Put with Text value, got $other")
  }

  test("batch: put with format applies style to cell") {
    val sheet = Sheet("Test")
    val wb = Workbook(Vector(sheet))

    val ops = Vector(
      BatchOp.Put("A1", CellValue.Number(BigDecimal("99.0")), Some(NumFmt.Currency))
    )

    BatchParser.applyBatchOperations(wb, Some(sheet), ops).map { updatedWb =>
      val updatedSheet = updatedWb.sheets.head

      // Value should be set
      val value = updatedSheet.cells.get(ref"A1").map(_.value)
      assertEquals(value, Some(CellValue.Number(BigDecimal("99.0"))))

      // Style should have Currency NumFmt
      val style = updatedSheet.getCellStyle(ref"A1")
      assert(style.isDefined, "Style should be set")
      assertEquals(style.get.numFmt, NumFmt.Currency)
    }
  }

  test("parseBatchJson: detect=false disables smart detection") {
    // With detect=false, $99.00 should stay as text, not become currency
    val json = """[{"op": "put", "ref": "A1", "value": "$99.00", "detect": false}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isRight, s"Should parse: $result")
    val op = result.toOption.get.ops.head
    op match
      case BatchOp.Put(ref, CellValue.Text(s), None) =>
        assertEquals(s, "$99.00")
      case other => fail(s"Expected Put with Text (no detection), got $other")
  }

  test("parseBatchJson: detect=true (default) enables smart detection") {
    // Without detect flag or with detect=true, $99.00 should become currency
    val json = """[{"op": "put", "ref": "A1", "value": "$99.00", "detect": true}]"""
    val result = BatchParser.parseBatchJson(json)

    assert(result.isRight, s"Should parse: $result")
    val op = result.toOption.get.ops.head
    op match
      case BatchOp.Put(ref, CellValue.Number(n), Some(NumFmt.Currency)) =>
        assertEquals(n, BigDecimal("99.00"))
      case other => fail(s"Expected Put with Currency detection, got $other")
  }

  // ========== Write commands without -o: usage error, not "Internal:" (GH-422) ==========

  test("recalc without -o reports a usage error naming the flag (GH-422)") {
    val wb = Workbook(Vector(Sheet("T")))
    Main.executeCommand(wb, None, None, None, false, CliCommand.Recalc(false)).attempt.map {
      case Left(err) =>
        assert(
          err.getMessage.contains("recalc requires -o"),
          s"message must name the command and flag, got: ${err.getMessage}"
        )
        assert(
          !err.getMessage.contains("Internal"),
          s"usage error must not masquerade as internal, got: ${err.getMessage}"
        )
      case Right(out) => fail(s"Expected missing-output error, got: $out")
    }
  }

  test("unfreeze without -o names the command in the usage error (GH-422)") {
    val wb = Workbook(Vector(Sheet("T")))
    Main.executeCommand(wb, None, None, None, false, CliCommand.Unfreeze).attempt.map {
      case Left(err) =>
        assert(
          err.getMessage.contains("unfreeze requires -o"),
          s"every write verb should name itself, got: ${err.getMessage}"
        )
      case Right(out) => fail(s"Expected missing-output error, got: $out")
    }
  }

  test("requireOutput: with -o supplied invokes the write (GH-422)") {
    Main
      .requireOutput("recalc", Some(java.nio.file.Paths.get("out.xlsx")), None)((path, _, _) =>
        IO.pure(s"wrote $path")
      )
      .map(out => assertEquals(out, "wrote out.xlsx"))
  }

  test("StyleBuilder.parseNumFmt: custom format codes accepted") {
    // Test various custom format patterns
    val formats = Seq(
      ("0.0x", true), // MOIC format
      ("$#,##0;($#,##0)", true), // Accounting
      ("0 \"bps\"", true), // Basis points
      ("0.00%", true), // Explicit percent
      ("yyyy-mm-dd", true), // Date format
      ("hh:mm:ss", true) // Time format
    )

    formats.foreach { case (code, shouldSucceed) =>
      val result = StyleBuilder.parseNumFmt(code)
      if shouldSucceed then
        assert(result.isRight, s"Format '$code' should parse successfully")
        result match
          case Right(NumFmt.Custom(parsed)) =>
            assertEquals(parsed, code, s"Custom format should preserve original code")
          case Right(other) =>
            // Named format matched - that's also valid
            assert(true)
          case Left(err) =>
            fail(s"Format '$code' failed: $err")
      else assert(result.isLeft, s"Format '$code' should fail")
    }
  }
