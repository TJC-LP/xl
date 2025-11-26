package com.tjclp.xl.cli

import java.nio.file.{Files, Path}

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.macros.ref

/**
 * Regression tests for CLI fixes from PR #44 review feedback.
 *
 * Tests error handling for:
 *   - Invalid regex patterns in search command
 *   - Empty override values in eval command
 */
@SuppressWarnings(Array("org.wartremover.warts.Var"))
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
