package com.tjclp.xl.cli

import java.nio.file.{Files, Path}

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.macros.ref

/**
 * Tests for the --in-place (-i) flag.
 *
 * Verifies:
 *   - `-i` resolves output to the input file path
 *   - `-i` and `-o` together produce an error
 *   - Neither `-i` nor `-o` returns None (no output)
 *   - In-place write actually modifies the original file
 */
@SuppressWarnings(
  Array(
    "org.wartremover.warts.Var",
    "org.wartremover.warts.OptionPartial",
    "org.wartremover.warts.IterableOps",
    "org.wartremover.warts.IsInstanceOf"
  )
)
class InPlaceSpec extends CatsEffectSuite:

  private val excel = ExcelIO.instance[IO]

  /** Create a temp xlsx with a single sheet containing A1="Hello", A2=42. */
  private def withTempExcelFile[A](test: Path => IO[A]): IO[A] =
    IO.blocking {
      val tempFile = Files.createTempFile("xl-inplace-test-", ".xlsx")
      tempFile.toFile.deleteOnExit()
      tempFile
    }.flatMap { tempFile =>
      val sheet = Sheet("Test")
        .put(ref"A1", CellValue.Text("Hello"))
        .put(ref"A2", CellValue.Number(BigDecimal(42)))
      val wb = Workbook(Vector(sheet))
      excel.write(wb, tempFile) *> test(tempFile)
    }

  test("resolveOutput: -i sets output to input file") {
    val file = Path.of("/tmp/test.xlsx")
    Main.resolveOutput(None, inPlace = true, file).assertEquals(Some(file))
  }

  test("resolveOutput: -o takes precedence when -i is false") {
    val file = Path.of("/tmp/input.xlsx")
    val out = Path.of("/tmp/output.xlsx")
    Main.resolveOutput(Some(out), inPlace = false, file).assertEquals(Some(out))
  }

  test("resolveOutput: -i and -o together is an error") {
    val file = Path.of("/tmp/input.xlsx")
    val out = Path.of("/tmp/output.xlsx")
    Main.resolveOutput(Some(out), inPlace = true, file).attempt.map { result =>
      assert(result.isLeft, "Should be an error")
      val msg = result.left.toOption.get.getMessage
      assert(
        msg.contains("mutually exclusive"),
        s"Error should mention mutual exclusivity, got: $msg"
      )
    }
  }

  test("resolveOutput: neither -i nor -o returns None") {
    val file = Path.of("/tmp/test.xlsx")
    Main.resolveOutput(None, inPlace = false, file).assertEquals(None)
  }

  test("in-place write modifies original file") {
    withTempExcelFile { tempFile =>
      for
        // Read → modify → write back to same path (this is what -i does)
        wb <- excel.read(tempFile)
        sheet = wb.sheets.head.put(ref"A1", CellValue.Text("Updated"))
        modifiedWb = Workbook(Vector(sheet))
        _ <- excel.write(modifiedWb, tempFile)
        // Read back and verify the file was modified
        wb2 <- excel.read(tempFile)
        sheet2 = wb2.sheets.head
        value = sheet2.cells.get(ref"A1").map(_.value)
      yield assertEquals(value, Some(CellValue.Text("Updated")))
    }
  }
