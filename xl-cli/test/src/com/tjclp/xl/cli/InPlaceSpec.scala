package com.tjclp.xl.cli

import java.nio.file.{Files, Path}

import cats.effect.{ExitCode, IO}
import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.macros.ref

/**
 * Tests for the `--in-place` (`-i`) flag and its atomic-write semantics.
 *
 * Verifies:
 *   - `-i` alone writes to a sibling temp file, then atomically moves onto the input
 *   - `-i` + `-o` together is an error (mutually exclusive)
 *   - A failed execution leaves the original file untouched and cleans up the temp
 *   - A successful execution produces the expected output in place
 */
@SuppressWarnings(
  Array(
    "org.wartremover.warts.Var",
    "org.wartremover.warts.OptionPartial",
    "org.wartremover.warts.IterableOps"
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

  test("runWithOutput: -o alone passes the output path through") {
    val file = Path.of("/tmp/input.xlsx")
    val out = Path.of("/tmp/output.xlsx")
    var captured: Option[Path] = None
    Main
      .runWithOutput(Some(out), inPlace = false, file) { received =>
        captured = received
        IO.pure(ExitCode.Success)
      }
      .map { code =>
        assertEquals(code, ExitCode.Success)
        assertEquals(captured, Some(out))
      }
  }

  test("runWithOutput: neither flag passes None through") {
    val file = Path.of("/tmp/input.xlsx")
    var captured: Option[Path] = Some(Path.of("/sentinel"))
    Main
      .runWithOutput(None, inPlace = false, file) { received =>
        captured = received
        IO.pure(ExitCode.Success)
      }
      .map { code =>
        assertEquals(code, ExitCode.Success)
        assertEquals(captured, None)
      }
  }

  test("runWithOutput: -i and -o together exits with error code") {
    val file = Path.of("/tmp/input.xlsx")
    val out = Path.of("/tmp/output.xlsx")
    Main
      .runWithOutput(Some(out), inPlace = true, file)(_ => IO.pure(ExitCode.Success))
      .map { code => assertEquals(code, ExitCode.Error) }
  }

  test("runWithOutput: -i writes to temp, atomically moves to input on success") {
    withTempExcelFile { tempFile =>
      // Capture the actual path we were asked to write to (should NOT be tempFile)
      var writePath: Option[Path] = None
      val result = Main.runWithOutput(None, inPlace = true, tempFile) { outOpt =>
        val out = outOpt.get
        writePath = Some(out)
        // Verify we're writing to a temp file, not the original
        assert(out != tempFile, s"In-place should write to temp, got: $out")
        assert(
          out.getFileName.toString.startsWith(".xl-inplace-"),
          s"Temp should have .xl-inplace- prefix, got: ${out.getFileName}"
        )
        // Simulate a successful write
        val wb = Workbook(
          Vector(Sheet("Test").put(ref"A1", CellValue.Text("Updated")))
        )
        excel.write(wb, out).as(ExitCode.Success)
      }

      for
        code <- result
        // After success, original file should have new content
        wb <- excel.read(tempFile)
        value = wb.sheets.head.cells.get(ref"A1").map(_.value)
        // And temp file should be gone
        tempExists = writePath.exists(p => Files.exists(p))
      yield
        assertEquals(code, ExitCode.Success)
        assertEquals(value, Some(CellValue.Text("Updated")))
        assert(!tempExists, "Temp file should be cleaned up after atomic move")
    }
  }

  test("runWithOutput: -i leaves original untouched on error exit") {
    withTempExcelFile { tempFile =>
      var writePath: Option[Path] = None
      val result = Main.runWithOutput(None, inPlace = true, tempFile) { outOpt =>
        writePath = outOpt
        // Simulate a failure: return Error exit code without writing anything useful
        IO.pure(ExitCode.Error)
      }

      for
        code <- result
        // Original file should be unchanged (still has original "Hello" / 42)
        wb <- excel.read(tempFile)
        a1 = wb.sheets.head.cells.get(ref"A1").map(_.value)
        // And temp file should be gone
        tempExists = writePath.exists(p => Files.exists(p))
      yield
        assertEquals(code, ExitCode.Error)
        assertEquals(a1, Some(CellValue.Text("Hello")))
        assert(!tempExists, "Temp file should be cleaned up on failure")
    }
  }

  test("runWithOutput: -i leaves original untouched when execute throws") {
    withTempExcelFile { tempFile =>
      var writePath: Option[Path] = None
      val result = Main.runWithOutput(None, inPlace = true, tempFile) { outOpt =>
        writePath = outOpt
        IO.raiseError(new RuntimeException("simulated crash"))
      }

      for
        outcome <- result.attempt
        wb <- excel.read(tempFile)
        a1 = wb.sheets.head.cells.get(ref"A1").map(_.value)
        tempExists = writePath.exists(p => Files.exists(p))
      yield
        assert(outcome.isLeft, "Expected error to propagate")
        assertEquals(a1, Some(CellValue.Text("Hello")))
        assert(!tempExists, "Temp file should be cleaned up on exception")
    }
  }
