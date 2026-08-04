package com.tjclp.xl.cli

import java.io.{ByteArrayOutputStream, PrintStream}
import java.nio.charset.StandardCharsets
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
 *   - Success messages name the user-visible target, never the temp file (GH-464)
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

  test("runWithOutput: -o alone passes the output path through as write and display") {
    val file = Path.of("/tmp/input.xlsx")
    val out = Path.of("/tmp/output.xlsx")
    var captured: Option[Path] = None
    var capturedDisplay: Option[Path] = None
    Main
      .runWithOutput(Some(out), inPlace = false, file) { (received, display) =>
        captured = received
        capturedDisplay = display
        IO.pure((ExitCode.Success, "ok"))
      }
      .map { code =>
        assertEquals(code, ExitCode.Success)
        assertEquals(captured, Some(out))
        assertEquals(capturedDisplay, Some(out))
      }
  }

  test("runWithOutput: neither flag passes None through") {
    val file = Path.of("/tmp/input.xlsx")
    var captured: Option[Path] = Some(Path.of("/sentinel"))
    var capturedDisplay: Option[Path] = Some(Path.of("/sentinel"))
    Main
      .runWithOutput(None, inPlace = false, file) { (received, display) =>
        captured = received
        capturedDisplay = display
        IO.pure((ExitCode.Success, "ok"))
      }
      .map { code =>
        assertEquals(code, ExitCode.Success)
        assertEquals(captured, None)
        assertEquals(capturedDisplay, None)
      }
  }

  test("runWithOutput: -i and -o together exits with error code") {
    val file = Path.of("/tmp/input.xlsx")
    val out = Path.of("/tmp/output.xlsx")
    Main
      .runWithOutput(Some(out), inPlace = true, file)((_, _) => IO.pure((ExitCode.Success, "ok")))
      .map { code => assertEquals(code, ExitCode.Error) }
  }

  test("runWithOutput: -i writes to temp, atomically moves to input on success") {
    withTempExcelFile { tempFile =>
      // Capture the actual path we were asked to write to (should NOT be tempFile)
      var writePath: Option[Path] = None
      val result = Main.runWithOutput(None, inPlace = true, tempFile) { (outOpt, displayOpt) =>
        val out = outOpt.get
        writePath = Some(out)
        // Messages must be rendered against the original file, not the temp (GH-464)
        assertEquals(displayOpt, Some(tempFile))
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
        excel.write(wb, out).as((ExitCode.Success, "saved"))
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
      val result = Main.runWithOutput(None, inPlace = true, tempFile) { (outOpt, _) =>
        writePath = outOpt
        // Simulate a failure: return Error exit code without writing anything useful
        IO.pure((ExitCode.Error, "failed"))
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
      val result = Main.runWithOutput(None, inPlace = true, tempFile) { (outOpt, _) =>
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

  // ========== Success message names the target, not the temp (GH-464) ==========

  private def xlCommand = com.monovore.decline.Command("xl", "test")(Main.main)

  private def captureStdout[A](io: IO[A]): IO[(A, String)] =
    IO.blocking(new ByteArrayOutputStream()).flatMap { baos =>
      IO.blocking {
        val prev = System.out
        System.setOut(new PrintStream(baos, true, StandardCharsets.UTF_8))
        prev
      }.bracket(_ => io)(prev => IO.blocking(System.setOut(prev)))
        .map(result => (result, baos.toString(StandardCharsets.UTF_8)))
    }

  /** Parse a full CLI invocation and run it, capturing everything printed to stdout. */
  private def runCliCaptured(args: String*): IO[(ExitCode, String)] =
    IO.fromEither(
      xlCommand.parse(args, Map.empty).left.map(help => new Exception(help.toString))
    ).flatMap(captureStdout)

  test("runWithOutput: failed final replacement prints no success message") {
    IO.blocking {
      val destination = Files.createTempDirectory("xl-inplace-move-failure-")
      Files.writeString(destination.resolve("keep"), "keep")
      destination
    }.bracket { destination =>
      val attempted = Main
        .runWithOutput(None, inPlace = true, destination) { (outOpt, _) =>
          val out = outOpt.get
          IO.blocking(Files.writeString(out, "replacement"))
            .as((ExitCode.Success, s"Saved: $destination"))
        }
        .attempt
      captureStdout(attempted).map { case (outcome, stdout) =>
        assert(outcome.isLeft, "replacing a non-empty directory must fail")
        assert(!stdout.contains("Saved:"), s"success was printed before replacement:\n$stdout")
        assert(Files.isDirectory(destination), "failed replacement must preserve the destination")
      }
    } { destination =>
      IO.blocking {
        Files.deleteIfExists(destination.resolve("keep"))
        Files.deleteIfExists(destination)
      }.void
    }
  }

  test("-i recalc: success message names the original file, not the temp (GH-464)") {
    withTempExcelFile { tempFile =>
      runCliCaptured("-f", tempFile.toString, "-i", "recalc").map { case (code, out) =>
        assertEquals(code, ExitCode.Success)
        assert(out.contains(s"Saved: $tempFile"), s"expected 'Saved: $tempFile' in:\n$out")
        assert(!out.contains(".xl-inplace-"), s"message leaks the temp path:\n$out")
      }
    }
  }

  test("-i put: success message names the original file, not the temp (GH-464)") {
    withTempExcelFile { tempFile =>
      runCliCaptured("-f", tempFile.toString, "-s", "Test", "-i", "put", "A2", "99").map {
        case (code, out) =>
          assertEquals(code, ExitCode.Success)
          assert(out.contains(s"Saved: $tempFile"), s"expected 'Saved: $tempFile' in:\n$out")
          assert(!out.contains(".xl-inplace-"), s"message leaks the temp path:\n$out")
      }
    }
  }

  test("-o put: success message still names the -o target (GH-464 regression guard)") {
    withTempExcelFile { tempFile =>
      IO.blocking {
        val out = Files.createTempFile("xl-inplace-test-out-", ".xlsx")
        Files.deleteIfExists(out)
        out.toFile.deleteOnExit()
        out
      }.flatMap { outFile =>
        runCliCaptured(
          "-f",
          tempFile.toString,
          "-s",
          "Test",
          "-o",
          outFile.toString,
          "put",
          "A2",
          "99"
        ).map { case (code, out) =>
          assertEquals(code, ExitCode.Success)
          assert(out.contains(s"Saved: $outFile"), s"expected 'Saved: $outFile' in:\n$out")
        }
      }
    }
  }

  test("GH-483: error message names the target, never the deleted temp") {
    val temp = Path.of("/data/.xl-inplace-482910.xlsx")
    val target = Path.of("/data/book.xlsx")
    val rendered = Main.renderErrorMessage(
      new Exception(s"No space left on device writing $temp"),
      Some(temp),
      Some(target)
    )
    assert(rendered.contains(target.toString), s"error must name the target:\n$rendered")
    assert(!rendered.contains(".xl-inplace-"), s"error leaks the temp path:\n$rendered")
  }

  test("GH-483: null exception message renders the exception class, not 'null'") {
    val temp = Path.of("/data/.xl-inplace-482910.xlsx")
    val target = Path.of("/data/book.xlsx")
    val rendered =
      Main.renderErrorMessage(new IllegalStateException(), Some(temp), Some(target))
    assert(rendered.contains("IllegalStateException"), s"expected exception class in:\n$rendered")
    assert(!rendered.contains("null"), s"error prints literal 'null':\n$rendered")
  }
