package com.tjclp.xl.cli.contract

import java.nio.file.{NoSuchFileException, Paths}

import cats.effect.ExitCode
import munit.FunSuite

import com.tjclp.xl.cli.{MemoryGuard, StrictFailure}
import com.tjclp.xl.cli.raster.RasterError
import com.tjclp.xl.error.{XLError, XLException}

class CliErrorSpec extends FunSuite:

  private val codePattern = "^[A-Z][A-Z0-9_]*$".r

  test("ErrorCode.all carries every XLError code and every CLI code, each exactly once") {
    XLError.codes.foreach(c => assert(ErrorCode.all.contains(c), s"$c missing from ErrorCode.all"))
    ErrorCode.cli.foreach(c => assert(ErrorCode.all.contains(c), s"$c missing from ErrorCode.all"))
    assertEquals(ErrorCode.all.distinct, ErrorCode.all)
    assertEquals(
      ErrorCode.all.size,
      ErrorCode.cli.size + XLError.codes.size,
      "a CLI-only code collides with an XLError case code"
    )
    ErrorCode.all.foreach(c => assert(codePattern.matches(c), s"'$c' is not SCREAMING_SNAKE"))
  }

  test("fromXLError is a projection: code, message, hint, candidates, location, cause") {
    val e = XLError.SheetRequired("view", Vector("Data", "Summary"))
    val at = Some(Location.file("book.xlsx"))
    val err = CliError.fromXLError(e, at)
    assertEquals(err.code, "SHEET_REQUIRED")
    assertEquals(err.message, e.message)
    assertEquals(err.hint, e.hint)
    assertEquals(err.candidates, Vector("Data", "Summary"))
    assertEquals(err.location, at)
    assertEquals(err.cause, Some(e))
    assertEquals(err.exitCode, ExitCode(3))
  }

  test("fromThrowable: a CliException yields its error unchanged") {
    val err = CliError(ErrorCode.OUTPUT_REQUIRED, "put requires -o")
    assertEquals(CliError.fromThrowable(CliException(err)), err)
  }

  test("fromThrowable: StrictFailure is the RECALC_GATE carrying the summary, exit 1") {
    val err = CliError.fromThrowable(StrictFailure("Recalculated 3 formulas\n1 error\nSaved: x"))
    assertEquals(err.code, ErrorCode.RECALC_GATE)
    assertEquals(err.message, "Recalculated 3 formulas\n1 error\nSaved: x")
    assertEquals(err.exitCode, ExitCode(1))
  }

  test("fromThrowable: XLException projects the wrapped XLError") {
    val err = CliError.fromThrowable(XLException(XLError.FormulaError("=SUM(", "unexpected end")))
    assertEquals(err.code, "FORMULA_ERROR")
    assertEquals(err.message, "Formula error in '=SUM(': unexpected end")
    assertEquals(err.hint, Some("check the formula with `xl eval`"))
    assertEquals(err.exitCode, ExitCode(3))
  }

  test("fromThrowable: a raster export with no backend is RASTERIZER_UNAVAILABLE, exit 3") {
    val none = CliError.fromThrowable(RasterError.NoRasterizerAvailable(List("batik", "resvg")))
    assertEquals(none.code, ErrorCode.RASTERIZER_UNAVAILABLE)
    assert(none.message.contains("No SVG rasterizer available"), none.message)
    assertEquals(none.exitCode, ExitCode(3))
    val missing =
      CliError.fromThrowable(RasterError.RasterizerNotFound("resvg", "brew install resvg"))
    assertEquals(missing.code, ErrorCode.RASTERIZER_UNAVAILABLE)
    assert(missing.message.contains("resvg"), missing.message)
  }

  test("fromThrowable: NoSuchFileException is IO_READ naming the file") {
    val err = CliError.fromThrowable(new NoSuchFileException("/nonexistent/nope.xlsx"))
    assertEquals(err.code, ErrorCode.IO_READ)
    assert(err.message.contains("/nonexistent/nope.xlsx"), err.message)
    assertEquals(err.exitCode, ExitCode(3))
  }

  test("fromThrowable: anything else is INTERNAL with getMessage; a null message is safe") {
    val plain =
      CliError.fromThrowable(new Exception("Range A1:B2 has 4 cells but 3 values provided"))
    assertEquals(plain.code, ErrorCode.INTERNAL)
    assertEquals(plain.message, "Range A1:B2 has 4 cells but 3 values provided")
    assertEquals(plain.exitCode, ExitCode(3))
    val nullMessage = new RuntimeException() // no-arg: getMessage is null
    assertEquals(Option(nullMessage.getMessage), None)
    val safe = CliError.fromThrowable(nullMessage)
    assertEquals(safe.code, ErrorCode.INTERNAL)
    assertEquals(safe.message, nullMessage.toString)
    assert(safe.message.nonEmpty)
  }

  test("fromThrowable: an OutOfMemoryError is RESOURCE_LIMIT naming the heap, exit 3 (GH-636)") {
    val err = CliError.fromThrowable(new OutOfMemoryError("Garbage-collected heap size exceeded"))
    assertEquals(err.code, ErrorCode.RESOURCE_LIMIT)
    assertEquals(err, MemoryGuard.exhausted)
    assert(err.message.contains(MemoryGuard.human(MemoryGuard.maxHeapBytes)), err.message)
    assert(err.hint.exists(h => h.contains("--stream") && h.contains("-Xmx")), err.hint.toString)
    assertEquals(err.exitCode, ExitCode(3))
  }

  test("fromThrowable: every other Error stays INTERNAL — only memory exhaustion is classified") {
    assertEquals(CliError.fromThrowable(new StackOverflowError("deep")).code, ErrorCode.INTERNAL)
    assertEquals(CliError.fromThrowable(new AssertionError("boom")).code, ErrorCode.INTERNAL)
    assertEquals(CliError.fromThrowable(new LinkageError("link")).code, ErrorCode.INTERNAL)
  }

  test("usage errors carry USAGE and exit 2") {
    val err = CliError.usage("--in-place (-i) and --output (-o) are mutually exclusive", None)
    assertEquals(err.code, ErrorCode.USAGE)
    assertEquals(err.hint, None)
    assertEquals(err.exitCode, ExitCode(2))
    assertEquals(CliError.usage("x", Some("y")).hint, Some("y"))
  }

  test("CliException: getMessage is the error message and there is no stack trace") {
    val e = CliException(CliError(ErrorCode.INTERNAL, "boom"))
    assertEquals(e.getMessage, "boom")
    assertEquals(e.getStackTrace.length, 0)
    assertEquals(e.error.code, ErrorCode.INTERNAL)
  }

  test("GH-870: an IO_READ on a relative path names the absolute path and the working directory") {
    val cwd = Paths.get("/work/books")
    val missing = CliError(
      ErrorCode.IO_READ,
      "No such file: q3/model.xlsx",
      hint = Some("check the path; the previous write may have failed"),
      location = Some(Location.file("q3/model.xlsx"))
    )
    val resolved = CliError.resolvingInputPath(missing, cwd)
    assertEquals(
      resolved.message,
      "No such file: q3/model.xlsx (resolved to /work/books/q3/model.xlsx " +
        "from working directory /work/books)"
    )
    // the rest of the error is untouched: location stays the path as given
    assertEquals(resolved.copy(message = missing.message), missing)
    // `..` is kept, not collapsed: through a symlinked directory a lexical `..` names a different
    // file from the one the OS opened, so the message shows the path exactly as resolved
    val up = missing.copy(
      message = "IO error: link/../model.xlsx (Permission denied)",
      location = Some(Location.file("link/../model.xlsx"))
    )
    assertEquals(
      CliError.resolvingInputPath(up, cwd).message,
      "IO error: link/../model.xlsx (Permission denied) (resolved to " +
        "/work/books/link/../model.xlsx from working directory /work/books)"
    )
  }

  test("GH-870: `.` segments are dropped from the resolved path; `..` is kept") {
    val cwd = Paths.get("/work/books")
    def resolved(typed: String): String =
      CliError
        .resolvingInputPath(
          CliError(
            ErrorCode.IO_READ,
            s"No such file: $typed",
            location = Some(Location.file(typed))
          ),
          cwd
        )
        .message
    assertEquals(
      resolved("./model.xlsx"),
      "No such file: ./model.xlsx (resolved to /work/books/model.xlsx " +
        "from working directory /work/books)"
    )
    assertEquals(
      resolved("./q3/./link/../model.xlsx"),
      "No such file: ./q3/./link/../model.xlsx (resolved to /work/books/q3/link/../model.xlsx " +
        "from working directory /work/books)"
    )
  }

  test("GH-870: resolving an error twice adds the resolution once") {
    val cwd = Paths.get("/work/books")
    val missing = CliError(
      ErrorCode.IO_READ,
      "No such file: model.xlsx",
      location = Some(Location.file("model.xlsx"))
    )
    val once = CliError.resolvingInputPath(missing, cwd)
    assertEquals(CliError.resolvingInputPath(once, cwd), once)
  }

  test("GH-870: an absolute path, a pathless IO_READ, and every other code are left as they are") {
    val cwd = Paths.get("/work/books")
    val absolute =
      CliError(
        ErrorCode.IO_READ,
        "No such file: /tmp/a.xlsx",
        location = Some(Location.file("/tmp/a.xlsx"))
      )
    val pathless = CliError(ErrorCode.IO_READ, "Parse error: bad (line 1, column 2)")
    val write =
      CliError(
        ErrorCode.IO_WRITE,
        "cannot write out.xlsx: denied",
        location = Some(Location.file("out.xlsx"))
      )
    List(absolute, pathless, write).foreach(e =>
      assertEquals(CliError.resolvingInputPath(e, cwd), e)
    )
  }
