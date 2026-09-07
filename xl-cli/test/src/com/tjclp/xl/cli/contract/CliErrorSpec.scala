package com.tjclp.xl.cli.contract

import java.nio.file.NoSuchFileException

import cats.effect.ExitCode
import munit.FunSuite

import com.tjclp.xl.cli.StrictFailure
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
