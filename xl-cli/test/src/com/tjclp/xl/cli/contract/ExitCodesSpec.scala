package com.tjclp.xl.cli.contract

import cats.effect.ExitCode
import munit.FunSuite

import com.tjclp.xl.error.XLError

class ExitCodesSpec extends FunSuite:

  test("the table: ok 0, signal 1, usage 2, failed 3") {
    assertEquals(ExitCodes.ok, ExitCode.Success)
    assertEquals(ExitCodes.signal, ExitCode(1))
    assertEquals(ExitCodes.usage, ExitCode(2))
    assertEquals(ExitCodes.failed, ExitCode(3))
  }

  test("every code maps to exactly one of 1, 2, 3") {
    ErrorCode.all.foreach { code =>
      val exit = ExitCodes.forCode(code).code
      assert(Set(1, 2, 3).contains(exit), s"$code maps to $exit")
    }
  }

  test("exit 1 is reserved for findings and gates — never a failure") {
    val ones = ErrorCode.all.filter(c => ExitCodes.forCode(c) == ExitCodes.signal).toSet
    assertEquals(
      ones,
      Set(
        ErrorCode.RECALC_GATE,
        ErrorCode.DIFFERENCES_FOUND,
        ErrorCode.LINT_FINDINGS,
        ErrorCode.AUDIT_FINDINGS
      )
    )
  }

  test("exit 2 is the command line being wrong") {
    val twos = ErrorCode.all.filter(c => ExitCodes.forCode(c) == ExitCodes.usage).toSet
    assertEquals(
      twos,
      Set(
        ErrorCode.USAGE,
        ErrorCode.UNKNOWN_VERB,
        ErrorCode.OUTPUT_REQUIRED,
        ErrorCode.UNSUPPORTED_IN_STREAM,
        ErrorCode.BATCH_JSON_INVALID,
        ErrorCode.BATCH_OP_UNKNOWN,
        ErrorCode.BATCH_OP_INVALID
      )
    )
  }

  test("every domain code, INTERNAL, I/O and an unknown code exit 3") {
    XLError.codes.foreach(c => assertEquals(ExitCodes.forCode(c), ExitCodes.failed, c))
    assertEquals(ExitCodes.forCode(ErrorCode.INTERNAL), ExitCodes.failed)
    assertEquals(ExitCodes.forCode(ErrorCode.IO_READ), ExitCodes.failed)
    assertEquals(ExitCodes.forCode(ErrorCode.IO_WRITE), ExitCodes.failed)
    assertEquals(ExitCodes.forCode(ErrorCode.BATCH_OP_FAILED), ExitCodes.failed)
    assertEquals(ExitCodes.forCode(ErrorCode.RASTERIZER_UNAVAILABLE), ExitCodes.failed)
    assertEquals(ExitCodes.forCode("NOT_A_CODE"), ExitCodes.failed)
  }
