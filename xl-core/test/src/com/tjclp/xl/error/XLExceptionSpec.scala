package com.tjclp.xl.error

import munit.FunSuite

/**
 * GH-589: `XLException` is the script-edge twin of the CLI's `CliError` — it carries the structured
 * [[XLError]] and its message IS the error's message, byte for byte, so there is exactly one source
 * of text per error. Richer diagnostics (the available sheets of a `SheetNotFound`, their nearest
 * as "did you mean") travel in the error itself, never as prose the exception alone knows.
 */
class XLExceptionSpec extends FunSuite:

  test("getMessage is the error's own message") {
    val ex = XLException(XLError.SheetNotFound("Sumary"))
    assertEquals(ex.getMessage, "Sheet not found: 'Sumary'")
    assertEquals(ex.error, XLError.SheetNotFound("Sumary"))
    assertEquals(ex.error.code, "SHEET_NOT_FOUND")
  }

  test("a SheetNotFound with the available sheets keeps its candidates structural") {
    val ex = XLException(XLError.SheetNotFound("Sumary", Vector("Data", "Summary")))
    assertEquals(ex.getMessage, "Sheet not found: 'Sumary'. Available: Data, Summary")
    assertEquals(ex.getMessage, ex.error.message)
    assertEquals(ex.error.candidates, Vector("Summary"))
    assertEquals(ex.error.hint, Some("list sheets with `xl -f <file> sheets`"))
    assertEquals(
      ex.error.renderDiagnostic,
      """Error: Sheet not found: 'Sumary'. Available: Data, Summary
        |  code: SHEET_NOT_FOUND
        |  did you mean: Summary
        |  hint: list sheets with `xl -f <file> sheets`""".stripMargin
    )
  }
